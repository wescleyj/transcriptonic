import { downloadTranscript } from './exporters.js'
import { generateAtaAtIndex } from './ata-generator.js'

// Download transcript and optionally generate Ata
// Fails if transcript is empty or no meetings in storage
/** @throws error codes: 009, 010, 013, 014 */
export function processLastMeeting() {
    return new Promise((resolve, reject) => {
        pickupLastMeetingFromStorage()
            .then(() => {
                chrome.storage.local.get(["meetings", "geminiApiKey"], function (resultLocalUntyped) {
                    const resultLocal = /** @type {ResultLocal} */ (resultLocalUntyped)
                    chrome.storage.sync.get(["autoDownloadFileAfterMeeting", "autoGenerateAtaAfterMeeting"], function (resultSyncUntyped) {
                        const resultSync = /** @type {ResultSync} */ (resultSyncUntyped)

                        // Create an array of promises to execute in parallel
                        /** @type {Promise<any>[]} */
                        const promises = []

                        // Meeting index to download
                        // @ts-ignore - Because this line exists in the resolved promise from pickupLastMeetingFromStorage, which clearly means that at least one meeting exists and resultLocal.meetings cannot be undefined.
                        const lastIndex = resultLocal.meetings.length - 1

                        // Promise to download transcript
                        if (resultSync.autoDownloadFileAfterMeeting !== false) {
                            promises.push(
                                downloadTranscript(lastIndex)
                            )
                        }

                        // Execute download sequentially before optional Ata generation to prevent storage write races
                        Promise.all(promises)
                            .then(async () => {
                                if (resultSync.autoGenerateAtaAfterMeeting && resultLocal.geminiApiKey) {
                                    try {
                                        await generateAtaAtIndex(lastIndex)
                                    } catch (ataErr) {
                                        console.error("Auto Ata generation failed:", ataErr)
                                    }
                                }
                                resolve("Meeting processing and download complete")
                            })
                            .catch(error => {
                                // Fails with error codes: 009, 010
                                const parsedError = /** @type {ErrorObject} */ (error)
                                console.error("Operation failed:", parsedError.errorMessage)
                                reject({ errorCode: parsedError.errorCode, errorMessage: parsedError.errorMessage })
                            })
                    })
                })
            })
            .catch((error) => {
                // Fails with error codes: 013, 014
                const parsedError = /** @type {ErrorObject} */ (error)
                reject({ errorCode: parsedError.errorCode, errorMessage: parsedError.errorMessage })
            })
    })
}

/**
 * @throws error codes: 013, 014
 */
// Process transcript and chat messages of the meeting that just ended from storage, format them into strings, and save as a new entry in meetings history
export function pickupLastMeetingFromStorage() {
    return new Promise((resolve, reject) => {
        chrome.storage.local.get([
            "meetingSoftware",
            "meetingTitle",
            "meetingStartTimestamp",
            "transcript",
            "chatMessages",
        ], function (resultUntyped) {
            const result = /** @type {ResultLocal} */ (resultUntyped)

            if (result.meetingStartTimestamp) {
                if ((result.transcript.length > 0) || (result.chatMessages.length > 0)) {
                    // Create new transcript entry
                    /** @type {Meeting} */
                    const newMeetingEntry = {
                        meetingSoftware: result.meetingSoftware || "Google Meet",
                        meetingTitle: result.meetingTitle,
                        meetingStartTimestamp: result.meetingStartTimestamp,
                        meetingEndTimestamp: new Date().toISOString(),
                        transcript: result.transcript,
                        chatMessages: result.chatMessages
                    }

                    // Get existing recent meetings and add the new meeting
                    chrome.storage.local.get(["meetings"], function (resultLocalUntyped) {
                        const resultLocal = /** @type {ResultLocal} */ (resultLocalUntyped)
                        let meetings = resultLocal.meetings || []
                        meetings.push(newMeetingEntry)

                        // Save full local meeting history without arbitrary 10-meeting truncation
                        chrome.storage.local.set({ meetings: meetings }, function () {
                            console.log("Last meeting picked up")
                            resolve("Last meeting picked up")
                        })
                    })
                }
                else {
                    reject({ errorCode: "014", errorMessage: "Empty transcript and empty chatMessages" })
                }
            }
            else {
                reject({ errorCode: "013", errorMessage: "No meetings found. May be attend one?" })
            }
        })
    })
}


/** @throws error codes: 009, 010, 011, 012, 013, 014 */
export function recoverLastMeeting() {
    return new Promise((resolve, reject) => {
        chrome.storage.local.get(["meetings", "meetingStartTimestamp"], function (resultLocalUntyped) {
            const resultLocal = /** @type {ResultLocal} */ (resultLocalUntyped)
            // Check if user ever attended a meeting
            if (resultLocal.meetingStartTimestamp) {
                /** @type {Meeting | undefined} */
                let lastSavedMeeting
                if ((resultLocal.meetings) && (resultLocal.meetings.length > 0)) {
                    lastSavedMeeting = resultLocal.meetings[resultLocal.meetings.length - 1]
                }

                // Last meeting was not processed for some reason. Need to recover that data, process and download it.
                if ((!lastSavedMeeting) || (resultLocal.meetingStartTimestamp !== lastSavedMeeting.meetingStartTimestamp)) {
                    processLastMeeting().then(() => {
                        resolve("Recovered last meeting to the best possible extent")
                    }).catch((error) => {
                        // Fails with error codes: 009, 010, 011, 013, 014
                        const parsedError = /** @type {ErrorObject} */ (error)
                        reject({ errorCode: parsedError.errorCode, errorMessage: parsedError.errorMessage })
                    })
                }
                else {
                    resolve("No recovery needed")
                }
            }
            else {
                reject({ errorCode: "013", errorMessage: "No meetings found. May be attend one?" })
            }
        })
    })
}

