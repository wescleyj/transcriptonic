import { getTranscriptString, getChatMessagesString, getTranscriptFileName, isFirefox } from './utils.js'
import { TIMEFORMAT } from './config.js'

/**
 * Converts a Blob into a Base64 string using ArrayBuffer (works in MV3 Service Workers, Firefox, and Node).
 * @param {Blob} blob
 * @returns {Promise<string>}
 */
export async function blobToBase64(blob) {
    const buffer = await blob.arrayBuffer()
    const bytes = new Uint8Array(buffer)
    const chunkSize = 0x8000
    let binary = ""
    for (let i = 0; i < bytes.length; i += chunkSize) {
        const chunk = bytes.subarray(i, i + chunkSize)
        binary += String.fromCharCode(...chunk)
    }
    return btoa(binary)
}

/**
 * Fallback download handler that delegates a Blob download to meetings.html.
 * @param {{ fileName: string, base64Data: string, mimeType: string }} params
 * @returns {Promise<string>}
 */
export function sendDownloadToMeetingsPage({ fileName, base64Data, mimeType }) {
    return new Promise((resolve, reject) => {
        const payload = {
            type: "download_blob_via_page",
            fileName: fileName,
            base64Data: base64Data,
            mimeType: mimeType
        }

        const sendToTab = (/** @type {number} */ tabId) => {
            chrome.tabs.sendMessage(tabId, payload, function (response) {
                if (chrome.runtime?.lastError) {
                    reject(new Error(chrome.runtime.lastError.message))
                    return
                }
                if (response && response.success) {
                    resolve("Downloaded via meetings page")
                } else {
                    reject(new Error("Failed to trigger download in meetings.html"))
                }
            })
        }

        chrome.tabs.query({}, function (tabs) {
            const meetingsTab = tabs.find(tab => tab.url && tab.url.includes("meetings.html"))
            if (meetingsTab && meetingsTab.id) {
                const tabId = meetingsTab.id
                chrome.tabs.update(tabId, { active: true }, function () {
                    sendToTab(tabId)
                })
                return
            }

            chrome.tabs.create({ url: chrome.runtime.getURL("meetings.html"), active: true }, function (newTab) {
                if (!newTab || !newTab.id) {
                    reject(new Error("Could not open meetings.html for download"))
                    return
                }
                const newTabId = newTab.id
                const timeoutId = setTimeout(() => {
                    chrome.tabs.onUpdated.removeListener(listener)
                    reject(new Error("Timed out waiting for meetings.html to load"))
                }, 15000)

                /**
                 * @param {number} tabId
                 * @param {chrome.tabs.TabChangeInfo} changeInfo
                 */
                const listener = function (tabId, changeInfo) {
                    if (tabId === newTabId && changeInfo.status === "complete") {
                        clearTimeout(timeoutId)
                        chrome.tabs.onUpdated.removeListener(listener)
                        sendToTab(newTabId)
                    }
                }
                chrome.tabs.onUpdated.addListener(listener)
            })
        })
    })
}

/**
 * Cross-browser download helper supporting Firefox blob URLs, Chromium MV3 data URLs, and meetings.html fallback.
 * @param {{ blob: Blob, fileName: string, fallbackFileName?: string, mimeType?: string }} params
 * @returns {Promise<string>}
 */
export async function downloadBlobOrDataUrl({ blob, fileName, fallbackFileName, mimeType }) {
    const g = /** @type {any} */ (globalThis)
    const downloadsApi = g.browser?.downloads || (typeof chrome !== "undefined" ? chrome.downloads : undefined)
    const contentMime = mimeType || blob.type || "application/octet-stream"
    const safeFallback = fallbackFileName || "TranscripTonic/Transcript.txt"

    if (isFirefox() && typeof URL !== "undefined" && typeof URL.createObjectURL === "function" && downloadsApi?.download) {
        const objectUrl = URL.createObjectURL(blob)
        const scheduleRevoke = () => {
            setTimeout(() => {
                try {
                    URL.revokeObjectURL(objectUrl)
                } catch (_err) {
                    // Ignore revocation errors
                }
            }, 10000)
        }

        try {
            await downloadsApi.download({
                url: objectUrl,
                filename: fileName,
                conflictAction: "uniquify"
            })
            scheduleRevoke()
            return "File downloaded successfully"
        } catch (err) {
            console.warn("Firefox primary download failed, retrying fallback filename:", err)
            try {
                await downloadsApi.download({
                    url: objectUrl,
                    filename: safeFallback,
                    conflictAction: "uniquify"
                })
                scheduleRevoke()
                return "File downloaded successfully with default file name"
            } catch (fallbackErr) {
                console.warn("Firefox downloads API fallback failed, routing via meetings.html:", fallbackErr)
                try {
                    URL.revokeObjectURL(objectUrl)
                } catch (_revokeErr) {
                    // Ignore
                }
                const base64Data = await blobToBase64(blob)
                return sendDownloadToMeetingsPage({ fileName: safeFallback, base64Data, mimeType: contentMime })
            }
        }
    }

    const base64Data = await blobToBase64(blob)
    if (downloadsApi?.download) {
        const dataUrl = `data:${contentMime};base64,${base64Data}`
        try {
            await downloadsApi.download({
                url: dataUrl,
                filename: fileName,
                conflictAction: "uniquify"
            })
            return "File downloaded successfully"
        } catch (err) {
            console.error("Primary download failed, retrying fallback filename:", err)
            try {
                await downloadsApi.download({
                    url: dataUrl,
                    filename: safeFallback,
                    conflictAction: "uniquify"
                })
                return "File downloaded successfully with default file name"
            } catch (fallbackErr) {
                console.warn("Downloads API fallback failed, routing via meetings.html:", fallbackErr)
                return sendDownloadToMeetingsPage({ fileName: safeFallback, base64Data, mimeType: contentMime })
            }
        }
    }

    return sendDownloadToMeetingsPage({ fileName: safeFallback, base64Data, mimeType: contentMime })
}

/**
 * @param {number} index
 * @param {boolean} [isWebhookEnabled]
 * @throws error codes: 009, 010
 */
export function downloadTranscript(index, isWebhookEnabled) {
    return new Promise((resolve, reject) => {
        chrome.storage.local.get(["meetings"], function (resultLocalUntyped) {
            const resultLocal = /** @type {ResultLocal} */ (resultLocalUntyped)

            if (resultLocal.meetings && resultLocal.meetings[index]) {
                const meeting = resultLocal.meetings[index]
                const fileName = getTranscriptFileName(meeting)

                // Format transcript and chatMessages content
                let content = getTranscriptString(meeting.transcript)
                content += `\n\n---------------\nCHAT MESSAGES\n---------------\n\n`
                content += getChatMessagesString(meeting.chatMessages)

                // Add branding
                content += "\n\n---------------\n"
                content += "Transcript saved using TranscripTonic extension (https://github.com/wescleyj/transcriptonic)"
                content += "\n---------------"

                const blob = new Blob([content], { type: "text/plain;charset=utf-8" })

                downloadBlobOrDataUrl({
                    blob,
                    fileName,
                    fallbackFileName: "TranscripTonic/Transcript.txt",
                    mimeType: "text/plain;charset=utf-8"
                })
                    .then((msg) => {
                        console.log("Transcript downloaded")
                        resolve(msg)
                    })
                    .catch((err) => {
                        console.error("Failed to download transcript:", err)
                        reject({ errorCode: "009", errorMessage: err?.message || "Failed to download transcript" })
                    })
            }
            else {
                reject({ errorCode: "010", errorMessage: "Meeting at specified index not found" })
            }
        })
    })
}

/**
 * @param {number} index
 * @throws error code: 010, 011, 012
 */
export function postTranscriptToWebhook(index) {
    return new Promise((resolve, reject) => {
        // Get webhook URL and meetings
        chrome.storage.local.get(["meetings"], function (resultLocalUntyped) {
            const resultLocal = /** @type {ResultLocal} */ (resultLocalUntyped)
            chrome.storage.sync.get(["webhookUrl", "webhookBodyType"], function (resultSyncUntyped) {
                const resultSync = /** @type {ResultSync} */ (resultSyncUntyped)

                if (resultSync.webhookUrl) {
                    if (resultLocal.meetings && resultLocal.meetings[index]) {
                        const meeting = resultLocal.meetings[index]

                        /** @type {WebhookBody} */
                        let webhookData
                        if (resultSync.webhookBodyType === "advanced") {
                            webhookData = {
                                webhookBodyType: "advanced",
                                meetingSoftware: meeting.meetingSoftware ? meeting.meetingSoftware : "",
                                meetingTitle: meeting.meetingTitle || meeting.title || "",
                                meetingStartTimestamp: new Date(meeting.meetingStartTimestamp).toISOString(),
                                meetingEndTimestamp: new Date(meeting.meetingEndTimestamp).toISOString(),
                                transcript: meeting.transcript,
                                chatMessages: meeting.chatMessages
                            }
                        }
                        else {
                            webhookData = {
                                webhookBodyType: "simple",
                                meetingSoftware: meeting.meetingSoftware ? meeting.meetingSoftware : "",
                                meetingTitle: meeting.meetingTitle || meeting.title || "",
                                meetingStartTimestamp: new Date(meeting.meetingStartTimestamp).toLocaleString("default", TIMEFORMAT).toUpperCase(),
                                meetingEndTimestamp: new Date(meeting.meetingEndTimestamp).toLocaleString("default", TIMEFORMAT).toUpperCase(),
                                transcript: getTranscriptString(meeting.transcript),
                                chatMessages: getChatMessagesString(meeting.chatMessages)
                            }
                        }

                        // Post to webhook
                        fetch(resultSync.webhookUrl, {
                            method: "POST",
                            headers: {
                                "Content-Type": "application/json"
                            },
                            body: JSON.stringify(webhookData)
                        }).then(async response => {
                            if (!response.ok) {
                                throw new Error(`Webhook request failed with HTTP status code ${response.status} ${response.statusText}`)
                            }
                            return response.json().catch(() => null)
                        }).then((responseData) => {
                            // @ts-ignore - Pointless type error about resultLocal.meetings being undefined, which is already checked above.
                            const meetingEntry = resultLocal.meetings[index]
                            meetingEntry.webhookPostStatus = "successful"
                            const candidateUrl = responseData?.transcriptUrl || responseData?.url || responseData?.webViewLink || responseData?.html_url
                            if (typeof candidateUrl === "string" && /^https?:\/\//i.test(candidateUrl.trim())) {
                                meetingEntry.transcriptUrl = candidateUrl.trim()
                            }
                            chrome.storage.local.set({ meetings: resultLocal.meetings }, function () {
                                resolve("Webhook posted successfully")
                            })
                        }).catch(error => {
                            console.error(error)
                            // Update failure status.
                            // @ts-ignore - Pointless type error about resultLocal.meetings being undefined, which is already checked above.
                            resultLocal.meetings[index].webhookPostStatus = "failed"
                            chrome.storage.local.set({ meetings: resultLocal.meetings }, function () {
                                // Create notification and open webhooks page if notifications API is available
                                if (chrome.notifications && typeof chrome.notifications.create === "function") {
                                    chrome.notifications.create({
                                        type: "basic",
                                        iconUrl: chrome.runtime.getURL("icon.png"),
                                        title: "Could not post webhook!",
                                        message: "Click to view status and retry. Check console for more details."
                                    }, function (notificationId) {
                                        chrome.notifications.onClicked.addListener(function (clickedNotificationId) {
                                            if (clickedNotificationId === notificationId) {
                                                chrome.tabs.create({ url: chrome.runtime.getURL("meetings.html") })
                                            }
                                        })
                                    })
                                }
                                reject({ errorCode: "011", errorMessage: error })
                            })
                        })
                    }
                    else {
                        reject({ errorCode: "010", errorMessage: "Meeting at specified index not found" })
                    }
                }
                else {
                    reject({ errorCode: "012", errorMessage: "No webhook URL configured" })
                }
            })
        })
    })
}
