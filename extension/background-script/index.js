import { processLastMeeting, recoverLastMeeting } from "./meetings.js"
import { downloadTranscript } from "./exporters.js"
import { generateAtaAtIndex, downloadAtaAtIndex, resolveGeminiModel } from "./ata-generator.js"
import {
    registerContentScript,
    deregisterContentScript,
    reRegisterContentScripts,
} from "./platforms.js"
import {
    clearTabIdAndApplyUpdate,
    openSidePanel
} from "./utils.js"

const ASYNC_RESPONSE_TYPES = new Set([
    "meeting_ended",
    "download_transcript_at_index",
    "generate_ata_at_index",
    "download_ata_at_index",
    "recover_last_meeting",
    "open_side_panel"
])

chrome.runtime.onMessage.addListener(function (messageUnTyped, sender, sendResponse) {
    const message = /** @type {ExtensionMessage} */ (messageUnTyped)
    console.log(message.type)

    if (message.type === "new_meeting_started") {
        // Saving current tab id, to download transcript when this tab is closed
        if (sender?.tab?.id) {
            chrome.storage.local.set({ meetingTabId: sender.tab.id }, function () {
                console.log("Meeting tab id saved")
            })
        } else {
            chrome.tabs.query({ active: true, currentWindow: true }, function (tabs) {
                const tabId = tabs?.[0]?.id
                if (tabId) {
                    chrome.storage.local.set({ meetingTabId: tabId }, function () {
                        console.log("Meeting tab id saved")
                    })
                }
            })
        }
    }

    if (message.type === "meeting_ended") {
        // Prevents double downloading of transcript from tab closed event listener. Also prevents available update from being applied, during meeting post processing.
        chrome.storage.local.set({ meetingTabId: "processing" }, function () {
            console.log("Meeting tab id set to processing meeting")

            processLastMeeting()
                .then(() => {
                    /** @type {ExtensionResponse} */
                    const response = { success: true }
                    sendResponse(response)
                })
                .catch((error) => {
                    // Fails with error codes: 009, 010, 011, 012, 013, 014
                    const parsedError = /** @type {ErrorObject} */ (error)

                    /** @type {ExtensionResponse} */
                    const response = { success: false, message: parsedError }
                    sendResponse(response)
                })
                .finally(() => {
                    clearTabIdAndApplyUpdate()
                })
        })
    }

    if (message.type === "download_transcript_at_index") {
        if ((typeof message.index === "number") && (message.index >= 0)) {
            // Download the requested item
            downloadTranscript(message.index)
                .then(() => {
                    /** @type {ExtensionResponse} */
                    const response = { success: true }
                    sendResponse(response)
                })
                .catch((error) => {
                    // Fails with error codes: 009, 010
                    const parsedError = /** @type {ErrorObject} */ (error)

                    /** @type {ExtensionResponse} */
                    const response = { success: false, message: parsedError }
                    sendResponse(response)
                })
        }
        else {
            /** @type {ExtensionResponse} */
            const response = { success: false, message: { errorCode: "015", errorMessage: "Invalid index" } }
            sendResponse(response)
        }
    }

    if (message.type === "generate_ata_at_index") {
        if ((typeof message.index === "number") && (message.index >= 0)) {
            generateAtaAtIndex(message.index, message.options || {})
                .then((result) => {
                    /** @type {ExtensionResponse} */
                    const response = { success: true, message: result }
                    sendResponse(response)
                })
                .catch((error) => {
                    const parsedError = error?.errorCode
                        ? /** @type {ErrorObject} */ (error)
                        : { errorCode: "019", errorMessage: error?.message || String(error) }
                    /** @type {ExtensionResponse} */
                    const response = { success: false, message: parsedError }
                    sendResponse(response)
                })
        }
        else {
            /** @type {ExtensionResponse} */
            const response = { success: false, message: { errorCode: "015", errorMessage: "Invalid index" } }
            sendResponse(response)
        }
    }

    if (message.type === "download_ata_at_index") {
        if ((typeof message.index === "number") && (message.index >= 0)) {
            downloadAtaAtIndex(message.index, message.format || "docx")
                .then((result) => {
                    /** @type {ExtensionResponse} */
                    const response = { success: true, message: result }
                    sendResponse(response)
                })
                .catch((error) => {
                    const parsedError = error?.errorCode
                        ? /** @type {ErrorObject} */ (error)
                        : { errorCode: "020", errorMessage: error?.message || String(error) }
                    /** @type {ExtensionResponse} */
                    const response = { success: false, message: parsedError }
                    sendResponse(response)
                })
        }
        else {
            /** @type {ExtensionResponse} */
            const response = { success: false, message: { errorCode: "015", errorMessage: "Invalid index" } }
            sendResponse(response)
        }
    }

    if (message.type === "recover_last_meeting") {
        recoverLastMeeting().then((message) => {
            /** @type {ExtensionResponse} */
            const response = { success: true, message: message }
            sendResponse(response)
        })
            .catch((error) => {
                // Fails with error codes: 009, 010, 011, 012, 013, 014
                const parsedError = /** @type {ErrorObject} */ (error)

                /** @type {ExtensionResponse} */
                const response = { success: false, message: parsedError }
                sendResponse(response)
            })
    }

    if (message.type === "open_side_panel") {
        openSidePanel(sender?.tab?.id).then((msg) => {
            /** @type {ExtensionResponse} */
            const response = { success: true, message: msg }
            sendResponse(response)
        }).catch((error) => {
            const parsedError = /** @type {ErrorObject} */ (error)

            /** @type {ExtensionResponse} */
            const response = { success: false, message: parsedError }
            sendResponse(response)
        })
    }

    return ASYNC_RESPONSE_TYPES.has(message.type)
})

// Download last meeting if meeting tab is closed
chrome.tabs?.onRemoved?.addListener(function (tabId) {
    chrome.storage.local.get(["meetingTabId"], function (resultLocalUntyped) {
        const resultLocal = /** @type {ResultLocal} */ (resultLocalUntyped)

        if (tabId === resultLocal.meetingTabId) {
            console.log("Successfully intercepted tab close")

            // Prevent misfires of onRemoved until next meeting. Also prevents available update from being applied, during meeting post processing.
            chrome.storage.local.set({ meetingTabId: "processing" }, function () {
                console.log("Meeting tab id set to processing meeting")

                processLastMeeting().finally(() => {
                    clearTabIdAndApplyUpdate()
                })
            })
        }
    })
})

// Listen for extension updates
chrome.runtime?.onUpdateAvailable?.addListener(() => {
    // Check if there is an active meeting
    chrome.storage.local.get(["meetingTabId"], function (resultUntyped) {
        const result = /** @type {ResultLocal} */ (resultUntyped)

        if (result.meetingTabId) {
            // There is an active meeting(values: tabId or processing), defer the update
            chrome.storage.local.set({ isDeferredUpdatedAvailable: true }, function () {
                console.log("Deferred update flag set")
            })
        } else {
            // No active meeting, apply the update immediately. Meeting tab id is nullified only post meeting operations are done, so no race conditions.
            console.log("No active meeting, applying update immediately")
            chrome.runtime.reload()
        }
    })
})

// Register content scripts whenever runtime permissions change—serialized after de-registration to avoid race conditions.
chrome.permissions?.onAdded?.addListener(() => {
    setTimeout(() => {
        /** @type {Platform[]} */
        const platforms = ["google_meet"]

        deregisterContentScript(platforms)
            .then(() => {
                console.log("De-registered all content scripts. Starting re-registration...")
                reRegisterContentScripts()
            })
            .catch((err) => {
                console.error("Failed to de-register content scripts:", err)
                reRegisterContentScripts()
            })
    }, 2000)
})


chrome.runtime?.onInstalled?.addListener(() => {
    // Set default values
    chrome.storage.sync.get(null, function (resultSyncUntyped) {
        const resultSync = /** @type {ResultSync} */ (resultSyncUntyped || {})

        chrome.storage.sync.remove([
            "webhookUrl",
            "autoPostWebhookAfterMeeting",
            "webhookBodyType",
            "ataClientName",
            "ataSubdiretoria",
            "wantTeams",
            "wantZoom"
        ])

        chrome.storage.sync.set({
            autoDownloadFileAfterMeeting: resultSync.autoDownloadFileAfterMeeting === false ? false : true,
            autoGenerateAtaAfterMeeting: resultSync.autoGenerateAtaAfterMeeting === false ? false : true,
            geminiModel: resolveGeminiModel(resultSync.geminiModel),
            operationMode: "auto",
            hideCaptions: resultSync.hideCaptions === true ? true : false,
            wantGoogleMeet: true
        }, function () {
            // Re-register content scripts whenever extension is installed or updated, provided permissions are available.
            reRegisterContentScripts()
        })
    })
})

chrome.runtime?.onStartup?.addListener(() => {
    reRegisterContentScripts()
})
