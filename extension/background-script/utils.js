import { ALARM_NAME, INTERVAL_IN_MINUTES, TIMEFORMAT } from "./config.js"
import { getPermissionStatus } from "./platforms.js"

const INVALID_FILENAME_REGEX = /[:?"*<>|~/\\\u{1}-\u{1f}\u{7f}\u{80}-\u{9f}\p{Cf}\p{Cn}]|^[.\u{0}\p{Zl}\p{Zp}\p{Zs}]|[.\u{0}\p{Zl}\p{Zp}\p{Zs}]$|^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?=\.|$)/gui

/**
 * Returns true when running inside Mozilla Firefox.
 * @returns {boolean}
 */
export function isFirefox() {
    const g = /** @type {any} */ (globalThis)
    return typeof g.browser !== "undefined" && typeof navigator !== "undefined" && /firefox/i.test(navigator.userAgent)
}

/**
 * Sanitises a meeting title for safe use in downloaded filenames.
 * @param {Meeting} meeting
 * @returns {string}
 */
export function getSanitisedMeetingTitle(meeting) {
    if (meeting?.meetingTitle) {
        return meeting.meetingTitle.replaceAll(INVALID_FILENAME_REGEX, "_")
    }
    if (meeting?.title) {
        return meeting.title.replaceAll(INVALID_FILENAME_REGEX, "_")
    }
    return "Meeting"
}

/**
 * Formats a meeting start timestamp for filenames.
 * @param {Meeting} meeting
 * @returns {string}
 */
export function getFormattedFileTimestamp(meeting) {
    const timestamp = new Date(meeting?.meetingStartTimestamp || Date.now())
    return timestamp.toLocaleString("default", TIMEFORMAT).replace(/[\/:]/g, "-")
}

/**
 * Builds the default transcript text filename for a meeting.
 * @param {Meeting} meeting
 * @returns {string}
 */
export function getTranscriptFileName(meeting) {
    const sanitisedMeetingTitle = getSanitisedMeetingTitle(meeting)
    const formattedTimestamp = getFormattedFileTimestamp(meeting)
    const prefix = meeting?.meetingSoftware ? `${meeting.meetingSoftware} transcript` : "Transcript"
    return `TranscripTonic/${prefix}-${sanitisedMeetingTitle} at ${formattedTimestamp} on.txt`
}

export function checkAndCreateAlarm() {
    chrome.alarms.get(ALARM_NAME, (alarm) => {
        if (!alarm) {
            chrome.alarms.create(ALARM_NAME, {
                periodInMinutes: INTERVAL_IN_MINUTES
            })
            console.log(`Alarm "${ALARM_NAME}" created successfully.`)
        } else {
            console.log(`Alarm "${ALARM_NAME}" already exists. Next scheduled run:`, new Date(alarm.scheduledTime))
        }
    })
}

/**
 * Format transcript entries into string
 * @param {TranscriptBlock[] | []} transcript
 */
export function getTranscriptString(transcript) {
    let transcriptString = ""
    if (transcript && transcript.length > 0) {
        transcript.forEach(transcriptBlock => {
            transcriptString += `${transcriptBlock.personName} (${new Date(transcriptBlock.timestamp).toLocaleString("default", TIMEFORMAT).toUpperCase()})\n`
            transcriptString += transcriptBlock.transcriptText
            transcriptString += "\n\n"
        })
        return transcriptString
    }
    return transcriptString
}

/**
 * Format chat messages into string
 * @param {ChatMessage[] | []} chatMessages
 */
export function getChatMessagesString(chatMessages) {
    let chatMessagesString = ""
    if (chatMessages && chatMessages.length > 0) {
        chatMessages.forEach(chatMessage => {
            chatMessagesString += `${chatMessage.personName} (${new Date(chatMessage.timestamp).toLocaleString("default", TIMEFORMAT).toUpperCase()})\n`
            chatMessagesString += chatMessage.chatMessageText
            chatMessagesString += "\n\n"
        })
    }
    return chatMessagesString
}

export function clearTabIdAndApplyUpdate() {
    // Nullify to indicate end of meeting processing
    chrome.storage.local.set({ meetingTabId: null }, function () {
        console.log("Meeting tab id cleared for next meeting")

        // Check if there's a deferred update
        chrome.storage.local.get(["isDeferredUpdatedAvailable"], function (resultLocalUntyped) {
            const resultLocal = /** @type {ResultLocal} */ (resultLocalUntyped)

            if (resultLocal.isDeferredUpdatedAvailable) {
                console.log("Applying deferred update")
                chrome.storage.local.set({ isDeferredUpdatedAvailable: false }, function () {
                    chrome.runtime.reload()
                })
            }
        })
    })
}

/**
 * Opens the extension popup programmatically.
 */
export function openExtensionPopup() {
    return new Promise((resolve, reject) => {
        if (!chrome.action || typeof chrome.action.openPopup !== "function") {
            reject("Popup API unavailable")
            return
        }
        chrome.action.openPopup()
            .then(() => {
                console.log("Popup opened successfully")
                resolve("Popup opened")
            })
            .catch((error) => {
                console.error("Failed to open popup:", error)
                reject("Failed to open popup")
            })
    })
}

/**
 * Fallback to open the side panel in a dedicated popup window or tab (e.g. Firefox or browsers without sidePanel.open).
 * @returns {Promise<string>}
 */
function openSidePanelWindow() {
    const sidePanelUrl = chrome.runtime.getURL("side-panel/index.html")
    if (chrome.windows && typeof chrome.windows.create === "function") {
        return chrome.windows.create({
            url: sidePanelUrl,
            type: "popup",
            width: 420,
            height: 720
        }).then(() => "Side panel opened in window")
            .catch(() => chrome.tabs.create({ url: sidePanelUrl }).then(() => "Side panel opened in tab"))
    }
    return chrome.tabs.create({ url: sidePanelUrl }).then(() => "Side panel opened in tab")
}

/**
 * Opens the side panel programmatically for the active or sender tab.
 * @param {number} [senderTabId]
 */
export function openSidePanel(senderTabId) {
    return new Promise((resolve, reject) => {
        const g = /** @type {any} */ (globalThis)

        if (chrome.sidePanel && typeof chrome.sidePanel.open === "function") {
            if (typeof senderTabId === "number") {
                chrome.sidePanel.open({ tabId: senderTabId })
                    .then(() => {
                        console.log("Side panel opened successfully")
                        resolve("Side panel opened")
                    })
                    .catch((error) => {
                        console.error("Failed to open side panel:", error)
                        openSidePanelWindow().then(resolve).catch(() => reject("Failed to open side panel"))
                    })
                return
            }

            chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
                const activeTab = tabs[0]
                if (!activeTab || !activeTab.id) {
                    console.error("No active tab found.")
                    reject("No active tab found")
                    return
                }

                chrome.sidePanel.open({ tabId: activeTab.id })
                    .then(() => {
                        console.log("Side panel opened successfully")
                        resolve("Side panel opened")
                    })
                    .catch((error) => {
                        console.error("Failed to open side panel:", error)
                        openSidePanelWindow().then(resolve).catch(() => reject("Failed to open side panel"))
                    })
            })
            return
        }

        if (g.browser?.sidebarAction && typeof g.browser.sidebarAction.open === "function") {
            Promise.resolve(g.browser.sidebarAction.open())
                .then(() => {
                    console.log("Firefox sidebar opened successfully")
                    resolve("Side panel opened")
                })
                .catch((error) => {
                    console.warn("Firefox sidebarAction.open fell back to window:", error)
                    openSidePanelWindow().then(resolve).catch(() => reject("Failed to open side panel"))
                })
            return
        }

        openSidePanelWindow()
            .then(resolve)
            .catch((error) => {
                console.error("Failed to open side panel window:", error)
                reject("Failed to open side panel")
            })
    })
}

export function checkPermissionsAndOpenMeetingsPage() {
    console.log("Check permissions")
    chrome.storage.sync.get(["wantGoogleMeet", "wantTeams", "wantZoom"], function (resultSyncUntyped) {
        const resultSync = /** @type {ResultSync} */ (resultSyncUntyped)

        /** @type {Platform[]} */
        const wantedPlatforms = []
        // Consider enabled if user has not explicitly opted out
        if (resultSync.wantGoogleMeet) {
            wantedPlatforms.push("google_meet")
        }
        if (resultSync.wantTeams) {
            wantedPlatforms.push("teams")
        }
        if (resultSync.wantZoom) {
            wantedPlatforms.push("zoom")
        }

        getPermissionStatus(wantedPlatforms).then((result) => {
            console.log(result)

            /** @type {Platform[]} */
            const permissionMissingPlatforms = []

            for (let i = 0; i < wantedPlatforms.length; i++) {
                if (Array.isArray(result) && result[i] === "Disabled") {
                    permissionMissingPlatforms.push(wantedPlatforms[i])
                }
            }

            if (permissionMissingPlatforms.length > 0) {
                console.log(permissionMissingPlatforms)
                chrome.tabs.create({
                    url: chrome.runtime.getURL("meetings.html")
                })
            }
        })
    })
}
