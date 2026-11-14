/** @type {Intl.DateTimeFormatOptions} */
export const TIMEFORMAT = {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: true
}

/** @type {Object<string, { id: string, js: string[], matches: string[], excludeMatches: string[], permissions: chrome.runtime.ManifestPermissions[] }>} */
export const PLATFORM_CONFIGS = {
    "google_meet": {
        id: "content-google-meet",
        js: [
            "content-scripts/meet-utils.js",
            "content-scripts/meet.js"
        ],
        matches: ["https://meet.google.com/*"],
        excludeMatches: ["https://meet.google.com/", "https://meet.google.com/landing"],
        permissions: ["notifications"]
    }
}

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
    return `TranscripTonic/${prefix}-${sanitisedMeetingTitle} at ${formattedTimestamp}.txt`
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
