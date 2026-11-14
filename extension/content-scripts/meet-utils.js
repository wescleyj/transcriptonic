// === Global config & styles ===

/** @type {ExtensionStatusJSON} */
const extensionStatusJSON_bug = {
    "status": 400,
    "message": `<strong style="color: #f69191;">CommitJr Atas encontrou um erro inesperado</strong> <br /> Por favor, reporte em <a href="https://github.com/wescleyj/transcriptonic/issues" target="_blank" style="color: #fac835; text-decoration: underline;">GitHub Issues</a>.`
}
const reportErrorMessage = "Ocorreu um erro no CommitJr Atas. Reporte em https://github.com/wescleyj/transcriptonic/issues"


/** @type {MutationObserverInit} */
const mutationConfig = { childList: true, attributes: true, subtree: true, characterData: true }

const LOGO_URL = (typeof chrome !== "undefined" && chrome.runtime && typeof chrome.runtime.getURL === "function")
    ? chrome.runtime.getURL("icon.png")
    : ""
const commonCSS = `background: rgba(12, 21, 45, 0.96); 
    position: fixed;
    left: 0; 
    right: 0; 
    margin-left: auto; 
    margin-right: auto;
    max-width: 780px;  
    z-index: 1000; 
    padding: 0.25rem 1.15rem;
    border-radius: 10px; 
    border: 1.5px solid #fac835;
    display: flex; 
    justify-content: center; 
    align-items: center; 
    gap: 14px;  
    font-size: 0.95rem; 
    line-height: 1.45; 
    font-family: "Poppins", Roboto, Arial, sans-serif; 
    box-shadow: rgba(0, 0, 0, 0.45) 0px 12px 32px 0px;`

const NOTIFICATION_PLATFORM_CONFIGS = {
    "google_meet": {
        notificationText: `<strong style="color: #fac835;">CommitJr Atas ativo</strong> <br /> Não desative as legendas do Google Meet`
    }
}

// === Google Meet DOM selectors ===

const SELECTORS_GOOGLE_MEET = {
    USER_NAME: ".awLEm",
    MEETING_TITLE: ".u6vdEc",
    GOOGLE_SYMBOLS: ".google-symbols",
    TEXT_CALL_END: "call_end",
    TEXT_CAPTIONS: ["closed_caption_off", "translate_cc"],
    TEXT_CHAT: "chat",
    TRANSCRIPT_REGION: 'div[role="region"][tabindex="0"]',
    CHAT_LIVE_REGION: 'div[aria-live="polite"].Ge9Kpc'
}

// === Core content-script helpers & FAB ===

/**
 * @description State Factory: Returns a pristine, isolated meeting state block.
 * @param {MeetingSoftware} meetingSoftware
 * @param {Platform} platform
 * @returns {ContentScriptState}
 */
function createContentScriptState(meetingSoftware, platform) {
    return {
        meetingSoftware: meetingSoftware,
        platform: platform,
        userName: "You",
        transcript: [],
        chatMessages: [],
        stateTranscriptBlock: {
            timestamp: "",
            mutationTargetElement: null,
            personName: "",
            transcriptTextBuffer: "",

        },
        meetingStartTimestamp: new Date().toISOString(),
        meetingTitle: document.title,
        transcriptTargetNode: null,
        transcriptObserver: null,
        chatMessagesTargetNode: null,
        chatMessagesObserver: null,
        isTranscriptDomErrorCaptured: false,
        isChatMessagesDomErrorCaptured: false,
        hasMeetingStarted: false,
        hasMeetingEnded: false,
        extensionStatusJSON: {
            status: 200,
            message: `<strong style="color: #fac835;">CommitJr Atas ativo</strong> <br /> Não desative as legendas do Google Meet`
        }
    }
}

/**
 * @description Resolves extension status locally without remote kill-switch or external HTML injection.
 * @param {ContentScriptState} state
 */
function checkExtensionStatus(state) {
    state.extensionStatusJSON = {
        status: 200,
        message: state.meetingSoftware ? NOTIFICATION_PLATFORM_CONFIGS[state.platform].notificationText : ""
    }
    return Promise.resolve("Extension status resolved locally")
}

/**
 * @description Overwrite state to chrome storage
 * @param {ContentScriptState} state
 * @param {Array<"meetingSoftware"  | "meetingTitle" | "meetingStartTimestamp" | "transcript" | "chatMessages">} keys
 * @param {boolean} sendDownloadMessage
 */
function overWriteChromeStorage(state, keys, sendDownloadMessage) {
    const objectToSave = {}
    if (keys.includes("meetingSoftware")) objectToSave.meetingSoftware = state.meetingSoftware
    if (keys.includes("meetingTitle")) objectToSave.meetingTitle = state.meetingTitle
    if (keys.includes("meetingStartTimestamp")) objectToSave.meetingStartTimestamp = state.meetingStartTimestamp
    if (keys.includes("transcript")) objectToSave.transcript = state.transcript
    if (keys.includes("chatMessages")) objectToSave.chatMessages = state.chatMessages

    chrome.storage.local.set(objectToSave, function () {
        if (sendDownloadMessage) {
            /** @type {ExtensionMessage} */
            const message = { type: "meeting_ended" }
            chrome.runtime.sendMessage(message, (responseUntyped) => {
                const response = /** @type {ExtensionResponse} */ (responseUntyped)
                if ((!response.success)) {
                    const parsedError = /** @type {ErrorObject} */ (response.message)
                    if (parsedError.errorCode === "010") {
                        console.error(parsedError.errorMessage)
                    }
                }
            })
        }
    })
}

/**
 * @description Attempts to recover last meeting to the best possible extent.
 */
function recoverLastMeeting() {
    return new Promise((resolve, reject) => {
        /** @type {ExtensionMessage} */
        const message = {
            type: "recover_last_meeting",
        }
        chrome.runtime.sendMessage(message, function (responseUntyped) {
            const response = /** @type {ExtensionResponse} */ (responseUntyped)
            if (response.success) {
                resolve("Last meeting recovered successfully or recovery not needed")
            }
            else {
                reject(response.message)
            }
        })
    })
}

/**
 * @description Efficiently waits until the element of the specified selector and textContent appears in the DOM. Polls only on animation frame change
 * @param {string} selector
 * @param {string | string[]} [text]
 * @param {HTMLIFrameElement | null} iframe
 */
async function waitForElement(selector, text, iframe = null) {
    // If an iframe is provided, use its content document; otherwise, default to top-level document
    const targetDoc = iframe ? /** @type {Document} */ (iframe.contentDocument) : document
    const texts = text !== undefined ? (Array.isArray(text) ? text : [text]) : []

    if (texts.length) {
        // loops for every animation frame change, until the required element is found
        while (!Array.from(targetDoc.querySelectorAll(selector)).find(element => texts.includes(element.textContent))) {
            await new Promise((resolve) => requestAnimationFrame(resolve))
        }
    }
    else {
        // loops for every animation frame change, until the required element is found
        while (!targetDoc.querySelector(selector)) {
            await new Promise((resolve) => requestAnimationFrame(resolve))
        }
    }
    return targetDoc.querySelector(selector)
}

/**
 * @description Waits until an element matching the selector has the specified computed CSS property value.
 * @param {string} selector - The selector to query (e.g., 'div[role="region"]')
 * @param {string} cssProp - The camelCase or kebab-case CSS property (e.g., 'containerName')
 * @param {string} cssPropValue - The expected value of the CSS property (e.g., 'captions-history')
 */
async function waitForElementByStyle(selector, cssProp, cssPropValue) {
    while (true) {
        const elements = Array.from(document.querySelectorAll(selector))
        const matchedElement = elements.find(element => {
            const computedStyle = window.getComputedStyle(element)
            // Cast the string to a valid key type of CSSStyleDeclaration to satisfy the compiler
            return computedStyle[/** @type {keyof CSSStyleDeclaration} */ (cssProp)] === cssPropValue
        })

        if (matchedElement) {
            return matchedElement
        }

        await new Promise((resolve) => requestAnimationFrame(resolve))
    }
}

/**
 * @param {string} key
 * @param {string} code
 * @param {KeyboardEventInit | undefined} modifiers
 */
function dispatchKeyboardShortcut(key, code, modifiers) {
    let event = new KeyboardEvent('keydown', {
        key: key,
        code: code,
        ...modifiers
    })
    document.dispatchEvent(event)
}


/**
 * Applies or clears hidden styles on the Google Meet captions container.
 * @param {Element | null} targetNode
 * @param {boolean} shouldHide
 */
function applyCaptionsVisibility(targetNode, shouldHide) {
    if (!targetNode) return
    if (shouldHide) {
        targetNode.setAttribute("style", "opacity:0; height:0px; overflow:hidden; pointer-events:none")
    } else {
        targetNode.removeAttribute("style")
    }
}

/** 
 * @description Single, flat polling monitor that handles initial attachment and all re-attachments.
 * @param {ContentScriptState} state
 */
function startTranscriptMonitor(state) {
    state.transcriptTargetNode = null
    /**
     * @type {number | undefined}
     */
    let monitorInterval = undefined
    let hasHideCaptionsListener = false

    if (typeof chrome !== "undefined" && chrome.storage?.onChanged && !hasHideCaptionsListener) {
        hasHideCaptionsListener = true
        chrome.storage.onChanged.addListener((changes, areaName) => {
            if (areaName === "sync" && changes.hideCaptions && state.transcriptTargetNode) {
                applyCaptionsVisibility(state.transcriptTargetNode, changes.hideCaptions.newValue === true)
            }
        })
    }

    // Call immediately
    transcriptMonitor()
    // Start monitoring
    monitorInterval = setInterval(transcriptMonitor, 2000)

    function transcriptMonitor() {
        if (state.hasMeetingEnded) {
            clearInterval(monitorInterval)
            return
        }

        const activeNode = document.querySelector(SELECTORS_GOOGLE_MEET.TRANSCRIPT_REGION)

        if (!activeNode) {
            return
        }

        // If the active node is new, replaced, or disconnected, re-attach the observer
        if (!state.transcriptTargetNode || activeNode !== state.transcriptTargetNode || !state.transcriptTargetNode.isConnected) {
            if (!state.transcriptTargetNode) {
                console.log("Captions region detected. Attaching observer...")
            }
            else if (activeNode !== state.transcriptTargetNode) {
                console.log("Captions region replaced. Re-attaching observer...")
            }

            // Flush any in-flight buffer to prevent losing text on transitions
            pushBufferToTranscript(state)
            state.stateTranscriptBlock.personName = ""
            state.stateTranscriptBlock.transcriptTextBuffer = ""
            state.stateTranscriptBlock.timestamp = ""

            if (state.transcriptObserver) {
                state.transcriptObserver.disconnect()
            }

            state.transcriptTargetNode = activeNode
            state.transcriptObserver = new MutationObserver((mutations) => {
                transcriptMutationCallbackGoogleMeet(state, mutations)
            })
            state.transcriptObserver.observe(activeNode, mutationConfig)

            // If specified, hide the whole transcript node
            chrome.storage.sync.get(["hideCaptions"], function (resultSyncUntyped) {
                const resultSync = /** @type {ResultSync} */ (resultSyncUntyped)
                applyCaptionsVisibility(state.transcriptTargetNode, resultSync?.hideCaptions === true)
            })
        }
    }
}

/**
 * @param {ContentScriptState} state
 */
function broadcastLiveBuffer(state) {
    /** @type {ExtensionMessage} */
    const message = {
        type: "broadcast_live_buffer",
        stateTranscriptBlock: {
            mutationTargetElement: null,
            personName: state.stateTranscriptBlock.personName,
            timestamp: state.stateTranscriptBlock.timestamp,
            transcriptTextBuffer: state.stateTranscriptBlock.transcriptTextBuffer
        }
    }
    chrome.runtime.sendMessage(message, () => { })
}

/**
 * @param {ContentScriptState} state
 */
function pushBufferToTranscript(state) {
    if ((state.stateTranscriptBlock.personName !== "") && (state.stateTranscriptBlock.transcriptTextBuffer !== "")) {
        state.transcript.push({
            "personName": state.stateTranscriptBlock.personName === "You" ? state.userName : state.stateTranscriptBlock.personName,
            "timestamp": state.stateTranscriptBlock.timestamp,
            "transcriptText": state.stateTranscriptBlock.transcriptTextBuffer
        })
        overWriteChromeStorage(state, ["transcript"], false)
    }
}

/**
 * @description Waits and grabs meeting title from document title
 * @param {ContentScriptState} state
 */
function updateMeetingTitle(state) {
    setTimeout(() => {
        // NON CRITICAL DOM DEPENDENCY
        state.meetingTitle = document.title
        overWriteChromeStorage(state, ["meetingTitle"], false)
    }, 5000)
}

function pulseStatus() {
    const statusActivityCSS = `position: fixed;
    top: 0px;
    width: 100%;
    height: 4px;
    z-index: 100;
    transition: background-color 0.3s ease-in
  `
    /** @type {HTMLDivElement | null}*/
    let activityStatus = document.querySelector(`#transcriptonic-status`)
    if (!activityStatus) {
        let html = document.querySelector("html")
        activityStatus = document.createElement("div")
        activityStatus.setAttribute("id", "transcriptonic-status")
        activityStatus.style.cssText = `background-color: #fac835; ${statusActivityCSS}`
        html?.appendChild(activityStatus)
    }
    else {
        activityStatus.style.cssText = `background-color: #fac835; ${statusActivityCSS}`
    }

    setTimeout(() => {
        activityStatus.style.cssText = `background-color: transparent; ${statusActivityCSS}`
    }, 3000)
}

function renderFab() {
    const fabCss = `
        position: fixed;
        top: 50%;
        bottom: 50%;
        right: 8px;
        height: 36px;
        width: 36px;
        border-radius: 36px;
        z-index: 100;
        display: flex;
        align-items: center;
        justify-content: center;
        background-color: #0c152d;
        box-shadow: 0 4px 14px rgba(0, 0, 0, 0.45), 0 0 0 2px #fac835;
        cursor: pointer;
        border: none;
        padding: 0;
        overflow: visible;
    `

    const html = document.querySelector("html")
    const fab = document.createElement("button")
    fab.id = "transcriptonic-fab"
    fab.ariaLabel = "CommitJr Atas"
    fab.title = "CommitJr Atas - Transcrição ao Vivo"
    fab.style.cssText = fabCss

    const logoUrl = LOGO_URL

    fab.innerHTML = `
        <div id="fab-main-content" style="display: flex; align-items: center; justify-content: center; width: 100%; height: 100%;">
            <img id="fab-default-logo" src="${logoUrl}" alt="CommitJr Atas" draggable="false" style="width: 24px; height: 24px; border-radius: 6px; object-fit: contain;" />
            <span id="fab-letter-mark" style="display: none; color: #fac835; font-weight: 700; font-size: 16px; text-transform: uppercase; font-family: 'Poppins', sans-serif;"></span>
        </div>

        <img id="fab-mini-badge" src="${logoUrl}" alt="Active Badge" draggable="false" style="
            display: none;
            position: absolute;
            bottom: -2px;
            right: -2px;
            width: 14px;
            height: 14px;
            border-radius: 50%;
            background-color: #0c152d;
            box-shadow: 0 0 2px rgba(0,0,0,0.65);
            object-fit: contain;
            pointer-events: none;
        " />
    `

    html?.appendChild(fab)
    makeVerticallyDraggable(fab)

    fab.addEventListener("click", () => {
        /** @type {ExtensionMessage} */
        const message = { type: "open_side_panel" }
        chrome.runtime.sendMessage(message, () => { })
    })

    // Initial storage query on load
    chrome.storage.local.get(["transcript"], (resultUntyped) => {
        const result = /** @type {ResultLocal} */ (resultUntyped)
        updateFabState(/** @type {ContentScriptState} */ ({ transcript: result.transcript || [] }))
    })
}

/**
 * @param {HTMLButtonElement} fab
 */
function makeVerticallyDraggable(fab) {
    let isDragging = false
    let startY = 0
    let initialTop = 0
    let hasMoved = false

    const onPointerDown = (e) => {
        isDragging = true
        hasMoved = false

        const clientY = e.touches ? e.touches[0].clientY : e.clientY
        startY = clientY
        initialTop = fab.getBoundingClientRect().top

        // Attach movement listeners to document so fast drags aren't lost
        document.addEventListener("mousemove", onPointerMove)
        document.addEventListener("mouseup", onPointerUp)
        document.addEventListener("touchmove", onPointerMove, { passive: false })
        document.addEventListener("touchend", onPointerUp)
    }

    const onPointerMove = (e) => {
        if (!isDragging) return

        const clientY = e.touches ? e.touches[0].clientY : e.clientY
        const deltaY = clientY - startY

        // Threshold (3px) to differentiate click from drag
        if (Math.abs(deltaY) > 3) {
            hasMoved = true
            if (e.cancelable) e.preventDefault() // Prevent scrolling on touch
        }

        let newTop = initialTop + deltaY

        // Bound vertical position inside the visible viewport
        const maxTop = window.innerHeight - fab.offsetHeight
        newTop = Math.max(0, Math.min(newTop, maxTop))

        fab.style.top = `${newTop}px`
    }

    const onPointerUp = () => {
        isDragging = false
        document.removeEventListener("mousemove", onPointerMove)
        document.removeEventListener("mouseup", onPointerUp)
        document.removeEventListener("touchmove", onPointerMove)
        document.removeEventListener("touchend", onPointerUp)
    }

    fab.addEventListener("mousedown", onPointerDown)
    fab.addEventListener("touchstart", onPointerDown, { passive: true })

    // Block the 'click' event if the user dragged the button
    fab.addEventListener("click", (e) => {
        if (hasMoved) {
            e.stopImmediatePropagation()
            e.preventDefault()
            hasMoved = false
        }
    }, true) // Capture phase ensures it runs before the side-panel click handler
}

/**
 * Updates the FAB visual state based on the current transcript data.
 * @param {ContentScriptState} state
 */
function updateFabState(state) {
    const fab = document.querySelector("#transcriptonic-fab")

    const defaultLogo = fab?.querySelector("#fab-default-logo")
    const letterMark = fab?.querySelector("#fab-letter-mark")
    const miniBadge = fab?.querySelector("#fab-mini-badge")

    if (state && Array.isArray(state.transcript) && state.transcript.length > 0) {
        const currentSpeaker = state.stateTranscriptBlock?.personName

        if (currentSpeaker && currentSpeaker.trim() !== "") {
            const initial = currentSpeaker.trim().charAt(0)

            // Active Speaker State: Show letter mark + corner badge, hide central logo
            if (defaultLogo) defaultLogo.style.display = "none"
            if (letterMark) {
                letterMark.textContent = initial
                letterMark.style.display = "inline"
            }
            if (miniBadge) miniBadge.style.display = "block"
            return
        }
    }

    // Default State: Fallback to central logo, hide mark + badge
    if (defaultLogo) defaultLogo.style.display = "block"
    if (letterMark) letterMark.style.display = "none"
    if (miniBadge) miniBadge.style.display = "none"
}

function unmountFab() {
    const fab = document.querySelector("#transcriptonic-fab")
    if (fab) {
        fab.remove()
    }
}

/**
   * @description Logs active transcript to console
   * @param {ContentScriptState} state
   */
function logTranscriptToConsole(state) {
    if (state.stateTranscriptBlock.transcriptTextBuffer.length > 125) {
        console.log(state.stateTranscriptBlock.transcriptTextBuffer.slice(0, 50) + "   ...   " + state.stateTranscriptBlock.transcriptTextBuffer.slice(-50))
    }
    else {
        console.log(state.stateTranscriptBlock.transcriptTextBuffer)
    }
}

/**
   * @description Logs errors locally without sending external telemetry.
   * @param {ContentScriptState} state
   * @param {string} code
   * @param {any} err
   */
function logError(state, code, err) {
    console.error(`[CommitJr Atas] Error ${code} (${state?.meetingSoftware || "unknown"}):`, err)
}

/**
   * @description Checks if the installed extension version meets the minimum required version.
   * @param {string} oldVer
   * @param {string} newVer
   */
function meetsMinVersion(oldVer, newVer) {
    const oldParts = oldVer.split('.')
    const newParts = newVer.split('.')
    for (var i = 0; i < newParts.length; i++) {
        const a = ~~newParts[i] // parse int
        const b = ~~oldParts[i] // parse int
        if (a > b) return false
        if (a < b) return true
    }
    return true
}

// === Google Meet captions & notification helpers ===

/**
 * @description Returns all elements of the specified selector type and specified textContent. Return array contains the actual element as well as all the parents.
 * @param {string} selector
 * @param {string | RegExp} text
 */
function selectElements(selector, text) {
    var elements = document.querySelectorAll(selector)
    return Array.prototype.filter.call(elements, function (element) {
        return RegExp(text).test(element.textContent)
    })
}

/**
   * @description Shows a responsive notification of specified type and message
   * @param {ExtensionStatusJSON} extensionStatusJSON
   */
function showNotificationGoogleMeet(extensionStatusJSON) {
    // Banner CSS
    let html = document.querySelector("html")
    let obj = document.createElement("div")
    let logo = document.createElement("img")
    let text = document.createElement("p")

    logo.setAttribute(
        "src",
        LOGO_URL
    )
    logo.setAttribute("height", "32px")
    logo.setAttribute("width", "32px")
    logo.style.cssText = "border-radius: 4px"

    // Remove banner after 5s
    setTimeout(() => {
        obj.style.display = "none"
    }, 5000)

    if (extensionStatusJSON.status === 200) {
        obj.style.cssText = `color: #ffffff; top: 5%; ${commonCSS}`
        text.innerHTML = extensionStatusJSON.message
    }
    else {
        obj.style.cssText = `color: #f69191; top: 5%; ${commonCSS} border-color: #f69191;`
        text.innerHTML = extensionStatusJSON.message
    }

    obj.prepend(text)
    obj.prepend(logo)
    if (html)
        html.append(obj)
}

/**
 * @description Grabs updated meeting title, if available
 * @param {ContentScriptState} state
 */
function updateMeetingTitleGoogleMeet(state) {
    waitForElement(SELECTORS_GOOGLE_MEET.MEETING_TITLE).then((element) => {
        const meetingTitleElement = /** @type {HTMLDivElement} */ (element)
        meetingTitleElement?.setAttribute("contenteditable", "true")
        meetingTitleElement.title = "Editar título da reunião para CommitJr Atas"
        meetingTitleElement.style.cssText = `text-decoration: underline #fac835; text-underline-offset: 4px;`

        meetingTitleElement?.addEventListener("input", handleMeetingTitleElementChange)

        // Pick up meeting name after a delay, since Google meet updates meeting name after a delay
        setTimeout(() => {
            handleMeetingTitleElementChange()
            if (location.pathname === `/${meetingTitleElement.innerText}`) {
                showNotificationGoogleMeet({ status: 200, message: `<strong style="color: #fac835;">Deseja nomear esta reunião?</strong><br/>Edite o texto sublinhado no canto inferior esquerdo` })
            }
        }, 7000)

        function handleMeetingTitleElementChange() {
            state.meetingTitle = meetingTitleElement.innerText
            overWriteChromeStorage(state, ["meetingTitle"], false)
        }
    })
}

/**
 * @description Attempt to get username before meeting starts. Abort interval if valid username is found or if meeting starts and default to "You"
 * @param {ContentScriptState} state
 */
function captureUserName(state) {
    waitForElement(SELECTORS_GOOGLE_MEET.USER_NAME).then(() => {
        // Poll the element until the textContent loads from network or until meeting starts
        const captureUserNameInterval = setInterval(() => {
            if (!state.hasMeetingStarted) {
                const capturedUserName = document.querySelector(SELECTORS_GOOGLE_MEET.USER_NAME)?.textContent
                if (capturedUserName) {
                    state.userName = capturedUserName
                    clearInterval(captureUserNameInterval)
                }
            }
            else {
                clearInterval(captureUserNameInterval)
            }
        }, 100)
    })
}

/**
 * @description Pushes object to array only if it doesn't already exist.
 * @param {ContentScriptState} state
 * @param {ChatMessage} chatBlock
 */
function pushUniqueChatBlock(state, chatBlock) {
    const isExisting = state.chatMessages.some(item =>
        (item.personName === chatBlock.personName) &&
        (item.chatMessageText === chatBlock.chatMessageText)
    )
    if (!isExisting) {
        console.log(chatBlock)
        state.chatMessages.push(chatBlock)
        overWriteChromeStorage(state, ["chatMessages"], false)
    }
}

/**
 * @description Returns the pressed state of the captions button: "true", "false", or null if the button has no aria-pressed attribute
 * @returns {string | null}
 */
function getCaptionsPressedStateGoogleMeet() {
    const captionsIcon = Array.from(document.querySelectorAll(SELECTORS_GOOGLE_MEET.GOOGLE_SYMBOLS))
        .find(element => SELECTORS_GOOGLE_MEET.TEXT_CAPTIONS.includes(element.textContent ?? ""))
    return captionsIcon?.closest("button")?.getAttribute("aria-pressed") ?? null
}

/**
 * @description Turns on captions using the "c" keyboard shortcut. Meet may not have its shortcut handler ready when the captions icon first appears, so the first keypress can be silently dropped. Verify via aria-pressed and retry. Never sends "c" when captions are already on, since the shortcut toggles them.
 */
async function turnOnCaptionsGoogleMeet() {
    const maxAttempts = 10
    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
        const pressedState = getCaptionsPressedStateGoogleMeet()

        // Captions are on ("true"), or state cannot be verified (null) -> do not risk toggling
        if ((pressedState === "true") || (pressedState === null)) {
            return
        }
        dispatchKeyboardShortcut("c", "KeyC", { bubbles: true })

        await new Promise((resolve) => setTimeout(resolve, 1500))
    }
    console.error("Could not turn on captions")
}
