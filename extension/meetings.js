// @ts-check
/// <reference path="../types/chrome.d.ts" />
/// <reference path="../types/index.js" />

import {
    DEFAULT_GEMINI_MODEL,
    resolveGeminiModel,
    escapeHtml,
    buildAtaHtml,
    buildTranscriptAccessInfo,
    formatDateTimeBr,
    extractTextFromDocxBytes
} from "./background-script/ata-generator.js"
import { getTranscriptFileName } from "./background-script/utils.js"

const MAX_TEMPLATE_FILE_BYTES = 3 * 1024 * 1024

let isMeetingsTableExpanded = false
/** @type {number | null} */
let activeTranscriptModalIndex = null
/** @type {number | null} */
let activeAtaPreviewIndex = null

/**
 * Triggers a local file download from a Blob in the page context.
 * @param {Blob} blob
 * @param {string} fileName
 */
function triggerPageBlobDownload(blob, fileName) {
    const safeName = String(fileName || "TranscripTonic-export.bin")
        .split("/")
        .pop()
        ?.split("\\")
        .pop() || "TranscripTonic-export.bin"
    const url = URL.createObjectURL(blob)
    const link = document.createElement("a")
    link.href = url
    link.download = safeName
    link.style.display = "none"
    document.body.appendChild(link)
    link.click()
    link.remove()
    setTimeout(() => {
        try {
            URL.revokeObjectURL(url)
        } catch (_err) {
            // Ignore cleanup errors
        }
    }, 10000)
}

/**
 * Encodes a Uint8Array as Base64 without blowing the call stack on large templates.
 * @param {Uint8Array} bytes
 * @returns {string}
 */
function uint8ArrayToBase64(bytes) {
    const chunkSize = 0x8000
    let binary = ""
    for (let i = 0; i < bytes.length; i += chunkSize) {
        binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize))
    }
    return btoa(binary)
}

/**
 * Decodes a Base64 string into a Uint8Array.
 * @param {string} base64Data
 * @returns {Uint8Array}
 */
function base64ToUint8Array(base64Data) {
    const binary = atob(base64Data)
    const bytes = new Uint8Array(binary.length)
    for (let i = 0; i < binary.length; i++) {
        bytes[i] = binary.charCodeAt(i)
    }
    return bytes
}

/**
 * Formats duration between two ISO timestamps.
 * @param {string} meetingStartTimestamp
 * @param {string} meetingEndTimestamp
 * @returns {string}
 */
export function getDuration(meetingStartTimestamp, meetingEndTimestamp) {
    const duration = new Date(meetingEndTimestamp).getTime() - new Date(meetingStartTimestamp).getTime()
    if (Number.isNaN(duration)) {
        return "0m"
    }
    const durationMinutes = Math.max(0, Math.round(duration / (1000 * 60)))
    const durationHours = Math.floor(durationMinutes / 60)
    const remainingMinutes = durationMinutes % 60
    return durationHours > 0
        ? `${durationHours}h ${remainingMinutes}m`
        : `${durationMinutes}m`
}

/**
 * Builds escaped HTML markup for a single row in #meetings-table.
 * @param {Meeting} meeting
 * @param {number} index
 * @returns {string}
 */
export function buildMeetingRowHtml(meeting, index) {
    const safeIndex = Number(index)
    const rawTitle = meeting.meetingTitle || meeting.title || "Google Meet call"
    const safeTitle = escapeHtml(rawTitle)
    const safeSoftware = escapeHtml(meeting.meetingSoftware ? meeting.meetingSoftware : "")
    const timestampValue = new Date(meeting.meetingStartTimestamp)
    const safeTimestamp = escapeHtml(
        Number.isNaN(timestampValue.getTime()) ? String(meeting.meetingStartTimestamp || "") : timestampValue.toLocaleString()
    )
    const safeDuration = escapeHtml(getDuration(meeting.meetingStartTimestamp, meeting.meetingEndTimestamp))

    let webhookStatusHtml = `<span class="status-new">Unknown</span>`
    switch (meeting.webhookPostStatus) {
        case "successful":
            webhookStatusHtml = `<span class="status-success">Successful</span>`
            break
        case "failed":
            webhookStatusHtml = `<span class="status-failed">Failed</span>`
            break
        case "new":
            webhookStatusHtml = `<span class="status-new">New</span>`
            break
    }

    const hasAta = Boolean(meeting.ataData)
    const ataStatusHtml = hasAta
        ? `<span class="status-success">Ata gerada</span>${meeting.ataModel ? `<div class="sub-text" style="font-size: 0.78rem;">${escapeHtml(meeting.ataModel)}</div>` : ""}`
        : `<span class="sub-text">Não gerada</span>`

    const ataActionsHtml = `
        <div style="display: flex; flex-wrap: wrap; gap: 0.4rem; margin-top: 0.4rem;">
            <button type="button" class="action-pill-btn primary generate-ata-btn" data-index="${safeIndex}">
                ${hasAta ? "Regerar Ata" : "Gerar Ata"}
            </button>
            ${hasAta ? `
                <button type="button" class="action-pill-btn download-ata-docx-btn" data-index="${safeIndex}">Baixar .docx</button>
                <button type="button" class="action-pill-btn preview-ata-btn" data-index="${safeIndex}">Visualizar / PDF</button>
            ` : ""}
            <button type="button" class="action-pill-btn view-transcript-btn" data-index="${safeIndex}">Ver Transcrição</button>
        </div>
    `

    const postTitle = meeting.webhookPostStatus === "new" ? "Post webhook" : "Repost webhook"

    return `
        <td>
            <div contenteditable="true" class="meeting-title" data-index="${safeIndex}" title="Rename">${safeTitle}</div>
        </td>
        <td>${safeSoftware}</td>
        <td>${safeTimestamp} &nbsp; &#9679; &nbsp; ${safeDuration}</td>
        <td>${webhookStatusHtml}</td>
        <td>
            ${ataStatusHtml}
            ${ataActionsHtml}
        </td>
        <td>
            <div style="display: flex; gap: 1rem; justify-content: end">
                <button type="button" class="download-button" data-index="${safeIndex}" title="Download" aria-label="Download this meeting transcript">
                    <img src="./icons/download.svg" alt="">
                </button>
                <button type="button" class="post-button" data-index="${safeIndex}" title="${postTitle}" aria-label="${postTitle}">
                    <img src="./icons/webhook.svg" alt="">
                </button>
                &nbsp;
                <button type="button" class="delete-button" data-index="${safeIndex}" title="Delete" aria-label="Delete this meeting">
                    <img src="./icons/delete.svg" alt="">
                </button>
            </div>
        </td>
    `
}

if (typeof chrome !== "undefined" && chrome.runtime?.onMessage) {
    chrome.runtime.onMessage.addListener(function (messageUntyped, _sender, sendResponse) {
        const message = /** @type {any} */ (messageUntyped)
        if (!message) {
            return false
        }

        if (message.type === "download_blob_via_page") {
            try {
                const bytes = base64ToUint8Array(message.base64Data || "")
                const blob = new Blob([bytes], { type: message.mimeType || "application/octet-stream" })
                triggerPageBlobDownload(blob, message.fileName || "TranscripTonic-export.bin")
                sendResponse({ success: true })
            } catch (err) {
                console.error("Failed download_blob_via_page:", err)
                sendResponse({ success: false, message: err instanceof Error ? err.message : String(err) })
            }
            return false
        }

        if (message.type === "download_transcript_blob") {
            try {
                const blob = new Blob([String(message.content || "")], { type: "text/plain;charset=utf-8" })
                triggerPageBlobDownload(blob, message.fileName || "Transcript.txt")
                sendResponse({ success: true })
            } catch (err) {
                console.error("Failed download_transcript_blob:", err)
                sendResponse({ success: false, message: err instanceof Error ? err.message : String(err) })
            }
            return false
        }

        return false
    })
}

if (typeof document !== "undefined") {
    document.addEventListener("DOMContentLoaded", function () {
        const webhookUrlForm = document.querySelector("#webhook-url-form")
        const webhookUrlInput = document.querySelector("#webhook-url")
        const saveButton = document.querySelector("#save-webhook")
        const autoPostCheckbox = document.querySelector("#auto-post-webhook")
        const autoDownloadCheckbox = document.querySelector("#auto-download-file")
        const simpleWebhookBodyRadio = document.querySelector("#simple-webhook-body")
        const advancedWebhookBodyRadio = document.querySelector("#advanced-webhook-body")
        const recoverLastMeetingButton = document.querySelector("#recover-last-meeting")
        const showAllButton = document.querySelector("#show-all")

        // Initial load of transcripts and Ata/Gemini settings
        loadMeetings()
        initGeminiAndTemplateSettings()
        initModals()

        const queryParams = new URLSearchParams(window.location.search)
        const transcriptParam = queryParams.get("transcript")
        if (transcriptParam !== null && /^\d+$/.test(transcriptParam)) {
            openTranscriptModal(parseInt(transcriptParam, 10))
        }

        // Reload transcripts when page becomes visible
        document.addEventListener("visibilitychange", function () {
            if (document.visibilityState === "visible") {
                loadMeetings()
            }
        })

        chrome.storage.onChanged.addListener((changes, areaName) => {
            if (areaName === "local" && (changes.meetings || changes.geminiApiKey || changes.customAtaTemplate)) {
                loadMeetings()
                refreshGeminiLocalUi()
            } else if (areaName === "sync") {
                loadMeetings()
            }
        })

        if (recoverLastMeetingButton instanceof HTMLButtonElement) {
            recoverLastMeetingButton.addEventListener("click", function () {
                /** @type {ExtensionMessage} */
                const message = {
                    type: "recover_last_meeting"
                }
                chrome.runtime.sendMessage(message, function (responseUntyped) {
                    const response = /** @type {ExtensionResponse} */ (responseUntyped)
                    loadMeetings()
                    scrollTo({ top: 0, behavior: "smooth" })
                    if (response?.success) {
                        if (response.message === "No recovery needed") {
                            alert("Nothing to recover—you're on top of the world!")
                        } else {
                            alert("Last meeting recovered successfully!")
                        }
                    } else {
                        const parsedError = /** @type {ErrorObject} */ (response?.message)
                        if (parsedError?.errorCode === "013") {
                            alert(parsedError.errorMessage)
                        } else if (parsedError?.errorCode === "014") {
                            alert("Nothing to recover—you're on top of the world!")
                        } else {
                            alert("Could not recover last meeting!")
                            console.error(parsedError?.errorMessage)
                        }
                    }
                })
            })
        }

        if (
            saveButton instanceof HTMLButtonElement &&
            webhookUrlForm instanceof HTMLFormElement &&
            webhookUrlInput instanceof HTMLInputElement &&
            autoPostCheckbox instanceof HTMLInputElement &&
            simpleWebhookBodyRadio instanceof HTMLInputElement &&
            advancedWebhookBodyRadio instanceof HTMLInputElement
        ) {
            // Initially disable the save button
            saveButton.disabled = true

            // Load saved webhook URL, auto-post setting, and webhook body type
            chrome.storage.sync.get(
                ["webhookUrl", "autoPostWebhookAfterMeeting", "autoDownloadFileAfterMeeting", "webhookBodyType"],
                function (resultSyncUntyped) {
                    const resultSync = /** @type {ResultSync} */ (resultSyncUntyped)

                    if (resultSync.webhookUrl) {
                        webhookUrlInput.value = resultSync.webhookUrl
                        saveButton.disabled = !webhookUrlInput.checkValidity()
                    }

                    // Set checkbox state
                    autoPostCheckbox.checked = Boolean(resultSync.autoPostWebhookAfterMeeting)
                    if (autoDownloadCheckbox instanceof HTMLInputElement) {
                        autoDownloadCheckbox.checked = resultSync.autoDownloadFileAfterMeeting !== false
                    }
                    updateAutoDownloadCheckBox()

                    // Set radio button state
                    if (resultSync.webhookBodyType === "advanced") {
                        advancedWebhookBodyRadio.checked = true
                    } else {
                        simpleWebhookBodyRadio.checked = true
                    }
                }
            )

            // Handle URL input changes
            webhookUrlInput.addEventListener("input", function () {
                saveButton.disabled = !webhookUrlInput.checkValidity()
            })

            // Save webhook URL, auto-post setting, and webhook body type
            webhookUrlForm.addEventListener("submit", function (e) {
                e.preventDefault()
                const webhookUrl = webhookUrlInput.value
                if (webhookUrl === "") {
                    chrome.storage.sync.set({
                        webhookUrl: webhookUrl
                    }, function () {
                        alert("Webhook URL saved!")
                    })
                } else if (webhookUrl && webhookUrlInput.checkValidity()) {
                    requestWebhookAndNotificationPermission(webhookUrl)
                        .then(() => {
                            chrome.storage.sync.set({
                                webhookUrl: webhookUrl
                            }, function () {
                                alert("Webhook URL saved!")
                            })
                        })
                        .catch((error) => {
                            alert("Fine! No webhooks for you!")
                            console.error("Webhook permission error:", error)
                        })
                }
            })

            // Auto save auto-post setting
            autoPostCheckbox.addEventListener("change", function () {
                chrome.storage.sync.set({
                    autoPostWebhookAfterMeeting: autoPostCheckbox.checked
                }, function () {
                    updateAutoDownloadCheckBox()
                })
            })

            if (autoDownloadCheckbox instanceof HTMLInputElement) {
                autoDownloadCheckbox.addEventListener("change", function () {
                    if (!autoDownloadCheckbox.checked) {
                        if (!confirm("Text file serves as a harmless backup, you sure you don't need it?")) {
                            autoDownloadCheckbox.checked = true
                            return
                        }
                    }
                    chrome.storage.sync.set({
                        autoDownloadFileAfterMeeting: autoDownloadCheckbox.checked
                    }, function () { })
                })
            }

            function updateAutoDownloadCheckBox() {
                if (autoDownloadCheckbox?.parentElement instanceof HTMLDivElement && autoPostCheckbox instanceof HTMLInputElement) {
                    autoDownloadCheckbox.parentElement.style.display = autoPostCheckbox.checked ? "flex" : "none"
                    if (!autoPostCheckbox.checked && autoDownloadCheckbox instanceof HTMLInputElement) {
                        autoDownloadCheckbox.checked = true
                        chrome.storage.sync.set({
                            autoDownloadFileAfterMeeting: true
                        }, function () { })
                    }
                }
            }

            // Auto save webhook body type
            simpleWebhookBodyRadio.addEventListener("change", function () {
                chrome.storage.sync.set({ webhookBodyType: "simple" }, function () { })
            })

            advancedWebhookBodyRadio.addEventListener("change", function () {
                chrome.storage.sync.set({ webhookBodyType: advancedWebhookBodyRadio.checked ? "advanced" : "simple" }, function () { })
            })
        }

        if (showAllButton instanceof HTMLButtonElement) {
            showAllButton.addEventListener("click", () => {
                const meetingsTableContainer = document.querySelector("#meetings-table-container")
                meetingsTableContainer?.classList.remove("fade-mask")
                showAllButton.setAttribute("style", "display:none;")
                isMeetingsTableExpanded = true
            })
        }
    })
}

/**
 * Refreshes local Gemini key badge and active Ata template label.
 */
function refreshGeminiLocalUi() {
    const geminiKeyInput = /** @type {HTMLInputElement | null} */ (document.querySelector("#gemini-api-key"))
    const clearGeminiKeyButton = /** @type {HTMLButtonElement | null} */ (document.querySelector("#clear-gemini-key"))
    const geminiKeyStatus = document.querySelector("#gemini-key-status")
    const ataTemplateStatus = document.querySelector("#ata-template-status")

    chrome.storage.local.get(["geminiApiKey", "customAtaTemplate"], function (resultLocalUntyped) {
        const resultLocal = /** @type {ResultLocal} */ (resultLocalUntyped || {})
        const hasKey = Boolean(resultLocal.geminiApiKey && String(resultLocal.geminiApiKey).trim())

        if (geminiKeyStatus) {
            geminiKeyStatus.textContent = hasKey ? "Chave configurada ✓" : "Chave não configurada"
            geminiKeyStatus.className = hasKey ? "status-success" : "sub-text"
        }
        if (clearGeminiKeyButton) {
            clearGeminiKeyButton.style.display = hasKey ? "inline-block" : "none"
        }
        if (geminiKeyInput) {
            geminiKeyInput.placeholder = hasKey ? "Chave salva (digite outra para substituir)" : "AIza..."
        }
        if (ataTemplateStatus) {
            const custom = resultLocal.customAtaTemplate
            ataTemplateStatus.textContent = custom?.fileName
                ? `Personalizado: ${custom.fileName}`
                : "Padrão embutido (MODELO ATA.docx)"
        }
    })
}

/**
 * Initializes Gemini API settings, defaults, and custom .pdf/.docx template upload/reset.
 */
function initGeminiAndTemplateSettings() {
    const geminiForm = /** @type {HTMLFormElement | null} */ (document.querySelector("#gemini-config-form"))
    const geminiKeyInput = /** @type {HTMLInputElement | null} */ (document.querySelector("#gemini-api-key"))
    const clearGeminiKeyButton = /** @type {HTMLButtonElement | null} */ (document.querySelector("#clear-gemini-key"))
    const geminiFeedback = document.querySelector("#gemini-feedback")
    const geminiModelSelect = /** @type {HTMLSelectElement | null} */ (document.querySelector("#gemini-model"))
    const autoGenerateAtaCheckbox = /** @type {HTMLInputElement | null} */ (document.querySelector("#auto-generate-ata"))
    const ataClientNameInput = /** @type {HTMLInputElement | null} */ (document.querySelector("#ata-client-name"))
    const ataSubdiretoriaInput = /** @type {HTMLInputElement | null} */ (document.querySelector("#ata-subdiretoria"))
    const ataTemplateFileInput = /** @type {HTMLInputElement | null} */ (document.querySelector("#ata-template-file"))
    const resetAtaTemplateButton = /** @type {HTMLButtonElement | null} */ (document.querySelector("#reset-ata-template"))

    refreshGeminiLocalUi()

    chrome.storage.sync.get(
        ["geminiModel", "autoGenerateAtaAfterMeeting", "ataClientName", "ataSubdiretoria"],
        function (resultSyncUntyped) {
            const resultSync = /** @type {ResultSync} */ (resultSyncUntyped || {})
            if (geminiModelSelect) {
                const resolvedModel = resolveGeminiModel(resultSync.geminiModel)
                geminiModelSelect.value = resolvedModel
                if (resultSync.geminiModel !== resolvedModel) {
                    chrome.storage.sync.set({ geminiModel: resolvedModel })
                }
            }
            if (autoGenerateAtaCheckbox) {
                autoGenerateAtaCheckbox.checked = resultSync.autoGenerateAtaAfterMeeting === true
            }
            if (ataClientNameInput) {
                ataClientNameInput.value = resultSync.ataClientName || ""
            }
            if (ataSubdiretoriaInput) {
                ataSubdiretoriaInput.value = resultSync.ataSubdiretoria || ""
            }
        }
    )

    if (geminiForm && geminiKeyInput) {
        geminiForm.addEventListener("submit", function (event) {
            event.preventDefault()
            const key = geminiKeyInput.value.trim()
            if (!key) {
                if (geminiFeedback) {
                    geminiFeedback.textContent = "Informe uma chave API válida do Google Gemini."
                }
                return
            }
            chrome.storage.local.set({ geminiApiKey: key }, function () {
                geminiKeyInput.value = ""
                refreshGeminiLocalUi()
                if (geminiFeedback) {
                    geminiFeedback.textContent = "Chave API do Google Gemini salva localmente com segurança."
                }
            })
        })
    }

    if (clearGeminiKeyButton) {
        clearGeminiKeyButton.addEventListener("click", function () {
            chrome.storage.local.remove(["geminiApiKey"], function () {
                refreshGeminiLocalUi()
                if (geminiFeedback) {
                    geminiFeedback.textContent = "Chave API removida."
                }
            })
        })
    }

    if (geminiModelSelect) {
        geminiModelSelect.addEventListener("change", function () {
            chrome.storage.sync.set({ geminiModel: geminiModelSelect.value || DEFAULT_GEMINI_MODEL })
        })
    }

    if (autoGenerateAtaCheckbox) {
        autoGenerateAtaCheckbox.addEventListener("change", function () {
            chrome.storage.sync.set({ autoGenerateAtaAfterMeeting: autoGenerateAtaCheckbox.checked })
        })
    }

    /** @type {ReturnType<typeof setTimeout> | null} */
    let clientSaveTimer = null
    if (ataClientNameInput) {
        const persistClient = () => {
            chrome.storage.sync.set({ ataClientName: ataClientNameInput.value.trim() })
        }
        ataClientNameInput.addEventListener("change", persistClient)
        ataClientNameInput.addEventListener("input", () => {
            if (clientSaveTimer) clearTimeout(clientSaveTimer)
            clientSaveTimer = setTimeout(persistClient, 400)
        })
    }

    /** @type {ReturnType<typeof setTimeout> | null} */
    let subdirSaveTimer = null
    if (ataSubdiretoriaInput) {
        const persistSubdir = () => {
            chrome.storage.sync.set({ ataSubdiretoria: ataSubdiretoriaInput.value.trim() })
        }
        ataSubdiretoriaInput.addEventListener("change", persistSubdir)
        ataSubdiretoriaInput.addEventListener("input", () => {
            if (subdirSaveTimer) clearTimeout(subdirSaveTimer)
            subdirSaveTimer = setTimeout(persistSubdir, 400)
        })
    }

    if (ataTemplateFileInput) {
        ataTemplateFileInput.addEventListener("change", async function () {
            const file = ataTemplateFileInput.files?.[0]
            if (!file) return

            const lowerName = file.name.toLowerCase()
            const isPdf = lowerName.endsWith(".pdf") || file.type === "application/pdf"
            const isDocx =
                lowerName.endsWith(".docx") ||
                file.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

            if (!isPdf && !isDocx) {
                alert("Formato não suportado. Envie um arquivo .pdf ou .docx.")
                ataTemplateFileInput.value = ""
                return
            }

            if (file.size > MAX_TEMPLATE_FILE_BYTES) {
                alert("O arquivo de modelo deve ter no máximo 3 MB.")
                ataTemplateFileInput.value = ""
                return
            }

            try {
                const buffer = await file.arrayBuffer()
                const bytes = new Uint8Array(buffer)
                let extractedText = undefined
                if (isDocx) {
                    extractedText = await extractTextFromDocxBytes(bytes)
                }

                /** @type {CustomAtaTemplate} */
                const customAtaTemplate = {
                    fileName: file.name,
                    mimeType: isPdf
                        ? "application/pdf"
                        : "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    base64Data: uint8ArrayToBase64(bytes),
                    extractedText
                }

                chrome.storage.local.set({ customAtaTemplate }, function () {
                    ataTemplateFileInput.value = ""
                    refreshGeminiLocalUi()
                    if (geminiFeedback) {
                        geminiFeedback.textContent = `Modelo personalizado "${file.name}" carregado.`
                    }
                })
            } catch (err) {
                console.error("Erro ao processar modelo personalizado:", err)
                alert("Não foi possível ler o arquivo de modelo selecionado.")
                ataTemplateFileInput.value = ""
            }
        })
    }

    if (resetAtaTemplateButton) {
        resetAtaTemplateButton.addEventListener("click", function () {
            chrome.storage.local.remove(["customAtaTemplate"], function () {
                refreshGeminiLocalUi()
                if (geminiFeedback) {
                    geminiFeedback.textContent = "Modelo padrão restaurado (MODELO ATA.docx)."
                }
            })
        })
    }
}

/**
 * Wires controls inside the transcript viewer and Ata HTML/PDF preview modals.
 */
function initModals() {
    const transcriptModal = /** @type {HTMLDialogElement | null} */ (document.querySelector("#transcript-modal"))
    const transcriptCloseBtn = document.querySelector("#transcript-modal-close")
    const transcriptDownloadBtn = document.querySelector("#transcript-modal-download")

    if (transcriptCloseBtn && transcriptModal) {
        transcriptCloseBtn.addEventListener("click", () => {
            transcriptModal.close()
        })
    }

    if (transcriptDownloadBtn) {
        transcriptDownloadBtn.addEventListener("click", () => {
            if (activeTranscriptModalIndex === null) return
            /** @type {ExtensionMessage} */
            const message = {
                type: "download_transcript_at_index",
                index: activeTranscriptModalIndex
            }
            chrome.runtime.sendMessage(message, (responseUntyped) => {
                const response = /** @type {ExtensionResponse} */ (responseUntyped)
                if (!response?.success) {
                    alert("Não foi possível baixar a transcrição.")
                }
            })
        })
    }

    const ataPreviewModal = /** @type {HTMLDialogElement | null} */ (document.querySelector("#ata-preview-modal"))
    const ataPreviewCloseBtn = document.querySelector("#ata-preview-close")
    const ataPreviewPrintBtn = document.querySelector("#ata-preview-print")
    const ataPreviewDocxBtn = document.querySelector("#ata-preview-docx")
    const ataPreviewFrame = /** @type {HTMLIFrameElement | null} */ (document.querySelector("#ata-preview-frame"))

    if (ataPreviewCloseBtn && ataPreviewModal) {
        ataPreviewCloseBtn.addEventListener("click", () => {
            ataPreviewModal.close()
        })
    }

    if (ataPreviewPrintBtn && ataPreviewFrame) {
        ataPreviewPrintBtn.addEventListener("click", () => {
            try {
                ataPreviewFrame.contentWindow?.focus()
                ataPreviewFrame.contentWindow?.print()
            } catch (err) {
                console.error("Failed to print Ata preview:", err)
                alert("Use o menu Imprimir do navegador para salvar em PDF.")
            }
        })
    }

    if (ataPreviewDocxBtn) {
        ataPreviewDocxBtn.addEventListener("click", () => {
            if (activeAtaPreviewIndex === null) return
            /** @type {ExtensionMessage} */
            const message = {
                type: "download_ata_at_index",
                index: activeAtaPreviewIndex,
                format: "docx"
            }
            chrome.runtime.sendMessage(message, (responseUntyped) => {
                const response = /** @type {ExtensionResponse} */ (responseUntyped)
                if (!response?.success) {
                    const parsedError = /** @type {ErrorObject} */ (response?.message)
                    alert(parsedError?.errorMessage || "Não foi possível baixar a Ata (.docx).")
                }
            })
        })
    }
}

/**
 * Opens the full verbatim transcript modal safely using DOM textContent nodes.
 * @param {number} index
 */
function openTranscriptModal(index) {
    const transcriptModal = /** @type {HTMLDialogElement | null} */ (document.querySelector("#transcript-modal"))
    const modalTitle = document.querySelector("#transcript-modal-title")
    const modalMeta = document.querySelector("#transcript-modal-meta")
    const modalBody = document.querySelector("#transcript-modal-body")
    if (!transcriptModal || !modalTitle || !modalMeta || !modalBody) return

    chrome.storage.local.get(["meetings"], function (resultLocalUntyped) {
        const resultLocal = /** @type {ResultLocal} */ (resultLocalUntyped || {})
        const meetings = resultLocal.meetings || []
        const meeting = meetings[index]
        if (!meeting) {
            alert("Reunião não encontrada no armazenamento local.")
            return
        }

        activeTranscriptModalIndex = index
        const title = meeting.meetingTitle || meeting.title || "Transcrição Completa"
        modalTitle.textContent = title
        modalMeta.textContent = `${meeting.meetingSoftware || "Reunião"} • Início: ${formatDateTimeBr(meeting.meetingStartTimestamp)} • Duração: ${getDuration(meeting.meetingStartTimestamp, meeting.meetingEndTimestamp)} • Arquivo: ${getTranscriptFileName(meeting)}`

        modalBody.innerHTML = ""
        const transcriptBlocks = Array.isArray(meeting.transcript) ? meeting.transcript : []
        const chatBlocks = Array.isArray(meeting.chatMessages) ? meeting.chatMessages : []

        if (transcriptBlocks.length === 0 && chatBlocks.length === 0) {
            const empty = document.createElement("p")
            empty.textContent = "Nenhuma fala ou mensagem de chat foi registrada nesta reunião."
            modalBody.appendChild(empty)
        } else {
            for (const block of transcriptBlocks) {
                const entry = document.createElement("div")
                entry.className = "transcript-entry"

                const meta = document.createElement("div")
                meta.className = "transcript-entry-meta"
                meta.textContent = `${block.personName || "Participante"} (${formatDateTimeBr(block.timestamp)})`

                const text = document.createElement("div")
                text.textContent = block.transcriptText || ""

                entry.appendChild(meta)
                entry.appendChild(text)
                modalBody.appendChild(entry)
            }

            if (chatBlocks.length > 0) {
                const heading = document.createElement("h3")
                heading.textContent = "Mensagens do Chat"
                heading.style.marginTop = "1.25rem"
                modalBody.appendChild(heading)

                for (const msg of chatBlocks) {
                    const entry = document.createElement("div")
                    entry.className = "transcript-entry"

                    const meta = document.createElement("div")
                    meta.className = "transcript-entry-meta"
                    meta.textContent = `${msg.personName || "Participante"} (${formatDateTimeBr(msg.timestamp)})`

                    const text = document.createElement("div")
                    text.textContent = msg.chatMessageText || ""

                    entry.appendChild(meta)
                    entry.appendChild(text)
                    modalBody.appendChild(entry)
                }
            }
        }

        if (!transcriptModal.open) {
            transcriptModal.showModal()
        }
    })
}

/**
 * Opens the interactive HTML preview / print-to-PDF modal for an already-generated Ata.
 * @param {number} index
 */
function openAtaPreviewModal(index) {
    const ataPreviewModal = /** @type {HTMLDialogElement | null} */ (document.querySelector("#ata-preview-modal"))
    const ataPreviewTitle = document.querySelector("#ata-preview-title")
    const ataPreviewFrame = /** @type {HTMLIFrameElement | null} */ (document.querySelector("#ata-preview-frame"))
    if (!ataPreviewModal || !ataPreviewTitle || !ataPreviewFrame) return

    chrome.storage.local.get(["meetings"], function (resultLocalUntyped) {
        const resultLocal = /** @type {ResultLocal} */ (resultLocalUntyped || {})
        const meetings = resultLocal.meetings || []
        const meeting = meetings[index]
        if (!meeting || !meeting.ataData) {
            alert("Gere a Ata primeiro antes de visualizar.")
            return
        }

        activeAtaPreviewIndex = index
        ataPreviewTitle.textContent = `Ata — ${meeting.meetingTitle || meeting.title || "Reunião"}`
        const accessInfo = buildTranscriptAccessInfo(meeting, index)
        ataPreviewFrame.srcdoc = buildAtaHtml(meeting.ataData, accessInfo)

        if (!ataPreviewModal.open) {
            ataPreviewModal.showModal()
        }
    })
}

/**
 * Request runtime permission for webhook URL.
 * @param {string} url
 */
function requestWebhookAndNotificationPermission(url) {
    return new Promise((resolve, reject) => {
        try {
            const urlObj = new URL(url)
            const originPattern = `${urlObj.protocol}//${urlObj.hostname}/*`

            chrome.permissions.request({
                origins: [originPattern],
                permissions: ["notifications"]
            }).then((granted) => {
                if (granted) {
                    resolve("Permission granted")
                } else {
                    reject(new Error("Permission denied"))
                }
            }).catch((error) => {
                reject(error)
            })
        } catch (error) {
            reject(error)
        }
    })
}

/**
 * Loads and displays recent transcripts safely in #meetings-table.
 */
function loadMeetings() {
    const meetingsTable = document.querySelector("#meetings-table")

    chrome.storage.local.get(["meetings"], function (resultLocalUntyped) {
        const resultLocal = /** @type {ResultLocal} */ (resultLocalUntyped || {})
        if (!meetingsTable) return

        meetingsTable.innerHTML = ""

        if (resultLocal.meetings && resultLocal.meetings.length > 0) {
            const meetings = resultLocal.meetings
            for (let i = meetings.length - 1; i >= 0; i--) {
                const meeting = meetings[i]
                const row = document.createElement("tr")
                row.innerHTML = buildMeetingRowHtml(meeting, i)
                meetingsTable.appendChild(row)

                const meetingTitleInput = row.querySelector(".meeting-title")
                if (meetingTitleInput instanceof HTMLDivElement) {
                    meetingTitleInput.addEventListener("blur", function () {
                        const updatedMeeting = /** @type {Meeting} */ ({
                            ...meeting,
                            meetingTitle: meetingTitleInput.innerText
                        })
                        meetings[i] = updatedMeeting
                        chrome.storage.local.set({ meetings: meetings }, function () {
                            console.log("Meeting title updated")
                        })
                    })
                }

                const generateAtaBtn = row.querySelector(".generate-ata-btn")
                if (generateAtaBtn instanceof HTMLButtonElement) {
                    generateAtaBtn.addEventListener("click", function () {
                        const index = parseInt(generateAtaBtn.getAttribute("data-index") ?? "-1", 10)
                        const clientInput = /** @type {HTMLInputElement | null} */ (document.querySelector("#ata-client-name"))
                        const subdirInput = /** @type {HTMLInputElement | null} */ (document.querySelector("#ata-subdiretoria"))
                        const originalLabel = generateAtaBtn.textContent || "Gerar Ata"

                        generateAtaBtn.disabled = true
                        generateAtaBtn.textContent = "Gerando..."

                        /** @type {ExtensionMessage} */
                        const message = {
                            type: "generate_ata_at_index",
                            index,
                            options: {
                                clientName: clientInput ? clientInput.value.trim() : undefined,
                                subdiretoria: subdirInput ? subdirInput.value.trim() : undefined
                            }
                        }
                        chrome.runtime.sendMessage(message, (responseUntyped) => {
                            const response = /** @type {ExtensionResponse} */ (responseUntyped)
                            generateAtaBtn.disabled = false
                            generateAtaBtn.textContent = originalLabel
                            if (response?.success) {
                                loadMeetings()
                            } else {
                                const parsedError = /** @type {ErrorObject} */ (response?.message)
                                alert(parsedError?.errorMessage || "Não foi possível gerar a Ata com Gemini.")
                            }
                        })
                    })
                }

                const downloadAtaDocxBtn = row.querySelector(".download-ata-docx-btn")
                if (downloadAtaDocxBtn instanceof HTMLButtonElement) {
                    downloadAtaDocxBtn.addEventListener("click", function () {
                        const index = parseInt(downloadAtaDocxBtn.getAttribute("data-index") ?? "-1", 10)
                        /** @type {ExtensionMessage} */
                        const message = {
                            type: "download_ata_at_index",
                            index,
                            format: "docx"
                        }
                        chrome.runtime.sendMessage(message, (responseUntyped) => {
                            const response = /** @type {ExtensionResponse} */ (responseUntyped)
                            if (!response?.success) {
                                const parsedError = /** @type {ErrorObject} */ (response?.message)
                                alert(parsedError?.errorMessage || "Não foi possível baixar a Ata (.docx).")
                            }
                        })
                    })
                }

                const previewAtaBtn = row.querySelector(".preview-ata-btn")
                if (previewAtaBtn instanceof HTMLButtonElement) {
                    previewAtaBtn.addEventListener("click", function () {
                        const index = parseInt(previewAtaBtn.getAttribute("data-index") ?? "-1", 10)
                        openAtaPreviewModal(index)
                    })
                }

                const viewTranscriptBtn = row.querySelector(".view-transcript-btn")
                if (viewTranscriptBtn instanceof HTMLButtonElement) {
                    viewTranscriptBtn.addEventListener("click", function () {
                        const index = parseInt(viewTranscriptBtn.getAttribute("data-index") ?? "-1", 10)
                        openTranscriptModal(index)
                    })
                }

                const downloadButton = row.querySelector(".download-button")
                if (downloadButton instanceof HTMLButtonElement) {
                    downloadButton.addEventListener("click", function () {
                        const index = parseInt(downloadButton.getAttribute("data-index") ?? "-1", 10)
                        /** @type {ExtensionMessage} */
                        const message = {
                            type: "download_transcript_at_index",
                            index: index
                        }
                        chrome.runtime.sendMessage(message, (responseUntyped) => {
                            const response = /** @type {ExtensionResponse} */ (responseUntyped)
                            if (!response?.success) {
                                alert("Could not download transcript")
                                const parsedError = /** @type {ErrorObject} */ (response?.message)
                                if (typeof parsedError === "object") {
                                    console.error(parsedError.errorMessage)
                                }
                            }
                        })
                    })
                }

                const webhookPostButton = row.querySelector(".post-button")
                if (webhookPostButton instanceof HTMLButtonElement) {
                    webhookPostButton.addEventListener("click", function () {
                        chrome.storage.sync.get(["webhookUrl"], function (resultSyncUntyped) {
                            const resultSync = /** @type {ResultSync} */ (resultSyncUntyped || {})
                            if (resultSync.webhookUrl) {
                                requestWebhookAndNotificationPermission(resultSync.webhookUrl)
                                    .then(() => {
                                        webhookPostButton.disabled = true
                                        webhookPostButton.textContent = meeting.webhookPostStatus === "new" ? "Posting..." : "Reposting..."

                                        const index = parseInt(webhookPostButton.getAttribute("data-index") ?? "-1", 10)
                                        /** @type {ExtensionMessage} */
                                        const message = {
                                            type: "post_webhook_at_index",
                                            index: index
                                        }
                                        chrome.runtime.sendMessage(message, (responseUntyped) => {
                                            const response = /** @type {ExtensionResponse} */ (responseUntyped)
                                            loadMeetings()
                                            if (response?.success) {
                                                alert("Posted successfully!")
                                            } else {
                                                const parsedError = /** @type {ErrorObject} */ (response?.message)
                                                if (typeof parsedError === "object") {
                                                    console.error(parsedError.errorMessage)
                                                }
                                            }
                                        })
                                    })
                                    .catch((error) => {
                                        alert("Fine! No webhooks for you!")
                                        console.error("Webhook permission error:", error)
                                    })
                            } else {
                                alert("Please provide a webhook URL")
                            }
                        })
                    })
                }

                const deleteButton = row.querySelector(".delete-button")
                if (deleteButton instanceof HTMLButtonElement) {
                    deleteButton.addEventListener("click", function () {
                        if (confirm("Delete this meeting?")) {
                            meetings.splice(i, 1)
                            chrome.storage.local.set({ meetings: meetings }, function () {
                                console.log("Meeting deleted")
                            })
                        }
                    })
                }
            }

            const meetingsTableContainer = document.querySelector("#meetings-table-container")
            if (!isMeetingsTableExpanded && meetingsTableContainer && meetingsTableContainer.clientHeight > 280) {
                meetingsTableContainer.classList.add("fade-mask")
                document.querySelector("#show-all")?.setAttribute("style", "display: block")
            }
        } else {
            meetingsTable.innerHTML = `<tr><td colspan="6">Your next meeting will show up here</td></tr>`
        }
    })
}
