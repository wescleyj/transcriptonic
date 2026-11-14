const PLATFORM_PERMISSION_ORIGINS = {
  google_meet: ["https://meet.google.com/*"],
  teams: [
    "https://teams.live.com/*",
    "https://teams.microsoft.com/*",
    "https://teams.cloud.microsoft/*",
    "https://teams.microsoft.com.mcas.ms/*",
    "https://teams.cloud.microsoft.mcas.ms/*",
    "https://teams.microsoft.com.mcas-gov.us/*",
    "https://teams.cloud.microsoft.mcas-gov.us/*",
    "https://teams.microsoft.com.mcas-gov.ms/*",
    "https://teams.cloud.microsoft.mcas-gov.ms/*"
  ],
  zoom: ["https://*.zoom.us/*"]
}

const DEFAULT_GEMINI_MODEL = "gemini-3.5-flash-lite"
const LEGACY_GEMINI_MODELS = new Set([
  "gemini-2.5-flash-lite",
  "gemini-2.0-flash",
  "gemini-1.5-flash"
])

/**
 * Normalizes and migrates legacy Gemini model identifiers to the active default.
 * @param {string | undefined | null} model
 * @returns {string}
 */
function resolveGeminiModel(model) {
  const clean = String(model || "").trim().replace(/^models\//i, "")
  if (!clean || LEGACY_GEMINI_MODELS.has(clean)) {
    return DEFAULT_GEMINI_MODEL
  }
  return clean
}

window.onload = function () {
  if (chrome.scripting && typeof chrome.scripting.getRegisteredContentScripts === "function") {
    chrome.scripting
      .getRegisteredContentScripts()
      .then((scripts) => {
        console.log(scripts)
      })
      .catch((err) => {
        console.warn("Could not list registered content scripts:", err)
      })
  }

  const autoModeRadio = document.querySelector("#auto-mode")
  const manualModeRadio = document.querySelector("#manual-mode")
  const versionElement = document.querySelector("#version")

  // Platform Checkboxes
  const googleMeetToggle = /** @type {HTMLInputElement} */ (document.querySelector("#enable-google-meet"))
  const teamsToggle = /** @type {HTMLInputElement} */ (document.querySelector("#enable-teams"))
  const zoomToggle = /** @type {HTMLInputElement} */ (document.querySelector("#enable-zoom"))
  const hideCaptionsToggle = /** @type {HTMLInputElement} */ (document.querySelector("#hide-captions"))

  // Gemini Ata controls
  const geminiKeyInput = /** @type {HTMLInputElement | null} */ (document.querySelector("#gemini-api-key"))
  const saveGeminiKeyBtn = /** @type {HTMLButtonElement | null} */ (document.querySelector("#save-gemini-key"))
  const clearGeminiKeyBtn = /** @type {HTMLButtonElement | null} */ (document.querySelector("#clear-gemini-key"))
  const geminiKeyStatus = document.querySelector("#gemini-key-status")
  const geminiModelSelect = /** @type {HTMLSelectElement | null} */ (document.querySelector("#gemini-model"))
  const autoGenerateAtaToggle = /** @type {HTMLInputElement | null} */ (document.querySelector("#auto-generate-ata"))
  const generateLastAtaBtn = /** @type {HTMLButtonElement | null} */ (document.querySelector("#generate-last-ata-btn"))
  const ataActionStatus = document.querySelector("#ata-action-status")

  if (versionElement) {
    versionElement.textContent = `v${chrome.runtime.getManifest().version}`
  }

  chrome.storage.sync.get(["operationMode"], function (resultSyncUntyped) {
    const resultSync = /** @type {ResultSync} */ (resultSyncUntyped)
    if (autoModeRadio instanceof HTMLInputElement && manualModeRadio instanceof HTMLInputElement) {
      if (resultSync.operationMode === "manual") {
        manualModeRadio.checked = true
      }
      else {
        autoModeRadio.checked = true
      }

      autoModeRadio.addEventListener("change", function () {
        chrome.storage.sync.set({ operationMode: "auto" }, function () { })
      })
      manualModeRadio.addEventListener("change", function () {
        chrome.storage.sync.set({ operationMode: "manual" }, function () { })
      })
    }
  })

  chrome.storage.sync.get(["hideCaptions", "geminiModel", "autoGenerateAtaAfterMeeting"], function (resultSyncUntyped) {
    const resultSync = /** @type {ResultSync} */ (resultSyncUntyped)
    if (hideCaptionsToggle instanceof HTMLInputElement) {
      hideCaptionsToggle.checked = Boolean(resultSync.hideCaptions)

      hideCaptionsToggle.addEventListener("change", function () {
        chrome.storage.sync.set({ hideCaptions: hideCaptionsToggle.checked }, function () { })
      })
    }

    if (geminiModelSelect) {
      const resolvedModel = resolveGeminiModel(resultSync.geminiModel)
      geminiModelSelect.value = resolvedModel
      if (resultSync.geminiModel !== resolvedModel) {
        chrome.storage.sync.set({ geminiModel: resolvedModel }, function () { })
      }
      geminiModelSelect.addEventListener("change", function () {
        chrome.storage.sync.set({ geminiModel: geminiModelSelect.value || DEFAULT_GEMINI_MODEL }, function () { })
      })
    }

    if (autoGenerateAtaToggle) {
      autoGenerateAtaToggle.checked = resultSync.autoGenerateAtaAfterMeeting === true
      autoGenerateAtaToggle.addEventListener("change", function () {
        chrome.storage.sync.set({ autoGenerateAtaAfterMeeting: autoGenerateAtaToggle.checked }, function () { })
      })
    }
  })

  function refreshGeminiKeyStatus(/** @type {boolean} */ hasKey) {
    if (geminiKeyStatus) {
      geminiKeyStatus.textContent = hasKey ? "Chave configurada" : "Chave não configurada"
      geminiKeyStatus.classList.toggle("status-ok", hasKey)
      geminiKeyStatus.classList.toggle("status-err", !hasKey)
    }
    if (clearGeminiKeyBtn) {
      clearGeminiKeyBtn.style.display = hasKey ? "inline-block" : "none"
    }
    if (geminiKeyInput) {
      geminiKeyInput.placeholder = hasKey
        ? "Chave salva (cole outra para substituir)"
        : "Cole sua chave API do Google Gemini (AIza...)"
    }
  }

  chrome.storage.local.get(["geminiApiKey"], function (resultLocalUntyped) {
    const resultLocal = /** @type {ResultLocal} */ (resultLocalUntyped)
    refreshGeminiKeyStatus(Boolean(resultLocal?.geminiApiKey))
  })

  if (saveGeminiKeyBtn && geminiKeyInput) {
    saveGeminiKeyBtn.addEventListener("click", () => {
      const apiKey = geminiKeyInput.value.trim()
      if (!apiKey) {
        if (ataActionStatus) {
          ataActionStatus.textContent = "Informe uma chave API válida do Google Gemini."
          ataActionStatus.className = "sub-text status-err"
        }
        return
      }
      chrome.storage.local.set({ geminiApiKey: apiKey }, () => {
        geminiKeyInput.value = ""
        refreshGeminiKeyStatus(true)
        if (ataActionStatus) {
          ataActionStatus.textContent = "Chave API do Gemini salva com segurança no armazenamento local."
          ataActionStatus.className = "sub-text status-ok"
        }
      })
    })
  }

  if (clearGeminiKeyBtn) {
    clearGeminiKeyBtn.addEventListener("click", () => {
      chrome.storage.local.remove(["geminiApiKey"], () => {
        refreshGeminiKeyStatus(false)
        if (ataActionStatus) {
          ataActionStatus.textContent = "Chave API removida."
          ataActionStatus.className = "sub-text"
        }
      })
    })
  }

  if (generateLastAtaBtn) {
    generateLastAtaBtn.addEventListener("click", () => {
      chrome.storage.local.get(["meetings", "geminiApiKey"], (resultLocalUntyped) => {
        const resultLocal = /** @type {ResultLocal} */ (resultLocalUntyped)
        if (!resultLocal?.geminiApiKey) {
          if (ataActionStatus) {
            ataActionStatus.textContent = "Configure sua chave API do Google Gemini primeiro."
            ataActionStatus.className = "sub-text status-err"
          }
          return
        }
        const meetings = resultLocal.meetings || []
        if (meetings.length === 0) {
          if (ataActionStatus) {
            ataActionStatus.textContent = "Nenhuma reunião gravada ainda para gerar Ata."
            ataActionStatus.className = "sub-text status-err"
          }
          return
        }

        const lastIndex = meetings.length - 1
        generateLastAtaBtn.disabled = true
        generateLastAtaBtn.textContent = "Gerando Ata..."
        if (ataActionStatus) {
          ataActionStatus.textContent = "Enviando transcrição para o Gemini..."
          ataActionStatus.className = "sub-text"
        }

        /** @type {ExtensionMessage} */
        const message = {
          type: "generate_ata_at_index",
          index: lastIndex
        }
        chrome.runtime.sendMessage(message, (responseUntyped) => {
          const response = /** @type {ExtensionResponse} */ (responseUntyped)
          generateLastAtaBtn.disabled = false
          generateLastAtaBtn.textContent = "Gerar Ata da última reunião"
          if (response && response.success) {
            if (ataActionStatus) {
              ataActionStatus.textContent = "Ata gerada e baixada (.docx) com sucesso!"
              ataActionStatus.className = "sub-text status-ok"
            }
          } else {
            const errMsg = response?.message?.errorMessage || response?.message || "Erro ao gerar Ata."
            if (ataActionStatus) {
              ataActionStatus.textContent = String(errMsg)
              ataActionStatus.className = "sub-text status-err"
            }
          }
        })
      })
    })
  }

  /**
   * Syncs checkbox UI with actual background script registration status.
   * Requests optional browser permissions directly inside the user-gesture event handler for Firefox compatibility.
   * @param {HTMLInputElement} element
   * @param {Platform} platform
   */
  function syncPlatformStatus(element, platform) {
    /** @type {ExtensionMessage} */
    const message = {
      type: "get_platform_enablement_status",
      platform: platform
    }
    chrome.runtime.sendMessage(message, (responseUntyped) => {
      const response = /** @type {ExtensionResponse} */ (responseUntyped)
      if (response && response.success) {
        element.checked = response.message === "Enabled"
      }
    })

    element.addEventListener("change", () => {
      const isEnabling = element.checked

      /** @param {boolean} enabled */
      const persistPlatformChoice = (enabled) => {
        switch (platform) {
          case "google_meet":
            chrome.storage.sync.set({ wantGoogleMeet: enabled }, function () { })
            break
          case "teams":
            chrome.storage.sync.set({ wantTeams: enabled }, function () { })
            break
          case "zoom":
            chrome.storage.sync.set({ wantZoom: enabled }, function () { })
            break
          default:
            break
        }
      }

      const sendToggleMessage = (/** @type {"enable_platform" | "disable_platform"} */ type) => {
        /** @type {ExtensionMessage} */
        const toggleMsg = { type, platform }
        chrome.runtime.sendMessage(toggleMsg, (responseUntyped) => {
          const response = /** @type {ExtensionResponse} */ (responseUntyped)
          if (!response || !response.success) {
            element.checked = !isEnabling
            persistPlatformChoice(!isEnabling)
            console.error(`Failed to toggle ${platform}:`, response?.message)
          }
        })
      }

      if (isEnabling && chrome.permissions && typeof chrome.permissions.request === "function") {
        // Call permissions.request directly in user gesture for Firefox MV3 compatibility
        const origins = PLATFORM_PERMISSION_ORIGINS[platform] || []
        chrome.permissions.request({
          origins,
          permissions: ["notifications"]
        })
          .then((granted) => {
            if (!granted) {
              element.checked = false
              persistPlatformChoice(false)
              return
            }
            persistPlatformChoice(true)
            sendToggleMessage("enable_platform")
          })
          .catch((err) => {
            console.error(`Permission request failed for ${platform}:`, err)
            element.checked = false
            persistPlatformChoice(false)
          })
        return
      }

      persistPlatformChoice(isEnabling)
      sendToggleMessage(isEnabling ? "enable_platform" : "disable_platform")
    })
  }

  // Initialize Toggles
  if (googleMeetToggle) {
    syncPlatformStatus(googleMeetToggle, "google_meet")
  }
  if (teamsToggle) {
    syncPlatformStatus(teamsToggle, "teams")
  }
  if (zoomToggle) {
    syncPlatformStatus(zoomToggle, "zoom")
  }
}
