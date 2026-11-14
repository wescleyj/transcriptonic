// @ts-check
/// <reference path="../../types/chrome.d.ts" />
/// <reference path="../../types/index.js" />

window.onload = function () {
  const hideCaptionsToggle = /** @type {HTMLInputElement | null} */ (document.querySelector("#hide-captions"))

  chrome.storage.sync.get(["hideCaptions"], function (resultSyncUntyped) {
    const resultSync = /** @type {ResultSync} */ (resultSyncUntyped || {})
    if (hideCaptionsToggle instanceof HTMLInputElement) {
      hideCaptionsToggle.checked = Boolean(resultSync.hideCaptions)

      hideCaptionsToggle.addEventListener("change", function () {
        chrome.storage.sync.set({ hideCaptions: hideCaptionsToggle.checked }, function () { })
      })
    }
  })
}
