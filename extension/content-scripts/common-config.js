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