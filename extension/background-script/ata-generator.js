import {
    getTranscriptString,
    getChatMessagesString,
    getSanitisedMeetingTitle,
    getFormattedFileTimestamp,
    getTranscriptFileName
} from './utils.js'
import { downloadBlobOrDataUrl } from './exporters.js'

export const DEFAULT_GEMINI_MODEL = "gemini-2.5-flash-lite"
export const GEMINI_MODELS = [
    "gemini-2.5-flash-lite",
    "gemini-2.5-flash",
    "gemini-2.0-flash"
]

const DEFAULT_TEMPLATE_SPEC = `Estrutura padrão da ATA DE REUNIÃO (baseada no MODELO ATA.docx):
- Banner superior (Tabela 1): "ATA DE REUNIÃO" (fonte Antonio, azul #004AAD, negrito, borda inferior #047176)
- Subtítulo em negrito (fonte Arimo #000000): "Cliente - Nome da Empresa" ou título da reunião
- Tabela de metadados com 5 colunas (bordas turquesa #00D0DD, cabeçalho #004AAD texto branco Canva Sans Bold, corpo #D4F8F8 Canva Sans):
  LOCAL | DATA E HORÁRIO | PARTICIPANTES | AUSENTES | SUBDIRETORIA
- Seção "I. Pauta da Reunião" (azul #004AAD, negrito) com lista dos principais itens de pauta
- "1. OBJETIVO DA REUNIÃO": parágrafo objetivo resumindo o propósito da reunião
- "2. PAUTA E PONTOS DISCUTIDOS": subseções numeradas (2.1., 2.2., ... em #047176) com tópicos em marcador sólido (● Título: descrição) e subtópicos (○ detalhe) + tabela-resumo opcional ("Aspecto" #004AAD | "Detalhes da Execução" #D4F8F8) quando houver escopo/cronograma/investimento estruturado
- "3. ENCAMINHAMENTOS E PRÓXIMOS PASSOS": lista numerada por responsável (1. Nome:) com ações em sub-marcador (○ ação)
- "4. ACESSO À TRANSCRIÇÃO COMPLETA": referência e link direto para a transcrição integral
- Assinatura final: "Ass" e "Diretor:" (azul #004AAD, negrito)`

/** @type {Uint32Array | null} */
let crc32Table = null

/**
 * Computes CRC-32 for ZIP archive generation.
 * @param {Uint8Array} bytes
 * @returns {number}
 */
function crc32(bytes) {
    if (!crc32Table) {
        crc32Table = new Uint32Array(256)
        for (let i = 0; i < 256; i++) {
            let c = i
            for (let j = 0; j < 8; j++) {
                c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1)
            }
            crc32Table[i] = c >>> 0
        }
    }
    let crc = 0xFFFFFFFF
    for (let i = 0; i < bytes.length; i++) {
        crc = crc32Table[(crc ^ bytes[i]) & 0xFF] ^ (crc >>> 8)
    }
    return (crc ^ 0xFFFFFFFF) >>> 0
}

/**
 * Converts a Date into DOS time/date fields for ZIP headers.
 * @param {Date} date
 * @returns {{ dosTime: number, dosDate: number }}
 */
function toDosDateTime(date) {
    const year = Math.max(1980, date.getFullYear())
    const dosTime = ((date.getHours() & 0x1F) << 11) |
        ((date.getMinutes() & 0x3F) << 5) |
        ((date.getSeconds() >> 1) & 0x1F)
    const dosDate = (((year - 1980) & 0x7F) << 9) |
        (((date.getMonth() + 1) & 0x0F) << 5) |
        (date.getDate() & 0x1F)
    return { dosTime, dosDate }
}

/**
 * Creates an uncompressed (STORE method 0) ZIP byte array from file entries.
 * Zero-dependency and compatible with Service Workers, Firefox, Word, LibreOffice, and Node.
 * @param {Array<{ name: string, data: Uint8Array }>} entries
 * @returns {Uint8Array}
 */
export function createZipBytes(entries) {
    const encoder = new TextEncoder()
    const now = toDosDateTime(new Date())
    /** @type {Uint8Array[]} */
    const localParts = []
    /** @type {Uint8Array[]} */
    const centralParts = []
    let offset = 0

    for (const entry of entries) {
        const nameBytes = encoder.encode(entry.name)
        const data = entry.data
        const crc = crc32(data)

        const localHeader = new Uint8Array(30 + nameBytes.length)
        const lv = new DataView(localHeader.buffer)
        lv.setUint32(0, 0x04034b50, true)
        lv.setUint16(4, 20, true) // version needed to extract (2.0)
        lv.setUint16(6, 0x0800, true) // UTF-8 filename flag
        lv.setUint16(8, 0, true) // compression method: 0 (store)
        lv.setUint16(10, now.dosTime, true)
        lv.setUint16(12, now.dosDate, true)
        lv.setUint32(14, crc, true)
        lv.setUint32(18, data.length, true)
        lv.setUint32(22, data.length, true)
        lv.setUint16(26, nameBytes.length, true)
        lv.setUint16(28, 0, true) // extra length
        localHeader.set(nameBytes, 30)

        localParts.push(localHeader, data)

        const centralHeader = new Uint8Array(46 + nameBytes.length)
        const cv = new DataView(centralHeader.buffer)
        cv.setUint32(0, 0x02014b50, true)
        cv.setUint16(4, 20, true) // version made by
        cv.setUint16(6, 20, true) // version needed
        cv.setUint16(8, 0x0800, true) // UTF-8 flag
        cv.setUint16(10, 0, true) // store
        cv.setUint16(12, now.dosTime, true)
        cv.setUint16(14, now.dosDate, true)
        cv.setUint32(16, crc, true)
        cv.setUint32(20, data.length, true)
        cv.setUint32(24, data.length, true)
        cv.setUint16(28, nameBytes.length, true)
        cv.setUint16(30, 0, true) // extra
        cv.setUint16(32, 0, true) // comment
        cv.setUint16(34, 0, true) // disk number
        cv.setUint16(36, 0, true) // internal attrs
        cv.setUint32(38, 0, true) // external attrs
        cv.setUint32(42, offset, true)
        centralHeader.set(nameBytes, 46)

        centralParts.push(centralHeader)
        offset += localHeader.length + data.length
    }

    const centralDirSize = centralParts.reduce((sum, part) => sum + part.length, 0)
    const eocd = new Uint8Array(22)
    const ev = new DataView(eocd.buffer)
    ev.setUint32(0, 0x06054b50, true)
    ev.setUint16(4, 0, true)
    ev.setUint16(6, 0, true)
    ev.setUint16(8, entries.length, true)
    ev.setUint16(10, entries.length, true)
    ev.setUint32(12, centralDirSize, true)
    ev.setUint32(16, offset, true)
    ev.setUint16(20, 0, true)

    const totalSize = offset + centralDirSize + eocd.length
    const result = new Uint8Array(totalSize)
    let cursor = 0
    for (const part of localParts) {
        result.set(part, cursor)
        cursor += part.length
    }
    for (const part of centralParts) {
        result.set(part, cursor)
        cursor += part.length
    }
    result.set(eocd, cursor)
    return result
}

/**
 * Decompresses raw DEFLATE bytes using Web DecompressionStream.
 * @param {Uint8Array} compressed
 * @returns {Promise<Uint8Array>}
 */
async function inflateRawBytes(compressed) {
    if (typeof DecompressionStream === "undefined") {
        throw new Error("DecompressionStream unavailable for reading compressed DOCX")
    }
    const ds = new DecompressionStream("deflate-raw")
    const writer = ds.writable.getWriter()
    const copy = new Uint8Array(compressed.byteLength)
    copy.set(compressed)
    writer.write(copy)
    writer.close()
    const ab = await new Response(ds.readable).arrayBuffer()
    return new Uint8Array(ab)
}

/**
 * Reads entries from a ZIP archive (supports STORE method 0 and DEFLATE method 8).
 * @param {Uint8Array} zipBytes
 * @returns {Promise<Map<string, Uint8Array>>}
 */
export async function readZipEntries(zipBytes) {
    const view = new DataView(zipBytes.buffer, zipBytes.byteOffset, zipBytes.byteLength)
    let eocdOffset = -1
    const minOffset = Math.max(0, zipBytes.length - 65557)

    for (let i = zipBytes.length - 22; i >= minOffset; i--) {
        if (view.getUint32(i, true) === 0x06054b50) {
            eocdOffset = i
            break
        }
    }
    if (eocdOffset === -1) {
        throw new Error("Invalid ZIP/DOCX archive: EOCD header not found")
    }

    const totalEntries = view.getUint16(eocdOffset + 10, true)
    const centralDirOffset = view.getUint32(eocdOffset + 16, true)
    const decoder = new TextDecoder("utf-8")
    /** @type {Map<string, Uint8Array>} */
    const files = new Map()

    let ptr = centralDirOffset
    for (let i = 0; i < totalEntries; i++) {
        if (ptr + 46 > zipBytes.length || view.getUint32(ptr, true) !== 0x02014b50) {
            throw new Error("Invalid ZIP/DOCX central directory entry")
        }
        const method = view.getUint16(ptr + 10, true)
        const compressedSize = view.getUint32(ptr + 20, true)
        const nameLen = view.getUint16(ptr + 28, true)
        const extraLen = view.getUint16(ptr + 30, true)
        const commentLen = view.getUint16(ptr + 32, true)
        const localHeaderOffset = view.getUint32(ptr + 42, true)

        const fileName = decoder.decode(zipBytes.subarray(ptr + 46, ptr + 46 + nameLen))
        ptr += 46 + nameLen + extraLen + commentLen

        if (localHeaderOffset + 30 > zipBytes.length || view.getUint32(localHeaderOffset, true) !== 0x04034b50) {
            continue
        }
        const localNameLen = view.getUint16(localHeaderOffset + 26, true)
        const localExtraLen = view.getUint16(localHeaderOffset + 28, true)
        const dataStart = localHeaderOffset + 30 + localNameLen + localExtraLen
        const compressedSlice = zipBytes.subarray(dataStart, dataStart + compressedSize)

        if (method === 0) {
            files.set(fileName, new Uint8Array(compressedSlice))
        } else if (method === 8) {
            files.set(fileName, await inflateRawBytes(compressedSlice))
        }
    }

    return files
}

/**
 * Decodes basic XML entities.
 * @param {string} xml
 * @returns {string}
 */
function decodeXmlEntities(xml) {
    return xml
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&#(\d+);/g, (_m, dec) => String.fromCharCode(Number(dec)))
        .replace(/&#x([0-9a-fA-F]+);/g, (_m, hex) => String.fromCharCode(parseInt(hex, 16)))
        .replace(/&amp;/g, "&")
}

/**
 * Extracts readable structured text from a DOCX file's bytes.
 * @param {Uint8Array} docxBytes
 * @returns {Promise<string>}
 */
export async function extractTextFromDocxBytes(docxBytes) {
    const entries = await readZipEntries(docxBytes)
    const documentXmlBytes = entries.get("word/document.xml")
    if (!documentXmlBytes) {
        throw new Error("Invalid DOCX template: word/document.xml not found")
    }
    const xml = new TextDecoder("utf-8").decode(documentXmlBytes)
    const withBreaks = xml
        .replace(/<w:tab\b[^>]*\/>/g, "\t")
        .replace(/<w:br\b[^>]*\/>/g, "\n")
        .replace(/<\/w:tc>/g, " | ")
        .replace(/<\/w:tr>/g, "\n")
        .replace(/<\/w:p>/g, "\n")
        .replace(/<[^>]+>/g, "")

    return decodeXmlEntities(withBreaks)
        .replace(/[ \t]+\n/g, "\n")
        .replace(/\n{3,}/g, "\n\n")
        .trim()
}

/**
 * Escapes unsafe XML characters and strips invalid XML 1.0 control characters.
 * @param {any} value
 * @returns {string}
 */
export function escapeXml(value) {
    return String(value ?? "")
        .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&apos;")
}

/**
 * Escapes HTML special characters.
 * @param {any} value
 * @returns {string}
 */
export function escapeHtml(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;")
}

/**
 * Formats an ISO timestamp into DD/MM/YYYY; HH:MM.
 * @param {string} isoTimestamp
 * @returns {string}
 */
export function formatDateTimeBr(isoTimestamp) {
    const date = new Date(isoTimestamp || Date.now())
    if (Number.isNaN(date.getTime())) {
        return ""
    }
    const day = String(date.getDate()).padStart(2, "0")
    const month = String(date.getMonth() + 1).padStart(2, "0")
    const year = date.getFullYear()
    const hours = String(date.getHours()).padStart(2, "0")
    const minutes = String(date.getMinutes()).padStart(2, "0")
    return `${day}/${month}/${year}; ${hours}:${minutes}`
}

/**
 * Extracts unique participant names observed in transcript and chat messages.
 * @param {Meeting} meeting
 * @returns {string[]}
 */
export function extractUniqueParticipants(meeting) {
    const names = new Set()
    for (const block of meeting?.transcript || []) {
        if (block.personName && block.personName.trim()) {
            names.add(block.personName.trim())
        }
    }
    for (const msg of meeting?.chatMessages || []) {
        if (msg.personName && msg.personName.trim()) {
            names.add(msg.personName.trim())
        }
    }
    return [...names]
}

/**
 * Encodes raw bytes to Base64 without blowing the call stack on embedded fonts.
 * @param {Uint8Array} bytes
 * @returns {string}
 */
function bytesToBase64(bytes) {
    let binary = ""
    const chunkSize = 0x8000
    for (let i = 0; i < bytes.length; i += chunkSize) {
        binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize))
    }
    return btoa(binary)
}

/**
 * Decodes a Base64 string into Uint8Array bytes.
 * @param {string} base64Data
 * @returns {Uint8Array}
 */
function base64ToBytes(base64Data) {
    const binary = atob(base64Data)
    const bytes = new Uint8Array(binary.length)
    for (let i = 0; i < binary.length; i++) {
        bytes[i] = binary.charCodeAt(i)
    }
    return bytes
}

/**
 * Fetches a bundled file from the extension package if available.
 * @param {string} assetPath
 * @param {typeof fetch | undefined} fetchFn
 * @returns {Promise<Uint8Array | null>}
 */
async function fetchBundledAssetBytes(assetPath, fetchFn) {
    if (!fetchFn || typeof chrome === "undefined" || !chrome.runtime?.getURL) {
        return null
    }
    try {
        const assetUrl = chrome.runtime.getURL(assetPath)
        const response = await fetchFn(assetUrl)
        if (!response.ok) {
            return null
        }
        return new Uint8Array(await response.arrayBuffer())
    } catch (_err) {
        return null
    }
}

/**
 * Loads the active Ata template (custom uploaded template from storage or bundled MODELO ATA.docx / modelo-ata.pdf).
 * @param {{ fetchImpl?: typeof fetch }} [deps]
 * @returns {Promise<{ source: string, fileName: string, mimeType: string, base64Data?: string, extractedText?: string, docxBytes?: Uint8Array | null }>}
 */
export async function loadAtaTemplate(deps = {}) {
    const fetchFn = deps.fetchImpl || (typeof fetch === "function" ? fetch : undefined)

    const localData = await new Promise((resolve) => {
        if (typeof chrome !== "undefined" && chrome.storage?.local) {
            chrome.storage.local.get(["customAtaTemplate"], (res) => resolve(res || {}))
        } else {
            resolve({})
        }
    })

    const custom = /** @type {CustomAtaTemplate | undefined} */ (localData.customAtaTemplate)
    if (custom && (custom.base64Data || custom.extractedText)) {
        let extractedText = custom.extractedText
        /** @type {Uint8Array | null} */
        let docxBytes = null
        const isDocx = /\.docx$/i.test(custom.fileName || "") ||
            (custom.mimeType || "").includes("wordprocessingml")
        if (isDocx && custom.base64Data) {
            try {
                docxBytes = base64ToBytes(custom.base64Data)
                if (!extractedText) {
                    extractedText = await extractTextFromDocxBytes(docxBytes)
                }
            } catch (err) {
                console.warn("Could not extract text from custom DOCX template:", err)
            }
        } else {
            // Keep bundled MODELO ATA.docx styling when user only uploads a PDF reference
            docxBytes = await fetchBundledAssetBytes("templates/modelo-ata.docx", fetchFn)
        }
        return {
            source: "custom",
            fileName: custom.fileName || "custom-template",
            mimeType: custom.mimeType || (isDocx
                ? "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                : "application/pdf"),
            base64Data: custom.base64Data,
            extractedText,
            docxBytes
        }
    }

    const defaultDocxBytes = await fetchBundledAssetBytes("templates/modelo-ata.docx", fetchFn)
    if (defaultDocxBytes) {
        let extractedText = DEFAULT_TEMPLATE_SPEC
        try {
            extractedText = await extractTextFromDocxBytes(defaultDocxBytes)
        } catch (err) {
            console.warn("Could not extract text from bundled modelo-ata.docx:", err)
        }
        return {
            source: "default",
            fileName: "MODELO ATA.docx",
            mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            extractedText,
            docxBytes: defaultDocxBytes
        }
    }

    const defaultPdfBytes = await fetchBundledAssetBytes("templates/modelo-ata.pdf", fetchFn)
    if (defaultPdfBytes) {
        return {
            source: "default",
            fileName: "modelo-ata.pdf",
            mimeType: "application/pdf",
            base64Data: bytesToBase64(defaultPdfBytes),
            docxBytes: null
        }
    }

    return {
        source: "builtin-spec",
        fileName: "MODELO ATA.docx",
        mimeType: "text/plain",
        extractedText: DEFAULT_TEMPLATE_SPEC,
        docxBytes: null
    }
}

/**
 * Extracts JSON object from raw Gemini text response.
 * @param {string} rawText
 * @returns {any}
 */
function parseGeminiJsonResponse(rawText) {
    const trimmed = String(rawText || "").trim()
    const fenceMatch = trimmed.match(/```(?:json)?\s*([\s\S]*?)\s*```/i)
    const jsonText = fenceMatch ? fenceMatch[1].trim() : trimmed
    return JSON.parse(jsonText)
}

/**
 * Normalizes Gemini output against meeting metadata so all required Ata sections exist.
 * @param {any} raw
 * @param {Meeting} meeting
 * @param {{ clientName?: string, subdiretoria?: string }} [options]
 * @returns {AtaData}
 */
export function normalizeAtaData(raw, meeting, options = {}) {
    const detectedParticipants = extractUniqueParticipants(meeting)
    const fallbackLocal = meeting?.meetingSoftware === "Google Meet"
        ? "Meet"
        : (meeting?.meetingSoftware || "Meet")
    const fallbackTitle = meeting?.meetingTitle || meeting?.title || "Nome da Empresa"

    const participantes = Array.isArray(raw?.participantes) && raw.participantes.length > 0
        ? raw.participantes.map((/** @type {any} */ p) => String(p).trim()).filter(Boolean)
        : (detectedParticipants.length > 0 ? detectedParticipants : ["Participantes não identificados"])

    const pauta = Array.isArray(raw?.pauta) && raw.pauta.length > 0
        ? raw.pauta.map((/** @type {any} */ sec, /** @type {number} */ idx) => ({
            numero: String(sec?.numero || `2.${idx + 1}.`).trim(),
            titulo: String(sec?.titulo || `Tópico ${idx + 1}`).trim(),
            topicos: Array.isArray(sec?.topicos)
                ? sec.topicos.map((/** @type {any} */ t) => ({
                    titulo: String(t?.titulo || "").trim(),
                    descricao: String(t?.descricao || "").trim(),
                    subtopicos: Array.isArray(t?.subtopicos)
                        ? t.subtopicos.map((/** @type {any} */ s) => String(s).trim()).filter(Boolean)
                        : []
                }))
                : []
        }))
        : [{
            numero: "2.1.",
            titulo: "Pontos Principais",
            topicos: [{
                titulo: "Resumo da Discussão",
                descricao: "Consulte a transcrição completa vinculada na Seção 4 para detalhes.",
                subtopicos: []
            }]
        }]

    const encaminhamentos = Array.isArray(raw?.encaminhamentos) && raw.encaminhamentos.length > 0
        ? raw.encaminhamentos.map((/** @type {any} */ item) => ({
            responsavel: String(item?.responsavel || "Equipe").trim(),
            acoes: Array.isArray(item?.acoes)
                ? item.acoes.map((/** @type {any} */ a) => String(a).trim()).filter(Boolean)
                : []
        }))
        : []

    const tabelaResumo = Array.isArray(raw?.tabelaResumo)
        ? raw.tabelaResumo
            .map((/** @type {any} */ row) => ({
                aspecto: String(row?.aspecto || "").trim(),
                detalhes: String(row?.detalhes || "").trim()
            }))
            .filter((/** @type {AtaTabelaResumoItem} */ row) => row.aspecto || row.detalhes)
        : []

    const secoesExtras = Array.isArray(raw?.secoesExtras)
        ? raw.secoesExtras
            .map((/** @type {any} */ sec) => ({
                titulo: String(sec?.titulo || "").trim(),
                conteudo: Array.isArray(sec?.conteudo)
                    ? sec.conteudo.map((/** @type {any} */ c) => String(c).trim()).filter(Boolean)
                    : []
            }))
            .filter((/** @type {{ titulo: string, conteudo: string[] }} */ sec) => sec.titulo && sec.conteudo.length > 0)
        : []

    return {
        cliente: String(raw?.cliente || options.clientName || fallbackTitle).trim(),
        local: String(raw?.local || fallbackLocal).trim(),
        dataHorario: String(raw?.dataHorario || formatDateTimeBr(meeting?.meetingStartTimestamp)).trim(),
        participantes,
        ausentes: String(raw?.ausentes || "-----").trim() || "-----",
        subdiretoria: String(raw?.subdiretoria ?? options.subdiretoria ?? "").trim(),
        objetivo: String(raw?.objetivo || "Registro dos pontos discutidos e encaminhamentos da reunião.").trim(),
        pauta,
        encaminhamentos,
        tabelaResumo,
        secoesExtras
    }
}

/**
 * Calls the Google Gemini API to synthesize a meeting transcript into an Ata de Reunião matching the template.
 * Sends the API key strictly via the x-goog-api-key HTTP header.
 * @param {{
 *   apiKey: string,
 *   model?: string,
 *   meeting: Meeting,
 *   meetingIndex?: number,
 *   template?: { source: string, fileName: string, mimeType: string, base64Data?: string, extractedText?: string },
 *   clientName?: string,
 *   subdiretoria?: string,
 *   fetchImpl?: typeof fetch
 * }} params
 * @returns {Promise<AtaData>}
 */
export async function callGeminiForAta({
    apiKey,
    model = DEFAULT_GEMINI_MODEL,
    meeting,
    template,
    clientName = "",
    subdiretoria = "",
    fetchImpl
}) {
    const trimmedKey = String(apiKey || "").trim()
    if (!trimmedKey) {
        throw new Error("Chave API do Google Gemini não configurada.")
    }

    const fetchFn = fetchImpl || fetch
    const chosenModel = (model || DEFAULT_GEMINI_MODEL).trim()
    const transcriptText = getTranscriptString(meeting?.transcript || [])
    const chatText = getChatMessagesString(meeting?.chatMessages || [])

    if (!transcriptText.trim() && !chatText.trim()) {
        throw new Error("A reunião selecionada não possui transcrição nem mensagens de chat.")
    }

    const detectedParticipants = extractUniqueParticipants(meeting)
    const defaultDateTime = formatDateTimeBr(meeting?.meetingStartTimestamp)
    const defaultLocal = meeting?.meetingSoftware === "Google Meet"
        ? "Meet"
        : (meeting?.meetingSoftware || "Meet")

    const promptText = `Você é um(a) secretário(a) executivo(a) especialista em redigir Atas de Reunião em Português do Brasil.
Analise a transcrição e mensagens de chat da reunião abaixo e gere uma Ata de Reunião completa, clara, profissional e fiel à estrutura e estilo do modelo fornecido (${template?.fileName || "modelo-ata.pdf"}).

DADOS DA REUNIÃO:
- Título capturado: ${meeting?.meetingTitle || meeting?.title || "Reunião"}
- Plataforma / Local sugerido: ${defaultLocal}
- Data e horário sugerido: ${defaultDateTime}
- Cliente padrão informado: ${clientName || "(extrair do título ou contexto, senão usar o título da reunião)"}
- Subdiretoria padrão informada: ${subdiretoria || "(extrair se mencionado na reunião, senão deixar vazio)"}
- Participantes detectados: ${detectedParticipants.join(", ") || "Não identificados"}

${template?.extractedText ? `MODELO DE REFERÊNCIA EM TEXTO:\n${template.extractedText}\n` : ""}

TRANSCRIÇÃO DA REUNIÃO:
${transcriptText || "(Sem falas registradas)"}

MENSAGENS DE CHAT:
${chatText || "(Sem mensagens no chat)"}

REGRAS DE PREENCHIMENTO:
1. Retorne EXCLUSIVAMENTE um JSON válido seguindo exatamente este schema:
{
  "cliente": "Nome do Cliente ou Projeto",
  "local": "Meet | Teams | Zoom | outro",
  "dataHorario": "DD/MM/YYYY; HH:MM",
  "participantes": ["Nome 1", "Nome 2"],
  "ausentes": "----- ou nomes mencionados como ausentes",
  "subdiretoria": "Nome da subdiretoria ou vazio",
  "objetivo": "Parágrafo objetivo sintetizando a finalidade da reunião.",
  "pauta": [
    {
      "numero": "2.1.",
      "titulo": "Nome do bloco temático",
      "topicos": [
        {
          "titulo": "Rótulo curto do ponto",
          "descricao": "Resumo fiel do que foi discutido e decidido.",
          "subtopicos": ["Detalhe específico ou decisão complementar"]
        }
      ]
    }
  ],
  "encaminhamentos": [
    {
      "responsavel": "Nome da pessoa ou equipe",
      "acoes": ["Ação concreta combinada com prazo/contexto se houver"]
    }
  ],
  "tabelaResumo": [
    {
      "aspecto": "Escopo / Prazo / Investimento (preencher apenas se houver parâmetros objetivos definidos, senão [])",
      "detalhes": "Descrição consolidada do item"
    }
  ],
  "secoesExtras": []
}
2. Não invente fatos que não estejam na transcrição ou no chat.
3. Agrupe os assuntos logicamente nas subseções 2.1, 2.2, etc., mantendo linguagem executiva concisa e bem organizada.`

    /** @type {any[]} */
    const parts = [{ text: promptText }]
    if (template?.base64Data && template?.mimeType === "application/pdf") {
        parts.push({
            inlineData: {
                mimeType: "application/pdf",
                data: template.base64Data
            }
        })
    }

    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(chosenModel)}:generateContent`
    const response = await fetchFn(endpoint, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "x-goog-api-key": trimmedKey
        },
        body: JSON.stringify({
            contents: [{ role: "user", parts }],
            generationConfig: {
                responseMimeType: "application/json",
                temperature: 0.2
            }
        })
    })

    const data = await response.json().catch(() => null)
    if (!response.ok) {
        const apiError = data?.error?.message || `HTTP ${response.status} ${response.statusText}`
        throw new Error(`Erro na API do Gemini (${chosenModel}): ${apiError}`)
    }

    const rawText = data?.candidates?.[0]?.content?.parts
        ?.map((/** @type {any} */ part) => part?.text || "")
        .join("\n")

    if (!rawText) {
        throw new Error("A API do Gemini retornou uma resposta vazia.")
    }

    const parsed = parseGeminiJsonResponse(rawText)
    return normalizeAtaData(parsed, meeting, { clientName, subdiretoria })
}

/**
 * Checks whether a URL is safe to embed as an OpenXML / HTML hyperlink.
 * @param {string | undefined} url
 * @returns {boolean}
 */
function isSafeLinkUrl(url) {
    if (!url || typeof url !== "string") return false
    return /^(https?:|chrome-extension:|moz-extension:)/i.test(url.trim())
}

/**
 * Injects or refreshes the rIdTranscript external hyperlink relationship inside document.xml.rels.
 * @param {string} relsXml
 * @param {string | undefined} viewUrl
 * @returns {string}
 */
function injectTranscriptHyperlinkRel(relsXml, viewUrl) {
    const cleaned = relsXml.replace(/<Relationship\b[^>]*\bId="rIdTranscript"[^>]*\/>/g, "")
    if (!isSafeLinkUrl(viewUrl)) {
        return cleaned
    }
    const relTag = `<Relationship Id="rIdTranscript" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${escapeXml(viewUrl)}" TargetMode="External"/>`
    if (cleaned.includes("</Relationships>")) {
        return cleaned.replace("</Relationships>", `${relTag}</Relationships>`)
    }
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${relTag}
</Relationships>`
}

/**
 * Builds the OpenXML Word (.docx) bytes matching MODELO ATA.docx layout, colors, and fonts.
 * When templateEntries (from readZipEntries) is provided, embedded fonts, theme, numbering, and styles are preserved.
 * @param {AtaData} ataData
 * @param {{ viewUrl?: string, transcriptFileName?: string, transcriptCount?: number, chatCount?: number }} [transcriptAccessInfo]
 * @param {Map<string, Uint8Array> | null} [templateEntries]
 * @returns {Uint8Array}
 */
export function buildAtaDocxBytes(ataData, transcriptAccessInfo = {}, templateEntries = null) {
    const encoder = new TextEncoder()
    const hasHyperlink = isSafeLinkUrl(transcriptAccessInfo.viewUrl)
    const transcriptFileLabel = transcriptAccessInfo.transcriptFileName || "TranscripTonic/Transcript.txt"
    const transcriptCount = transcriptAccessInfo.transcriptCount ?? 0
    const chatCount = transcriptAccessInfo.chatCount ?? 0

    // Exact column widths from Table2 of MODELO ATA.docx (sum = 10632 dxa)
    const colWidths = [1586, 1731, 2018, 1775, 3522]

    const makeHeaderCell = (/** @type {string} */ label, /** @type {number} */ width, /** @type {boolean} */ isFirst, /** @type {boolean} */ isLast) => `
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="${width}" w:type="dxa"/>
            <w:tcBorders>
              <w:top w:val="single" w:sz="6" w:space="0" w:color="f6ffff"/>
              ${isFirst ? `<w:left w:val="single" w:sz="6" w:space="0" w:color="ffffff"/>` : ""}
              <w:bottom w:val="single" w:sz="6" w:space="0" w:color="047176"/>
              ${isLast ? `<w:right w:val="single" w:sz="6" w:space="0" w:color="ffffff"/>` : ""}
            </w:tcBorders>
            <w:shd w:val="clear" w:color="auto" w:fill="004aad"/>
            <w:tcMar>
              <w:top w:w="120" w:type="dxa"/>
              <w:left w:w="120" w:type="dxa"/>
              <w:bottom w:w="120" w:type="dxa"/>
              <w:right w:w="120" w:type="dxa"/>
            </w:tcMar>
            <w:vAlign w:val="center"/>
          </w:tcPr>
          <w:p>
            <w:pPr><w:jc w:val="center"/><w:spacing w:before="120" w:after="120" w:line="336" w:lineRule="auto"/></w:pPr>
            <w:r>
              <w:rPr><w:rFonts w:ascii="Canva Sans Bold" w:cs="Canva Sans Bold" w:eastAsia="Canva Sans Bold" w:hAnsi="Canva Sans Bold"/><w:b/><w:bCs/><w:color w:val="ffffff"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>
              <w:t>${escapeXml(label)}</w:t>
            </w:r>
          </w:p>
        </w:tc>`

    const makeValueCell = (/** @type {string[]} */ lines, /** @type {number} */ width, /** @type {boolean} */ isFirst, /** @type {boolean} */ isLast) => {
        const paragraphs = (lines.length ? lines : [""]).map(line => `
          <w:p>
            <w:pPr><w:jc w:val="center"/><w:spacing w:before="40" w:after="60" w:line="300" w:lineRule="auto"/></w:pPr>
            <w:r>
              <w:rPr><w:rFonts w:ascii="Canva Sans" w:cs="Canva Sans" w:eastAsia="Canva Sans" w:hAnsi="Canva Sans"/><w:color w:val="000000"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>
              <w:t xml:space="preserve">${escapeXml(line)}</w:t>
            </w:r>
          </w:p>`).join("")

        return `
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="${width}" w:type="dxa"/>
            <w:tcBorders>
              <w:top w:val="single" w:sz="6" w:space="0" w:color="047176"/>
              ${isFirst ? `<w:left w:val="single" w:sz="6" w:space="0" w:color="ffffff"/>` : ""}
              <w:bottom w:val="single" w:sz="6" w:space="0" w:color="f6ffff"/>
              ${isLast ? `<w:right w:val="single" w:sz="6" w:space="0" w:color="ffffff"/>` : ""}
            </w:tcBorders>
            <w:shd w:val="clear" w:color="auto" w:fill="d4f8f8"/>
            <w:tcMar>
              <w:top w:w="120" w:type="dxa"/>
              <w:left w:w="120" w:type="dxa"/>
              <w:bottom w:w="120" w:type="dxa"/>
              <w:right w:w="120" w:type="dxa"/>
            </w:tcMar>
            <w:vAlign w:val="top"/>
          </w:tcPr>
          ${paragraphs}
        </w:tc>`
    }

    const participantsLines = [
        "Membros presentes:",
        ...(ataData.participantes || []).map(p => (String(p).startsWith("-") ? String(p) : `- ${p}`))
    ]

    const pautaXml = (ataData.pauta || []).map((sec) => {
        const numPrefix = sec.numero.endsWith(".") ? sec.numero : `${sec.numero}.`
        const headerXml = `
    <w:p>
      <w:pPr><w:spacing w:before="140" w:after="80" w:line="300" w:lineRule="auto"/></w:pPr>
      <w:r>
        <w:rPr><w:rFonts w:ascii="Arimo" w:cs="Arimo" w:eastAsia="Arimo" w:hAnsi="Arimo"/><w:b/><w:bCs/><w:color w:val="047176"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
        <w:t xml:space="preserve">${escapeXml(numPrefix)} ${escapeXml(sec.titulo)}</w:t>
      </w:r>
    </w:p>`

        const topicosXml = (sec.topicos || []).map((topic) => {
            const label = topic.titulo ? `${topic.titulo}: ` : ""
            const mainBullet = `
    <w:p>
      <w:pPr><w:ind w:left="720" w:hanging="360"/><w:spacing w:before="80" w:after="80" w:line="300" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Arimo" w:cs="Arimo" w:eastAsia="Arimo" w:hAnsi="Arimo"/><w:color w:val="000000"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t xml:space="preserve">● </w:t></w:r>
      ${label ? `<w:r><w:rPr><w:rFonts w:ascii="Arimo" w:cs="Arimo" w:eastAsia="Arimo" w:hAnsi="Arimo"/><w:b/><w:bCs/><w:color w:val="000000"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t xml:space="preserve">${escapeXml(label)}</w:t></w:r>` : ""}
      <w:r><w:rPr><w:rFonts w:ascii="Arimo" w:cs="Arimo" w:eastAsia="Arimo" w:hAnsi="Arimo"/><w:color w:val="000000"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t xml:space="preserve">${escapeXml(topic.descricao)}</w:t></w:r>
    </w:p>`

            const subBullets = (topic.subtopicos || []).map((sub) => `
    <w:p>
      <w:pPr><w:ind w:left="1440" w:hanging="360"/><w:spacing w:before="40" w:after="60" w:line="300" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Arimo" w:cs="Arimo" w:eastAsia="Arimo" w:hAnsi="Arimo"/><w:color w:val="000000"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t xml:space="preserve">○ ${escapeXml(sub)}</w:t></w:r>
    </w:p>`).join("")

            return mainBullet + subBullets
        }).join("")

        return headerXml + topicosXml
    }).join("")

    // Table3 layout from MODELO ATA.docx (widths 1983 + 8471 = 10454 dxa)
    const tabelaResumoXml = (ataData.tabelaResumo && ataData.tabelaResumo.length > 0)
        ? `
    <w:tbl>
      <w:tblPr>
        <w:tblStyle w:val="Table3"/>
        <w:tblW w:w="10454" w:type="dxa"/>
        <w:jc w:val="left"/>
        <w:tblLayout w:type="fixed"/>
      </w:tblPr>
      <w:tblGrid>
        <w:gridCol w:w="1983"/>
        <w:gridCol w:w="8471"/>
      </w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="1983" w:type="dxa"/>
            <w:tcBorders><w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tcBorders>
            <w:shd w:val="clear" w:color="auto" w:fill="004aad"/>
            <w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>
          </w:tcPr>
          <w:p><w:r><w:rPr><w:rFonts w:ascii="Arimo" w:cs="Arimo" w:eastAsia="Arimo" w:hAnsi="Arimo"/><w:b/><w:color w:val="f6ffff"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t>Aspecto</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="8471" w:type="dxa"/>
            <w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tcBorders>
            <w:shd w:val="clear" w:color="auto" w:fill="d4f8f8"/>
            <w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>
          </w:tcPr>
          <w:p><w:r><w:rPr><w:rFonts w:ascii="Arimo" w:cs="Arimo" w:eastAsia="Arimo" w:hAnsi="Arimo"/><w:b/><w:color w:val="000000"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t>Detalhes da Execução</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
      ${ataData.tabelaResumo.map(row => `
      <w:tr>
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="1983" w:type="dxa"/>
            <w:tcBorders><w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tcBorders>
            <w:shd w:val="clear" w:color="auto" w:fill="004aad"/>
            <w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>
          </w:tcPr>
          <w:p><w:r><w:rPr><w:rFonts w:ascii="Arimo" w:cs="Arimo" w:eastAsia="Arimo" w:hAnsi="Arimo"/><w:b/><w:color w:val="ffffff"/><w:sz w:val="19"/><w:szCs w:val="19"/></w:rPr><w:t xml:space="preserve">${escapeXml(row.aspecto)}</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="8471" w:type="dxa"/>
            <w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tcBorders>
            <w:shd w:val="clear" w:color="auto" w:fill="d4f8f8"/>
            <w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>
          </w:tcPr>
          <w:p><w:r><w:rPr><w:rFonts w:ascii="Arimo" w:cs="Arimo" w:eastAsia="Arimo" w:hAnsi="Arimo"/><w:color w:val="000000"/><w:sz w:val="19"/><w:szCs w:val="19"/></w:rPr><w:t xml:space="preserve">${escapeXml(row.detalhes)}</w:t></w:r></w:p>
        </w:tc>
      </w:tr>`).join("")}
    </w:tbl>`
        : ""

    const encaminhamentosXml = (ataData.encaminhamentos && ataData.encaminhamentos.length > 0)
        ? ataData.encaminhamentos.map((enc, idx) => {
            const respXml = `
    <w:p>
      <w:pPr><w:ind w:left="360" w:hanging="220"/><w:spacing w:before="80" w:after="40" w:line="300" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Arimo" w:cs="Arimo" w:eastAsia="Arimo" w:hAnsi="Arimo"/><w:b/><w:bCs/><w:color w:val="000000"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t xml:space="preserve">${idx + 1}. ${escapeXml(enc.responsavel)}:</w:t></w:r>
    </w:p>`
            const acoesXml = (enc.acoes || []).map((acao) => `
    <w:p>
      <w:pPr><w:ind w:left="780" w:hanging="220"/><w:spacing w:before="20" w:after="40" w:line="300" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Arimo" w:cs="Arimo" w:eastAsia="Arimo" w:hAnsi="Arimo"/><w:color w:val="000000"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t xml:space="preserve">○ ${escapeXml(acao)}</w:t></w:r>
    </w:p>`).join("")
            return respXml + acoesXml
        }).join("")
        : `
    <w:p>
      <w:pPr><w:spacing w:before="40" w:after="80" w:line="300" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Arimo" w:cs="Arimo" w:eastAsia="Arimo" w:hAnsi="Arimo"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t>Nenhum encaminhamento específico registrado.</w:t></w:r>
    </w:p>`

    const extrasXml = (ataData.secoesExtras || []).map((sec) => {
        const heading = `
    <w:p>
      <w:pPr><w:spacing w:before="140" w:after="60" w:line="300" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Arimo" w:cs="Arimo" w:eastAsia="Arimo" w:hAnsi="Arimo"/><w:b/><w:color w:val="047176"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t>${escapeXml(sec.titulo)}</w:t></w:r>
    </w:p>`
        const body = (sec.conteudo || []).map((line) => `
    <w:p>
      <w:pPr><w:ind w:left="420" w:hanging="220"/><w:spacing w:before="20" w:after="40" w:line="300" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Arimo" w:cs="Arimo" w:eastAsia="Arimo" w:hAnsi="Arimo"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t xml:space="preserve">○ ${escapeXml(line)}</w:t></w:r>
    </w:p>`).join("")
        return heading + body
    }).join("")

    const transcriptLinkParagraph = hasHyperlink
        ? `
    <w:p>
      <w:pPr><w:spacing w:before="40" w:after="60" w:line="300" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Arimo" w:cs="Arimo" w:eastAsia="Arimo" w:hAnsi="Arimo"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t xml:space="preserve">Link direto na extensão: </w:t></w:r>
      <w:hyperlink r:id="rIdTranscript" w:history="1">
        <w:r>
          <w:rPr><w:rFonts w:ascii="Arimo" w:cs="Arimo" w:eastAsia="Arimo" w:hAnsi="Arimo"/><w:color w:val="004aad"/><w:u w:val="single"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>
          <w:t>${escapeXml(transcriptAccessInfo.viewUrl)}</w:t>
        </w:r>
      </w:hyperlink>
    </w:p>`
        : ""

    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:o="urn:schemas-microsoft-com:office:office">
  <w:body>
    <w:tbl>
      <w:tblPr>
        <w:tblStyle w:val="Table1"/>
        <w:tblW w:w="10659" w:type="dxa"/>
        <w:jc w:val="left"/>
        <w:tblInd w:w="180" w:type="dxa"/>
        <w:tblBorders>
          <w:top w:val="single" w:sz="6" w:space="0" w:color="ffffff"/>
          <w:left w:val="single" w:sz="6" w:space="0" w:color="ffffff"/>
          <w:bottom w:val="single" w:sz="6" w:space="0" w:color="ffffff"/>
          <w:right w:val="single" w:sz="6" w:space="0" w:color="ffffff"/>
          <w:insideH w:val="single" w:sz="6" w:space="0" w:color="ffffff"/>
          <w:insideV w:val="single" w:sz="6" w:space="0" w:color="ffffff"/>
        </w:tblBorders>
        <w:tblLayout w:type="fixed"/>
      </w:tblPr>
      <w:tblGrid>
        <w:gridCol w:w="8060"/>
        <w:gridCol w:w="2599"/>
      </w:tblGrid>
      <w:tr>
        <w:trPr><w:trHeight w:val="2000" w:hRule="atLeast"/></w:trPr>
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="8060" w:type="dxa"/>
            <w:tcBorders><w:bottom w:val="single" w:sz="6" w:space="0" w:color="047176"/></w:tcBorders>
            <w:shd w:val="clear" w:color="auto" w:fill="ffffff"/>
            <w:tcMar><w:top w:w="180" w:type="dxa"/><w:left w:w="180" w:type="dxa"/><w:bottom w:w="180" w:type="dxa"/><w:right w:w="180" w:type="dxa"/></w:tcMar>
            <w:vAlign w:val="center"/>
          </w:tcPr>
          <w:p>
            <w:pPr><w:spacing w:before="120" w:after="120" w:line="276" w:lineRule="auto"/></w:pPr>
            <w:r>
              <w:rPr><w:rFonts w:ascii="Antonio" w:cs="Antonio" w:eastAsia="Antonio" w:hAnsi="Antonio"/><w:b/><w:bCs/><w:color w:val="004aad"/><w:sz w:val="112"/><w:szCs w:val="112"/></w:rPr>
              <w:t>ATA DE REUNIÃO</w:t>
            </w:r>
          </w:p>
        </w:tc>
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="2599" w:type="dxa"/>
            <w:tcBorders><w:bottom w:val="single" w:sz="6" w:space="0" w:color="047176"/></w:tcBorders>
            <w:shd w:val="clear" w:color="auto" w:fill="ffffff"/>
            <w:tcMar><w:top w:w="180" w:type="dxa"/><w:left w:w="180" w:type="dxa"/><w:bottom w:w="180" w:type="dxa"/><w:right w:w="180" w:type="dxa"/></w:tcMar>
          </w:tcPr>
          <w:p/>
        </w:tc>
      </w:tr>
    </w:tbl>

    <w:p>
      <w:pPr><w:spacing w:before="120" w:after="120" w:line="300" w:lineRule="auto"/></w:pPr>
      <w:r>
        <w:rPr><w:rFonts w:ascii="Arimo" w:cs="Arimo" w:eastAsia="Arimo" w:hAnsi="Arimo"/><w:b/><w:bCs/><w:color w:val="000000"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>
        <w:t xml:space="preserve">Cliente - ${escapeXml(ataData.cliente)}</w:t>
      </w:r>
    </w:p>

    <w:tbl>
      <w:tblPr>
        <w:tblStyle w:val="Table2"/>
        <w:tblW w:w="10632" w:type="dxa"/>
        <w:jc w:val="left"/>
        <w:tblInd w:w="120" w:type="dxa"/>
        <w:tblBorders>
          <w:top w:val="single" w:sz="6" w:space="0" w:color="00d0dd"/>
          <w:left w:val="single" w:sz="6" w:space="0" w:color="00d0dd"/>
          <w:bottom w:val="single" w:sz="6" w:space="0" w:color="00d0dd"/>
          <w:right w:val="single" w:sz="6" w:space="0" w:color="00d0dd"/>
          <w:insideH w:val="single" w:sz="6" w:space="0" w:color="00d0dd"/>
          <w:insideV w:val="single" w:sz="6" w:space="0" w:color="00d0dd"/>
        </w:tblBorders>
        <w:tblLayout w:type="fixed"/>
      </w:tblPr>
      <w:tblGrid>
        ${colWidths.map(w => `<w:gridCol w:w="${w}"/>`).join("")}
      </w:tblGrid>
      <w:tr>
        ${makeHeaderCell("LOCAL", colWidths[0], true, false)}
        ${makeHeaderCell("DATA E HORÁRIO", colWidths[1], false, false)}
        ${makeHeaderCell("PARTICIPANTES", colWidths[2], false, false)}
        ${makeHeaderCell("AUSENTES", colWidths[3], false, true)}
        ${makeHeaderCell("SUBDIRETORIA", colWidths[4], false, true)}
      </w:tr>
      <w:tr>
        ${makeValueCell([ataData.local], colWidths[0], true, false)}
        ${makeValueCell([ataData.dataHorario], colWidths[1], false, false)}
        ${makeValueCell(participantsLines, colWidths[2], false, false)}
        ${makeValueCell([ataData.ausentes || "-----"], colWidths[3], false, true)}
        ${makeValueCell([ataData.subdiretoria || ""], colWidths[4], false, true)}
      </w:tr>
    </w:tbl>

    <w:p>
      <w:pPr><w:spacing w:before="120" w:after="120" w:line="300" w:lineRule="auto"/></w:pPr>
      <w:r>
        <w:pict>
          <v:rect style="width:0.0pt;height:1.5pt" o:hr="t" o:hrstd="t" o:hralign="center" fillcolor="#A0A0A0" stroked="f"/>
        </w:pict>
      </w:r>
    </w:p>

    <w:p>
      <w:pPr><w:spacing w:before="120" w:after="120" w:line="300" w:lineRule="auto"/></w:pPr>
      <w:r>
        <w:rPr><w:rFonts w:ascii="Arimo" w:cs="Arimo" w:eastAsia="Arimo" w:hAnsi="Arimo"/><w:b/><w:bCs/><w:color w:val="004aad"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
        <w:t>I. Pauta da Reunião</w:t>
      </w:r>
    </w:p>

    <w:p>
      <w:pPr><w:spacing w:before="100" w:after="80" w:line="300" w:lineRule="auto"/></w:pPr>
      <w:r>
        <w:rPr><w:rFonts w:ascii="Arimo" w:cs="Arimo" w:eastAsia="Arimo" w:hAnsi="Arimo"/><w:b/><w:bCs/><w:color w:val="047176"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
        <w:t>1. OBJETIVO DA REUNIÃO</w:t>
      </w:r>
    </w:p>

    <w:p>
      <w:pPr><w:spacing w:before="20" w:after="140" w:line="300" w:lineRule="auto"/></w:pPr>
      <w:r>
        <w:rPr><w:rFonts w:ascii="Arimo" w:cs="Arimo" w:eastAsia="Arimo" w:hAnsi="Arimo"/><w:color w:val="000000"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>
        <w:t xml:space="preserve">${escapeXml(ataData.objetivo)}</w:t>
      </w:r>
    </w:p>

    <w:p>
      <w:pPr><w:spacing w:before="100" w:after="80" w:line="300" w:lineRule="auto"/></w:pPr>
      <w:r>
        <w:rPr><w:rFonts w:ascii="Arimo" w:cs="Arimo" w:eastAsia="Arimo" w:hAnsi="Arimo"/><w:b/><w:bCs/><w:color w:val="047176"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
        <w:t>2. PAUTA E PONTOS DISCUTIDOS</w:t>
      </w:r>
    </w:p>
    ${pautaXml}
    ${tabelaResumoXml}

    <w:p>
      <w:pPr><w:spacing w:before="160" w:after="80" w:line="300" w:lineRule="auto"/></w:pPr>
      <w:r>
        <w:rPr><w:rFonts w:ascii="Arimo" w:cs="Arimo" w:eastAsia="Arimo" w:hAnsi="Arimo"/><w:b/><w:bCs/><w:color w:val="047176"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
        <w:t>3. ENCAMINHAMENTOS E PRÓXIMOS PASSOS</w:t>
      </w:r>
    </w:p>
    ${encaminhamentosXml}
    ${extrasXml}

    <w:p>
      <w:pPr><w:spacing w:before="180" w:after="80" w:line="300" w:lineRule="auto"/></w:pPr>
      <w:r>
        <w:rPr><w:rFonts w:ascii="Arimo" w:cs="Arimo" w:eastAsia="Arimo" w:hAnsi="Arimo"/><w:b/><w:bCs/><w:color w:val="047176"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
        <w:t>4. ACESSO À TRANSCRIÇÃO COMPLETA</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr><w:spacing w:before="20" w:after="40" w:line="300" w:lineRule="auto"/></w:pPr>
      <w:r>
        <w:rPr><w:rFonts w:ascii="Arimo" w:cs="Arimo" w:eastAsia="Arimo" w:hAnsi="Arimo"/><w:color w:val="000000"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>
        <w:t xml:space="preserve">Arquivo local da transcrição integral (${transcriptCount} falas, ${chatCount} mensagens de chat): ${escapeXml(transcriptFileLabel)}</w:t>
      </w:r>
    </w:p>
    ${transcriptLinkParagraph}

    <w:p>
      <w:pPr><w:spacing w:before="280" w:after="120" w:line="300" w:lineRule="auto"/></w:pPr>
      <w:r>
        <w:rPr><w:rFonts w:ascii="Arimo" w:cs="Arimo" w:eastAsia="Arimo" w:hAnsi="Arimo"/><w:b/><w:bCs/><w:color w:val="004aad"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>
        <w:t>Ass</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr><w:spacing w:before="40" w:after="120" w:line="300" w:lineRule="auto"/></w:pPr>
      <w:r>
        <w:rPr><w:rFonts w:ascii="Arimo" w:cs="Arimo" w:eastAsia="Arimo" w:hAnsi="Arimo"/><w:b/><w:bCs/><w:color w:val="004aad"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>
        <w:t>Diretor:</w:t>
      </w:r>
    </w:p>

    <w:sectPr>
      <w:pgSz w:h="16845" w:w="11910" w:orient="portrait"/>
      <w:pgMar w:bottom="720" w:top="720" w:left="720" w:right="720" w:header="720" w:footer="720"/>
      <w:pgNumType w:start="1"/>
    </w:sectPr>
  </w:body>
</w:document>`

    if (templateEntries && templateEntries.size > 0) {
        /** @type {Array<{ name: string, data: Uint8Array }>} */
        const outputEntries = []
        let hasRels = false
        let hasDoc = false

        for (const [entryName, entryBytes] of templateEntries.entries()) {
            if (entryName === "word/document.xml") {
                hasDoc = true
                outputEntries.push({ name: entryName, data: encoder.encode(documentXml) })
            } else if (entryName === "word/_rels/document.xml.rels") {
                hasRels = true
                const existingRels = new TextDecoder("utf-8").decode(entryBytes)
                outputEntries.push({
                    name: entryName,
                    data: encoder.encode(injectTranscriptHyperlinkRel(existingRels, transcriptAccessInfo.viewUrl))
                })
            } else {
                outputEntries.push({ name: entryName, data: entryBytes })
            }
        }

        if (!hasDoc) {
            outputEntries.push({ name: "word/document.xml", data: encoder.encode(documentXml) })
        }
        if (!hasRels) {
            outputEntries.push({
                name: "word/_rels/document.xml.rels",
                data: encoder.encode(injectTranscriptHyperlinkRel("", transcriptAccessInfo.viewUrl))
            })
        }

        return createZipBytes(outputEntries)
    }

    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Arimo" w:hAnsi="Arimo" w:cs="Arimo"/>
        <w:sz w:val="20"/>
        <w:szCs w:val="20"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="character" w:styleId="Hyperlink">
    <w:name w:val="Hyperlink"/>
    <w:rPr>
      <w:color w:val="004aad"/>
      <w:u w:val="single"/>
    </w:rPr>
  </w:style>
  <w:style w:type="table" w:styleId="Table1"><w:name w:val="Table1"/></w:style>
  <w:style w:type="table" w:styleId="Table2"><w:name w:val="Table2"/></w:style>
  <w:style w:type="table" w:styleId="Table3"><w:name w:val="Table3"/></w:style>
</w:styles>`

    const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`

    const documentRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  ${hasHyperlink ? `<Relationship Id="rIdTranscript" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${escapeXml(transcriptAccessInfo.viewUrl)}" TargetMode="External"/>` : ""}
</Relationships>`

    const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`

    return createZipBytes([
        { name: "[Content_Types].xml", data: encoder.encode(contentTypesXml) },
        { name: "_rels/.rels", data: encoder.encode(relsXml) },
        { name: "word/document.xml", data: encoder.encode(documentXml) },
        { name: "word/styles.xml", data: encoder.encode(stylesXml) },
        { name: "word/_rels/document.xml.rels", data: encoder.encode(documentRelsXml) }
    ])
}

/**
 * Generates DOCX bytes by cloning the base DOCX template (e.g. MODELO ATA.docx) when supplied.
 * Preserves embedded fonts, styles, numbering, and themes from the template package.
 * @param {AtaData} ataData
 * @param {{ viewUrl?: string, transcriptFileName?: string, transcriptCount?: number, chatCount?: number }} [transcriptAccessInfo]
 * @param {Uint8Array | null} [templateDocxBytes]
 * @returns {Promise<Uint8Array>}
 */
export async function buildAtaDocxBytesFromTemplate(ataData, transcriptAccessInfo = {}, templateDocxBytes = null) {
    /** @type {Map<string, Uint8Array> | null} */
    let templateEntries = null
    if (templateDocxBytes && templateDocxBytes.length > 0) {
        try {
            templateEntries = await readZipEntries(templateDocxBytes)
        } catch (err) {
            console.warn("Could not read DOCX template entries, falling back to standalone OOXML:", err)
        }
    }
    return buildAtaDocxBytes(ataData, transcriptAccessInfo, templateEntries)
}

/**
 * Builds a DOCX Blob for a given AtaData object.
 * @param {AtaData} ataData
 * @param {{ viewUrl?: string, transcriptFileName?: string, transcriptCount?: number, chatCount?: number }} [transcriptAccessInfo]
 * @returns {Blob}
 */
export function buildAtaDocxBlob(ataData, transcriptAccessInfo = {}) {
    const bytes = buildAtaDocxBytes(ataData, transcriptAccessInfo)
    const bufferCopy = new Uint8Array(bytes.byteLength)
    bufferCopy.set(bytes)
    return new Blob([bufferCopy], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    })
}

/**
 * Builds a DOCX Blob cloned from the base DOCX template when available.
 * @param {AtaData} ataData
 * @param {{ viewUrl?: string, transcriptFileName?: string, transcriptCount?: number, chatCount?: number }} [transcriptAccessInfo]
 * @param {Uint8Array | null} [templateDocxBytes]
 * @returns {Promise<Blob>}
 */
export async function buildAtaDocxBlobFromTemplate(ataData, transcriptAccessInfo = {}, templateDocxBytes = null) {
    const bytes = await buildAtaDocxBytesFromTemplate(ataData, transcriptAccessInfo, templateDocxBytes)
    const bufferCopy = new Uint8Array(bytes.byteLength)
    bufferCopy.set(bytes)
    return new Blob([bufferCopy], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    })
}

/**
 * Builds a print-ready HTML representation of the Ata de Reunião matching the PDF template.
 * @param {AtaData} ataData
 * @param {{ viewUrl?: string, transcriptFileName?: string, transcriptCount?: number, chatCount?: number }} [transcriptAccessInfo]
 * @returns {string}
 */
export function buildAtaHtml(ataData, transcriptAccessInfo = {}) {
    const hasHyperlink = isSafeLinkUrl(transcriptAccessInfo.viewUrl)
    const participantsHtml = (ataData.participantes || [])
        .map(p => `<div>${escapeHtml(String(p).startsWith("-") ? String(p) : `- ${p}`)}</div>`)
        .join("")

    const pautaHtml = (ataData.pauta || []).map((sec) => {
        const numPrefix = sec.numero.endsWith(".") ? sec.numero : `${sec.numero}.`
        const topicosHtml = (sec.topicos || []).map((topic) => {
            const subHtml = (topic.subtopicos && topic.subtopicos.length > 0)
                ? `<ul class="sub-list">${topic.subtopicos.map(s => `<li>${escapeHtml(s)}</li>`).join("")}</ul>`
                : ""
            return `
              <li>
                ${topic.titulo ? `<strong>${escapeHtml(topic.titulo)}:</strong> ` : ""}${escapeHtml(topic.descricao)}
                ${subHtml}
              </li>`
        }).join("")

        return `
          <div class="pauta-subsecao">
            <h3>${escapeHtml(numPrefix)} ${escapeHtml(sec.titulo)}</h3>
            <ul class="main-list">${topicosHtml}</ul>
          </div>`
    }).join("")

    const tabelaResumoHtml = (ataData.tabelaResumo && ataData.tabelaResumo.length > 0)
        ? `<table class="resumo-table">
            <thead>
              <tr>
                <th>Aspecto</th>
                <th>Detalhes da Execução</th>
              </tr>
            </thead>
            <tbody>
              ${ataData.tabelaResumo.map(row => `
                <tr>
                  <td class="resumo-aspecto">${escapeHtml(row.aspecto)}</td>
                  <td class="resumo-detalhes">${escapeHtml(row.detalhes)}</td>
                </tr>`).join("")}
            </tbody>
          </table>`
        : ""

    const encaminhamentosHtml = (ataData.encaminhamentos && ataData.encaminhamentos.length > 0)
        ? `<ol class="encaminhamentos-list">${ataData.encaminhamentos.map((enc) => `
            <li>
              <strong>${escapeHtml(enc.responsavel)}:</strong>
              <ul class="sub-list">${(enc.acoes || []).map(a => `<li>${escapeHtml(a)}</li>`).join("")}</ul>
            </li>`).join("")}</ol>`
        : `<p>Nenhum encaminhamento específico registrado.</p>`

    const extrasHtml = (ataData.secoesExtras || []).map((sec) => `
      <div class="section-block">
        <h2>${escapeHtml(sec.titulo)}</h2>
        <ul class="sub-list">${(sec.conteudo || []).map(line => `<li>${escapeHtml(line)}</li>`).join("")}</ul>
      </div>`).join("")

    const transcriptLinkHtml = hasHyperlink
        ? `<p><strong>Link direto na extensão:</strong> <a href="${escapeHtml(transcriptAccessInfo.viewUrl)}" target="_blank">${escapeHtml(transcriptAccessInfo.viewUrl)}</a></p>`
        : ""

    return `<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8" />
  <title>ATA DE REUNIÃO - ${escapeHtml(ataData.cliente)}</title>
  <style>
    @page { size: A4; margin: 16mm; }
    body { font-family: Arial, Helvetica, sans-serif; color: #000000; margin: 0 auto; max-width: 820px; padding: 32px 40px; line-height: 1.5; background: #ffffff; }
    .ata-title { color: #004AAD; font-size: 36px; font-weight: 800; letter-spacing: 0.5px; margin: 0 0 8px 0; border-bottom: 2px solid #047176; padding-bottom: 10px; text-transform: uppercase; }
    .ata-client { font-size: 15px; font-weight: 700; margin: 12px 0 16px 0; color: #000000; }
    table.meta-table { width: 100%; border-collapse: collapse; margin-bottom: 18px; font-size: 13px; border: 1px solid #00D0DD; }
    table.meta-table th { background: #004AAD; color: #ffffff; padding: 10px 8px; text-align: center; font-weight: 700; text-transform: uppercase; border: 1px solid #00D0DD; border-bottom: 1px solid #047176; }
    table.meta-table td { background: #D4F8F8; color: #000000; padding: 12px 8px; text-align: center; vertical-align: top; border: 1px solid #00D0DD; }
    table.resumo-table { width: 100%; border-collapse: collapse; margin: 14px 0 18px 0; font-size: 13px; border: 1px solid #000000; }
    table.resumo-table th:first-child, td.resumo-aspecto { width: 24%; background: #004AAD; color: #ffffff; font-weight: 700; padding: 8px 10px; border: 1px solid #000000; text-align: left; vertical-align: top; }
    table.resumo-table th:last-child, td.resumo-detalhes { width: 76%; background: #D4F8F8; color: #000000; padding: 8px 10px; border: 1px solid #000000; text-align: left; vertical-align: top; }
    .section-divider { border: none; border-top: 1.5px solid #A0A0A0; margin: 18px 0 14px 0; }
    .pauta-header { color: #004AAD; font-size: 15px; font-weight: 700; margin: 0 0 14px 0; }
    h2 { color: #047176; font-size: 14px; font-weight: 700; margin: 18px 0 8px 0; text-transform: uppercase; }
    h3 { color: #047176; font-size: 13.5px; font-weight: 700; margin: 12px 0 6px 0; }
    ul.main-list { list-style-type: disc; padding-left: 24px; margin: 6px 0; }
    ul.sub-list { list-style-type: circle; padding-left: 24px; margin: 4px 0; }
    ol.encaminhamentos-list { padding-left: 22px; margin: 8px 0; }
    li { margin-bottom: 4px; }
    .signature { margin-top: 32px; color: #004AAD; font-weight: 700; font-size: 15px; }
    .signature div { margin-bottom: 6px; }
    @media print { .no-print { display: none; } body { padding: 0; } }
  </style>
</head>
<body>
  <h1 class="ata-title">ATA DE REUNIÃO</h1>
  <div class="ata-client">Cliente - ${escapeHtml(ataData.cliente)}</div>

  <table class="meta-table">
    <thead>
      <tr>
        <th>LOCAL</th>
        <th>DATA E HORÁRIO</th>
        <th>PARTICIPANTES</th>
        <th>AUSENTES</th>
        <th>SUBDIRETORIA</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td>${escapeHtml(ataData.local)}</td>
        <td>${escapeHtml(ataData.dataHorario)}</td>
        <td><div>Membros presentes:</div>${participantsHtml}</td>
        <td>${escapeHtml(ataData.ausentes || "-----")}</td>
        <td>${escapeHtml(ataData.subdiretoria || "")}</td>
      </tr>
    </tbody>
  </table>

  <hr class="section-divider" />
  <div class="pauta-header">I. Pauta da Reunião</div>

  <h2>1. OBJETIVO DA REUNIÃO</h2>
  <p>${escapeHtml(ataData.objetivo)}</p>

  <h2>2. PAUTA E PONTOS DISCUTIDOS</h2>
  ${pautaHtml}
  ${tabelaResumoHtml}

  <h2>3. ENCAMINHAMENTOS E PRÓXIMOS PASSOS</h2>
  ${encaminhamentosHtml}
  ${extrasHtml}

  <h2>4. ACESSO À TRANSCRIÇÃO COMPLETA</h2>
  <p>Arquivo local da transcrição integral (${transcriptAccessInfo.transcriptCount ?? 0} falas, ${transcriptAccessInfo.chatCount ?? 0} mensagens de chat): <code>${escapeHtml(transcriptAccessInfo.transcriptFileName || "TranscripTonic/Transcript.txt")}</code></p>
  ${transcriptLinkHtml}

  <div class="signature">
    <div>Ass</div>
    <div>Diretor:</div>
  </div>
</body>
</html>`
}

/**
 * Builds the output filename for an Ata de Reunião document.
 * @param {Meeting} meeting
 * @param {"docx" | "html"} [format]
 * @returns {string}
 */
export function getAtaFileName(meeting, format = "docx") {
    const sanitisedTitle = getSanitisedMeetingTitle(meeting)
    const formattedTimestamp = getFormattedFileTimestamp(meeting)
    return `TranscripTonic/ATA-${sanitisedTitle} at ${formattedTimestamp}.${format}`
}

/**
 * Builds transcript access metadata linked inside the generated Ata.
 * @param {Meeting} meeting
 * @param {number} index
 * @returns {{ viewUrl: string, transcriptFileName: string, transcriptCount: number, chatCount: number }}
 */
export function buildTranscriptAccessInfo(meeting, index) {
    const viewUrl = typeof chrome !== "undefined" && chrome.runtime?.getURL
        ? chrome.runtime.getURL(`meetings.html?transcript=${index}`)
        : ""
    return {
        viewUrl,
        transcriptFileName: getTranscriptFileName(meeting),
        transcriptCount: Array.isArray(meeting?.transcript) ? meeting.transcript.length : 0,
        chatCount: Array.isArray(meeting?.chatMessages) ? meeting.chatMessages.length : 0
    }
}

/**
 * Generates an Ata de Reunião via Gemini for meetings[index], stores it, and downloads the DOCX file.
 * @param {number} index
 * @param {{ clientName?: string, subdiretoria?: string, model?: string, skipDownload?: boolean }} [options]
 * @returns {Promise<{ ataData: AtaData, fileName: string, model: string }>}
 */
export async function generateAtaAtIndex(index, options = {}) {
    const [resultLocal, resultSync] = await Promise.all([
        new Promise((resolve) => chrome.storage.local.get(["meetings", "geminiApiKey"], (res) => resolve(res || {}))),
        new Promise((resolve) => chrome.storage.sync.get(["geminiModel", "ataClientName", "ataSubdiretoria"], (res) => resolve(res || {})))
    ])

    const meetings = resultLocal.meetings || []
    if (!meetings[index]) {
        throw { errorCode: "010", errorMessage: "Meeting at specified index not found" }
    }

    const meeting = meetings[index]
    const apiKey = resultLocal.geminiApiKey
    if (!apiKey || !String(apiKey).trim()) {
        throw { errorCode: "017", errorMessage: "Chave API do Google Gemini não configurada." }
    }

    const model = options.model || resultSync.geminiModel || DEFAULT_GEMINI_MODEL
    const clientName = options.clientName ?? resultSync.ataClientName ?? ""
    const subdiretoria = options.subdiretoria ?? resultSync.ataSubdiretoria ?? ""

    const template = await loadAtaTemplate()
    const ataData = await callGeminiForAta({
        apiKey,
        model,
        meeting,
        template,
        clientName,
        subdiretoria
    })

    meetings[index] = {
        ...meeting,
        ataData,
        ataGeneratedAt: new Date().toISOString(),
        ataModel: model
    }

    await new Promise((resolve) => {
        chrome.storage.local.set({ meetings }, () => resolve())
    })

    const fileName = getAtaFileName(meeting, "docx")
    if (!options.skipDownload) {
        const accessInfo = buildTranscriptAccessInfo(meeting, index)
        const docxBlob = await buildAtaDocxBlobFromTemplate(ataData, accessInfo, template.docxBytes)
        await downloadBlobOrDataUrl({
            blob: docxBlob,
            fileName,
            fallbackFileName: "TranscripTonic/ATA.docx",
            mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        })
    }

    return { ataData, fileName, model }
}

/**
 * Downloads an already-generated Ata de Reunião for meetings[index] in DOCX or HTML format.
 * @param {number} index
 * @param {"docx" | "html"} [format]
 * @returns {Promise<string>}
 */
export async function downloadAtaAtIndex(index, format = "docx") {
    const resultLocal = await new Promise((resolve) => {
        chrome.storage.local.get(["meetings"], (res) => resolve(res || {}))
    })

    const meetings = resultLocal.meetings || []
    const meeting = meetings[index]
    if (!meeting) {
        throw { errorCode: "010", errorMessage: "Meeting at specified index not found" }
    }
    if (!meeting.ataData) {
        throw { errorCode: "018", errorMessage: "Ata ainda não foi gerada para esta reunião." }
    }

    const accessInfo = buildTranscriptAccessInfo(meeting, index)
    if (format === "html") {
        const html = buildAtaHtml(meeting.ataData, accessInfo)
        const blob = new Blob([html], { type: "text/html;charset=utf-8" })
        const fileName = getAtaFileName(meeting, "html")
        return downloadBlobOrDataUrl({
            blob,
            fileName,
            fallbackFileName: "TranscripTonic/ATA.html",
            mimeType: "text/html;charset=utf-8"
        })
    }

    const template = await loadAtaTemplate()
    const blob = await buildAtaDocxBlobFromTemplate(meeting.ataData, accessInfo, template.docxBytes)
    const fileName = getAtaFileName(meeting, "docx")
    return downloadBlobOrDataUrl({
        blob,
        fileName,
        fallbackFileName: "TranscripTonic/ATA.docx",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    })
}
