import { describe, it } from "node:test"
import assert from "node:assert/strict"
import fs from "node:fs"
import path from "node:path"
import vm from "node:vm"
import { execFileSync } from "node:child_process"
import { fileURLToPath } from "node:url"

import {
    DEFAULT_GEMINI_MODEL,
    buildAtaDocxBytes,
    buildAtaHtml,
    buildTranscriptAccessInfo,
    callGeminiForAta,
    extractTextFromDocxBytes,
    readZipEntries
} from "../extension/background-script/ata-generator.js"
import { downloadBlobOrDataUrl } from "../extension/background-script/exporters.js"
import { requestPlatformPermission } from "../extension/background-script/platforms.js"
import { isFirefox, openSidePanel } from "../extension/background-script/utils.js"
import { buildMeetingRowHtml } from "../extension/meetings.js"

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const ROOT_DIR = path.resolve(__dirname, "..")
const EXTENSION_DIR = path.join(ROOT_DIR, "extension")
const DIST_DIR = path.join(ROOT_DIR, "dist")

/**
 * Temporarily overrides global properties for isolated testing.
 * @param {Record<string, any>} overrides
 * @param {() => Promise<void> | void} fn
 */
async function withMockedGlobals(overrides, fn) {
    /** @type {Array<{ key: string, desc: PropertyDescriptor | undefined }>} */
    const saved = []
    for (const [key, value] of Object.entries(overrides)) {
        saved.push({
            key,
            desc: Object.getOwnPropertyDescriptor(globalThis, key)
        })
        Object.defineProperty(globalThis, key, {
            value,
            configurable: true,
            writable: true
        })
    }

    try {
        await fn()
    } finally {
        for (const { key, desc } of saved) {
            if (desc) {
                Object.defineProperty(globalThis, key, desc)
            } else {
                delete globalThis[key]
            }
        }
    }
}

/**
 * Recursively lists files in a directory.
 * @param {string} dir
 * @returns {string[]}
 */
function walkFiles(dir) {
    /** @type {string[]} */
    const files = []
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
        const fullPath = path.join(dir, entry.name)
        if (entry.isDirectory()) {
            files.push(...walkFiles(fullPath))
        } else {
            files.push(fullPath)
        }
    }
    return files
}

describe("Cross-browser packaging & manifests", () => {
    it("builds Chrome and Firefox MV3 bundles and ZIP packages with templates", async () => {
        execFileSync("bash", [path.join(ROOT_DIR, "scripts/build-cross.sh")], {
            cwd: ROOT_DIR,
            stdio: "pipe"
        })

        const chromeManifestPath = path.join(DIST_DIR, "chrome", "manifest.json")
        const firefoxManifestPath = path.join(DIST_DIR, "firefox", "manifest.json")
        const chromeZipPath = path.join(DIST_DIR, "transcriptonic-chrome-v3.4.5.zip")
        const firefoxZipPath = path.join(DIST_DIR, "transcriptonic-firefox-v3.4.5.zip")

        assert.ok(fs.existsSync(chromeManifestPath), "Chrome dist manifest should exist")
        assert.ok(fs.existsSync(firefoxManifestPath), "Firefox dist manifest should exist")
        assert.ok(fs.existsSync(chromeZipPath), "Chrome ZIP should exist")
        assert.ok(fs.existsSync(firefoxZipPath), "Firefox ZIP should exist")
        assert.equal(fs.existsSync(path.join(DIST_DIR, "chrome", "manifest-firefox.json")), false)
        assert.equal(fs.existsSync(path.join(DIST_DIR, "firefox", "manifest-firefox.json")), false)

        const chromeManifest = JSON.parse(fs.readFileSync(chromeManifestPath, "utf8"))
        assert.equal(chromeManifest.manifest_version, 3)
        assert.equal(chromeManifest.background.service_worker, "background-script/index.js")
        assert.equal(chromeManifest.background.scripts, undefined)
        assert.equal(chromeManifest.sidebar_action, undefined)
        assert.equal(chromeManifest.browser_specific_settings, undefined)
        assert.ok(chromeManifest.permissions.includes("sidePanel"))
        assert.ok(chromeManifest.host_permissions.includes("https://generativelanguage.googleapis.com/*"))
        assert.ok(chromeManifest.optional_host_permissions.includes("https://*.zoom.us/*"))
        assert.ok(chromeManifest.web_accessible_resources?.[0]?.resources?.includes("templates/*"))

        const firefoxManifest = JSON.parse(fs.readFileSync(firefoxManifestPath, "utf8"))
        assert.equal(firefoxManifest.manifest_version, 3)
        assert.deepEqual(firefoxManifest.background.scripts, ["background-script/index.js"])
        assert.equal(firefoxManifest.background.service_worker, undefined)
        assert.equal(firefoxManifest.sidebar_action?.default_panel, "side-panel/index.html")
        assert.ok(firefoxManifest.browser_specific_settings?.gecko?.id)
        assert.equal(firefoxManifest.permissions.includes("sidePanel"), false)
        assert.ok(firefoxManifest.host_permissions.includes("https://generativelanguage.googleapis.com/*"))

        const chromeEntries = await readZipEntries(new Uint8Array(fs.readFileSync(chromeZipPath)))
        const firefoxEntries = await readZipEntries(new Uint8Array(fs.readFileSync(firefoxZipPath)))
        for (const required of [
            "manifest.json",
            "background-script/index.js",
            "background-script/ata-generator.js",
            "meetings.html",
            "meetings.js",
            "templates/modelo-ata.pdf",
            "templates/modelo-ata.docx"
        ]) {
            assert.ok(chromeEntries.has(required), `Chrome zip missing ${required}`)
            assert.ok(firefoxEntries.has(required), `Firefox zip missing ${required}`)
        }
    })
})

describe("Cross-browser downloads, sidePanel, and Firefox compatibility", () => {
    it("detects Firefox cleanly only when browser API and Firefox UA are present", async () => {
        await withMockedGlobals(
            {
                browser: undefined,
                navigator: { userAgent: "Mozilla/5.0 Chrome/130.0.0.0" }
            },
            () => {
                assert.equal(isFirefox(), false)
            }
        )

        await withMockedGlobals(
            {
                browser: { runtime: {} },
                navigator: { userAgent: "Mozilla/5.0 (X11; Linux x86_64; rv:140.0) Gecko/20100101 Firefox/140.0" }
            },
            () => {
                assert.equal(isFirefox(), true)
            }
        )
    })

    it("uses URL.createObjectURL in Firefox and data URL in Chromium, falling back to meetings.html", async () => {
        /** @type {any[]} */
        const firefoxCalls = []
        let revokedUrl = null
        await withMockedGlobals(
            {
                browser: {
                    downloads: {
                        download: async (opts) => {
                            firefoxCalls.push(opts)
                            return 1
                        }
                    }
                },
                navigator: { userAgent: "Mozilla/5.0 Firefox/140.0" },
                chrome: { runtime: {} },
                URL: {
                    createObjectURL: () => "blob:moz-extension://uuid/123",
                    revokeObjectURL: (u) => {
                        revokedUrl = u
                    }
                },
                setTimeout: () => 0
            },
            async () => {
                const msg = await downloadBlobOrDataUrl({
                    blob: new Blob(["hello"], { type: "text/plain" }),
                    fileName: "TranscripTonic/Test.txt"
                })
                assert.equal(msg, "File downloaded successfully")
                assert.equal(firefoxCalls.length, 1)
                assert.equal(firefoxCalls[0].url, "blob:moz-extension://uuid/123")
                assert.equal(revokedUrl, null)
            }
        )

        /** @type {any[]} */
        const chromeCalls = []
        await withMockedGlobals(
            {
                browser: undefined,
                navigator: { userAgent: "Mozilla/5.0 Chrome/130.0.0.0" },
                chrome: {
                    downloads: {
                        download: async (opts) => {
                            chromeCalls.push(opts)
                            return 2
                        }
                    }
                }
            },
            async () => {
                const msg = await downloadBlobOrDataUrl({
                    blob: new Blob(["hello"], { type: "text/plain" }),
                    fileName: "TranscripTonic/Test.txt"
                })
                assert.equal(msg, "File downloaded successfully")
                assert.equal(chromeCalls.length, 1)
                assert.ok(chromeCalls[0].url.startsWith("data:text/plain"))
            }
        )

        /** @type {any[]} */
        const sentMessages = []
        await withMockedGlobals(
            {
                browser: {
                    downloads: {
                        download: async () => {
                            throw new Error("Download API blocked")
                        }
                    }
                },
                navigator: { userAgent: "Mozilla/5.0 Firefox/140.0" },
                chrome: {
                    runtime: {
                        getURL: (p) => `moz-extension://id/${p}`,
                        lastError: null
                    },
                    tabs: {
                        query: (_q, cb) => cb([{ id: 77, url: "moz-extension://id/meetings.html" }]),
                        update: (_id, _opts, cb) => cb(),
                        sendMessage: (_tabId, payload, cb) => {
                            sentMessages.push(payload)
                            cb({ success: true })
                        }
                    }
                },
                URL: {
                    createObjectURL: () => "blob:moz-extension://uuid/fallback",
                    revokeObjectURL: () => { }
                }
            },
            async () => {
                const msg = await downloadBlobOrDataUrl({
                    blob: new Blob(["ata"], { type: "application/octet-stream" }),
                    fileName: "TranscripTonic/ATA.docx",
                    fallbackFileName: "TranscripTonic/ATA.docx"
                })
                assert.equal(msg, "Downloaded via meetings page")
                assert.equal(sentMessages.length, 1)
                assert.equal(sentMessages[0].type, "download_blob_via_page")
                assert.equal(sentMessages[0].base64Data, Buffer.from("ata").toString("base64"))
            }
        )
    })

    it("opens Chrome sidePanel, Firefox sidebarAction, or popup window fallback", async () => {
        /** @type {any[]} */
        const calls = []

        await withMockedGlobals(
            {
                chrome: {
                    sidePanel: {
                        open: async (opts) => {
                            calls.push(["sidePanel", opts])
                        }
                    }
                }
            },
            async () => {
                await openSidePanel(99)
                assert.deepEqual(calls, [["sidePanel", { tabId: 99 }]])
            }
        )

        calls.length = 0
        await withMockedGlobals(
            {
                chrome: {
                    runtime: { getURL: (p) => `moz-extension://id/${p}` },
                    windows: { create: async (o) => calls.push(["window", o]) }
                },
                browser: {
                    sidebarAction: {
                        open: async () => {
                            calls.push(["sidebarAction"])
                        }
                    }
                }
            },
            async () => {
                await openSidePanel()
                assert.deepEqual(calls, [["sidebarAction"]])
            }
        )

        calls.length = 0
        await withMockedGlobals(
            {
                chrome: {
                    runtime: { getURL: (p) => `chrome-extension://id/${p}` },
                    windows: {
                        create: async (opts) => {
                            calls.push(["window", opts])
                            return { id: 1 }
                        }
                    },
                    tabs: { query: (_q, cb) => cb([]) }
                },
                browser: undefined
            },
            async () => {
                await openSidePanel()
                assert.equal(calls.length, 1)
                assert.equal(calls[0][0], "window")
                assert.equal(calls[0][1].url, "chrome-extension://id/side-panel/index.html")
            }
        )
    })
})

describe("Permission & FAB bug regressions", () => {
    it("filters non-optional permissions and skips request when already granted", async () => {
        /** @type {any[]} */
        const requested = []
        let containsResult = false

        await withMockedGlobals(
            {
                chrome: {
                    permissions: {
                        contains: async () => containsResult,
                        request: async (req) => {
                            requested.push(req)
                            return true
                        }
                    }
                }
            },
            async () => {
                await requestPlatformPermission("zoom")
                assert.equal(requested.length, 1)
                assert.deepEqual(requested[0].origins, ["https://*.zoom.us/*"])
                assert.deepEqual(requested[0].permissions, ["notifications"])
                assert.equal(requested[0].permissions.includes("declarativeNetRequestWithHostAccess"), false)

                containsResult = true
                await requestPlatformPermission("zoom")
                assert.equal(requested.length, 1, "Should not prompt again when already granted")
            }
        )
    })

    it("safely handles empty/undefined state in updateFabState without throwing", () => {
        const commonConfigSrc = fs.readFileSync(path.join(EXTENSION_DIR, "content-scripts/common-config.js"), "utf8")
        const commonUtilsSrc = fs.readFileSync(path.join(EXTENSION_DIR, "content-scripts/common-utils.js"), "utf8")

        const elements = {
            "#fab-default-logo": { style: { display: "" } },
            "#fab-letter-mark": { style: { display: "" }, textContent: "" },
            "#fab-mini-badge": { style: { display: "" } }
        }
        const fabMock = {
            querySelector: (sel) => elements[sel] || null
        }
        const sandbox = {
            chrome: { runtime: { getURL: (p) => `chrome-extension://id/${p}` } },
            document: {
                querySelector: (sel) => (sel === "#transcriptonic-fab" ? fabMock : elements[sel] || null)
            },
            navigator: { userAgent: "Chrome" },
            window: {},
            console
        }

        vm.runInNewContext(`${commonConfigSrc}\n${commonUtilsSrc}`, sandbox)

        assert.doesNotThrow(() => sandbox.updateFabState(undefined))
        assert.doesNotThrow(() => sandbox.updateFabState({ transcript: [] }))
        assert.equal(elements["#fab-default-logo"].style.display, "block")

        assert.doesNotThrow(() => sandbox.updateFabState({
            transcript: [{ personName: "Alice", timestamp: "2026-01-01T10:00:00Z", transcriptText: "Hi" }],
            stateTranscriptBlock: { personName: "Alice", timestamp: "", mutationTargetElement: null, transcriptTextBuffer: "" }
        }))
        assert.equal(elements["#fab-default-logo"].style.display, "none")
        assert.equal(elements["#fab-letter-mark"].textContent, "A")
        assert.equal(elements["#fab-mini-badge"].style.display, "block")
    })

    it("serializes permissions.onAdded content script registration and gates async return true", () => {
        const bgIndexSrc = fs.readFileSync(path.join(EXTENSION_DIR, "background-script/index.js"), "utf8")
        assert.match(
            bgIndexSrc,
            /deregisterContentScript\(platforms\)[\s\S]*?\.then\([\s\S]*?reRegisterContentScripts\(\)/
        )
        assert.match(bgIndexSrc, /return ASYNC_RESPONSE_TYPES\.has\(message\.type\)/)
    })
})

describe("Security audit & privacy regressions", () => {
    it("contains zero references to ejnana.github.io or script.google.com across extension/", () => {
        for (const filePath of walkFiles(EXTENSION_DIR)) {
            if (!/\.(js|mjs|html|json|css)$/i.test(filePath)) continue
            const content = fs.readFileSync(filePath, "utf8")
            assert.equal(content.includes("ejnana.github.io"), false, `Unexpected ejnana.github.io in ${filePath}`)
            assert.equal(content.includes("script.google.com"), false, `Unexpected script.google.com in ${filePath}`)
        }
    })

    it("escapes meeting title and software in meetings table rows and side panel (prevents DOM XSS)", () => {
        const xssPayload = `<img src=x onerror=alert("xss")>`
        const rowHtml = buildMeetingRowHtml({
            meetingSoftware: `<svg onload=alert(1)>`,
            meetingTitle: xssPayload,
            meetingStartTimestamp: "2026-10-21T14:00:00.000Z",
            meetingEndTimestamp: "2026-10-21T14:45:00.000Z",
            transcript: [],
            chatMessages: [],
            webhookPostStatus: "new"
        }, 0)

        assert.equal(rowHtml.includes("<img src=x"), false)
        assert.equal(rowHtml.includes("<svg onload"), false)
        assert.ok(rowHtml.includes("&lt;img src=x onerror=alert(&quot;xss&quot;)&gt;"))
        assert.ok(rowHtml.includes("&lt;svg onload=alert(1)&gt;"))

        const sidePanelSrc = fs.readFileSync(path.join(EXTENSION_DIR, "side-panel/side-panel.js"), "utf8")
        assert.equal(/meetingTitle\.innerHTML\s*=/.test(sidePanelSrc), false)
        assert.ok(sidePanelSrc.includes("meetingTitle.textContent ="))
    })

    it("sends Gemini API key strictly in x-goog-api-key header and never in query string", async () => {
        /** @type {Array<{ url: string, init: any }>} */
        const fetchCalls = []
        const mockFetch = async (url, init) => {
            fetchCalls.push({ url: String(url), init })
            return {
                ok: true,
                status: 200,
                json: async () => ({
                    candidates: [{
                        content: {
                            parts: [{
                                text: JSON.stringify({
                                    cliente: "Sofia Tecnologia",
                                    local: "Meet",
                                    dataHorario: "21/10/2026; 14:00",
                                    participantes: ["Sofia", "Wescley"],
                                    ausentes: "-----",
                                    subdiretoria: "Projetos",
                                    objetivo: "Alinhar escopo do projeto.",
                                    pauta: [
                                        {
                                            numero: "2.1.",
                                            titulo: "Escopo",
                                            topicos: [
                                                {
                                                    titulo: "Cronograma",
                                                    descricao: "Definido marco inicial.",
                                                    subtopicos: ["Validar entregas na sexta"]
                                                }
                                            ]
                                        }
                                    ],
                                    encaminhamentos: [
                                        { responsavel: "Wescley", acoes: ["Publicar versão de testes"] }
                                    ]
                                })
                            }]
                        }
                    }]
                })
            }
        }

        const result = await callGeminiForAta({
            apiKey: "AIza-TEST-SECRET-KEY",
            model: DEFAULT_GEMINI_MODEL,
            meeting: {
                meetingSoftware: "Google Meet",
                meetingTitle: "Alinhamento Sofia",
                meetingStartTimestamp: "2026-10-21T14:00:00.000Z",
                meetingEndTimestamp: "2026-10-21T14:30:00.000Z",
                transcript: [
                    { personName: "Sofia", timestamp: "2026-10-21T14:05:00.000Z", transcriptText: "Vamos definir o escopo." },
                    { personName: "Wescley", timestamp: "2026-10-21T14:10:00.000Z", transcriptText: "Vou publicar a versão de testes." }
                ],
                chatMessages: [],
                webhookPostStatus: "new"
            },
            fetchImpl: /** @type {any} */ (mockFetch)
        })

        assert.equal(fetchCalls.length, 1)
        assert.ok(fetchCalls[0].url.includes(`models/${DEFAULT_GEMINI_MODEL}:generateContent`))
        assert.equal(fetchCalls[0].url.includes("AIza-TEST-SECRET-KEY"), false)
        assert.equal(fetchCalls[0].url.includes("key="), false)
        assert.equal(fetchCalls[0].init.headers["x-goog-api-key"], "AIza-TEST-SECRET-KEY")
        assert.equal(result.cliente, "Sofia Tecnologia")
        assert.deepEqual(result.participantes, ["Sofia", "Wescley"])
        assert.equal(result.pauta[0].numero, "2.1.")
    })
})

describe("DOCX ZIP/XML & HTML Ata generation", () => {
    it("reads bundled templates/modelo-ata.docx cleanly", async () => {
        const docxBytes = new Uint8Array(fs.readFileSync(path.join(EXTENSION_DIR, "templates/modelo-ata.docx")))
        const extracted = await extractTextFromDocxBytes(docxBytes)
        assert.ok(extracted.includes("ATA DE REUNIÃO"))
        assert.ok(extracted.includes("LOCAL"))
        assert.ok(extracted.includes("DATA E HORÁRIO"))
        assert.ok(extracted.includes("PARTICIPANTES"))
        assert.ok(extracted.includes("AUSENTES"))
        assert.ok(extracted.includes("SUBDIRETORIA"))
        assert.ok(extracted.includes("1. OBJETIVO DA REUNIÃO"))
        assert.ok(extracted.includes("3. ENCAMINHAMENTOS E PRÓXIMOS PASSOS"))
    })

    it("generates valid DOCX archive with 5-column metadata table, sections 1-4, transcript link, and XML escaping", async () => {
        const sampleAta = {
            cliente: 'Empresa A & B <Filial "Sul">',
            local: "Meet",
            dataHorario: "21/10/2026; 14:00",
            participantes: ["Sofia <PM>", "Wescley"],
            ausentes: "-----",
            subdiretoria: "Projetos & QA",
            objetivo: "Revisar requisitos <críticos> & próximos passos.",
            pauta: [
                {
                    numero: "2.1.",
                    titulo: "Arquitetura & Segurança",
                    topicos: [
                        {
                            titulo: "Headers",
                            descricao: "Usar x-goog-api-key.",
                            subtopicos: ["Evitar chave em query string"]
                        }
                    ]
                }
            ],
            encaminhamentos: [
                { responsavel: "Wescley", acoes: ["Abrir PR no fork & rodar testes"] }
            ],
            secoesExtras: []
        }

        const accessInfo = {
            viewUrl: "chrome-extension://abcdef/meetings.html?transcript=2",
            transcriptFileName: "TranscripTonic/Google Meet transcript-Test.txt",
            transcriptCount: 12,
            chatCount: 3
        }

        const docxBytes = buildAtaDocxBytes(sampleAta, accessInfo)
        assert.equal(docxBytes[0], 0x50)
        assert.equal(docxBytes[1], 0x4b)

        const entries = await readZipEntries(docxBytes)
        for (const xmlPath of [
            "[Content_Types].xml",
            "_rels/.rels",
            "word/_rels/document.xml.rels",
            "word/styles.xml",
            "word/document.xml"
        ]) {
            assert.ok(entries.has(xmlPath), `DOCX missing ${xmlPath}`)
        }

        // Validate every XML entry parses cleanly with Python ElementTree
        const xmlPayload = {}
        for (const [name, bytes] of entries.entries()) {
            xmlPayload[name] = Buffer.from(bytes).toString("utf8")
        }
        execFileSync(
            "python3",
            [
                "-c",
                "import json, sys, xml.etree.ElementTree as ET\n" +
                "data = json.load(sys.stdin)\n" +
                "for k, v in data.items(): ET.fromstring(v)"
            ],
            { input: JSON.stringify(xmlPayload), stdio: ["pipe", "pipe", "pipe"] }
        )

        const docXml = xmlPayload["word/document.xml"]
        const relsXml = xmlPayload["word/_rels/document.xml.rels"]

        assert.ok(docXml.includes("Empresa A &amp; B &lt;Filial &quot;Sul&quot;&gt;"))
        assert.ok(docXml.includes("LOCAL"))
        assert.ok(docXml.includes("DATA E HORÁRIO"))
        assert.ok(docXml.includes("PARTICIPANTES"))
        assert.ok(docXml.includes("AUSENTES"))
        assert.ok(docXml.includes("SUBDIRETORIA"))
        assert.ok(docXml.includes("1. OBJETIVO DA REUNIÃO"))
        assert.ok(docXml.includes("2. PAUTA E PONTOS DISCUTIDOS"))
        assert.ok(docXml.includes("3. ENCAMINHAMENTOS E PRÓXIMOS PASSOS"))
        assert.ok(docXml.includes("4. ACESSO À TRANSCRIÇÃO COMPLETA"))
        assert.ok(docXml.includes("Ass"))
        assert.ok(docXml.includes("Diretor:"))
        assert.ok(relsXml.includes("Target=\"chrome-extension://abcdef/meetings.html?transcript=2\""))
    })

    it("escapes XSS and blocks unsafe javascript: links in HTML Ata preview", async () => {
        const html = buildAtaHtml(
            {
                cliente: `<script>alert(1)</script>`,
                local: "Meet",
                dataHorario: "21/10/2026; 14:00",
                participantes: [`<img src=x onerror=alert(2)>`],
                ausentes: "-----",
                subdiretoria: "",
                objetivo: `<b>Teste</b>`,
                pauta: [],
                encaminhamentos: []
            },
            {
                viewUrl: "javascript:alert(3)",
                transcriptFileName: "TranscripTonic/<bad>.txt",
                transcriptCount: 1,
                chatCount: 0
            }
        )

        assert.equal(html.includes("<script>alert(1)</script>"), false)
        assert.equal(html.includes("<img src=x"), false)
        assert.equal(html.includes(`href="javascript:`), false)
        assert.ok(html.includes("&lt;script&gt;alert(1)&lt;/script&gt;"))

        await withMockedGlobals(
            { chrome: { runtime: { getURL: (p) => `moz-extension://xyz/${p}` } } },
            () => {
                const access = buildTranscriptAccessInfo(
                    {
                        meetingSoftware: "Teams",
                        meetingTitle: "Sync",
                        meetingStartTimestamp: "2026-10-21T10:00:00Z",
                        meetingEndTimestamp: "2026-10-21T10:30:00Z",
                        transcript: [],
                        chatMessages: [],
                        webhookPostStatus: "new"
                    },
                    4
                )
                assert.equal(access.viewUrl, "moz-extension://xyz/meetings.html?transcript=4")
            }
        )
    })
})
