# CommitJr Atas

Extensão para **Google Chrome** e **Mozilla Firefox** que captura automaticamente as transcrições e mensagens do chat no **Google Meet** e gera a **Ata de Reunião (.docx)** no padrão oficial da **Commit Jr.** com auxílio do **Google Gemini**.

## Recursos principais

- **Captura automática no Google Meet**: ao entrar em uma reunião, as legendas e o chat são gravados automaticamente em armazenamento local do navegador (sem servidores intermediários ou webhooks).
- **Legendas discretas**: marque **Ocultar legendas na tela do Google Meet** no popup da extensão se não quiser ver as legendas nativas do Meet durante a chamada.
- **Ata padronizada com IA (`MODELO ATA.docx`)**: sintetiza cliente/projeto, assunto resumido, participantes, pauta numerada, encaminhamentos, nome inteligente de arquivo (`ATA - DD-MM-YYYY - [Cliente -] Objetivo.docx`) e anexa automaticamente o **Anexo I** com a transcrição completa após uma página em branco.
- **Privacidade local**: sua chave API do Google Gemini e todo o histórico de reuniões ficam salvos exclusivamente no navegador (`chrome.storage.local`), enviados apenas diretamente para `https://generativelanguage.googleapis.com` quando uma Ata é gerada.

## Instalação rápida para membros da Commit Jr.

Gere os pacotes prontos na pasta `dist/`:

```bash
bash scripts/build-cross.sh
```

### Google Chrome / Brave / Microsoft Edge
1. Acesse `chrome://extensions` e ative o **Modo do desenvolvedor** (Developer mode).
2. Clique em **Carregar sem compactação** (Load unpacked) e selecione a pasta `dist/chrome/` (ou instale o pacote `dist/transcriptonic-chrome-v3.4.5.zip`).

### Mozilla Firefox
1. Acesse `about:debugging#/runtime/this-firefox`.
2. Clique em **Carregar extensão temporária** e selecione `dist/firefox/manifest.json` (ou instale `dist/transcriptonic-firefox-v3.4.5.zip` em builds que aceitem XPI/ZIP locais).

## Como configurar e gerar Atas

1. Clique no ícone **CommitJr Atas** na barra do navegador e depois em **Histórico de Atas & Configurações**.
2. Em **Configuração do Gemini**, clique em **Obter chave gratuita no Google AI Studio**, crie sua chave (`AIza...`) e salve-a no campo indicado.
3. Escolha o modelo Gemini desejado (padrão recomendado: `Gemini 3.5 Flash-Lite`) e mantenha ativada a geração automática de Ata ao final da reunião.
4. Participe normalmente da reunião no Google Meet. Ao encerrar a chamada (ou clicando em **Gerar Ata** no histórico), os arquivos `.txt` da transcrição e `.docx` da Ata serão salvos na sua pasta de downloads (`Downloads/TranscripTonic/`).

## Estrutura do projeto

```text
extension/
├── manifest.json                 # Manifest MV3 base (adaptado pelo build para Chrome e Firefox)
├── icon.png / icons/             # Ícone oficial e mascote Pull em SVG
├── meetings.html / meetings.js   # Histórico completo, preview/PDF da Ata e configurações Gemini
├── popup/                        # Popup minimalista (ocultar legendas + acesso às configurações)
├── side-panel/                   # Painel lateral de transcrição ao vivo
├── content-scripts/
│   ├── meet-utils.js             # Helpers, FAB e captura de legendas/chat no Google Meet
│   └── meet.js                   # Ponto de entrada no Google Meet
├── background-script/
│   ├── index.js                  # Roteamento de eventos e ciclo de vida da extensão
│   ├── ata-generator.js          # Integração Gemini e geração do DOCX/HTML com Anexo I
│   ├── exporters.js              # Download cross-browser (Chrome e Firefox)
│   ├── meetings.js               # Persistência e recuperação do histórico de reuniões
│   ├── platforms.js              # Registro dinâmico dos content scripts no Google Meet
│   └── utils.js                  # Configurações do Google Meet, timestamps e side panel
└── templates/
    └── modelo-ata.docx           # Modelo oficial da Ata de Reunião da Commit Jr.
scripts/
└── build-cross.sh                # Empacota dist/chrome e dist/firefox (+ ZIPs)
tests/
└── extension.test.mjs            # Suíte de testes unitários e de empacotamento
```

## Build e testes

```bash
bash scripts/build-cross.sh
node --test tests/extension.test.mjs
```
