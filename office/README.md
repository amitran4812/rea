# Rea Office

An installable desktop office suite with Claude built in:

- **Rea Docs**: a word processor that opens and saves `.docx`, plus HTML, Markdown and plain text.
- **Rea Sheets**: a spreadsheet with formulas, multiple sheets, fill handle, sorting and `.xlsx`/`.csv` support.
- **Rea Slides**: a presentation editor with themes, layouts, shapes, images, speaker notes and a full-screen slide show. It saves `.reaslides` and exports `.pptx`.

Every app has a familiar ribbon toolbar, a File screen (New, Open, Save, Export, Print, Options), a status bar and a **Claude** tab and side pane.

## What Claude can do

| App | Claude features |
| --- | --- |
| Docs | Draft from a description, continue writing, improve or shorten a selection, change tone, summarize, proofread with one-click corrections, translate, chat about the document |
| Sheets | Write a formula for the active cell, explain a formula, analyze the data, generate a table, clean up a messy selection |
| Slides | Write a whole deck from a topic, improve the current slide, suggest the next slide, write speaker notes for every slide, and **What to say** during a slide show: a button on the bottom bar (or the S key) shows a short script for the slide on screen that follows along as you change slides |

Claude runs through your own Anthropic API key. Add it in **File > Options** or in the Claude pane the first time you open it. The key is encrypted with the operating system's keychain (Electron `safeStorage`) and only ever leaves your computer in requests to the Anthropic API. You can also set `ANTHROPIC_API_KEY` in the environment instead.

The default model is Claude Opus 5.5 with medium effort; you can switch to Sonnet 5.5 or Haiku 5.5 and change the effort in Options.

## Install

Installers are built by the **Rea Office desktop** GitHub Actions workflow (`.github/workflows/office-desktop.yml`) for Windows (`.exe`), macOS (`.dmg`) and Linux (`.AppImage`, `.deb`). Open the latest run and download the artifact for your platform.

The builds are not code-signed yet, so Windows SmartScreen will warn the first time ("More info" > "Run anyway") and macOS needs right-click > Open.

## Develop

```bash
cd office
npm install
npm start          # build the renderer bundles and launch the app
npm test           # file format and formula engine tests
npm run dist:win   # build a Windows installer into office/release/
```

`npm run watch` rebuilds the renderer bundles on change; reload the window with Ctrl+R.

### Layout

```
main/            Electron main process
  main.js        windows, menus, IPC, open/save/export dialogs
  preload.js     the window.office bridge exposed to the renderer
  claude.js      streams Claude responses (the API key stays in this process)
  settings.js    settings + encrypted API key
  formats/       .docx, .xlsx/.csv and .pptx conversion
common/          helpers shared by main and renderer (cell references, CSV)
src/
  shared/        ribbon, app frame, File screen, Claude pane, icons, styles
  home/          start screen
  docs/          word processor (TipTap / ProseMirror)
  sheets/        spreadsheet grid and formula engine
  slides/        presentation editor
test/            node:test suites
build/icon.*     app icon
```

## Built on

Open-source components only: [Electron](https://www.electronjs.org/), [TipTap](https://tiptap.dev/)/ProseMirror, [fast-formula-parser](https://github.com/LesterLyu/fast-formula-parser), [docx](https://github.com/dolanmiu/docx), [mammoth](https://github.com/mwilliamson/mammoth.js), [ExcelJS](https://github.com/exceljs/exceljs), [PptxGenJS](https://github.com/gitbrent/PptxGenJS), [marked](https://marked.js.org/) and the [Anthropic TypeScript SDK](https://github.com/anthropics/anthropic-sdk-typescript).

The look follows the general layout conventions of desktop office suites (ribbon tabs, a File screen, a status bar). All icons, the app icon, colours and code are original; nothing is taken from Microsoft Office or any other proprietary product. Microsoft, Word, Excel and PowerPoint are trademarks of Microsoft Corporation; Rea Office is not affiliated with Microsoft. File formats are handled through the open libraries above.

## Known limits (first version)

- Docs shows pages as one continuous sheet with page markers rather than true pagination; headers, footers, comments and track changes aren't supported yet.
- Sheets has no charts, merged cells or conditional formatting yet, and the grid starts at 200 rows by 26 columns (it grows as you add data).
- Slides can't import `.pptx` yet (export works) and has no animations or transitions.
