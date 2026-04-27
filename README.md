# Math Popup

A tray-resident notepad for quick math. Click the system-tray icon, get a small
floating panel with a math-aware notepad. Variables, line references,
percentages, basis points, custom unit suffixes, light markdown, syntax
highlighting and one-keystroke copy.

## Features

- **Live results** in a right-side gutter, like Obsidian's *Solve* plugin —
  results never modify your text.
- **Variables**: write `z = 5` then use `z` anywhere below.
- **Line refs**: `L1`, `L2`, ... resolve to that line's numeric result.
- **Percentages**: `150 + 20%` → `180`. `100 * 50%` → `50`.
- **Basis points**: `100M * 50bps` → `500,000`.
- **Custom suffixes**: configure `m → 1,000,000`, `k → 1,000`, etc. Type `1m`
  + space and the editor expands it to `1,000,000`.
- **Auto-comma**: `1000` + space → `1,000`. Re-formats as you edit.
- **Leading zero**: ` .25` + space → ` 0.25`.
- **Markdown lite**: `#`, `##`, `###` headers, `-`/`*` bullets,
  `**bold**`, `*italic*`.
- **Wrap-aware**: long lines wrap, and the line-number and result columns
  follow the wrap.
- **Syntax colors**: variables blue, operators grey, results green,
  line refs cyan, headers bold.
- **Drag anywhere**: the title bar is the drag handle, panel can sit anywhere
  on screen and remembers its position.
- **Shortcuts**:
  - `Ctrl + Shift + C` — copy the current line's result
  - `Ctrl + Shift + M` — copy the whole note as markdown with results inline
  - `Ctrl + Alt + M` — global toggle for the popup
  - `Esc` — hide the panel

## Setup

```bash
npm install
npm run build
npm start
```

For a single command that builds and launches in dev:

```bash
npm run dev
```

The tray icon should appear in the Windows system tray. Click it to toggle the
panel, right-click for the Settings / Quit menu.

> If you ever see `TypeError: Cannot read properties of undefined (reading 'isPackaged')`,
> your shell has `ELECTRON_RUN_AS_NODE=1` set (which forces Electron into a
> bare-Node mode). Unset it in that shell and re-run.

## Settings

Open the cog icon in the panel's title bar (or right-click the tray icon).
You can:

- Toggle auto-comma formatting
- Toggle in-editor suffix expansion
- Adjust result precision
- Add / edit / remove custom suffixes (symbol, multiplier, case sensitivity)

Settings (and your note text + last window position) persist to
`%APPDATA%/math-popup/settings.json`.

## Architecture

```
src/
  main/         Electron main process (tray, windows, IPC, settings store)
  renderer/     Popup + settings UI; evaluator and highlighter run here
  shared/       TypeScript types shared by main and renderer
build/          esbuild bundler + tray-icon generator
assets/         tray icon
```

Math is evaluated by [mathjs](https://mathjs.org/) on top of a small
preprocessor that handles line refs, percentages, basis points, custom
suffixes and number-comma stripping.

## Notes

- Reserved names you can't shadow: `pi`, `e`, `L1`, `L2`, ... and standard
  math functions (`sin`, `log`, `sqrt`, ...).
- Header lines (`# heading`) never produce a result.
- Bullet lines (`- 5 + 5`) evaluate the part after the marker.
- Plain text lines (no math) display in the editor without a result.
