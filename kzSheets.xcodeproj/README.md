# kzSheets

`kzSheets` is a SwiftUI prototype that brings local LLM-assisted spreadsheet analysis to iphone and Mac.

The current app targets mac and also runs on ios

The app can:
- Import `CSV` and `XLSX` files
- Build sheet context from headers and rows
- Run local MLX language models
- Let the model generate JavaScript in `<js>...</js>` blocks
- Execute that JavaScript against sheet data in-app

## Current Capabilities

- Local model loading with progress, speed, and ETA
- Multiple model options in picker
- Editable system prompt (via modal editor)
- JavaScript extraction from assistant responses
- Per-block `Run JS` buttons directly under visible JS code
- JavaScript runtime inputs:
  - `headers`: `[String]`
  - `allRows`: `[[String: String]]` (header-keyed row objects)
  - `sheetObjects`: `[[String: String]]` (same as `allRows`)
  - `rawRows`: `[[String]]` (index-based rows)

## Stack

- SwiftUI
- JavaScriptCore
- CoreXLSX
- MLX ecosystem (`MLXLLM`, `MLXLMCommon`)

## Models Configured in App

- `mlx-community/Phi-3.5-mini-instruct-4bit`
- `mlx-community/DeepSeek-R1-0528-Qwen3-8B-4bit-DWQ`
- `mlx-community/Qwen3-8B-4bit-DWQ-053125`
- `mlx-community/Josiefied-Qwen2.5-Coder-7B-Instruct-abliterated-v1`

## Setup

1. Open the project in Xcode.
2. Ensure Swift packages are resolved and linked to the app target:
   - `CoreXLSX`
   - `mlx-swift`
   - `mlx-swift-lm`
3. Build and run 
## Usage

1. Choose a model and tap `Load Model`.
2. Import a `CSV` or `XLSX` file.
3. Open `Prompt Settings` and adjust the system prompt if needed.
4. Ask for analysis in chat.
5. If the assistant returns `<js>...</js>`, review the shown code and tap `Run JS`.
6. Read JS output in the conversation.

## JavaScript Contract

- The assistant should return executable JavaScript inside `<js>...</js>`.
- JavaScript should analyze runtime data (`allRows`, `sheetObjects`, `rawRows`) and not hardcode dataset rows.
- Prefer returning a value or calling `emit(value)`.

Example:

```js
<js>
const djiCount = allRows.filter(r => (r.Symbol || "").toLowerCase() === "dji").length;
return { djiCount };
</js>
```

## Project Structure

- `kzSheets/kzSheets/ContentView.swift` - main UI + chat + model + JS execution
- `kzSheets/kzSheets/kzSheetsApp.swift` - app entry point
- `kzSheets/kzSheets/Assets.xcassets` - assets

## Notes

- This is a prototype focused on local analysis workflows.
- JS execution currently works in-memory on imported data and returns analysis results.
- Direct XLSX write-back is not implemented yet.
