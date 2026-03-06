//
//  ContentView.swift
//  kzSheets
//
//  Created by MARKETING AGROTECH on 3/3/26.
//

import SwiftUI
import Foundation
import UniformTypeIdentifiers
import JavaScriptCore
import CoreXLSX
import MLXLLM
import MLXLMCommon

private struct ChatMessage: Identifiable {
    let id = UUID()
    let role: String
    var text: String
}

private struct ModelOption: Identifiable {
    let id: String
    let label: String
}

struct ContentView: View {
    @State private var isImporterPresented = false
    @State private var selectedFileName: String = "No file selected"
    @State private var headers: [String] = []
    @State private var allRows: [[String]] = []
    @State private var sampleRows: [[String]] = []
    @State private var systemPrompt: String = """
You are a spreadsheet assistant. Analyze the sheet context and answer the user request.

If JavaScript execution would help, include code inside <js>...</js>.
The app exposes JavaScript variables `headers`, `allRows`, `sheetObjects`, and `rawRows`.
Do NOT hardcode spreadsheet rows, sample values, or copied table data inside JavaScript.
Always analyze using only `allRows`/`sheetObjects` at runtime.
Always put executable code only inside <js>...</js> and return a JSON-serializable value from that script.

JavaScript rules:
1) Never define inline arrays/objects containing spreadsheet row data.
2) Read from `allRows` or `sheetObjects` (objects keyed by column header). Use `rawRows` only if you need index-based arrays.
3) Write plain JavaScript only (no TypeScript types, no imports, no markdown, no backticks).
4) Return concise computed results only.
5) Do not output <think> or chain-of-thought.

Required response format:
a short answer plus exactly one <js>...</js> block containing executable JavaScript only.

Example valid block:
<js>const total = allRows.length; const unique = new Set(sheetObjects.map(r => r[headers[0]])); return { totalRows: total, uniqueFirstColumn: unique.size };</js>
"""
    @State private var userPrompt: String = "Hi im Kostas"
    @State private var fileStatusMessage: String = "Import a CSV or XLSX to begin."

    @State private var modelStatus: String = "Model not loaded"
    @State private var isModelLoading: Bool = false
    @State private var isGenerating: Bool = false
    @State private var downloadStartedAt: Date?
    @State private var lastProgressAt: Date?
    @State private var lastProgressFraction: Double = 0
    @State private var downloadProgressFraction: Double = 0
    @State private var smoothedUnitsPerSecond: Double = 0
    @State private var isRunningJavaScript: Bool = false
    @State private var runningJavaScriptMessageID: UUID?
    @State private var runningJavaScriptBlockIndex: Int?
    @State private var messages: [ChatMessage] = []
    @State private var selectedModelId: String = "mlx-community/Phi-3.5-mini-instruct-4bit"
    @State private var isSystemPromptEditorPresented: Bool = false
    @State private var isContextEditorPresented: Bool = false
    @State private var isContextPreviewPresented: Bool = false
    @State private var customContext: String = ""

    @State private var modelContainer: ModelContainer?
    private let modelOptions: [ModelOption] = [
        ModelOption(id: "mlx-community/Phi-3.5-mini-instruct-4bit", label: "Phi-3.5-mini-instruct-4bit"),
        ModelOption(id: "mlx-community/DeepSeek-R1-0528-Qwen3-8B-4bit-DWQ", label: "DeepSeek-R1-0528-Qwen3-8B-4bit-DWQ"),
        ModelOption(id: "mlx-community/Qwen3-8B-4bit-DWQ-053125", label: "Qwen3-8B-4bit-DWQ-053125"),
        ModelOption(id: "mlx-community/Josiefied-Qwen2.5-Coder-7B-Instruct-abliterated-v1", label: "Josiefied-Qwen2.5-Coder-7B-Instruct-abliterated-v1")
    ]
    var body: some View {
        GeometryReader { geometry in
            ZStack {
                LinearGradient(
                    colors: [
                        Color(red: 0.03, green: 0.03, blue: 0.05),
                        Color(red: 0.07, green: 0.07, blue: 0.10)
                    ],
                    startPoint: .topLeading,
                    endPoint: .bottomTrailing
                )
                .ignoresSafeArea()

                if geometry.size.width >= 900 {
                    HStack(spacing: 14) {
                        sidebarPanel
                            .frame(width: min(320, geometry.size.width * 0.30))
                        mainPanel
                    }
                    .padding(14)
                } else {
                    VStack(spacing: 12) {
                        compactTopBar
                        mainPanel
                    }
                    .padding(12)
                }
            }
        }
        .fileImporter(
            isPresented: $isImporterPresented,
            allowedContentTypes: [.commaSeparatedText, .spreadsheet],
            allowsMultipleSelection: false
        ) { result in
            handleImport(result)
        }
        .onChange(of: selectedModelId) { _, _ in
            modelContainer = nil
            modelStatus = "Model not loaded"
        }
        .sheet(isPresented: $isSystemPromptEditorPresented) {
            NavigationStack {
                TextEditor(text: $systemPrompt)
                    .font(.body.monospaced())
                    .foregroundColor(.white)
                    .scrollContentBackground(.hidden)
                    .background(Color.black.opacity(0.9))
                    .padding(12)
                    .navigationTitle("System Prompt")
                    .navigationBarTitleDisplayMode(.inline)
                    .toolbar {
                        ToolbarItem(placement: .topBarTrailing) {
                            Button("Done") {
                                isSystemPromptEditorPresented = false
                            }
                        }
                    }
            }
        }
        .sheet(isPresented: $isContextEditorPresented) {
            NavigationStack {
                TextEditor(text: $customContext)
                    .font(.body.monospaced())
                    .foregroundColor(.white)
                    .scrollContentBackground(.hidden)
                    .background(Color.black.opacity(0.9))
                    .padding(12)
                    .navigationTitle("Custom Context")
                    .navigationBarTitleDisplayMode(.inline)
                    .toolbar {
                        ToolbarItem(placement: .topBarTrailing) {
                            Button("Done") {
                                isContextEditorPresented = false
                }
            }
        }
        .sheet(isPresented: $isContextPreviewPresented) {
            NavigationStack {
                ScrollView {
                    VStack(alignment: .leading, spacing: 14) {
                        Text("File: \(selectedFileName)")
                            .font(.headline)
                            .foregroundColor(.white)
                        Text("Columns: \(headers.count)")
                            .font(.subheadline)
                            .foregroundColor(.white.opacity(0.9))
                        Text("Rows: \(allRows.count)")
                            .font(.subheadline)
                            .foregroundColor(.white.opacity(0.9))

                        VStack(alignment: .leading, spacing: 6) {
                            Text("Headers")
                                .font(.subheadline.weight(.semibold))
                                .foregroundColor(.white)
                            if headers.isEmpty {
                                Text("No headers loaded.")
                                    .foregroundColor(.white.opacity(0.7))
                            } else {
                                Text(headers.joined(separator: ", "))
                                    .font(.caption.monospaced())
                                    .foregroundColor(.white)
                                    .textSelection(.enabled)
                            }
                        }
                        .padding(10)
                        .background(Color.white.opacity(0.10), in: RoundedRectangle(cornerRadius: 10, style: .continuous))

                        VStack(alignment: .leading, spacing: 6) {
                            Text("Sample Rows")
                                .font(.subheadline.weight(.semibold))
                                .foregroundColor(.white)
                            if sampleRows.isEmpty {
                                Text("No sample rows loaded.")
                                    .foregroundColor(.white.opacity(0.7))
                            } else {
                                ForEach(sampleRows.indices, id: \.self) { index in
                                    Text(sampleRows[index].joined(separator: " | "))
                                        .font(.caption.monospaced())
                                        .foregroundColor(.white)
                                        .textSelection(.enabled)
                                        .frame(maxWidth: .infinity, alignment: .leading)
                                        .padding(.vertical, 2)
                                }
                            }
                        }
                        .padding(10)
                        .background(Color.white.opacity(0.10), in: RoundedRectangle(cornerRadius: 10, style: .continuous))
                    }
                    .padding(14)
                }
                .scrollContentBackground(.hidden)
                .background(Color.black.ignoresSafeArea())
                .navigationTitle("Extracted Context")
                .navigationBarTitleDisplayMode(.inline)
                .toolbar {
                    ToolbarItem(placement: .topBarTrailing) {
                        Button("Done") {
                            isContextPreviewPresented = false
                        }
                    }
                }
            }
        }
    }
        }
    }

    private var compactTopBar: some View {
        VStack(alignment: .leading, spacing: 6) {
            HStack(spacing: 10) {
                Menu {
                    Button("Import CSV/XLSX") { isImporterPresented = true }
                    Button(isModelLoading ? "Loading model..." : "Load Model") { loadModel() }
                        .disabled(isModelLoading || isGenerating)
                    Divider()
                    Button("Edit System Prompt") { isSystemPromptEditorPresented = true }
                    Button("Edit Context") { isContextEditorPresented = true }
                    Button("View Extracted Context") { isContextPreviewPresented = true }
                    Divider()
                    Button("Clear Chat") { messages.removeAll() }
                        .disabled(messages.isEmpty || isGenerating)
                } label: {
                    Image(systemName: "line.3.horizontal")
                        .font(.title3.weight(.semibold))
                        .foregroundColor(.white)
                        .frame(width: 32, height: 32)
                }
                Spacer()
                Text("kzSheets")
                    .font(.headline.weight(.semibold))
                    .foregroundColor(.white)
                Spacer()
                Circle()
                    .fill(LinearGradient(colors: [.green, .blue], startPoint: .topLeading, endPoint: .bottomTrailing))
                    .frame(width: 34, height: 34)
                    .overlay(Image(systemName: "person.fill").foregroundColor(.white))
            }
            VStack(alignment: .leading, spacing: 4) {
                Text(modelStatus)
                    .font(.caption)
                    .foregroundColor(.white.opacity(0.85))
                    .lineLimit(2)
                if isModelLoading {
                    ProgressView(value: downloadProgressFraction)
                        .tint(.blue)
                }
            }
            .padding(8)
            .background(Color.white.opacity(0.10), in: RoundedRectangle(cornerRadius: 10, style: .continuous))
        }
        .padding(.horizontal, 4)
    }

    private var sidebarPanel: some View {
        VStack(alignment: .leading, spacing: 14) {
            VStack(alignment: .leading, spacing: 2) {
                Text("kzSheets")
                    .font(.title3.weight(.semibold))
            }

            GroupBox("Model") {
                VStack(alignment: .leading, spacing: 8) {
                    Picker("Model", selection: $selectedModelId) {
                        ForEach(modelOptions) { option in
                            Text(option.label).tag(option.id)
                        }
                    }
                    .pickerStyle(.menu)
                    Button(isModelLoading ? "Loading..." : "Load Model") {
                        loadModel()
                    }
                    .buttonStyle(.borderedProminent)
                    .disabled(isModelLoading || isGenerating)
                    Text(modelStatus)
                        .font(.caption)
                        .foregroundColor(.black)
                    if isModelLoading {
                        ProgressView(value: downloadProgressFraction)
                            .tint(.blue)
                    }
                }
            }

            GroupBox("Sheet") {
                VStack(alignment: .leading, spacing: 8) {
                    Button("Import CSV/XLSX") {
                        isImporterPresented = true
                    }
                    .buttonStyle(.borderedProminent)
                    Text(selectedFileName)
                        .font(.footnote.weight(.medium))
                    Text(fileStatusMessage)
                        .font(.caption)
                        .foregroundStyle(.secondary)
                }
            }

            GroupBox("Context") {
                VStack(alignment: .leading, spacing: 6) {
                    Text("Columns: \(headers.count)")
                    Text("Rows: \(allRows.count)")
                    Text(headers.isEmpty ? "No headers loaded" : headers.joined(separator: ", "))
                        .lineLimit(4)
                        .foregroundStyle(.secondary)
                    Button("View Extracted Context") {
                        isContextPreviewPresented = true
                    }
                    .buttonStyle(.bordered)
                }
                .font(.caption)
            }

            GroupBox("Prompt Settings") {
                VStack(alignment: .leading, spacing: 8) {
                    Button("Edit System Prompt") {
                        isSystemPromptEditorPresented = true
                    }
                    .buttonStyle(.bordered)
                    Button("Edit Context") {
                        isContextEditorPresented = true
                    }
                    .buttonStyle(.bordered)
                    Text("Custom instructions are active.")
                        .font(.caption)
                        .foregroundStyle(.secondary)
                }
            }

            Spacer()
        }
        .padding(14)
        .background(.ultraThinMaterial, in: RoundedRectangle(cornerRadius: 18, style: .continuous))
    }

    private var mainPanel: some View {
        VStack(spacing: 10) {
            if messages.isEmpty {
                VStack(alignment: .leading, spacing: 16) {
                    Text("Start by loading a model and importing a sheet.")
                        .font(.title3.weight(.medium))
                        .foregroundColor(.white.opacity(0.9))
                    Spacer()
                }
                .padding(18)
                .frame(maxWidth: .infinity, maxHeight: .infinity, alignment: .topLeading)
                .background(Color.white.opacity(0.04), in: RoundedRectangle(cornerRadius: 18, style: .continuous))
            } else {
                ScrollView {
                    LazyVStack(alignment: .leading, spacing: 10) {
                        ForEach(messages) { message in
                            let scripts = message.role == "assistant" ? extractExecutableJavaScripts(in: message.text) : []
                            HStack {
                                if message.role == "assistant" {
                                    VStack(alignment: .leading, spacing: 6) {
                                        let assistantText = displayText(for: message)
                                        if !(scripts.isEmpty == false && assistantText == "Assistant generated JavaScript.") {
                                            Text(assistantText)
                                                .font(.subheadline)
                                                .foregroundColor(.white)
                                                .textSelection(.enabled)
                                        }
                                        if !scripts.isEmpty {
                                            ForEach(Array(scripts.enumerated()), id: \.offset) { index, script in
                                                VStack(alignment: .leading, spacing: 6) {
                                                    Text(script)
                                                        .font(.caption.monospaced())
                                                        .foregroundColor(.white)
                                                        .textSelection(.enabled)
                                                        .padding(10)
                                                        .frame(maxWidth: .infinity, alignment: .leading)
                                                        .background(Color.white.opacity(0.10), in: RoundedRectangle(cornerRadius: 10, style: .continuous))
                                                    Button(
                                                        runningJavaScriptMessageID == message.id &&
                                                        runningJavaScriptBlockIndex == index &&
                                                        isRunningJavaScript ? "Running..." : "Run JS \(index + 1)"
                                                    ) {
                                                        runJavaScript(script, sourceMessageID: message.id, sourceBlockIndex: index)
                                                    }
                                                    .buttonStyle(.bordered)
                                                    .controlSize(.small)
                                                    .disabled(isRunningJavaScript)
                                                }
                                            }
                                        }
                                    }
                                    .padding(12)
                                    .background(Color.white.opacity(0.08), in: RoundedRectangle(cornerRadius: 14, style: .continuous))
                                    Spacer(minLength: 24)
                                } else {
                                    Spacer(minLength: 24)
                                    Text(displayText(for: message))
                                        .font(.subheadline)
                                        .foregroundColor(.white)
                                        .textSelection(.enabled)
                                        .padding(12)
                                        .background(Color.blue.opacity(0.30), in: RoundedRectangle(cornerRadius: 14, style: .continuous))
                                }
                            }
                        }
                    }
                    .padding(12)
                }
                .background(Color.white.opacity(0.04), in: RoundedRectangle(cornerRadius: 18, style: .continuous))
            }

            VStack(spacing: 8) {
                TextField("Message kzSheets...", text: $userPrompt, axis: .vertical)
                    .textFieldStyle(.plain)
                    .foregroundColor(.white)
                    .padding(12)
                    .background(Color.white.opacity(0.10), in: RoundedRectangle(cornerRadius: 12, style: .continuous))
                HStack {
                    Button {
                        isImporterPresented = true
                    } label: {
                        Image(systemName: "plus")
                    }
                    .buttonStyle(.bordered)

                    Menu {
                        Picker("Model", selection: $selectedModelId) {
                            ForEach(modelOptions) { option in
                                Text(option.label).tag(option.id)
                            }
                        }
                    } label: {
                        Image(systemName: "slider.horizontal.3")
                    }
                    .buttonStyle(.bordered)

                    Button {
                        loadModel()
                    } label: {
                        Text(isModelLoading ? "Loading..." : "Load")
                    }
                    .buttonStyle(.bordered)
                    .disabled(isModelLoading || isGenerating)

                    Spacer()

                    Button("Clear") {
                        messages.removeAll()
                    }
                    .buttonStyle(.bordered)
                    .disabled(messages.isEmpty || isGenerating)

                    Button(isGenerating ? "Generating..." : "Send") {
                        sendMessage()
                    }
                    .buttonStyle(.borderedProminent)
                    .disabled(!canSendMessage)
                }
            }
            .padding(12)
            .background(Color.white.opacity(0.06), in: RoundedRectangle(cornerRadius: 18, style: .continuous))
        }
    }

    private var canSendMessage: Bool {
        !headers.isEmpty &&
        !isGenerating &&
        modelContainer != nil &&
        !isModelLoading &&
        !userPrompt.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty
    }

    private func buildContextBlock() -> String {
        var sections: [String] = []
        sections.append(systemPrompt)
        if !headers.isEmpty {
            sections.append("Headers: \(headers.joined(separator: ", "))")
        }
        if !allRows.isEmpty {
            sections.append("Row count: \(allRows.count)")
        }
        if !customContext.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty {
            sections.append("User Context:\n\(truncate(customContext, maxCharacters: isRunningOnIPhone ? 1500 : 3000))")
        }
        return sections.joined(separator: "\n\n")
    }

    private func buildConversationBlock() -> String {
        guard !messages.isEmpty else { return "" }
        let maxMessages = isRunningOnIPhone ? (isHeavyModelSelected ? 2 : 4) : 6
        let maxCharsPerMessage = isRunningOnIPhone ? (isHeavyModelSelected ? 400 : 700) : 1200
        let recentMessages = messages.suffix(maxMessages)
        return recentMessages.map { message in
            let limited = truncate(message.text, maxCharacters: maxCharsPerMessage)
            return "\(message.role): \(limited)"
        }.joined(separator: "\n")
    }

    private func loadModel() {
        Task { await loadModelImpl() }
    }

    private func sendMessage() {
        let trimmed = userPrompt.trimmingCharacters(in: .whitespacesAndNewlines)
        guard !trimmed.isEmpty else { return }
        let userMessage = ChatMessage(role: "user", text: trimmed)
        messages.append(userMessage)
        userPrompt = ""
        Task { await generateResponse(userMessage: trimmed) }
    }

    @MainActor
    private func loadModelImpl() async {
        if modelContainer != nil {
            modelStatus = "Model already loaded"
            return
        }
        guard !isModelLoading else { return }
        isModelLoading = true
        defer { isModelLoading = false }
        downloadStartedAt = Date()
        lastProgressAt = nil
        lastProgressFraction = 0
        downloadProgressFraction = 0
        smoothedUnitsPerSecond = 0
        modelStatus = "Loading model..."

        do {
            let configuration = ModelConfiguration(id: selectedModelId)
            let container = try await LLMModelFactory.shared.loadContainer(configuration: configuration) { progress in
                Task { @MainActor in
                    self.updateDownloadStatus(progress: progress)
                }
            }
            modelContainer = container
            downloadProgressFraction = 1
            if isRunningOnIPhone && isHeavyModelSelected {
                modelStatus = "Model loaded (Safe Mode: short responses to avoid memory kill)"
            } else {
                modelStatus = "Model loaded"
            }
        } catch {
            modelStatus = "Failed to load model: \(errorMessage(from: error))"
        }
    }

    @MainActor
    private func updateDownloadStatus(progress: Progress) {
        let now = Date()
        let fraction = max(0, min(progress.fractionCompleted, 1))
        downloadProgressFraction = fraction
        let percent = Int(fraction * 100)

        if let lastProgressAt {
            let dt = now.timeIntervalSince(lastProgressAt)
            if dt > 0 {
                let df = max(0, fraction - lastProgressFraction)
                let totalUnits = Double(max(progress.totalUnitCount, 0))
                let instantUnitsPerSecond = totalUnits > 0 ? (df * totalUnits / dt) : 0
                if instantUnitsPerSecond > 0 {
                    if smoothedUnitsPerSecond == 0 {
                        smoothedUnitsPerSecond = instantUnitsPerSecond
                    } else {
                        smoothedUnitsPerSecond = (0.25 * instantUnitsPerSecond) + (0.75 * smoothedUnitsPerSecond)
                    }
                }
            }
        }

        self.lastProgressAt = now
        self.lastProgressFraction = fraction

        let speedText = formatDownloadSpeed(unitsPerSecond: smoothedUnitsPerSecond, totalUnitCount: progress.totalUnitCount)
        let etaText = formatETA(fractionCompleted: fraction, unitsPerSecond: smoothedUnitsPerSecond, totalUnitCount: progress.totalUnitCount)
        modelStatus = "Downloading model... \(percent)%  \(speedText)  ETA \(etaText)"
    }

    private func formatDownloadSpeed(unitsPerSecond: Double, totalUnitCount: Int64) -> String {
        guard unitsPerSecond > 0 else { return "--/s" }
        if totalUnitCount > 1_000_000 {
            let mbps = unitsPerSecond / 1_000_000
            return String(format: "%.2f MB/s", mbps)
        }
        return String(format: "%.1f units/s", unitsPerSecond)
    }

    private func formatETA(fractionCompleted: Double, unitsPerSecond: Double, totalUnitCount: Int64) -> String {
        guard fractionCompleted < 1 else { return "0s" }
        let totalUnits = Double(max(totalUnitCount, 0))
        let etaSeconds: Double
        if totalUnits > 0, unitsPerSecond > 0 {
            let remainingUnits = (1 - fractionCompleted) * totalUnits
            etaSeconds = remainingUnits / unitsPerSecond
        } else if let downloadStartedAt {
            let elapsed = Date().timeIntervalSince(downloadStartedAt)
            if fractionCompleted > 0 {
                etaSeconds = max(0, elapsed * ((1 - fractionCompleted) / fractionCompleted))
            } else {
                return "--"
            }
        } else {
            return "--"
        }
        return formatDuration(seconds: etaSeconds)
    }

    private func formatDuration(seconds: Double) -> String {
        if !seconds.isFinite || seconds < 0 { return "--" }
        let total = Int(seconds.rounded())
        let h = total / 3600
        let m = (total % 3600) / 60
        let s = total % 60
        if h > 0 { return "\(h)h \(m)m" }
        if m > 0 { return "\(m)m \(s)s" }
        return "\(s)s"
    }

    @MainActor
    private func generateResponse(userMessage: String) async {
        guard let modelContainer else {
            messages.append(ChatMessage(role: "assistant", text: "Model not loaded."))
            return
        }
        guard !isGenerating else { return }

        isGenerating = true
        defer { isGenerating = false }
        modelStatus = "Generating..."

        let contextBlock = buildContextBlock()
        let conversationBlock = buildConversationBlock()
        let prompt = [
            contextBlock,
            "Conversation so far:\n\(conversationBlock)",
            "Now answer this user message:\n\(userMessage)"
        ].joined(separator: "\n\n")
        let boundedPrompt = truncate(
            prompt,
            maxCharacters: isRunningOnIPhone ? (isHeavyModelSelected ? 3500 : 5500) : 12000
        )

        let assistantMessage = ChatMessage(role: "assistant", text: "")
        messages.append(assistantMessage)
        let assistantIndex = messages.count - 1
        let maxTokens = recommendedMaxTokens()

        do {
            try await modelContainer.perform { context in
                let input = try await context.processor.prepare(input: UserInput(prompt: boundedPrompt))
                let parameters = GenerateParameters(
                    maxTokens: maxTokens,
                    temperature: 0.7,
                    topP: 0.9
                )
                let stream = try MLXLMCommon.generate(input: input, parameters: parameters, context: context)
                for await part in stream {
                    if let chunk = part.chunk {
                        await MainActor.run {
                            self.messages[assistantIndex].text += chunk
                        }
                    }
                }
            }
            if messages[assistantIndex].text.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty {
                messages[assistantIndex].text = "No output generated."
            }
            trimChatMemoryFootprint()
            modelStatus = "Model loaded"
        } catch {
            messages[assistantIndex].text = "Generation failed: \(errorMessage(from: error))"
            trimChatMemoryFootprint()
            modelStatus = "Generation failed"
        }
    }

    private func runJavaScript(_ script: String, sourceMessageID: UUID, sourceBlockIndex: Int) {
        guard !isRunningJavaScript else { return }
        isRunningJavaScript = true
        runningJavaScriptMessageID = sourceMessageID
        runningJavaScriptBlockIndex = sourceBlockIndex
        defer {
            isRunningJavaScript = false
            runningJavaScriptMessageID = nil
            runningJavaScriptBlockIndex = nil
        }

        let context = JSContext()
        var javaScriptError: String?
        var javaScriptLogs: [String] = []
        context?.exceptionHandler = { _, exception in
            javaScriptError = exception?.toString()
        }
        let logBlock: @convention(block) (JSValue?) -> Void = { value in
            if let value {
                javaScriptLogs.append(value.toString())
            }
        }
        context?.setObject(logBlock, forKeyedSubscript: "_kzLog" as NSString)

        let objects = buildSheetObjects(from: allRows)
        context?.setObject(headers, forKeyedSubscript: "headers" as NSString)
        context?.setObject(allRows, forKeyedSubscript: "rawRows" as NSString)
        context?.setObject(objects, forKeyedSubscript: "allRows" as NSString)
        context?.setObject(objects, forKeyedSubscript: "sheetObjects" as NSString)
        _ = context?.evaluateScript(
            """
            var __kzResult = undefined;
            function emit(value) { __kzResult = value; return value; }
            var console = { log: function(v) { _kzLog(v); } };
            """
        )

        let wrappedScript = """
        (function () {
        \(script)
        })()
        """

        let resultValue = context?.evaluateScript(wrappedScript)

        if let javaScriptError {
            messages.append(ChatMessage(role: "assistant", text: "JavaScript error: \(javaScriptError)"))
            trimChatMemoryFootprint()
            return
        }

        if let resultValue, !resultValue.isUndefined {
            messages.append(ChatMessage(role: "assistant", text: "JavaScript result: \(stringifyJavaScriptValue(resultValue))"))
            trimChatMemoryFootprint()
            return
        }

        if let emittedValue = context?.objectForKeyedSubscript("__kzResult"), !emittedValue.isUndefined {
            messages.append(ChatMessage(role: "assistant", text: "JavaScript result: \(stringifyJavaScriptValue(emittedValue))"))
            trimChatMemoryFootprint()
            return
        }

        if !javaScriptLogs.isEmpty {
            messages.append(ChatMessage(role: "assistant", text: "JavaScript logs: \(javaScriptLogs.joined(separator: " | "))"))
            trimChatMemoryFootprint()
            return
        }

        messages.append(ChatMessage(role: "assistant", text: "JavaScript executed with no output. Return a value or call emit(value)."))
        trimChatMemoryFootprint()
    }

    private func handleImport(_ result: Result<[URL], Error>) {
        switch result {
        case .success(let urls):
            guard let url = urls.first else { return }
            selectedFileName = url.lastPathComponent
            fileStatusMessage = "Reading file..."

            guard url.startAccessingSecurityScopedResource() else {
                fileStatusMessage = "Could not access the file."
                return
            }
            defer { url.stopAccessingSecurityScopedResource() }

            let fileExtension = url.pathExtension.lowercased()
            if fileExtension == "csv" {
                readCSV(from: url)
            } else if fileExtension == "xlsx" {
                readXLSX(from: url)
            } else {
                headers = []
                allRows = []
                sampleRows = []
                fileStatusMessage = "Unsupported file type. Please choose a CSV or XLSX."
            }
        case .failure(let error):
            fileStatusMessage = "Import failed: \(errorMessage(from: error))"
        }
    }

    private func readCSV(from url: URL) {
        do {
            let content: String
            if let utf8 = try? String(contentsOf: url, encoding: .utf8) {
                content = utf8
            } else if let utf16 = try? String(contentsOf: url, encoding: .utf16) {
                content = utf16
            } else if let utf16LE = try? String(contentsOf: url, encoding: .utf16LittleEndian) {
                content = utf16LE
            } else if let utf16BE = try? String(contentsOf: url, encoding: .utf16BigEndian) {
                content = utf16BE
            } else if let latin1 = try? String(contentsOf: url, encoding: .isoLatin1) {
                content = latin1
            } else {
                throw CocoaError(.fileReadInapplicableStringEncoding)
            }
            let rows = parseCSV(normalizeCSVText(content))
            guard let firstRow = rows.first else {
                headers = []
                allRows = []
                sampleRows = []
                fileStatusMessage = "CSV is empty."
                return
            }
            headers = firstRow
            allRows = Array(rows.dropFirst())
            sampleRows = Array(allRows.prefix(5))
            fileStatusMessage = "Loaded \(headers.count) columns and \(allRows.count) rows."
        } catch {
            headers = []
            allRows = []
            sampleRows = []
            fileStatusMessage = "Failed to read CSV: \(errorMessage(from: error))"
        }
    }

    private func normalizeCSVText(_ content: String) -> String {
        var normalized = content
            .replacingOccurrences(of: "\u{0000}", with: "")
            .replacingOccurrences(of: "\u{2028}", with: "\n")
            .replacingOccurrences(of: "\u{2029}", with: "\n")

        // Some sources store line breaks as literal escape sequences (\n / \r\n).
        // If no real newlines exist, decode escaped newlines so rows split correctly.
        if !normalized.contains("\n") && !normalized.contains("\r") {
            normalized = normalized
                .replacingOccurrences(of: "\\r\\n", with: "\n")
                .replacingOccurrences(of: "\\n", with: "\n")
                .replacingOccurrences(of: "\\r", with: "\n")
        }
        return normalized
    }

    private func readXLSX(from url: URL) {
        do {
            guard let file = XLSXFile(filepath: url.path) else {
                headers = []
                allRows = []
                sampleRows = []
                fileStatusMessage = "Could not open XLSX file."
                return
            }

            let workbooks = try file.parseWorkbooks()
            guard let workbook = workbooks.first else {
                headers = []
                allRows = []
                sampleRows = []
                fileStatusMessage = "No workbook found in XLSX."
                return
            }

            let worksheetInfo = try file.parseWorksheetPathsAndNames(workbook: workbook)
            guard let firstWorksheet = worksheetInfo.first else {
                headers = []
                allRows = []
                sampleRows = []
                fileStatusMessage = "No worksheets found in XLSX."
                return
            }

            let worksheet = try file.parseWorksheet(at: firstWorksheet.path)
            let sharedStrings = try? file.parseSharedStrings()
            let rows = worksheet.data?.rows ?? []
            let parsedRows = rows.map { row in
                row.cells.map { cellString($0, sharedStrings: sharedStrings) }
            }

            guard let firstRow = parsedRows.first else {
                headers = []
                allRows = []
                sampleRows = []
                fileStatusMessage = "XLSX is empty."
                return
            }

            headers = firstRow
            allRows = Array(parsedRows.dropFirst())
            sampleRows = Array(allRows.prefix(5))
            fileStatusMessage = "Loaded \(headers.count) columns and \(allRows.count) rows."
        } catch {
            headers = []
            allRows = []
            sampleRows = []
            fileStatusMessage = "Failed to read XLSX: \(errorMessage(from: error))"
        }
    }

    private func cellString(_ cell: Cell, sharedStrings: SharedStrings?) -> String {
        if let sharedStrings, let value = cell.stringValue(sharedStrings) {
            return value
        }
        return cell.value ?? ""
    }

    private func parseCSV(_ content: String) -> [[String]] {
        let delimiter = detectCSVDelimiter(content)
        var rows: [[String]] = []
        var currentRow: [String] = []
        var currentField = ""
        var isInsideQuotes = false
        var index = content.startIndex

        while index < content.endIndex {
            let character = content[index]

            if character == "\"" {
                let nextIndex = content.index(after: index)
                if isInsideQuotes, nextIndex < content.endIndex, content[nextIndex] == "\"" {
                    // Escaped quote inside quoted field.
                    currentField.append("\"")
                    index = content.index(after: nextIndex)
                    continue
                } else {
                    // Enter/exit quoted field.
                    isInsideQuotes.toggle()
                    index = content.index(after: index)
                    continue
                }
            }

            if character == delimiter && !isInsideQuotes {
                currentRow.append(currentField.trimmingCharacters(in: .whitespacesAndNewlines))
                currentField = ""
                index = content.index(after: index)
                continue
            }

            if (character == "\n" || character == "\r") && !isInsideQuotes {
                // Handle CRLF as a single newline.
                if character == "\r" {
                    let nextIndex = content.index(after: index)
                    if nextIndex < content.endIndex, content[nextIndex] == "\n" {
                        index = nextIndex
                    }
                }
                if !currentField.isEmpty || !currentRow.isEmpty {
                    currentRow.append(currentField.trimmingCharacters(in: .whitespacesAndNewlines))
                    rows.append(currentRow)
                    currentRow = []
                    currentField = ""
                }
                index = content.index(after: index)
                continue
            }

            currentField.append(character)
            index = content.index(after: index)
        }

        if !currentField.isEmpty || !currentRow.isEmpty {
            currentRow.append(currentField.trimmingCharacters(in: .whitespacesAndNewlines))
            rows.append(currentRow)
        }

        return rows
    }

    private func detectCSVDelimiter(_ content: String) -> Character {
        let candidates: [Character] = [",", ";", "\t"]
        var counts: [Character: Int] = [",": 0, ";": 0, "\t": 0]
        var inQuotes = false

        for character in content {
            if character == "\"" {
                inQuotes.toggle()
                continue
            }
            if character == "\n" || character == "\r" {
                if character == "\n" || character == "\r" {
                    break
                }
            }
            if !inQuotes, candidates.contains(character) {
                counts[character, default: 0] += 1
            }
        }

        let best = counts.max { lhs, rhs in lhs.value < rhs.value }
        return (best?.value ?? 0) > 0 ? (best?.key ?? ",") : ","
    }

    private func extractExecutableJavaScripts(in text: String) -> [String] {
        let normalized = normalizeJSTags(in: text)
        let tagged = extractAllMatches(
            pattern: "<js>([\\s\\S]*?)</js>",
            in: normalized
        ).map(stripCodeFenceIfNeeded)
        if !tagged.isEmpty {
            return tagged
        }

        let fenced = extractAllMatches(
            pattern: "```(?:javascript|js)?\\s*([\\s\\S]*?)```",
            in: normalized
        )
        return fenced
    }

    private func displayText(for message: ChatMessage) -> String {
        guard message.role == "assistant" else { return message.text }
        let normalized = normalizeJSTags(in: message.text)
        let withoutTagged = removeMatches(pattern: "<js>[\\s\\S]*?</js>", in: normalized)
        let cleaned = removeMatches(pattern: "```(?:javascript|js)?\\s*[\\s\\S]*?```", in: withoutTagged)
            .trimmingCharacters(in: .whitespacesAndNewlines)
        return cleaned.isEmpty ? "Assistant generated JavaScript." : cleaned
    }

    private func buildSheetObjects(from rows: [[String]]) -> [[String: String]] {
        rows.map { row in
            var object: [String: String] = [:]
            for index in headers.indices {
                let key = headers[index]
                let value = index < row.count ? row[index] : ""
                object[key] = value
            }
            return object
        }
    }

    private func normalizeJSTags(in text: String) -> String {
        text
            .replacingOccurrences(of: "&lt;js&gt;", with: "<js>", options: [.caseInsensitive])
            .replacingOccurrences(of: "&lt;/js&gt;", with: "</js>", options: [.caseInsensitive])
    }

    private func extractFirstMatch(pattern: String, in text: String) -> String? {
        guard let regex = try? NSRegularExpression(pattern: pattern, options: [.caseInsensitive]) else {
            return nil
        }
        let range = NSRange(text.startIndex..., in: text)
        guard let match = regex.firstMatch(in: text, options: [], range: range), match.numberOfRanges > 1 else {
            return nil
        }
        guard let matchedRange = Range(match.range(at: 1), in: text) else { return nil }
        let value = String(text[matchedRange]).trimmingCharacters(in: .whitespacesAndNewlines)
        return value.isEmpty ? nil : value
    }

    private func extractAllMatches(pattern: String, in text: String) -> [String] {
        guard let regex = try? NSRegularExpression(pattern: pattern, options: [.caseInsensitive]) else {
            return []
        }
        let range = NSRange(text.startIndex..., in: text)
        return regex.matches(in: text, options: [], range: range).compactMap { match in
            guard match.numberOfRanges > 1, let matchedRange = Range(match.range(at: 1), in: text) else {
                return nil
            }
            let value = String(text[matchedRange]).trimmingCharacters(in: .whitespacesAndNewlines)
            return value.isEmpty ? nil : value
        }
    }

    private func removeMatches(pattern: String, in text: String) -> String {
        guard let regex = try? NSRegularExpression(pattern: pattern, options: [.caseInsensitive]) else {
            return text
        }
        let range = NSRange(text.startIndex..., in: text)
        return regex.stringByReplacingMatches(in: text, options: [], range: range, withTemplate: "")
    }

    private func stripCodeFenceIfNeeded(_ code: String) -> String {
        if let fenced = extractFirstMatch(pattern: "```(?:javascript|js)?\\s*([\\s\\S]*?)```", in: code) {
            return fenced
        }
        return code
    }

    private func errorMessage(from error: Error) -> String {
        let message = (error as NSError).localizedDescription
        return message.isEmpty ? "Unknown error" : message
    }

    private func stringifyJavaScriptValue(_ value: JSValue) -> String {
        if let object = value.toObject() {
            if JSONSerialization.isValidJSONObject(object),
               let data = try? JSONSerialization.data(withJSONObject: object, options: [.prettyPrinted]),
               let json = String(data: data, encoding: .utf8) {
                return json
            }
            return String(describing: object)
        }
        if let string = value.toString() {
            return string
        }
        return "Unprintable value"
    }

    private func recommendedMaxTokens() -> Int {
        2000
    }

    private var isHeavyModelSelected: Bool {
        selectedModelId.contains("8B") || selectedModelId.contains("7B")
    }

    private var isRunningOnIPhone: Bool {
        #if os(iOS)
        return true
        #else
        return false
        #endif
    }

    private func truncate(_ text: String, maxCharacters: Int) -> String {
        guard text.count > maxCharacters else { return text }
        let end = text.index(text.startIndex, offsetBy: maxCharacters)
        return String(text[..<end]) + " …[truncated]"
    }

    private func trimChatMemoryFootprint() {
        let maxStoredMessages = isRunningOnIPhone ? (isHeavyModelSelected ? 4 : 10) : 30
        if messages.count > maxStoredMessages {
            messages = Array(messages.suffix(maxStoredMessages))
        }

        let maxMessageChars = isRunningOnIPhone ? (isHeavyModelSelected ? 900 : 1600) : 4000
        messages = messages.map { message in
            var trimmed = message
            trimmed.text = truncate(trimmed.text, maxCharacters: maxMessageChars)
            return trimmed
        }
    }
}

#Preview {
    ContentView()
}
