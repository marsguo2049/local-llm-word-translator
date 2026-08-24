# Local LLM Word Translator

A privacy-first, resumable Word (`.docx`) translation workflow powered by a local LLM through LM Studio.

基于 LM Studio 本地大模型的隐私优先 Word 英译中工具，支持自动术语提取、断点续传、中英对照输出和选择性重译。原始文档、译文、术语表、日志与进度默认不会进入 Git。

## Features / 功能

- Calls a locally loaded model through the LM Studio REST API
- Translates long `.docx` documents in manageable batches
- Saves progress after every completed batch
- Extracts and audits terminology for consistent translations
- Produces Chinese-only and bilingual review documents
- Preserves images, tables, bibliography text, and footnote content where possible
- Retranslates only passages affected by a corrected term
- Recovers from a known LM Studio `peg-native` parser error by splitting the failed batch

## Privacy / 隐私

The following directories are ignored by Git and must never be force-added:

```text
input/       source documents
output/      translated Word files
progress/    extracted text and translations
glossary/    document-specific terminology
logs/        prompts, responses, and runtime logs
```

`config.json` is also ignored. Only the sanitized `config.example.json` is public.

## Requirements / 环境要求

- Python 3.11 or 3.12
- LM Studio with a chat/instruct GGUF model loaded
- LM Studio local server running, normally at `http://127.0.0.1:1234`

## Local model setup / 本地模型部署

This project is model-agnostic. It does not download a model and does not require a particular model family. Any instruction-tuned GGUF model that can follow translation prompts and run through LM Studio's local API may be used.

本项目不会自动下载模型，也不绑定某个具体模型。选择能够在 LM Studio 中运行、适合指令对话和翻译任务的 GGUF 模型即可。Qwen 等支持中英文的 instruct/chat 模型可以作为候选，但请根据自己的硬件和实际试译结果选择。

### 1. Choose a model / 选择模型

- Prefer an **Instruct** or **Chat** model rather than a Base model.
- Prefer models with reliable Chinese and English capability.
- For consumer hardware, `Q4_K_M` is a practical starting quantization; use a smaller model or quantization if memory is insufficient.
- A larger context window consumes more memory. Start with `8192` tokens and increase it only when the document chunks require it.
- Test one representative section before translating the entire document.

### 2. Load it in LM Studio / 在 LM Studio 中加载

1. Download or import a compatible `.gguf` file in LM Studio.
2. For an NVIDIA GPU, select the CUDA llama.cpp runtime.
3. Start with the following load settings:

   ```text
   Context Length:       8192
   GPU Offload:          Auto, or the highest value that loads reliably
   Max Concurrent Predictions: 1
   Flash Attention:      On, when supported
   Speculative Decoding: Off
   ```

4. If loading fails, first close memory-heavy applications, then reduce GPU Offload or Context Length. Do not force a configuration that causes Windows to use excessive swap memory.
5. Thinking/reasoning is not needed for direct translation. Disable it when the selected chat template supports that option.

The exact layer count, batch size, KV-cache type, and offload value depend on the model, GPU memory, system RAM, and LM Studio runtime. Treat the values above as a safe starting point rather than universal requirements.

### 3. Start the local API / 启动本地 API

Open LM Studio's **Developer / Local Server** page and:

1. Load the selected model.
2. Start the server on port `1234`.
3. Keep **Serve on Local Network** disabled unless another trusted device must connect.
4. Authentication may remain disabled when the server is restricted to `127.0.0.1`.
5. Confirm that the model is marked `READY`.

Verify the server in PowerShell:

```powershell
Invoke-RestMethod -Uri "http://127.0.0.1:1234/v1/models" |
    ConvertTo-Json -Depth 5
```

Copy the returned `id` into the local `config.json` file:

```json
{
  "server_url": "http://127.0.0.1:1234",
  "model": "your-loaded-model-id"
}
```

The complete set of recommended translation parameters is already provided in `config.example.json`.

Install dependencies:

```bash
pip install -r requirements.txt
```

Create a local configuration:

```cmd
copy config.example.json config.json
```

Then edit `config.json` and replace `your-loaded-model-id` with the identifier returned by LM Studio's `/v1/models` endpoint.

## Quick start / 快速开始

Place a Word document in the ignored `input` directory, for example:

```text
input/document.docx
```

Inspect its structure:

```cmd
python translate_docx.py inspect input\document.docx
```

Extract terminology:

```cmd
python translate_docx.py terms input\document.docx
```

Audit the high-impact terminology set using a Pareto-style pass:

```cmd
python translate_docx.py audit-terms input\document.docx
```

Test a few translation batches:

```cmd
python translate_docx.py translate input\document.docx --max-chunks 3
python translate_docx.py render input\document.docx --allow-partial
```

Continue the full translation:

```cmd
python translate_docx.py translate input\document.docx
python translate_docx.py status input\document.docx
```

Render the final files:

```cmd
python translate_docx.py render input\document.docx
```

Outputs:

```text
output/document_zh.docx
output/document_bilingual.docx
```

## Correct a term / 修改术语并选择性重译

Example:

```cmd
python translate_docx.py retranslate-term input\document.docx "source term" "统一译名"
python translate_docx.py translate input\document.docx
python translate_docx.py render input\document.docx
```

Only completed passages containing that source term are invalidated. Unrelated translations remain intact.

## Known limitations / 已知边界

- This is a translation and review workflow, not a publication-grade DOCX typesetting engine.
- Chinese font, line spacing, and paragraph styles may require adjustment in Word.
- Complex inline formatting and footnote reference positions may shift after translation.
- Table-of-contents page numbers should be updated in Word after rendering.
- Text embedded inside images is not OCR-processed.
- The bilingual output is recommended for final quality review.

## License

MIT
