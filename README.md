# Local LLM Word Translator

**English** | [简体中文](README.zh-CN.md)

A privacy-first, resumable Word (`.docx`) translation workflow powered by a local LLM through LM Studio.

It supports automatic terminology extraction, terminology auditing, resumable translation, Chinese-only output, bilingual output, and selective retranslation. Source documents, translations, terminology, logs, and progress data are excluded from Git by default.

## Features

- Calls a locally loaded model through the LM Studio REST API
- Translates long `.docx` documents in manageable batches
- Saves progress after every completed batch
- Extracts and audits terminology for consistent translations
- Produces Chinese-only and bilingual review documents
- Preserves images, tables, bibliography text, and footnote content where possible
- Retranslates only passages affected by a corrected term
- Recovers from a known LM Studio `peg-native` parser error by splitting the failed batch

## Privacy

The following directories are ignored by Git and must never be force-added:

```text
input/       source documents
output/      translated Word files
progress/    extracted text and translations
glossary/    document-specific terminology
logs/        prompts, responses, and runtime logs
```

`config.json` is also ignored. Only the sanitized `config.example.json` is public.

## Requirements

- Python 3.11 or 3.12
- LM Studio with a chat/instruct GGUF model loaded
- LM Studio local server running, normally at `http://127.0.0.1:1234`

## Local model setup

This project is model-agnostic. It does not download a model and does not require a particular model family. Any instruction-tuned GGUF model that can follow translation prompts and run through LM Studio's local API may be used. Bilingual instruct/chat models such as the Qwen family are possible candidates, but model choice should be based on available hardware and a representative translation test.

### 1. Choose a model

- Prefer an **Instruct** or **Chat** model rather than a Base model.
- Prefer models with reliable Chinese and English capability.
- For consumer hardware, `Q4_K_M` is a practical starting quantization; use a smaller model or quantization if memory is insufficient.
- A larger context window consumes more memory. Start with `8192` tokens and increase it only when document chunks require it.
- Test one representative section before translating an entire document.

### 2. Load it in LM Studio

1. Download or import a compatible `.gguf` file in LM Studio.
2. For an NVIDIA GPU, select the CUDA llama.cpp runtime.
3. Start with these load settings:

   ```text
   Context Length:              8192
   GPU Offload:                 Auto, or the highest value that loads reliably
   Max Concurrent Predictions:  1
   Flash Attention:             On, when supported
   Speculative Decoding:        Off
   ```

4. If loading fails, close memory-heavy applications first, then reduce GPU Offload or Context Length.
5. Thinking/reasoning is not needed for direct translation. Disable it when the selected chat template supports that option.

The exact layer count, batch size, KV-cache type, and offload value depend on the model, GPU memory, system RAM, and LM Studio runtime. Treat these values as a starting point rather than universal requirements.

### 3. Start the local API

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

## Installation

Install dependencies:

```bash
pip install -r requirements.txt
```

Create a local configuration:

```cmd
copy config.example.json config.json
```

Edit `config.json` and replace `your-loaded-model-id` with the identifier returned by LM Studio's `/v1/models` endpoint. The example file contains the recommended translation parameters.

## Quick start

Place a Word document in the ignored `input` directory:

```text
input/document.docx
```

Inspect its structure:

```cmd
python translate_docx.py inspect input\document.docx
```

Extract and audit terminology:

```cmd
python translate_docx.py terms input\document.docx
python translate_docx.py audit-terms input\document.docx
```

Test a few translation batches:

```cmd
python translate_docx.py translate input\document.docx --max-chunks 3
python translate_docx.py render input\document.docx --allow-partial
```

Continue the full translation and check its status:

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

## Correct a term and selectively retranslate

```cmd
python translate_docx.py retranslate-term input\document.docx "source term" "统一译名"
python translate_docx.py translate input\document.docx
python translate_docx.py render input\document.docx
```

Only completed passages containing that source term are invalidated. Unrelated translations remain intact.

## Related projects and alternatives

This project focuses on private, resumable English-to-Chinese translation of Word documents through LM Studio. Another tool may be a better fit when your main requirement is a different document format, stronger layout preservation, a graphical interface, or a traditional machine-translation service.

| If you need... | Consider | Why it may fit better |
| --- | --- | --- |
| Scientific PDF translation with formulas and page layout preserved | [PDFMathTranslate](https://github.com/PDFMathTranslate/PDFMathTranslate) | Designed specifically for bilingual PDF translation and supports multiple translation backends. |
| A GUI and support for both DOCX and PDF with Ollama or an OpenAI-compatible endpoint | [TransDocs](https://github.com/codefitz/TransDocs) | Provides a Flask web interface, language detection, proofreading, and multiple document elements. |
| Bilingual EPUB books | [EPUB Translator](https://github.com/oomol-lab/epub-translator) | Preserves EPUB structure and presents the original and translation together. |
| A self-hosted translation API without an LLM workflow | [LibreTranslate](https://github.com/LibreTranslate/LibreTranslate) | Provides a general-purpose, offline-capable machine-translation API powered by Argos Translate. |
| A lightweight offline Python library, CLI, or desktop translator | [Argos Translate](https://github.com/argosopentech/argos-translate) | Uses installable language packages and does not require LM Studio. |
| DOCX editing, redlining, or tracked-change workflows for LLM agents | [Adeu](https://github.com/dealfluence/adeu) | Focuses on safe DOCX-to-LLM round trips and projecting edits back as tracked changes. |

These links are references, not endorsements. Review each project's current documentation, license, model/provider configuration, and privacy behavior before processing sensitive documents. A local interface does not guarantee local processing if the selected translation backend is a cloud service.

## Known limitations

- This is a translation and review workflow, not a publication-grade DOCX typesetting engine.
- Chinese font, line spacing, and paragraph styles may require adjustment in Word.
- Complex inline formatting and footnote reference positions may shift after translation.
- Table-of-contents page numbers should be updated in Word after rendering.
- Text embedded inside images is not OCR-processed.
- The bilingual output is recommended for final quality review.

## License

[MIT](LICENSE)
