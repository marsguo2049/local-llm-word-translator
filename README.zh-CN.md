# Local LLM Word Translator

[English](README.md) | **简体中文**

这是一个通过 LM Studio 调用本地大语言模型、以隐私保护为优先并支持断点续传的 Word（`.docx`）英译中工作流。

项目支持自动提取与审校术语、长文档分批翻译、纯中文与中英对照输出，以及按术语选择性重译。原始文档、译文、术语表、日志和翻译进度默认不会进入 Git。

## 功能

- 通过 LM Studio REST API 调用本地加载的模型
- 将长篇 `.docx` 文档拆分为可控批次翻译
- 每完成一个批次便保存进度，支持中断后继续
- 自动提取和审校术语，提高全文译名一致性
- 输出纯中文文档和中英对照审校文档
- 尽可能保留图片、表格、参考文献文字和脚注内容
- 修改术语后，仅重新翻译受影响的段落
- 遇到已知的 LM Studio `peg-native` 解析错误时，自动缩小失败批次后重试

## 隐私保护

以下目录已被 Git 忽略，切勿使用强制参数加入版本控制：

```text
input/       原始文档
output/      翻译后的 Word 文档
progress/    提取文本、翻译结果和断点进度
glossary/    文档专用术语表
logs/        提示词、模型回复和运行日志
```

本机实际配置 `config.json` 同样不会进入 Git。仓库只公开经过清理的 `config.example.json`。

## 项目结构

```text
local-llm-word-translator/
├── translate_docx.py       主命令行程序
├── config.example.json     不含私人信息的配置模板
├── config.json             本机实际配置（被 Git 忽略）
├── requirements.txt        Python 依赖
├── input/                  私人的原始 DOCX 文档
├── output/                 生成的纯中文和中英对照 DOCX
├── progress/               断点进度和已缓存的翻译结果
├── glossary/               自动提取、审校及修正后的术语
├── logs/                   本地请求与运行日志
├── README.md               英文说明
├── README.zh-CN.md         简体中文说明
└── LICENSE                 MIT 许可证
```

公开仓库中的五个数据目录只保留一个 `.gitkeep` 占位文件。真实内容由用户在本地放入或由程序生成，并持续受到 `.gitignore` 保护。

`translate_docx.py` 集中了文档检查、术语处理、翻译、错误恢复、进度查询、选择性重译和 Word 输出命令。用户需要把 `config.example.json` 复制为 `config.json`，再在本地填写 API 地址、模型标识和翻译参数。

## 环境要求

- Python 3.11 或 3.12
- LM Studio，以及一个已经加载的 chat/instruct GGUF 模型
- LM Studio 本地服务器，通常运行于 `http://127.0.0.1:1234`

## 本地模型部署

本项目不会自动下载模型，也不绑定某个具体模型。只要模型能够在 LM Studio 中运行、遵循翻译指令，并通过本地 API 接收请求，就可以使用。Qwen 等支持中英文的 instruct/chat 模型可以作为候选，但应当根据自己的硬件条件和代表性段落的试译结果进行选择。

### 1. 选择模型

- 优先选择 **Instruct** 或 **Chat** 模型，不要选择未经指令微调的 Base 模型。
- 优先选择中英文能力可靠的模型。
- 消费级电脑可以从 `Q4_K_M` 量化开始；如果内存不足，再选择更小的模型或量化版本。
- 上下文越长，占用的内存通常越多。建议从 `8192` tokens 开始，只有分段确实需要时才提高。
- 在翻译全文之前，先选择一个有代表性的章节试译。

### 2. 在 LM Studio 中加载模型

1. 在 LM Studio 中下载或导入兼容的 `.gguf` 文件。
2. 使用 NVIDIA 显卡时，选择 CUDA llama.cpp 运行时。
3. 可以从以下加载参数开始：

   ```text
   Context Length（上下文长度）:          8192
   GPU Offload（显卡卸载）:               Auto，或能够稳定加载的最高值
   Max Concurrent Predictions（并发数）:  1
   Flash Attention:                       支持时开启
   Speculative Decoding（推测解码）:      关闭
   ```

4. 如果模型无法加载，先关闭占用大量内存的程序，然后降低 GPU Offload 或 Context Length。
5. 直接翻译通常不需要 Thinking/Reasoning。若模型的聊天模板支持该选项，建议关闭。

具体的层数、批次大小、KV Cache 类型和 GPU Offload 数值取决于模型、显存、系统内存和 LM Studio 运行时。上述参数只是通用起点，不是所有电脑都必须照搬的固定值。

### 3. 启动本地 API

打开 LM Studio 的 **Developer / Local Server** 页面：

1. 加载所选模型。
2. 在 `1234` 端口启动服务器。
3. 除非需要让另一台可信设备连接，否则保持 **Serve on Local Network** 关闭。
4. 当服务器仅监听 `127.0.0.1` 时，可以不启用身份验证。
5. 确认模型显示为 `READY`。

在 PowerShell 中验证服务器：

```powershell
Invoke-RestMethod -Uri "http://127.0.0.1:1234/v1/models" |
    ConvertTo-Json -Depth 5
```

## 安装项目

安装依赖：

```bash
pip install -r requirements.txt
```

创建本地配置：

```cmd
copy config.example.json config.json
```

编辑 `config.json`，把 `your-loaded-model-id` 替换为 LM Studio `/v1/models` 接口返回的模型 `id`。推荐的翻译参数已经写在示例配置中。

## 快速开始

把 Word 文档放入已被 Git 忽略的 `input` 目录：

```text
input/document.docx
```

检查文档结构：

```cmd
python translate_docx.py inspect input\document.docx
```

提取并审校术语：

```cmd
python translate_docx.py terms input\document.docx
python translate_docx.py audit-terms input\document.docx
```

先试译少量批次：

```cmd
python translate_docx.py translate input\document.docx --max-chunks 3
python translate_docx.py render input\document.docx --allow-partial
```

继续全文翻译并检查进度：

```cmd
python translate_docx.py translate input\document.docx
python translate_docx.py status input\document.docx
```

生成最终文档：

```cmd
python translate_docx.py render input\document.docx
```

输出文件：

```text
output/document_zh.docx
output/document_bilingual.docx
```

## 修改术语并选择性重译

```cmd
python translate_docx.py retranslate-term input\document.docx "source term" "统一译名"
python translate_docx.py translate input\document.docx
python translate_docx.py render input\document.docx
```

程序只会让包含该英文术语的已完成段落失效并重新翻译，不影响其他翻译结果。

## 相关项目与替代方案

本项目主要解决“通过 LM Studio 在本地进行可断点续传的 Word 英译中”这一类需求。如果你更重视其他文档格式、复杂排版还原、图形界面，或者希望使用传统机器翻译服务，下面的项目可能更合适。

| 你的主要需求 | 可以参考 | 更适合的原因 |
| --- | --- | --- |
| 翻译科学论文 PDF，并尽可能保留公式和页面排版 | [PDFMathTranslate](https://github.com/PDFMathTranslate/PDFMathTranslate) | 专门面向双语 PDF 翻译，并支持多种翻译后端。 |
| 需要图形界面，同时处理 DOCX、PDF，并连接 Ollama 或 OpenAI-compatible 接口 | [TransDocs](https://github.com/codefitz/TransDocs) | 提供 Flask Web 界面、语言检测、校对模式及多种文档元素处理。 |
| 制作保留原书结构的中英对照 EPUB | [EPUB Translator](https://github.com/oomol-lab/epub-translator) | 专门面向 EPUB，并将原文和译文组合为双语电子书。 |
| 需要自托管的通用翻译 API，但不需要大模型工作流 | [LibreTranslate](https://github.com/LibreTranslate/LibreTranslate) | 提供由 Argos Translate 驱动、可离线运行的通用机器翻译 API。 |
| 需要轻量的离线 Python 库、命令行或桌面翻译工具 | [Argos Translate](https://github.com/argosopentech/argos-translate) | 使用可安装的语言包，无需 LM Studio。 |
| 需要 DOCX 修订、红线比较或 LLM Agent 的修订追踪工作流 | [Adeu](https://github.com/dealfluence/adeu) | 侧重 DOCX 与 LLM 之间的安全往返，并可把修改投射为修订记录。 |

以上链接仅供比较和参考，不代表背书。处理敏感文档前，请自行核对各项目的最新文档、许可证、模型或服务商配置及隐私行为。界面在本地运行，不代表翻译一定在本地完成；如果选择云端翻译后端，文档内容仍可能被发送到外部服务。

## 已知限制

- 本项目是翻译和审校工作流，不是出版级 Word 排版引擎。
- 中文字体、行距和段落样式可能需要在 Word 中调整。
- 复杂的行内格式和脚注引用标记的位置可能在翻译后发生偏移。
- 生成文档后，应当在 Word 中重新更新目录页码。
- 项目不会对图片内的文字执行 OCR。
- 建议使用中英对照版本完成最终质量检查。

## 许可证

[MIT](LICENSE)
