# PPT-Translator | PPT 翻译专家

[English](#english) | [中文](#chinese)

<a name="english"></a>
## English

A professional-grade PowerPoint translation tool powered by Large Language Models (DeepSeek, OpenAI, Gemini). It delivers high-fidelity translations while ensuring presentation stability.

### 🚀 Key Features

- **High Fidelity**: Preserves text-level formatting including **bold**, *italic*, and font colors.
- **Context-Aware Batching**: Translates all content on a single slide in one request to maintain semantic consistency.
- **Recursive Extraction**: Deeply extracts text from nested **Group Shapes** and complex **Tables**.
- **Stability-First Rendering**: Deliberately ignores geometry modifications (width/height) and fragile paragraph spacing to prevent PowerPoint "File Repair" errors.
- **ID-Based Mapping**: Uses unique ID anchors to ensure no text block is lost or misaligned.
- **Persistent Caching**: Automatically saves translations to `translation_cache.json` to reduce API costs and support resuming tasks.

### 🛠️ Installation

#### As a Standalone Tool
1. **Clone the repository**:
   ```bash
   git clone https://github.com/Thetail001/ppt-translator-skill.git
   cd ppt-translator-skill/scripts
   ```

2. **Set up Virtual Environment**:
   ```bash
   python -m venv .venv
   # Windows
   .venv\Scripts\activate
   # Linux/Mac
   source .venv/bin/activate
   
   pip install -r requirements.txt
   ```

3. **Configure API Keys**:
   - Create a `.env` file in the `scripts/` directory.
   - Add your keys: `DEEPSEEK_API_KEY=your_key_here`.

#### As a Gemini CLI Skill
If you are using the [Gemini CLI](https://github.com/google/gemini-cli), you can install this repository as a skill to enable AI-powered PPT translation directly in your chat:
```bash
gemini skills install https://github.com/Thetail001/ppt-translator-skill.git
```

### 📖 Usage

**Standard Translation**:
```bash
python main.py "path/to/file.pptx" --provider deepseek --source-lang en --target-lang zh
```

**Global Color Modification**:
If you need to unify font colors after translation (e.g., set all to white):
```bash
python change_color.py "path/to/translated.pptx" FFFFFF
```

### ⚙️ Parameters

| Parameter | Description |
| :--- | :--- |
| `input_path` | Target .pptx file path |
| `--provider` | Translation provider (Default: `deepseek`) |
| `--source-lang`| Source language (Default: `en`) |
| `--target-lang`| Target language (Default: `zh`) |
| `--max-workers`| Number of slides to process in parallel (Default: 4) |

### ⚠️ Limitations & Roadmap

Currently, to ensure the maximum stability of the generated `.pptx` files and avoid "Repair Needed" errors, the following features are temporarily disabled:
- **Geometry Modifications**: Automatic adjustment of shape width, height, or position.
- **Paragraph Spacing**: Modification of line spacing (`spcPct`).

**Future Plans**:
- [ ] **Smart Auto-scaling**: Dynamically adjust font size based on text length.
- [ ] **Stability Fixes**: Investigate safer ways to re-enable geometry and spacing modifications without corrupting the XML structure.
- [ ] **Table Enhancements**: Better support for complex merged cells.

### 🙏 Acknowledgments

This project is modified from the original work by [tristan-mcinnis/PPT-Translator-Formatting-Intact-with-LLMs](https://github.com/tristan-mcinnis/PPT-Translator-Formatting-Intact-with-LLMs). I have made some adjustments and improvements to handle nested group shapes and improve stability based on my practical use cases.

---

<a name="chinese"></a>
## 中文

这是一款基于大语言模型（DeepSeek、OpenAI、Gemini 等）开发的专业级 PowerPoint 翻译工具。它在提供高保真翻译的同时，通过一系列安全策略确保生成的 PPT 文件稳定、不损坏。

### 🚀 核心特性

- **高保真还原**：完整保留原有的**加粗**、*斜体*以及字体颜色等样式。
- **上下文感知**：采用“单页批处理”策略，将整页文字统一发送给 AI，确保翻译语境连贯。
- **深度提取**：支持递归遍历，能够找回隐藏在**组合形状 (Group)** 和复杂**表格**内部的文字。
- **稳定性优先**：针对 `python-pptx` 的特性，主动绕过易导致文件损坏的几何尺寸（宽/高）和行间距修改，根治“需要修复”的报错。
- **ID 锚点匹配**：通过唯一的 ID 锚点回填翻译结果，确保文字块不丢失、不位移。
- **智能缓存**：自动保存翻译结果到本地缓存，节省 API 成本并支持断点续传。

### 🛠️ 安装步骤

#### 作为独立工具使用
1. **克隆仓库**:
   ```bash
   git clone https://github.com/Thetail001/ppt-translator-skill.git
   cd ppt-translator-skill/scripts
   ```

2. **设置虚拟环境**:
   ```bash
   python -m venv .venv
   # Windows 系统
   .venv\Scripts\activate
   # Linux/Mac 系统
   source .venv/bin/activate
   
   pip install -r requirements.txt
   ```

3. **配置 API 密钥**:
   - 在 `scripts/` 目录下创建 `.env` 文件。
   - 填写您的密钥：`DEEPSEEK_API_KEY=您的密钥`。

#### 作为 Gemini CLI 技能安装
如果您正在使用 [Gemini CLI](https://github.com/google/gemini-cli)，可以直接将此仓库安装为技能，从而在对话中直接调用 PPT 翻译功能：
```bash
gemini skills install https://github.com/Thetail001/ppt-translator-skill.git
```

### 📖 使用说明

**标准翻译任务**:
```bash
python main.py "PPT文件路径.pptx" --provider deepseek --source-lang en --target-lang zh
```

**全局修改字体颜色**:
翻译完成后，如果由于背景原因需要统一修改字体颜色（例如全部设为白色）：
```bash
python change_color.py "翻译后的PPT路径.pptx" FFFFFF
```

### ⚙️ 参数详解 | Parameters

| 参数 (Parameter) | 说明 (Description) |
| :--- | :--- |
| `input_path` | PPT 文件路径 (Target .pptx file path) |
| `--provider` | 翻译服务商 (Default: `deepseek`) |
| `--source-lang`| 源语言 (Default: `en`) |
| `--target-lang`| 目标语言 (Default: `zh`) |
| `--max-workers`| 并行处理的幻灯片数量 (Default: 4) |

### ⚠️ 局限性与后续规划

目前为了确保生成的 PPT 文件具有最高的稳定性，避免出现“需要修复”的报错，我们暂时去除了以下功能：
- **几何属性修改**：自动调整形状的宽度、高度或位置。
- **行间距调整**：修改段落的行间距属性 (`line_spacing`)。

**后续计划**：
- [ ] **智能字体缩放**：根据翻译后的文字长度自动微调字体大小。
- [ ] **稳定性修复**：探索更安全的方式来重新开启几何属性和间距修改，同时不破坏 PPT 的 XML 结构。
- [ ] **表格增强**：优化对复杂合并单元格的处理逻辑。

### 🙏 致谢

本项目修改自 [tristan-mcinnis/PPT-Translator-Formatting-Intact-with-LLMs](https://github.com/tristan-mcinnis/PPT-Translator-Formatting-Intact-with-LLMs)。针对实际使用场景，我进行了一些适配和改进，包括递归处理组合形状、改进回填逻辑以及加入防止 PPT 文件损坏的安全策略。

## 📄 License

MIT License.
