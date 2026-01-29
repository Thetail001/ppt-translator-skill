---
name: ppt-translator
description: Translate PowerPoint (.pptx) files using a high-fidelity, context-aware batching strategy. Preserves formatting (bold, italic, colors) while ensuring file stability by avoiding geometry modifications. Supports DeepSeek (default), OpenAI, Gemini, and more.
allowed-tools:
  - run_shell_command
  - read_file
  - write_file
  - list_directory
---

# PowerPoint Translation Skill (High-Fidelity)

This skill provides a robust, professional-grade pipeline for translating PPTX files while maintaining complex layouts and internal formatting.

## Usage

When the user requests a translation (e.g., via `/ppt-translator <path> [options]`), follow these automated steps:

1. **Environment Initialization**:
   - Verify `./scripts/.venv` exists. If not, run `python -m venv .venv` and `pip install -r requirements.txt`.
2. **API Check**:
   - Ensure `DEEPSEEK_API_KEY` is present in `./scripts/.env`. If missing, ask the user for the key.
3. **Execution Logic**:
   - Invoke the translation via the virtual environment's python:
     ```bash
     cd scripts
     ./.venv/Scripts/python main.py $ARGUMENTS
     ```
   - **Default Arguments**: If not specified, default to `--provider deepseek --source-lang en --target-lang zh`.

## Technical Principles (The "True Flow")

Claude should understand the underlying logic to troubleshoot issues:

- **Context-Aware Batching**: Unlike simple tools, this skill gathers ALL text runs from a single slide and sends them as a structured JSON array. This allows the LLM to understand the slide's theme and maintain semantic consistency.
- **ID Anchoring**: Every text segment is assigned a unique ID. The mapping back from the LLM is based on these IDs, making the process immune to "dropped items" or "count mismatches."
- **Recursive Group Processing**: The tool performs a deep-first search to find text hidden inside nested Group Shapes, ensuring no diagrams or charts are left untranslated.
- **Stability-First Rendering**: To prevent the "Repair Needed" error in PowerPoint, the tool deliberately ignores geometry changes (width/height) and fragile paragraph spacing, focusing purely on high-fidelity text replacement.
- **Text Sanitization**: Automatically strips non-printable control characters from LLM responses to prevent XML corruption.

## Additional Utilities

### Global Color Modification
If the user wants to change font colors across the entire presentation:
- Run `./.venv/Scripts/python change_color.py <path> <HEX_COLOR>`.
- Example: `/ppt-translator-color path/to/file.pptx FFFFFF` (for white).

## Troubleshooting

- **"Repair Needed"**: If PPT still asks for repair, check if the source has extremely complex nested tables.
- **English content in XML**: This usually indicates a logic failure in the "Deferred Write" phase. Ensure the latest `pipeline.py` is being used.
- **Garbled Characters**: Check terminal encoding (the tool uses UTF-8 internally).
