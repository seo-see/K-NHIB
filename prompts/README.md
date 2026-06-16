# Prompt Templates

These are verbatim copies of the prompts used by the evaluation pipeline. The
authoritative definitions live in `code/benchmark_pipeline.py`:

- `system_prompt.txt` ← `FORWARD_SYSTEM_PROMPT` (benchmark_pipeline.py)
- `user_prompt_template.txt` ← `build_forward_user_prompt()` (benchmark_pipeline.py)
- `structure_guided_system_prompt.txt` ← `STRUCTURE_GUIDED_SYSTEM_PROMPT` (used only by the
  `pdf_structure` condition)

The user template fields are filled per benchmark row:
`{cancer_type}` (Korean cancer name, derived from the sheet), `{regimen}`, `{attributes}`
(clinical/administrative attributes; values are Korean).

## Conditions

The main analysis and the structured-text / web-search sensitivity conditions share the same
system + user prompt (`system_prompt.txt`); they differ only in **what guideline material is
supplied** to the model. The `pdf_structure` condition is the exception: it uses a distinct
system prompt (`structure_guided_system_prompt.txt`).

| Condition | System prompt | Guideline input | Web search | Role |
|---|---|---|---|---|
| `pdf` | base | Guideline PDF attached natively | off | **Main analysis** |
| `text_md` | base | Guideline as plain text extracted from the PDF (pdfplumber) | off | Sensitivity (structured text) |
| `pdf_websearch` | base | Guideline PDF attached natively | on | Sensitivity (web search) |
| `pdf_structure` | structure-guided | Guideline PDF attached natively | off | Sensitivity (structure-guided) |

(The pipeline also supports `websearch` and `text_md_websearch` as development ablations; these
were not reported in the manuscript.)

See `code/benchmark_pipeline.py` (`call_claude` / `call_gpt` / `call_gemini`) for how each
condition assembles its request.
