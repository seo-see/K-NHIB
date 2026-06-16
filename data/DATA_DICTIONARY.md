# Data Dictionary — K-NHIB Benchmark

File: `data/benchmark/K-NHIB_GY.xlsx`

The benchmark contains **222 cases** across **3 gynecologic cancers**, built from
**74 anticancer regimens** (each regimen contributes one `eligible`, one `ineligible`, and
one `undeterminable` case → 74 × 3 = 222). Cases are synthetic / processed; they contain no
patient identifiers.

## Sheets

One sheet per cancer type. All three share the same column schema.

| Sheet | Cancer type | Cases |
|---|---|---:|
| `cervical` | Cervical | 45 |
| `uterine` | Uterine (endometrial) | 51 |
| `ovarian` | Ovarian (incl. fallopian-tube / primary peritoneal) | 126 |
| **Total** | | **222** |

## Columns (identical across all three sheets)

Column headers are English; the **values** of `regimen` and `attributes` are Korean, matching
the Korean-language guideline documents (this is the actual model input — see note below).

| Column | Meaning | Notes / values |
|---|---|---|
| `ID` | Case identifier | Format `FW-<cancer>-<regimen>-<class>`, e.g. `FW-C-R1-pos` (FW = forward; C/… = cancer; R1 = regimen; class suffix) |
| `regimen_code` | Regimen reference | Regimen handle (e.g. `R1`, `R2`, …) |
| `regimen` | Anticancer regimen | Drug regimen evaluated for reimbursement; fed to the model as `{regimen}` (e.g. `ifosfamide + carboplatin/cisplatin`) |
| `attributes` | Clinical/administrative attributes | The case's structured attributes (Korean, e.g. `투여단계=1차, 질환상태=재발성`); fed to the model as `{attributes}` |
| `expected` | **Gold label** | One of `eligible` / `ineligible` / `undeterminable` (74 each) |
| `class` | Case construction class | `pos` / `neg` / `unk` (74 each) — the eligibility-condition state the case was designed to probe |

> The cancer type passed to the model is derived from the **sheet name** (cervical → 자궁경부암,
> uterine → 자궁내막암, ovarian → 난소암); the prompt uses the Korean cancer name, consistent with
> the Korean-language guideline documents in `guidelines/`.

## Tristate labels (`expected`)

- `eligible` — all required conditions are met.
- `ineligible` — ≥1 condition is explicitly not met (criterion present but violated).
- `undeterminable` — no condition is explicitly violated, but ≥1 condition is unevaluable due
  to absent clinical information, so a determination cannot be reached.

The pipeline reads the cancer type (from the sheet), `regimen`, and `attributes` into the prompt
template (see `prompts/`) and compares the model's `decision` against the `expected` gold label.
