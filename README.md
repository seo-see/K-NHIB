# K-NHIB: Korean National Health Insurance Benchmark

[![Paper](https://img.shields.io/badge/Paper-JMIR-blue)](https://doi.org/10.2196/95877)
[![Dataset](https://img.shields.io/badge/Dataset-Zenodo-blue)](https://doi.org/10.5281/zenodo.20563360)

A benchmark for **anticancer drug reimbursement eligibility determination** under South
Korea's National Health Insurance, with an LLM evaluation pipeline. It formalizes the
condition-level adjudication logic clinicians and utilization-review nurses apply when
evaluating incomplete clinical evidence against coverage rules. It covers **3 gynecologic
cancers (cervical, uterine, ovarian)**, **74 anticancer regimens / 222 cases**, and a
**tristate adjudication framework** (`eligible` / `ineligible` / `undeterminable`).

> Companion to: Seo J, Kim T, Kim J-H. *Assessing Eligibility for Anticancer Drug Health
> Insurance Reimbursement Using Large Language Models: Benchmark Development and Comparative
> Study.* J Med Internet Res. 2026;28:e95877. doi:10.2196/95877

The benchmark and evaluation pipeline are **model-agnostic** — you can evaluate any LLM. For
the specific models evaluated in our study and their comparative results, see the paper.

## Layout

```
data/        benchmark dataset (xlsx) + HIRA guideline source files + DATA_DICTIONARY.md
prompts/     system + user prompt templates and condition descriptions
code/        evaluation pipeline + scoring (run from the repo root)
docs/        REPRODUCE.md — end-to-end reproduction steps
```

## Quick start

```bash
pip install -r requirements.txt
cp .env.example .env          # then fill in your API keys
python code/benchmark_pipeline.py     # run the eval → generates results/ (gitignored)
python code/aggregate_results.py      # score → accuracy / per-class F1 under results/aggregated/
```

All scripts are run **from the repository root**. See `docs/REPRODUCE.md` for details.

## Data

- `data/benchmark/K-NHIB_GY.xlsx` — synthetic / processed cases (no patient identifiers),
  one sheet per cancer (`cervical` / `uterine` / `ovarian`). See `data/DATA_DICTIONARY.md`.
- `data/guidelines/` — HIRA (Health Insurance Review & Assessment Service) public reimbursement
  guideline documents (Korean) used as model input:
  - `Cervical.pdf`, `Uterine.pdf`, `Ovarian.pdf` — per-cancer guideline excerpts (the
    per-cancer PDF is the `pdf`-condition input).
  - `HIRA_full_booklet_2026-02-01.pdf` — full HIRA reimbursement-criteria booklet (2026-02-01).
  - These are official public notifications (고시/공고) from HIRA, included for reproducibility.
  - The text-input condition parses these PDFs at run time (`code/pdf_parser.py`); the parsed
    text is written to `results/parsed_guidelines/` (gitignored), not shipped.

## Citation

Archive DOI: [10.5281/zenodo.20563360](https://doi.org/10.5281/zenodo.20563360). See
`CITATION.cff` for full metadata, and cite the companion article (doi:10.2196/95877) as well.

## License

Dual-licensed (see `LICENSE`):
- **Code** (`code/`): MIT — `LICENSE-CODE`
- **Dataset & documentation** (`data/benchmark/`, `prompts/`, `docs/`, this README): CC BY 4.0 — `LICENSE-DATA`
- **HIRA guideline documents** (`data/guidelines/`): official public notifications (고시/공고); not subject to copyright under Korean law; included for reproducibility.
