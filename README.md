# RFQ PDF → n8n → Excel quotation draft

A small, deterministic prototype for electrical-parts RFQs. It accepts one fixed bordered text-PDF table, matches only exact normalized product code + description + unit against an approved catalogue and AED price list, and writes an English/Arabic Excel draft. Missing prices and conflicts remain `REVIEW`; the complete total is withheld until review.

This repository uses synthetic data only. It contains no customer PDF, credentials, external API call, LLM call, ERP connector, or automatic substitution.

## Run

```powershell
python build_demo.py
python -m unittest -v test_rfq.py
python rfq.py
```

The helper listens on `127.0.0.1:8795` and accepts `POST /quote` with the JSON shape in `sample/request.json`. Import `workflow.json` into n8n; it is intentionally inactive and uses the local helper URL. `n8n-evidence.json` records an actual n8n CLI run on 2026-09-15 using n8n 2.39.5.

## Sample result

The synthetic eight-row fixture produces three matched rows (AED 26.16 subtotal) and five review rows. The line `0.145 × 1.00` is AED 0.15 using decimal `ROUND_HALF_UP`; the complete total is blank while reviews remain. The workbook has `Quote EN`, `عرض السعر AR`, and `Exceptions` sheets, formula caches, freeze panes, filters, and RTL Arabic layout.

## Scope and estimate

A paid test should use an unseen PDF of the same table layout. A reasonable first test is USD 25 for one sample PDF and a short handover; a fuller pilot can be scoped after the test (roughly USD 150 for up to 100 rows, 1,000 catalogue rows, approved AED prices, and bug fixes within the agreed layout). This is an estimate, not a commitment: OCR, ERP/email integration, substitutes, tax/discount logic, and extra PDF layouts are separate work.

Implementation and test execution were performed by an AI agent for its owner. No independent human code review is claimed; the buyer can inspect the source and rerun the tests.
