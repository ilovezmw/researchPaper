# researchPaper

Research note workflow: structured `manuscript.yaml` → Word (optional PDF) with title/abstract (single column), two-column body, full-width table sections, AI disclosure, and references.

## Quick start

```powershell
pip install -r requirements.txt
python src/generate_publication.py --input "In Process/sample_BTC_ETH_decoupling"
```

Outputs are written under `published/` (created automatically; on case-insensitive Windows this may match an existing `Published/` folder).

See `In Process/sample_BTC_ETH_decoupling/manuscript.yaml` for the YAML schema.
