# researchPaper

Research note workflow: structured `manuscript.yaml` → Word (optional PDF) with title/abstract (single column), two-column body, full-width table sections, AI disclosure, and references.

## Quick start

```powershell
pip install -r requirements.txt
python src/generate_publication.py --input "In Process/sample_BTC_ETH_decoupling"
```

Run these commands from the **same folder that contains `src/`** (the project root you opened in Cursor). If the terminal `cd` points somewhere else, files go to that other copy’s `Published/` and you will not see them in the IDE.

If two topic folders use the same `title` in YAML, the default `.docx` name collides and one overwrites the other. Use a distinct name:

```powershell
python src/generate_publication.py --input "In Process/Research Paper Generation Request" --output-stem "Research_Paper_Generation_Request_March_2026"
```

Or set optional `output_stem: My_File_Stem` at the top level of `manuscript.yaml`.

**Filename vs content:** `--output-stem` / `output_stem` only change the **.docx file name**. The **body** always comes from `manuscript.yaml`. If you copied YAML from another topic, the title on disk can look correct while the text is still the old paper — rebuild YAML from your real Word draft:

```powershell
python src/docx_to_manuscript.py "In Process/Research Paper Generation Request/Research Paper Generation Request.docx" -o "In Process/Research Paper Generation Request/manuscript.yaml"
python src/generate_publication.py --input "In Process/Research Paper Generation Request" --no-done
```

(That importer matches the outline of `Research Paper Generation Request.docx`. If you change headings in Word, adjust slice logic in `src/docx_to_manuscript.py`.)

Outputs are written under **`Published/`** (created automatically). After each run the tool checks that section headings from your YAML appear in the DOCX; if you see `WARNING` lines, close the file in Word and regenerate.

See `In Process/sample_BTC_ETH_decoupling/manuscript.yaml` for the YAML schema.

## Folder layout (recommended)

- **`In Process/<课题文件夹>/`** — one folder per paper or work-in-progress. Put that paper’s `manuscript.yaml` **inside** that folder, plus any related `.docx` / PDFs you use while drafting.
- **`Published/`** — generated `.docx` / `.pdf` from the script (output file name comes from the YAML title + date).

Do **not** merge two different topics into a single `manuscript.yaml` unless they are literally one combined document. Otherwise you lose a clean 1:1 mapping: *one topic folder → one manuscript → one export*.

You **can** put **`In Process/manuscript.yaml` directly under `In Process`** (no subfolder) if you only ever work on one draft at a time. Then run:

```powershell
python src/generate_publication.py --input "In Process"
```

The script always looks for **`manuscript.yaml` inside the path you pass to `--input`**, then writes to **`Published/`**.

Examples:

| What you pass to `--input` | Manuscript path used |
|----------------------------|----------------------|
| `In Process/sample_BTC_ETH_decoupling` | `In Process/sample_BTC_ETH_decoupling/manuscript.yaml` |
| `In Process/Research Paper Generation Request` | `In Process/Research Paper Generation Request/manuscript.yaml` |
| `In Process` | `In Process/manuscript.yaml` |

## Workflows

### A — You already have a Word file (layout done in Word)

1. Put the `.docx` under `In Process/<topic>/` (or anywhere you like).
2. Install dependencies (`pip install -r requirements.txt`). You need **Microsoft Word** and **pywin32** on Windows.
3. Export PDF:

```powershell
python src/word_to_pdf.py "In Process/my-topic/your.docx"
```

By default the PDF is written next to the `.docx` with the same base name. To put it under `Published/`:

```powershell
python src/word_to_pdf.py "In Process/my-topic/your.docx" -o Published/your.pdf
```

Alternatively: open the file in Word → **File → Save As → PDF** (same rendering, no script).

### B — You use `manuscript.yaml` to auto-build the note layout

1. Edit `manuscript.yaml` in the topic folder.
2. Generate Word and PDF in one go:

```powershell
python src/generate_publication.py --input "In Process/my-topic" --pdf
```

Word/PDF go under **`Published/`**.
