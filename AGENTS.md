# AGENTS.md

## Entrypoint
- `traduzir_gui.py` — single-file Tkinter app. Run: `python traduzir_gui.py`
- No package manifest, no README, no tests, no git repo.

## Dependencies (not declared anywhere)
```
pip install openpyxl deep-translator pywin32
```
- `python-docx` is **not** used; `.doc` files are handled via win32com Word.
- `win32com` requires Windows + Microsoft Office installed.

## Data layout
- `CHECKLIST_CANTEIRO/` — 16 subdirectories, each with 1–2 checklist files.
- Only edit/process files **without** `_traduzido` in the name (that suffix marks already-translated outputs).
- `.xls` originals are auto-converted to `.xlsx` on output (COM Excel saves as FileFormat=51).

## Running the GUI
1. Select the root folder (`CHECKLIST_CANTEIRO/`).
2. Choose file types and target language.
3. Click **Escanear** → review list → **Iniciar Tradução**.
- Stop via **⏹ Parar** (threading.Event, checked between cells).
- No parallel execution; files processed one at a time.

## Cache & config
- `cache_traducao.json` — bilingual text cache, reused across sessions. Clear via **Limpar Cache** button.
- `config_tradutor.json` — saves last used directory only.
- Both live next to `traduzir_gui.py`.

## Translation engine
- `deep-translator` `GoogleTranslator.translate_batch()` with 0.4s delay between batches.
- Rate limiting is real; stalls are expected on large files.

## File size note
- Original `.xls` files: ~20 KB (data only).
- Retranslated `.xlsx` via win32com: 100–350 KB (full formatting preserved).
