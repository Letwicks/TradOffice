# AGENTS.md

## Entrypoint
- `traduzir_gui.py` — single-file Tkinter app. Run: `python traduzir_gui.py`
- Tests: `python -m unittest discover -s tests` (no Office/network needed; translator and COM are faked).

## Dependencies (not declared anywhere)
```
pip install openpyxl deep-translator pywin32
```
- `python-docx` is **not** used; Word files are handled via win32com Word.
- `win32com` requires Windows + Microsoft Office. The import is optional: without it only `.xlsx` works (openpyxl fallback, loses images/charts).
- Office is driven with `DispatchEx` (private, invisible instances reused across files). Never use `Dispatch`: it attaches to the user's open Excel/Word and `Quit()` would close their work.

## Data layout
- `CHECKLIST_CANTEIRO/` — 16 subdirectories, each with 1–2 checklist files.
- Outputs are written next to the original as `<name>_<LANG>.<ext>` (e.g. `_PT_BR`); `.xls` → `.xlsx` (FileFormat=51), `.docx` → FileFormat=16, `.doc` → FileFormat=0.
- The scanner skips outputs (`_traduzido` in the name, the target-language suffix, or another language suffix when the original sits in the same folder) and Office lock files (`~$*`).

## Running the GUI
1. Select the root folder (`CHECKLIST_CANTEIRO/`).
2. Choose file types, source language (or auto) and target language.
3. Click **Escanear** → review list → **Iniciar Tradução**.
- Stop via **⏹ Parar**: cooperative (`threading.Event`), the worker closes Office and saves the cache; the file in progress is not written.
- No parallel execution; files processed one at a time.

## Cache & config
- `cache_traducao.json` — `{"__versao__": 2, "idiomas": {"pt": {src: dst}}}`; legacy flat v1.1 caches are migrated to `pt`. Written atomically (tmp + replace); a corrupt file is renamed to `*.corrompido.json`.
- `config_tradutor.json` — last folder, languages and options.
- Both live next to `traduzir_gui.py` (or the frozen `.exe`, falling back to `%LOCALAPPDATA%\TradOffice`).

## Translation engine
- `deep-translator` `GoogleTranslator.translate()` (its `translate_batch` is just a loop of single requests).
- `Tradutor` groups short single-line texts into one request joined by `\n` (max ~1800 UTF-8 bytes, the text goes in the URL); if the line count differs, it falls back to one request per text. Lines returned unchanged are retried alone (mixed-language checklists).
- 0.4 s minimum between requests, exponential backoff on errors, `ServicoIndisponivel` after 3 consecutive failed requests.

## Preservation rules
- Excel: skip formula cells (reads `UsedRange.Value`/`.Formula` in bulk), prefix `'` when a translation would be parsed as formula/number, sanitize/uniquify sheet names, restore `Calculation` before saving.
- Word: translate per paragraph excluding the paragraph/cell mark; skip paragraphs containing fields, images, shapes, notes or content controls; turn off `TrackRevisions`.

## File size note
- Original `.xls` files: ~20 KB (data only).
- Retranslated `.xlsx` via win32com: 100–350 KB (full formatting preserved).
