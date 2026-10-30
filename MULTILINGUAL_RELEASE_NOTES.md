# SimpleSPC multilingual expansion — development build

**Date:** 30 October 2026  
**Base:** the complete, previously updated English/French SimpleSPC ZIP; original SPC formulae and vendor libraries retained.

## What is included

- Language selector now supports **English**, **French**, **Spanish (Spain)**, **German (Germany)**, **Italian (Italy)** and **Portuguese (Portugal)**.
- Four NEW draft dictionaries: `i18n/es.js`, `i18n/de.js`, `i18n/it.js`, `i18n/pt.js`.
- **443 translated entries per new language, out of 1,350 English reference keys (32.8%)**. Each dictionary intentionally contains only translated entries. The remaining **907 keys** automatically fall back to English, without displaying broken lookup keys. This is *not* a complete Spanish, German, Italian or Portuguese translation.
- Interface controls, SPC chart names, most chart-axis labels, several dialog messages, chart context menu actions and some helper questions have new-language translations.
- A visible draft notice appears only for the four additional languages; existing English/French user experience is unchanged.
- Number formatting supports `es-ES`, `de-DE`, `it-IT`, `pt-PT` in addition to `en-GB` and `fr-FR`; percentage display uses the browser's locale-aware percentage formatter.
- Numeric input parsing handles comma decimals, either punctuation ordering for thousands separators, nonbreaking-space groups and locale-specific ambiguous strings. It does not change mathematical calculations.
- Project JSON restore accepts the additional language codes for the original numeric-data locale.
- New test file `tests/multilingual.test.cjs` checks translation keys, placeholders, language registration, numbers, percentage formatting and draft warnings.
- `i18n/TRANSLATION_REVIEW_QUEUE.csv` is a complete 1,350-key review sheet: English source, French reference and four draft language columns. Blank cells indicate untranslated entries. It can be edited in a spreadsheet program or shared with native-speaking reviewers; it is a **review aid**, not loaded by SimpleSPC.

## Important: this is NOT a multilingual publication release

The French dictionary still contains all 1,350 English reference keys. Spanish, German, Italian and Portuguese currently cover **443/1,350** keys each. The rest (including longer statistical guidance, SPC helper responses, wizard descriptions, many validation messages and detailed chart interpretation) is in English. A native-speaking reviewer with statistical-process-control expertise should check the new translations, especially *special cause*, *baseline*, *control limits* and rule names. Additional translations must be incorporated into the `.js` dictionaries before removing the draft warning labels.

The draft status is **intentional**: it avoids presenting an inadequately reviewed technical translation as complete. For regional terminology, this build uses Spanish from Spain, standard German from Germany, standard Italian from Italy and European Portuguese.

## Interpreting imported numeric data

As in the existing French revision, when CSV data are imported the tool remembers the **interface language selected at import time** as the source-data numeric locale. This lets you switch the interface language later without changing numeric values. For example, select Deutsch before uploading a German CSV, then switch to English afterwards if desired.

The strings `1.234` and `1,234` are ambiguous across regions:

| Data language chosen when imported | `1.234` means | `1,234` means |
|---|---:|---:|
| English | 1.234 | 1234 |
| French | 1.234 | 1.234 |
| Spanish, German, Italian, Portuguese | 1234 | 1.234 |

Unambiguous forms such as `1234,56` or `1.234,56` are preferred. If importing numeric data in a language different from the interface, first choose the source-data language, import the data, then change the interface language. Numeric values in Excel workbooks that are actual numeric cells do not need textual separator parsing. The `pt-PT` parser accepts dot grouping for compatibility with common imported data formats even though some modern Portuguese systems group with spaces.

## Validation performed

- `node --check` completed for all six translation scripts, `i18n.js` and `spc.js`.
- `node --test tests/*.test.cjs` passed **9 of 9** tests (four prior French regression tests and five multilingual tests).
- Tested all six locale mappings and format/parse examples using the JavaScript runtime.
- Browser navigation was attempted, but blocked by the execution environment. **Full in-browser CSV, chart rendering, saved-project round-trip, offline HTML and exports have NOT been end-to-end verified for the four new languages**. Treat this as a development candidate and test on your GitHub Pages development site before wider circulation.

## To update via your usual Notepad/GitHub workflow

1. Copy your entire current development project to a safe backup folder.
2. Extract `SimpleSPC_Multilingual_Changed_Files_Only.zip`.
3. In your local project, replace **`i18n.js`**, **`spc.js`** and **`index.html`** with the corresponding new files.
4. Copy the **four new files** into your existing `i18n` folder: `es.js`, `de.js`, `it.js`, `pt.js`.
5. Optional but recommended: copy `tests/multilingual.test.cjs`, `i18n/TRANSLATION_REVIEW_QUEUE.csv`, and `MULTILINGUAL_RELEASE_NOTES.md` to the same relative paths for future use.
6. Upload the changed files to those same locations in GitHub. Do **not** upload the ZIP itself or delete other existing files.
7. After GitHub Pages updates, hard-refresh and test each language, run charts, import a CSV with comma decimals, switch language on an open chart and reopen a saved project.

All other original files are unchanged and can remain where they are. **Do not overwrite local edits you have made since the ZIP supplied earlier without comparing them first.**

## Recommended next translation stage

Use the review queue to complete the remaining 907 keys for each language. Focus first on `summary.*`, `helper.*`, `help.*`, `setup.*` and `wizard.*` because these include the tool's most substantive SPC advice. Keep English as the fallback until an expert reviews the completed technical text. When complete, update the in-code draft metadata and confirm that all six dictionaries have matching key sets.
