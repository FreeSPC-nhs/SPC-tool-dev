# SimpleSPC — multilingual release notes

**Prepared:** 30 October 2026  
**Status:** Six-language release candidate, suitable for user testing and native-speaker review  
**Scope:** Localisation and translation updates; statistical calculation algorithms are not intentionally changed by this release.

## Overview

SimpleSPC now offers six interface languages:

|Language|Locale|Dictionary|Translation status|
|-|-|-|-|
|English|`en-GB`|`i18n/en.js`|Reference language|
|French (France)|`fr-FR`|`i18n/fr.js`|Full AI-assisted first draft|
|Spanish (Spain)|`es-ES`|`i18n/es.js`|Full AI-assisted first draft|
|German (Germany)|`de-DE`|`i18n/de.js`|Full AI-assisted first draft|
|Italian (Italy)|`it-IT`|`i18n/it.js`|Full AI-assisted first draft|
|Portuguese (Portugal)|`pt-PT`|`i18n/pt.js`|Full AI-assisted first draft|

Each dictionary contains the **1,350 entries** present in the English reference dictionary. The Spanish, German, Italian and Portuguese translations were completed after the initial multilingual development build. Their earlier **draft labels and banners have been removed**. English remains the application's fallback language if a lookup fails.

**What “complete” means:** dictionary-key coverage is complete. It does **not** mean the translations have been certified by native-speaking SPC specialists, or that every browser, chart and export workflow has undergone exhaustive testing. The versions are intended for practical evaluation and feedback.

## User-facing localisation

* The language selector includes all six languages and uses their native names.
* Interface controls, chart selection and setup guidance, help material, SPC Helper content, generated interpretations and relevant warnings have translated dictionary entries.
* Formatting controls use language-appropriate labels and tooltips where implemented.
* The CSV/Excel file-selection control retains translated text but has been restyled to resemble a conventional grey file-picker button.
* Switching language updates translated interface content and, where supported, displayed chart annotations and summaries without changing underlying statistical values.
* The selected language and the source-data numeric locale are handled when saving/reopening projects.

## Numbers, dates and imported data

* Display formatting uses the chosen locale (including decimal/group separators and percentages).
* Textual numeric data imported from CSV may use decimal commas and locale-appropriate grouping separators; the importer uses the **language selected when the data are imported** to interpret ambiguous numeric strings.
* If the CSV's number convention differs from the display language, **select the source-data language before importing**, then switch the interface language afterwards if desired.
* Text such as `1.234` or `1,234` can mean different things in different regions. Prefer unambiguous source values and check a sample of imported data in the editor before interpreting charts.
* Excel cells stored as actual numbers do not require textual decimal-separator parsing; text-formatted numeric cells may still need checking.
* SPC calculations use numeric data values, not the formatted strings shown to users.

## Translation quality review

A follow-up AI-assisted SPC terminology and meaning audit was completed on **30 October 2026**, covering **6,750 translated dictionary entries** across the five non-English languages.

* All five translation dictionaries contained the full English reference key set.
* **18 targeted corrections** were made: French (2), German (11), Italian (3), Portuguese (2). No Spanish dictionary change was required in that review pass.
* Review priorities included control limits versus targets or specification limits; common-cause versus special-cause variation; stability versus performance; XmR screening; and avoiding unsupported causal conclusions.
* Two older French Pareto messages intentionally omit a `{count}` placeholder because they refer grammatically to the first category; this is a documented exception, not missing statistical content.
* Some chart abbreviations and symbols remain unchanged intentionally (for example `XmR`, `X̄–S`, `P′`, `U′` and `Sigma Z`).

Supporting review files were generated separately: `SimpleSPC\_Multilingual\_SPC\_QA\_Report.md`, `SimpleSPC\_QA\_Changes.csv` and `SimpleSPC\_Terminology\_Glossary.csv`. These are useful maintenance records but are **not required for the live application**.

## Verification carried out

For the completed translation builds and terminology corrections:

* JavaScript syntax and translation-key coverage checks passed.
* Translation placeholders and embedded HTML structure were checked, with the documented French exceptions above.
* The final multilingual localisation suite reported **11 of 11 automated tests passed**.
* Browser smoke tests during the language-completion work checked language selection, translated interface labels, draft-banner removal and decimal display in selected languages.
* An additional full browser run during the subsequent terminology audit did not complete reliably.

**Not yet established by these checks:** exhaustive validation of every chart type, signal rule, CSV/Excel import variant, JSON restore path, image/PDF export and offline portable HTML export in each language and browser. The existing SPC algorithms were not intentionally changed as part of translation work, but that is not a substitute for functional regression testing.

## Recommended checks before wider promotion

1. Confirm that the **latest** `i18n/\*.js`, root `i18n.js`, application files and the file-picker stylesheet have all been deployed together. Older dictionaries or a cached script may produce misleading results.
2. Hard-refresh the development site; switch through all six languages and check navigation, dialogs, chart selection, chart summaries, help and SPC Helper responses.
3. Import a small synthetic CSV containing decimal commas in a comma-decimal locale. Confirm the values shown in the data editor and plotted in the chart.
4. Generate at least a Run chart and an XmR chart in each language; confirm that translated labels do not obscure statistical values or chart controls.
5. Save and reopen a JSON project; test chart image/export and portable HTML where used.
6. Invite native-speaking users with SPC or quality-improvement experience to report ambiguous wording or misleading statistical interpretations.

**Release positioning:** appropriate to share as a **multilingual first-draft release for evaluation**, with a clear request for feedback. Do not describe the translations as professionally certified or the software as fully validated in every locale.

## Guidance for future SimpleSPC changes

The application code and SPC calculations are shared across languages. Any new or materially revised user-facing text should be added to the English reference dictionary and reviewed in the five other dictionaries before release.

A useful routine is to compare dictionary keys and placeholders, check changed text for meaning (not just presence), test locale-aware inputs and display, and run chart regression tests. Prefer automated key-set comparison to a hard-coded assumption that there will always be 1,350 entries.

## Repository maintenance notes

* Keep this file in the repository root as `MULTILINGUAL\_RELEASE\_NOTES.md`.
* The older `i18n/TRANSLATION\_REVIEW\_QUEUE.csv` was generated for incomplete dictionaries and may now contain stale coverage information. Update it or mark it as historical before using it as a current translation-status tracker.
* Keep a GitHub commit or local backup of the last known working version before replacing application files.
* These notes document the completed local translation work; they do not, by themselves, confirm that the same set of files has been deployed to the live/public site.

