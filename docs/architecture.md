# Architecture

Uzbek Orfo is a Microsoft Word VSTO add-in targeting .NET Framework 4.7.2. It is intentionally maintained as one deployable application because VSTO deployment and Word interop are tightly coupled to the add-in host.

## Repository layout

- `src/UzbekOrfoAddIn/` — VSTO project and product source.
- `eng/` — repository validation scripts.
- `tests/` — tests for logic that can run without Word.
- `docs/` — project documentation.

## Source responsibilities

- `Core/` defines service contracts.
- `Prediction/` implements Word-independent document extraction, encrypted collections,
  observed phrase statistics, asynchronous ranking and opt-in acceptance preferences.
  See [MatnAI document prediction](prediction/README.md) for storage, safety and validation boundaries.
- `Models/` holds application data structures.
- `Services/` implements spelling, grammar, dictionary, export, settings, and workflow behavior.
- `Forms/` and `UI/` contain WinForms presentation and controls.
- `Data/` contains versioned embedded dictionaries and rules.

`ThisAddIn` and the ribbon classes are the VSTO host boundary. `AddInRuntime` is the composition root for runtime services. New domain logic should avoid direct Word or WinForms dependencies where possible so it can be tested in `tests/`.

## Dictionary and morphology

`Data/uzbek_dictionary.json` is the canonical dictionary. Reviewed entries may carry `Lemma` and `PartOfSpeech`; the generator validates both fields together and copies them into the compact runtime metadata index. Unannotated entries stay `unknown` and are never assigned a guessed part of speech.

`Data/uzbek_suffixes.json` is a versioned executable data contract. It owns suffix families, category stages, occurrence limits, dependency requirements, allomorph groups, productive flags, vowels, and root-mutation maps. `MorphologyRuleSetLoader` validates the complete file before `UzbekSuffixParser` can use it. Invalid data disables productive acceptance as a unit.

Rejected words from explicit document checks are aggregated into `%AppData%/UzbekOrfo/unknown_words.tsv`. This is local-only review data containing the normalized word, count, and first/last UTC timestamps; document names, text context, and user identifiers are not stored or transmitted.
