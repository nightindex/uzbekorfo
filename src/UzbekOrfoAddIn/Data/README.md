# Bundled data

These files are the versioned seed data shipped with the add-in.

- `uzbek_dictionary.json` is the canonical source (`uzbekorfo-dictionary-v1`). Edit its `Entries` and metadata here; it is a build-time source and is not shipped to end users.
- `uzbek_main.dic` is the generated, UTF-8, one-word-per-line runtime spelling list. Do not edit it manually.
- `uzbek_dictionary_metadata.json` is the generated compact runtime index of only entries that have definitions, rules, grammar notes, or examples. Do not edit it manually.
- `grammar_rules.json`, `uzbek_suffixes.json`, and `proper_nouns.json` support grammar analysis.
- `translit_exceptions.json` and `explanations.json` seed their corresponding user-local stores.
- `uzbek_freq_seed.json` is the source seed for language-model work.

The project embeds and copies the runtime data specified in `UzbekOrfoAddIn.csproj`. When replacing a dataset, retain the stable filename and record its provenance and date in this document or in the pull request.

Dictionary export formats: JSON is the complete migration/backup format. DIC is a UTF-8, one-word-per-line spelling list; it intentionally excludes definitions and other metadata. Excel formats preserve the structured fields.

The build generates both runtime artifacts before packaging. Spell checking loads only the DIC; definition lookups load the compact metadata index only after a local explanation misses. Built-in artifacts are refreshed on an existing installation when their packaged bytes change, while user dictionaries and user explanations are never overwritten. Run `eng\format-dictionary-source.ps1` after a large import, `eng\sync-dictionary-data.ps1` after editing the source, then `eng\validate-dictionary-sync.ps1` before changing bundled data or in CI. The generator rejects duplicate, empty, invisible, and control-character words.
