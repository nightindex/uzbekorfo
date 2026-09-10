# Bundled data

These files are the versioned seed data shipped with the add-in.

- `uzbek_main.dic` is the primary word list used to seed the local application dictionary.
- `uzbek_dictionary.json` is the structured dictionary dataset, generated on 2026-03-12.
- `grammar_rules.json`, `uzbek_suffixes.json`, and `proper_nouns.json` support grammar analysis.
- `translit_exceptions.json` and `explanations.json` seed their corresponding user-local stores.
- `uzbek_freq_seed.json` is the source seed for language-model work.

The project embeds and copies the runtime data specified in `UzbekOrfoAddIn.csproj`. When replacing a dataset, retain the stable filename and record its provenance and date in this document or in the pull request.
