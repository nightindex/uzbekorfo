# Dictionary and morphology data contracts

`uzbek_dictionary.json` uses `uzbekorfo-dictionary-v2`. A reviewed lexical annotation is stored directly on an entry as a pair:

```json
{
  "Word": "kitob",
  "Lemma": "kitob",
  "PartOfSpeech": "noun"
}
```

Both fields are optional for migration, but declaring only one is invalid. A lemma must reference another bundled entry. Unannotated entries remain `unknown`; tooling must not infer metadata merely from spelling.

`uzbek_suffixes.json` uses `uzbekorfo-suffixes-v2`. It owns the executable morphology policy: families, category stages, occurrence limits, dependency requirements, productive flags, allomorph groups, vowel characters, and root-mutation maps. The application rejects the whole rule set if validation fails.

Compatible imported suffix entries additionally declare `standaloneOnly: true`
and `requiredRootFlags`. They are complete suffix forms: the parser never adds
them to a recursive chain. At least one case-sensitive character from
`requiredRootFlags` must occur in the root entry's optional `HunspellFlags`
metadata. Missing metadata rejects the imported suffix. These fields are read
by the runtime JSON loader as well as validated by the morphology model.

`HunspellFlags` is source metadata; it does not activate an external engine or
special Hunspell capitalization/suggestion directives. The source archive in
`docs/reference` is not a runtime resource. See `docs/hunspell-import.md` for the
converted subset and remaining source rules.

`uzbek_main.dic` and `uzbek_dictionary_metadata.json` are generated artifacts. Run `eng/sync-dictionary-data.ps1` after editing the canonical dictionary and never hand-edit either artifact.
