# Confirmed duplicate cleanup

Removed only two mixed-script keys from the canonical dictionary:

- `атакcия` (Latin c): retained existing `атаксия` (Cyrillic с).
- `юнeско` (Latin e): retained existing `юнеско` (Cyrillic е) and transferred
  the removed entry's definition verbatim. Existing source flags remain unchanged.

Evidence is the code-point mismatch and already-existing all-Cyrillic counterpart;
this is not a claim of independent linguistic review. All other screening findings,
including uncertain abbreviations and multiword phrases, were left unchanged.

The complete pre-edit canonical dictionary is recoverable locally at
`TestResults/dictionary-before-confirmed-removals.json`. Restore only the two entries
and the previous target definition from that copy if later unrelated edits must be
preserved, update counts and regenerate using `eng/sync-dictionary-data.ps1`.
The exact-key migration is recorded in `eng/remove-confirmed-dictionary-duplicates.ps1`.

New canonical count: 216,339 words. Definitions remain at 25,642.
Earlier audit hashes refer to the pre-cleanup snapshot; linguistic approval remains pending.
