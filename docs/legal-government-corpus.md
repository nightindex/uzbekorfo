# Government legal corpus

Raw files are stored locally under these Git-ignored directories:

- `corpus/raw/sud_uz_qonunchilik/`
- `corpus/raw/sud_uz_prezident-farmonlari-qarorlari/`
- `corpus/raw/sud_uz_vazirlar-mahkamasi-qarorlari/`

The files were downloaded on 2026-09-16 from the Uzbek Supreme Court pages:

- <https://old.sud.uz/%D2%9B%D0%BE%D0%BD%D1%83%D0%BD%D1%87%D0%B8%D0%BB%D0%B8%D0%BA/>
- <https://old.sud.uz/%D0%BF%D1%80%D0%B5%D0%B7%D0%B8%D0%B4%D0%B5%D0%BD%D1%82-%D1%84%D0%B0%D1%80%D0%BC%D0%BE%D0%BD%D0%BB%D0%B0%D1%80%D0%B8-%D0%B2%D0%B0-%D2%9B%D0%B0%D1%80%D0%BE%D1%80%D0%BB%D0%B0%D1%80%D0%B8/
- <https://old.sud.uz/%D0%B2%D0%B0%D0%B7%D0%B8%D1%80%D0%BB%D0%B0%D1%80-%D0%BC%D0%B0%D2%B3%D0%BA%D0%B0%D0%BC%D0%B0%D1%81%D0%B8-%D2%9B%D0%B0%D1%80%D0%BE%D1%80%D0%BB%D0%B0%D1%80%D0%B8/>

## LexUZ Word exports

Each DOC file uses the official endpoint `https://lex.uz/docs/{id}?type=doc`.
The export is labelled `application/msword` and contains Word-compatible HTML inside
the `.doc` filename.

| Source | LexUZ document IDs |
| --- | --- |
| Қонунчилик | 20596 |
| Президент фармонлари ва қарорлари | 3050491, 3107036, 3121087, 3252979, 3286042, 3323071, 3327998, 3432426, 3516847, 3725214, 3731060, 3735818, 4166958, 4172023, 4188795, 4188851, 4190059, 4192433, 4199119, 4201081, 4203364 |
| Вазирлар Маҳкамаси қарорлари | 698065 |

Total: 23 DOC exports. The presidential page also linked LexUZ PDF `4230944`; its
endpoint returned an HTML viewer page rather than a PDF file, so it was not retained
as a misleading `.pdf` artifact.

These are dated source snapshots. Before training, record each document's effective
date/version, remove duplicates and formatting artifacts, and compare the text with
the current authoritative publication. Do not include confidential client material
without documented permission and de-identification.
