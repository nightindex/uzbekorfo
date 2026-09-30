# Laws prediction corpus

The raw Word exports are stored locally under
`corpus/raw/sud_uz_qonunlar/`. The directory is ignored by Git because these are
source documents for corpus preparation, not application assets.

Source page:

<https://old.sud.uz/%D2%9B%D0%BE%D0%BD%D1%83%D0%BD%D0%BB%D0%B0%D1%80/>

On 2026-09-16, the 16 LexUZ links on the page were downloaded through the official
Word export endpoint `https://lex.uz/docs/{id}?type=doc`:

| Local file | LexUZ document |
| --- | ---: |
| `lexuz-1072079.doc` | 1072079 |
| `lexuz-1633102.doc` | 1633102 |
| `lexuz-1685726.doc` | 1685726 |
| `lexuz-1876877.doc` | 1876877 |
| `lexuz-2374282.doc` | 2374282 |
| `lexuz-2381133.doc` | 2381133 |
| `lexuz-2387357.doc` | 2387357 |
| `lexuz-26477.doc` | 26477 |
| `lexuz-3088008.doc` | 3088008 |
| `lexuz-3093543.doc` | 3093543 |
| `lexuz-3153668.doc` | 3153668 |
| `lexuz-3221763.doc` | 3221763 |
| `lexuz-3227588.doc` | 3227588 |
| `lexuz-4202732.doc` | 4202732 |
| `lexuz-68532.doc` | 68532 |
| `lexuz-955375.doc` | 955375 |

LexUZ labels these responses as Microsoft Word exports. As with the code corpus,
the `.doc` files contain Word-compatible HTML and can be opened in Word or parsed
as HTML/text.

The page also contains links to external sites and one direct PDF. Those four items
were not copied into this DOC corpus; they should be handled separately if their
source terms permit redistribution.

These are dated source snapshots. Before model training, record each document's
version/date, remove formatting artifacts and duplicates, and compare the text with
the current authoritative publication. Do not include confidential client material
without documented permission and de-identification.
