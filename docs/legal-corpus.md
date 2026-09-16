# Legal prediction corpus

The raw files downloaded from the Uzbek Supreme Court code page are stored locally
under `corpus/raw/sud_uz_kodekslar/`. They are ignored by Git because they are source
documents for model preparation, not application assets.

Source page:

<https://old.sud.uz/%D0%BA%D0%BE%D0%B4%D0%B5%D0%BA%D1%81%D0%BB%D0%B0%D1%80/>

The page links to LexUZ. Each file below was downloaded from the official LexUZ Word
export endpoint `https://lex.uz/docs/{id}?type=doc` on 2026-09-16. LexUZ labels the
response as `application/msword`; the downloaded `.doc` files use Word-compatible HTML
inside the DOC filename, so they can be opened in Word and parsed as HTML/text.

| File | LexUZ document |
| --- | ---: |
| `01-jinoyat-kodeksi.doc` | 111453 |
| `02-jinoyat-protsessual-kodeksi.doc` | 111460 |
| `03-fuqarolik-kodeksi-1-qism.doc` | 111189 |
| `04-fuqarolik-kodeksi-2-qism.doc` | 180552 |
| `05-fuqarolik-protsessual-kodeksi.doc` | 3517337 |
| `06-iqtisodiy-protsessual-kodeksi.doc` | 3523891 |
| `07-mamuriy-sud-ishlari-kodeksi.doc` | 3527353 |
| `08-mamuriy-javobgarlik-kodeksi.doc` | 97664 |
| `09-bojxona-kodeksi.doc` | 2876354 |
| `10-byudjet-kodeksi.doc` | 2304138 |
| `11-soliq-kodeksi.doc` | 1286558 |
| `12-shaharsozlik-kodeksi.doc` | 46868 |
| `13-uy-joy-kodeksi.doc` | 106136 |
| `14-yer-kodeksi.doc` | 152653 |
| `15-oila-kodeksi.doc` | 104720 |
| `16-jinoyat-ijroiya-kodeksi.doc` | 163629 |
| `17-mehnat-kodeksi.doc` | 142859 |
| `18-havo-kodeksi.doc` | 55594 |

These are dated source snapshots. Before using them for legal prediction, record the
document version/date from each file, remove non-content formatting, and compare the
text with the current authoritative publication. The corpus must not include
confidential client documents without documented permission and de-identification.
