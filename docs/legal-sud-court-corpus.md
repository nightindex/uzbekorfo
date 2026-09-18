# Supreme Court court-materials corpus

The downloader added the five requested Supreme Court index pages to the local
training corpus on 2026-09-17. Exact URLs and per-page results are preserved in:

- `corpus/raw_cyrillic/sud-court/manifest.json`
- `corpus/raw-latin/sud-court/manifest.json`

The pages exposed **99 unique LexUZ IDs**. **98 Cyrillic** and **98 Latin**
Word-compatible exports were downloaded. ID `1448644` returned HTTP 404 for both
variants and is recorded as unavailable. The economic-materials index page also
returned HTTP 404, so no links could be discovered from it. The other four pages
returned 53, 38, 2, and 6 links respectively.

The positive LexUZ ID export is used for Cyrillic and the negative ID export for
Uzbek Latin. Files remain separate; no translation or transliteration was done.

Run the resumable downloader again with:

```powershell
./eng/download-sud-court-corpus.ps1
```

Existing valid files are retained. `-Force` refreshes available exports. The
downloaded source documents are ignored by Git and are not included in the add-in.

These files have not been silently added to the built-in prediction artifacts.
After checking permissions, source quality, duplicate/related-version grouping,
and script content, rebuild and evaluate explicitly. Public availability does not
by itself establish permission to use every document for a commercial product.
The current corpus build intentionally excludes the `sud-court` subfolder, so
existing 57-source model builds remain reproducible. A future reviewed merge
should add these sources deliberately and update both manifests/reports together.
