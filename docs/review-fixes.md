# Document safety and review fixes

## Behavior

- Bulk corrections use each error's live Word range and require the active document to match its COM identity. Exact checked text must still match before replacement. Edited or closed ranges are skipped and reported, without being marked resolved.
- The errors dialog and modeless suggestions also validate the target before replacing. A one-character or whitespace selection remains a selection; only an insertion point/no selection falls back to document content.
- Spelling and grammar marks are tracked by this session. Existing underlines or mixed underline formatting are left intact; those errors remain available in the errors dialog. Cleanup restores owned marks and preserves subsequent user formatting changes where distinguishable. Legacy documents are not swept for wavy underlines because their ownership cannot be established safely.
- Whole-range transliteration preserves trailing whitespace and Word's mandatory final paragraph marker without adding a blank paragraph. Its confirmation names the selection/document scope.
- Shortcuts use a thread-specific `WH_KEYBOARD` hook. They do not register global hotkeys, run only with a Word editing window in the foreground, and dispatch actions after the hook returns. Word dialogs and add-in forms do not activate these shortcuts. Other add-ins may still assign the same shortcuts inside Word.
- Settings writes use unique temporary files, a named mutex across Word processes, and replacement with a `.bak` of the previous version. Last successful save wins.
- The ribbon emphasizes checking, reviewing, and script conversion. Import cards stack in narrow windows; About navigation switches to a selector. Custom controls expose keyboard focus and the export toggles expose their labels and checked state.

## Automated validation

Run the checks listed in the README. The UI harness additionally tests toggle keyboard/accessibility behavior and proves that installing the thread hook leaves a matching global hotkey available.

Use `./eng/test-unit.ps1` for unit tests. It selects Visual Studio's C# compiler, which avoids the .NET compiler's CET compatibility issue on affected Windows installations. The equivalent explicit command is:

```powershell
$testVs = & "${env:ProgramFiles(x86)}/Microsoft Visual Studio/Installer/vswhere.exe" -latest -products * -property installationPath
dotnet test tests/UzbekOrfoAddIn.UnitTests/UzbekOrfoAddIn.UnitTests.csproj /p:UseSharedCompilation=false "/p:CscToolPath=$testVs/MSBuild/Current/Bin/Roslyn" /p:CscToolExe=csc.exe
```

Unit tests use a deliberately small Word double. They verify the production helper/workflow logic, not Word's COM implementation, live-range behavior, or Undo behavior.

After building, `./eng/test-word.ps1` runs targeted integration checks against a separate hidden Word instance using newly created documents, closed without saving. This checks actual COM document identity, live-range movement, stale/closed ranges, selection boundaries, underline restoration, and Word's final paragraph behavior. It requires Word and is intentionally separate from CI's host-independent checks.

## Smoke tests in Word

1. Check document A, activate B with different content at the same positions, and invoke Replace All. B must remain unchanged; skipped corrections must be reported.
2. Check A, insert text before an error, then replace it. The original error should be corrected at its live position. Edit the error itself before replacing: the stale correction must be skipped. Repeat with an open suggestions dialog and after closing A.
3. Apply deliberate single, red wavy, and green wavy underlines. Check, ignore an error, replace another, save, and close. Deliberate formatting must survive; corrected text must not inherit the add-in's mark. Repeat with mixed formatting and Track Changes enabled.
4. Select one letter, a space, and a table row before conversion. Confirm the scope and verify surrounding text, paragraph/table boundaries, and Undo. Test a whole-document conversion ending in spaces and paragraph breaks.
5. Exercise shortcuts in Word, in another application, in a Word modal dialog, and in an add-in form. Test holding a key and two separate Word processes. Only the active Word editing window should run a command.
6. Use Tab/Space to change export toggles and check announced names/states with Narrator. Resize the import and About dialogs; verify that all content is reachable at 100–400% scaling.

Local certificate values belong in the ignored `Directory.Build.props`. A certificate is required for VSTO manifests; source compilation alone does not require signing. Do not commit the props file or certificate.
