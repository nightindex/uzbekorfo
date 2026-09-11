# Display compatibility

The add-in uses 96-DPI design coordinates and scales its own top-level windows, controls, fonts, and custom painting for the monitor DPI. Resolution alone is not the deciding factor: a 4K display at 200% scaling has a much smaller usable layout area than a 4K display at 100%.

## Implementation

- `UI/DpiForm.cs` handles initial scaling and monitor DPI changes. It scales explicit control geometry without enabling application-wide WinForms autoscaling.
- `UI/DpiLayout.cs` scopes the thread DPI context when creating and positioning add-in windows, then restores it. It does not change Word's process DPI policy or configuration.
- `UI/Controls/DpiDataGridView.cs` creates grid handles in their actual parent's DPI context and repairs early-created handles left in a hidden WinForms parking window. This matters when Office opens a dialog from a differently scaled callback thread.
- `UI/ModernForm.cs` lets content resize down to the dialog's minimum layout. Outer scrollbars are enabled only below that minimum. Scroll extents are recalculated after resizing, so previous DPI values cannot leave phantom scrollbars behind.
- Long message text can scroll vertically inside the message dialog. Data grids and information pages retain their own content scrolling. Short messages and normally sized dialog frames should not scroll.
- Popup positioning uses the relevant monitor's working area, including monitors to the left or above the primary monitor.

The Office hosting approach follows [Microsoft's high-DPI guidance for Office solutions](https://learn.microsoft.com/en-us/office/client-developer/ddpi/handle-high-dpi-and-dpi-scaling-in-your-office-solution).

## Automated checks

Run `./eng/test-ui.ps1` on Windows with Visual Studio installed. The harness links the actual UI source, creates transparent test windows, and runs layout/paint checks at 96, 120, 144, 192, 240, 288, and 384 DPI (100–400%). It covers short/long messages, import dialogs, the information dialog, exceptions with in-memory test data, and common controls. No Word document or user dictionary is opened or modified.

Checks include normal-size scrollbar absence, small-window scrolling, grid visibility, button widths, font sizes, DPI round trips, and bounds fitting on small and negative-coordinate/8K monitor rectangles. Rendered snapshots are saved under the ignored `tests/UiCompatibility/bin/` directory.

These are simulated scaling and geometry checks, not certification on physical 4K/8K hardware. Windows may cap test-window dimensions to the current monitor's maximum tracking size.

The harness also opens the errors and exceptions dialogs modally from DPI-unaware, system-aware, and per-monitor-aware thread contexts. Before any forced repaint or resize, it checks the native window parent/visibility, row counts, and spelling/grammar tab switching. The Word navigation boundary is guarded and test error entries have no COM document ranges.

The error-dialog tests verify that constructing 76 rows performs no suggestion lookups, opening loads only the selected entry, and revisiting an entry uses its cached suggestions. Its grouped footer is tested at an 800-logical-pixel width across 100–400% scaling, including button bounds and absence of unnecessary footer scrollbars.

Exception-brand typography is checked through repeated 125% → 200% → 150% → 400% → 125% → 100% transitions. Grid cell styles are replaced after native layout processing so WinForms cannot leave the old point-based font in place; brand names retain their intended 11-point logical size.

## Manual checks in Word

1. Rebuild the add-in and restart Word to load the new assembly.
2. Open the success message and exceptions manager at 100%, 125%, 150%, and 200%. Their outer frames should not scroll when the content fits; a populated grid may need its own vertical scrollbar.
3. Check long messages, many grid rows, dictionary editing, and all information pages. Reach the last item and all action buttons by mouse and keyboard.
4. Move resizable dialogs between monitors with different scaling, including a monitor left of the primary one. Check text sharpness, stable control sizes, and accessible close buttons.
5. On small work areas, confirm necessary scrollbars expose all content. Resize larger again and confirm unnecessary scrollbars disappear.
6. Repeat at the intended 4K/8K resolutions and 250–400% scaling before release. Also check the supported Office versions and Office's display-compatibility settings.
