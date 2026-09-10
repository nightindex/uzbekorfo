using System;
using System.Drawing;
using System.Drawing.Text;
using System.Globalization;
using System.Windows.Forms;

namespace UzbekOrfoAddIn.UI.Controls
{
    /// <summary>
    /// A label that properly renders color emoji on Windows 10+.
    /// 
    /// Standard WinForms Label uses GDI (TextRenderer.DrawText) which cannot
    /// display colour emoji — they appear as monochrome outlines or missing glyphs.
    /// 
    /// This control uses GDI+ (Graphics.DrawString) and automatically renders
    /// emoji code-points with the "Segoe UI Emoji" font while keeping the rest
    /// of the text in the control's regular font.
    /// </summary>
    public class EmojiLabel : Control
    {
        // Cached emoji font — created once, matches the control's current size.
        private Font _emojiFont;
        private float _lastEmojiFontSize;

        public EmojiLabel()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.UserPaint |
                     ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.SupportsTransparentBackColor |
                     ControlStyles.ResizeRedraw, true);

            BackColor = Color.Transparent;
            DoubleBuffered = true;
        }

        // =====================================================================
        //  AUTO-SIZE support
        // =====================================================================

        private bool _autoSizeEnabled = true;

        /// <summary>Label auto-sizes to fit its text (like the standard Label).</summary>
        public override bool AutoSize
        {
            get => _autoSizeEnabled;
            set
            {
                _autoSizeEnabled = value;
                if (value) RecalcSize();
            }
        }

        private Size _maximumSize = Size.Empty;
        public new Size MaximumSize
        {
            get => _maximumSize;
            set { _maximumSize = value; if (_autoSizeEnabled) RecalcSize(); }
        }

        private ContentAlignment _textAlign = ContentAlignment.TopLeft;
        /// <summary>
        /// Gets or sets the alignment of text within the control.
        /// Supports horizontal (Left/Center/Right) and vertical (Top/Middle/Bottom).
        /// </summary>
        public ContentAlignment TextAlign
        {
            get => _textAlign;
            set { _textAlign = value; Invalidate(); }
        }

        protected override void OnTextChanged(EventArgs e)
        {
            base.OnTextChanged(e);
            if (_autoSizeEnabled) RecalcSize();
            Invalidate();
        }

        protected override void OnFontChanged(EventArgs e)
        {
            base.OnFontChanged(e);
            if (_autoSizeEnabled) RecalcSize();
            Invalidate();
        }

        protected override void OnPaddingChanged(EventArgs e)
        {
            base.OnPaddingChanged(e);
            if (_autoSizeEnabled) RecalcSize();
            Invalidate();
        }

        private void RecalcSize()
        {
            if (string.IsNullOrEmpty(Text)) return;

            using (var g = CreateGraphics())
            {
                g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                var maxW = _maximumSize.Width > 0 ? _maximumSize.Width - Padding.Horizontal : 9999;
                var sz = MeasureTextMixed(g, Text, Font, maxW);

                int w = (int)Math.Ceiling(sz.Width) + Padding.Horizontal + 2;
                int h = (int)Math.Ceiling(sz.Height) + Padding.Vertical + 2;

                if (_maximumSize.Width > 0 && w > _maximumSize.Width)
                    w = _maximumSize.Width;

                Size = new Size(w, h);
            }
        }

        // =====================================================================
        //  PAINTING
        // =====================================================================

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

            if (string.IsNullOrEmpty(Text)) return;

            float maxW = Width - Padding.Horizontal;
            var textSize = MeasureTextMixed(g, Text, Font, maxW);

            // Horizontal position
            float x;
            switch (_textAlign)
            {
                case ContentAlignment.TopCenter:
                case ContentAlignment.MiddleCenter:
                case ContentAlignment.BottomCenter:
                    x = Padding.Left + (maxW - textSize.Width) / 2f;
                    break;
                case ContentAlignment.TopRight:
                case ContentAlignment.MiddleRight:
                case ContentAlignment.BottomRight:
                    x = Width - Padding.Right - textSize.Width;
                    break;
                default: // Left
                    x = Padding.Left;
                    break;
            }

            // Vertical position
            float y;
            float availH = Height - Padding.Vertical;
            switch (_textAlign)
            {
                case ContentAlignment.MiddleLeft:
                case ContentAlignment.MiddleCenter:
                case ContentAlignment.MiddleRight:
                    y = Padding.Top + (availH - textSize.Height) / 2f;
                    break;
                case ContentAlignment.BottomLeft:
                case ContentAlignment.BottomCenter:
                case ContentAlignment.BottomRight:
                    y = Height - Padding.Bottom - textSize.Height;
                    break;
                default: // Top
                    y = Padding.Top;
                    break;
            }

            DrawTextMixed(g, Text, Font, ForeColor, x, y, maxW);
        }

        // =====================================================================
        //  MIXED TEXT + EMOJI RENDERING
        // =====================================================================

        /// <summary>
        /// Draws text segment by segment. Emoji code-points are drawn with
        /// "Segoe UI Emoji" via GDI+ (DrawString), everything else with the
        /// regular font.  Both paths use GDI+ so color emoji layers work.
        /// </summary>
        private void DrawTextMixed(Graphics g, string text, Font font, Color color, float x, float y, float maxWidth)
        {
            EnsureEmojiFont(font.Size);

            float startX = x;
            float textLineH = font.GetHeight(g);
            float emojiLineH = _emojiFont.GetHeight(g);
            float lineHeight = Math.Max(textLineH, emojiLineH);

            using (var brush = new SolidBrush(color))
            {
                int i = 0;
                while (i < text.Length)
                {
                    // Find the next run boundary (emoji vs non-emoji)
                    int runStart = i;
                    bool runIsEmoji = IsEmojiAt(text, i);

                    while (i < text.Length && IsEmojiAt(text, i) == runIsEmoji)
                    {
                        if (char.IsHighSurrogate(text, i) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                            i += 2; // surrogate pair
                        else
                            i++;
                    }

                    string run = text.Substring(runStart, i - runStart);
                    Font runFont = runIsEmoji ? _emojiFont : font;

                    // Measure this run
                    var runSize = g.MeasureString(run, runFont);

                    // Word-wrap: if we exceed maxWidth, move to next line
                    if (x + runSize.Width > startX + maxWidth && x > startX + 1)
                    {
                        x = startX;
                        y += lineHeight;
                    }

                    // Vertically center each run in the shared line box.
                    float runHeight = runIsEmoji ? emojiLineH : textLineH;
                    float drawY = y + (lineHeight - runHeight) / 2f;

                    g.DrawString(run, runFont, brush, x, drawY);
                    x += runSize.Width;
                }
            }
        }

        private SizeF MeasureTextMixed(Graphics g, string text, Font font, float maxWidth)
        {
            EnsureEmojiFont(font.Size);

            float x = 0, y = 0;
            float maxX = 0;
            float textLineH = font.GetHeight(g);
            float emojiLineH = _emojiFont.GetHeight(g);
            float lineHeight = Math.Max(textLineH, emojiLineH);

            int i = 0;
            while (i < text.Length)
            {
                int runStart = i;
                bool runIsEmoji = IsEmojiAt(text, i);

                while (i < text.Length && IsEmojiAt(text, i) == runIsEmoji)
                {
                    if (char.IsHighSurrogate(text, i) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                        i += 2;
                    else
                        i++;
                }

                string run = text.Substring(runStart, i - runStart);
                Font runFont = runIsEmoji ? _emojiFont : font;

                var runSize = g.MeasureString(run, runFont);

                if (x + runSize.Width > maxWidth && x > 1)
                {
                    x = 0;
                    y += lineHeight;
                }

                x += runSize.Width;
                if (x > maxX) maxX = x;
            }

            return new SizeF(maxX, y + lineHeight);
        }

        // =====================================================================
        //  EMOJI DETECTION
        // =====================================================================

        /// <summary>
        /// Returns true if the character at position i is an emoji
        /// (surrogate-pair emoji, common emoji ranges, or variation selectors).
        /// </summary>
        private static bool IsEmojiAt(string text, int i)
        {
            char c = text[i];

            // Surrogate pair — decode and check
            if (char.IsHighSurrogate(c) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
            {
                int cp = char.ConvertToUtf32(c, text[i + 1]);
                return IsEmojiCodePoint(cp);
            }

            // BMP emoji ranges
            return IsEmojiCodePoint(c);
        }

        private static bool IsEmojiCodePoint(int cp)
        {
            // Common emoji / symbol ranges:
            // Emoticons, Dingbats, Symbols, Transport, Misc, etc.
            if (cp >= 0x1F600 && cp <= 0x1F64F) return true;  // Emoticons
            if (cp >= 0x1F300 && cp <= 0x1F5FF) return true;  // Misc Symbols & Pictographs
            if (cp >= 0x1F680 && cp <= 0x1F6FF) return true;  // Transport & Map
            if (cp >= 0x1F900 && cp <= 0x1F9FF) return true;  // Supplemental Symbols
            if (cp >= 0x1FA00 && cp <= 0x1FA6F) return true;  // Chess symbols
            if (cp >= 0x1FA70 && cp <= 0x1FAFF) return true;  // Symbols Extended-A
            if (cp >= 0x2600 && cp <= 0x26FF) return true;    // Misc Symbols
            if (cp >= 0x2700 && cp <= 0x27BF) return true;    // Dingbats
            if (cp >= 0x231A && cp <= 0x23FA) return true;    // Misc Technical (subset)
            if (cp >= 0xFE00 && cp <= 0xFE0F) return true;    // Variation selectors
            if (cp == 0x200D) return true;                     // ZWJ
            if (cp >= 0x2702 && cp <= 0x27B0) return true;    // Dingbats
            if (cp >= 0x1F1E0 && cp <= 0x1F1FF) return true;  // Regional indicators (flags)
            if (cp == 0x25B6 || cp == 0x25C0) return true;    // Play/reverse buttons
            if (cp == 0x2B05 || cp == 0x2B06 || cp == 0x2B07) return true; // Arrows
            if (cp == 0x2B1B || cp == 0x2B1C || cp == 0x2B50 || cp == 0x2B55) return true;
            if (cp == 0x3030 || cp == 0x303D) return true;
            if (cp == 0x2049 || cp == 0x203C) return true;    // Exclamation marks
            if (cp >= 0x2194 && cp <= 0x2199) return true;    // Arrows
            if (cp >= 0x21A9 && cp <= 0x21AA) return true;    // Arrows
            if (cp >= 0x23E9 && cp <= 0x23F3) return true;    // Media controls
            if (cp == 0x23CF) return true;                     // Eject
            if (cp == 0x24C2) return true;                     // Circled M
            if (cp >= 0x25AA && cp <= 0x25AB) return true;    // Squares
            if (cp == 0x25FB || cp == 0x25FC || cp == 0x25FD || cp == 0x25FE) return true;
            if (cp >= 0x2934 && cp <= 0x2935) return true;    // Arrows
            if (cp >= 0x2122 && cp <= 0x2139) return true;    // TM, info
            return false;
        }

        // =====================================================================
        //  FONT CACHE
        // =====================================================================

        private void EnsureEmojiFont(float size)
        {
            if (_emojiFont == null || Math.Abs(_lastEmojiFontSize - size) > 0.01f)
            {
                _emojiFont?.Dispose();
                _emojiFont = new Font("Segoe UI Emoji", size, FontStyle.Regular);
                _lastEmojiFontSize = size;
            }
        }

        // =====================================================================
        //  CLEANUP
        // =====================================================================

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _emojiFont?.Dispose();
                _emojiFont = null;
            }
            base.Dispose(disposing);
        }
    }
}
