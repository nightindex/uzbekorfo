using System;
using System.Drawing;
using Microsoft.Win32;

namespace UzbekOrfoAddIn.UI
{
    /// <summary>
    /// Manages the application theme — provides all colors, fonts, and spacing values
    /// from a central design token system. Supports Light and Dark themes automatically
    /// by detecting Word's current Office theme.
    /// </summary>
    public static class ThemeManager
    {
        // =====================================================================
        //  THEME DETECTION
        // =====================================================================

        private static bool? _isDarkTheme;

        /// <summary>
        /// Detects whether Word is currently using a dark theme.
        /// </summary>
        public static bool IsDarkTheme
        {
            get
            {
                if (!_isDarkTheme.HasValue)
                    _isDarkTheme = DetectDarkTheme();
                return _isDarkTheme.Value;
            }
        }

        /// <summary>
        /// Forces a theme refresh (call when Office theme may have changed).
        /// </summary>
        public static void RefreshTheme()
        {
            _isDarkTheme = DetectDarkTheme();
        }

        private static bool DetectDarkTheme()
        {
            try
            {
                // Office 2016+ theme setting
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\Common"))
                {
                    var theme = key?.GetValue("UI Theme");
                    if (theme is int themeInt)
                        return themeInt >= 4; // 4 = Dark Gray, 5 = Black
                }
            }
            catch { }
            return false;
        }

        // =====================================================================
        //  COLOR TOKENS
        // =====================================================================

        // --- Primary ---
        public static Color Primary => Color.FromArgb(37, 99, 235);           // #2563EB
        public static Color PrimaryHover => Color.FromArgb(29, 78, 216);      // #1D4ED8
        public static Color PrimaryPressed => Color.FromArgb(30, 64, 175);    // #1E40AF
        public static Color PrimaryLight => Color.FromArgb(219, 234, 254);    // #DBEAFE
        public static Color PrimaryText => Color.White;

        // --- Accent (Uzbek flag-inspired) ---
        public static Color Accent => Color.FromArgb(14, 165, 233);           // #0EA5E9

        // --- Semantic ---
        public static Color Success => Color.FromArgb(22, 163, 74);           // #16A34A
        public static Color SuccessLight => IsDarkTheme 
            ? Color.FromArgb(20, 83, 45) : Color.FromArgb(220, 252, 231);
        public static Color Error => Color.FromArgb(220, 38, 38);             // #DC2626
        public static Color ErrorLight => IsDarkTheme 
            ? Color.FromArgb(69, 10, 10) : Color.FromArgb(254, 242, 242);
        public static Color Warning => Color.FromArgb(245, 158, 11);          // #F59E0B
        public static Color WarningLight => IsDarkTheme 
            ? Color.FromArgb(113, 63, 18) : Color.FromArgb(254, 252, 232);

        // --- Surfaces ---
        public static Color Background => IsDarkTheme
            ? Color.FromArgb(30, 30, 30)       // #1E1E1E
            : Color.White;                      // #FFFFFF

        public static Color Surface => IsDarkTheme
            ? Color.FromArgb(45, 45, 45)        // #2D2D2D
            : Color.FromArgb(248, 250, 252);    // #F8FAFC

        public static Color SurfaceHover => IsDarkTheme
            ? Color.FromArgb(61, 61, 61)        // #3D3D3D
            : Color.FromArgb(241, 245, 249);    // #F1F5F9

        public static Color SurfaceElevated => IsDarkTheme
            ? Color.FromArgb(55, 55, 55)        // #373737
            : Color.White;

        // --- Borders ---
        public static Color Border => IsDarkTheme
            ? Color.FromArgb(64, 64, 64)        // #404040
            : Color.FromArgb(226, 232, 240);    // #E2E8F0

        public static Color BorderFocus => Primary;

        // --- Text ---
        public static Color TextPrimary => IsDarkTheme
            ? Color.FromArgb(241, 245, 249)     // #F1F5F9
            : Color.FromArgb(15, 23, 42);       // #0F172A

        public static Color TextSecondary => IsDarkTheme
            ? Color.FromArgb(148, 163, 184)     // #94A3B8
            : Color.FromArgb(100, 116, 139);    // #64748B

        public static Color TextDisabled => IsDarkTheme
            ? Color.FromArgb(82, 82, 82) 
            : Color.FromArgb(203, 213, 225);

        public static Color TextOnPrimary => Color.White;

        // --- Selection ---
        public static Color SelectedRow => IsDarkTheme
            ? Color.FromArgb(30, 58, 95)        // #1E3A5F
            : Color.FromArgb(239, 246, 255);    // #EFF6FF

        public static Color SelectedBorder => Primary;

        // --- Buttons ---
        public static Color ButtonSecondary => IsDarkTheme
            ? Color.FromArgb(55, 65, 81)        // #374151
            : Color.FromArgb(241, 245, 249);    // #F1F5F9

        public static Color ButtonSecondaryHover => IsDarkTheme
            ? Color.FromArgb(75, 85, 99)        // #4B5563
            : Color.FromArgb(226, 232, 240);    // #E2E8F0

        public static Color ButtonDanger => Color.FromArgb(239, 68, 68);       // #EF4444
        public static Color ButtonDangerHover => Color.FromArgb(220, 38, 38);  // #DC2626
        public static Color ButtonDangerPressed => Color.FromArgb(185, 28, 28); // #B91C1C

        // =====================================================================
        //  TYPOGRAPHY TOKENS
        // =====================================================================

        public static string FontFamily => "Segoe UI";
        public static string FontFamilyMono => "Cascadia Code";
        public static string FontFamilyUzbek => "Times New Roman";

        // Cached font instances — created once, reused everywhere.
        // IMPORTANT: Never call Dispose() on these shared instances.
        private static readonly Font _fontXS = new Font("Segoe UI", 8.25f, FontStyle.Regular);
        private static readonly Font _fontSM = new Font("Segoe UI", 9f, FontStyle.Regular);
        private static readonly Font _fontBase = new Font("Segoe UI", 9.75f, FontStyle.Regular);
        private static readonly Font _fontLG = new Font("Segoe UI", 11.25f, FontStyle.Regular);
        private static readonly Font _fontXL = new Font("Segoe UI", 13.5f, FontStyle.Regular);
        private static readonly Font _font2XL = new Font("Segoe UI", 18f, FontStyle.Regular);
        private static readonly Font _fontBaseBold = new Font("Segoe UI", 9.75f, FontStyle.Bold);
        private static readonly Font _fontLGBold = new Font("Segoe UI", 11.25f, FontStyle.Bold);
        private static readonly Font _fontXLBold = new Font("Segoe UI", 13.5f, FontStyle.Bold);
        private static readonly Font _fontMono = new Font("Cascadia Code", 9f, FontStyle.Regular);

        public static Font FontXS => _fontXS;
        public static Font FontSM => _fontSM;
        public static Font FontBase => _fontBase;
        public static Font FontLG => _fontLG;
        public static Font FontXL => _fontXL;
        public static Font Font2XL => _font2XL;

        public static Font FontBaseBold => _fontBaseBold;
        public static Font FontLGBold => _fontLGBold;
        public static Font FontXLBold => _fontXLBold;

        public static Font FontMono => _fontMono;

        // =====================================================================
        //  SPACING TOKENS (in pixels)
        // =====================================================================

        public static int SpaceXS => 4;
        public static int SpaceSM => 8;
        public static int SpaceMD => 12;
        public static int SpaceLG => 16;
        public static int SpaceXL => 24;
        public static int Space2XL => 32;

        // =====================================================================
        //  CORNER RADIUS TOKENS
        // =====================================================================

        public static int RadiusSM => 4;
        public static int RadiusMD => 8;
        public static int RadiusLG => 12;
        public static int RadiusXL => 16;

        // =====================================================================
        //  SHADOW HELPERS (applied via custom painting)
        // =====================================================================

        public static Color ShadowColor => IsDarkTheme
            ? Color.FromArgb(40, 0, 0, 0)
            : Color.FromArgb(18, 0, 0, 0);

        public static Color ShadowColorMedium => IsDarkTheme
            ? Color.FromArgb(60, 0, 0, 0)
            : Color.FromArgb(30, 0, 0, 0);

        // =====================================================================
        //  SCROLLBAR TOKENS
        // =====================================================================

        public static Color ScrollTrack => IsDarkTheme
            ? Color.FromArgb(38, 38, 38)        // #262626
            : Color.FromArgb(245, 245, 245);    // #F5F5F5

        public static Color ScrollThumb => IsDarkTheme
            ? Color.FromArgb(80, 80, 80)        // #505050
            : Color.FromArgb(200, 200, 200);    // #C8C8C8

        public static Color ScrollThumbHover => IsDarkTheme
            ? Color.FromArgb(110, 110, 110)     // #6E6E6E
            : Color.FromArgb(170, 170, 170);    // #AAAAAA

        public static Color ScrollThumbPressed => IsDarkTheme
            ? Color.FromArgb(130, 130, 130)     // #828282
            : Color.FromArgb(140, 140, 140);    // #8C8C8C
    }
}
