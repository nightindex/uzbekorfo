using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Reflection;
using System.Windows.Forms;
using UzbekOrfoAddIn.UI;
using UzbekOrfoAddIn.UI.Controls;

namespace UzbekOrfoAddIn.Forms
{
    /// <summary>
    /// In-app information dialog with sidebar navigation.
    /// Tabs: Дастур ҳақида, Қўлланма, Махфиялик, Қўллаб-қувватлаш, Фикр билдириш
    /// </summary>
    public class AppInfoForm : ModernForm
    {
        private const int InfoScrollBarWidth = 8;

        private Panel _sidebar;
        private Panel _contentArea;
        private Panel[] _pages;
        private Panel[] _navButtons;
        private int _activeIndex;

        private static readonly string[] TabLabels = {
            "Дастур ҳақида",
            "Қўлланма",
            "Тезкор тугмалар",
            "Махфиялик",
            "Қўллаб-қувватлаш",
            "Фикр билдириш"
        };

        // Segoe MDL2 Assets glyphs (Windows 10+)
        private static readonly string[] TabIcons = {
            "\uE946",  // Info
            "\uE736",  // Library / Book
            "\uE765",  // Keyboard
            "\uE72E",  // Lock / Shield
            "\uEB51",  // Heart
            "\uE8F2"   // Chat / Feedback
        };

        public AppInfoForm()
        {
            Title = "Маълумот";
            Size = new Size(1100, 780);
            MinimumSize = new Size(960, 680);
            ShowMinimizeButton = false;
            AllowResize = true;

            BuildUI();
        }

        private void BuildUI()
        {
            // Hide default action bar
            ActionBar.Height = 0;
            ActionBar.Visible = false;

            // Override content panel padding — we manage our own layout
            ContentPanel.Padding = new Padding(0);
            ContentPanel.AutoScroll = false;

            // === Sidebar ===
            _sidebar = new Panel
            {
                Dock = DockStyle.Left,
                Width = 270,
                BackColor = ThemeManager.Surface,
                Padding = new Padding(0, ThemeManager.SpaceSM, 0, ThemeManager.SpaceSM)
            };

            // Sidebar separator
            var sidebarSep = new Panel
            {
                Dock = DockStyle.Left,
                Width = 1,
                BackColor = ThemeManager.Border
            };

            // === Content area ===
            _contentArea = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ThemeManager.Background,
                Padding = new Padding(ThemeManager.SpaceXL, ThemeManager.SpaceLG,
                                      ThemeManager.SpaceXL, ThemeManager.SpaceLG),
                AutoScroll = true
            };

            ContentPanel.Controls.Add(_contentArea);
            ContentPanel.Controls.Add(sidebarSep);
            ContentPanel.Controls.Add(_sidebar);

            // Create pages
            _pages = new Panel[TabLabels.Length];
            _pages[0] = BuildAboutPage();
            _pages[1] = BuildHelpPage();
            _pages[2] = BuildHotkeysPage();
            _pages[3] = BuildPrivacyPage();
            _pages[4] = BuildDonatePage();
            _pages[5] = BuildFeedbackPage();

            foreach (var page in _pages)
            {
                page.Dock = DockStyle.Fill;
                page.Visible = false;
                _contentArea.Controls.Add(page);
            }

            // Create nav buttons
            _navButtons = new Panel[TabLabels.Length];
            for (int i = 0; i < TabLabels.Length; i++)
            {
                _navButtons[i] = CreateNavButton(TabIcons[i], TabLabels[i], i);
                _sidebar.Controls.Add(_navButtons[i]);
                // Stack top-to-bottom — add in reverse because Dock.Top stacks last-added at top
            }
            // Re-order so first tab is at top
            for (int i = TabLabels.Length - 1; i >= 0; i--)
                _navButtons[i].BringToFront();

            // Show first tab
            SetActiveTab(0);
        }

        // =================================================================
        //  SIDEBAR NAVIGATION
        // =================================================================

        private Panel CreateNavButton(string icon, string label, int index)
        {
            var btn = new Panel
            {
                Dock = DockStyle.Top,
                Height = 50,
                Cursor = Cursors.Hand,
                BackColor = ThemeManager.Surface,
                Padding = new Padding(ThemeManager.SpaceLG, 0, ThemeManager.SpaceSM, 0),
                Tag = index
            };

            btn.Paint += (s, e) =>
            {
                var g = e.Graphics;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                g.SmoothingMode = SmoothingMode.AntiAlias;

                bool isActive = (int)btn.Tag == _activeIndex;

                // Active indicator bar
                if (isActive)
                {
                    using (var brush = new SolidBrush(ThemeManager.Primary))
                        g.FillRectangle(brush, 0, 4, 3, btn.Height - 8);

                    using (var bgBrush = new SolidBrush(ThemeManager.SurfaceHover))
                        g.FillRectangle(bgBrush, 3, 0, btn.Width - 3, btn.Height);
                }

                // Icon (Segoe MDL2 Assets for crisp glyph rendering)
                Color textColor = isActive ? ThemeManager.Primary : ThemeManager.TextSecondary;
                using (var iconFont = new Font("Segoe MDL2 Assets", 15f, FontStyle.Regular))
                using (var iconBrush = new SolidBrush(textColor))
                {
                    g.DrawString(icon, iconFont, iconBrush, 12, 13);
                }

                // Label text
                Color labelColor = isActive ? ThemeManager.TextPrimary : ThemeManager.TextSecondary;
                using (var labelFont = new Font("Segoe UI", 12.5f,
                    isActive ? FontStyle.Bold : FontStyle.Regular))
                using (var labelBrush = new SolidBrush(labelColor))
                {
                    g.DrawString(label, labelFont, labelBrush, 44, 13);
                }
            };

            btn.MouseEnter += (s, e) =>
            {
                if ((int)btn.Tag != _activeIndex)
                    btn.BackColor = ThemeManager.SurfaceHover;
            };
            btn.MouseLeave += (s, e) =>
            {
                btn.BackColor = ThemeManager.Surface;
            };
            btn.Click += (s, e) => SetActiveTab(index);

            return btn;
        }

        private void SetActiveTab(int index)
        {
            _activeIndex = index;
            for (int i = 0; i < _pages.Length; i++)
            {
                _pages[i].Visible = i == index;
                _navButtons[i].BackColor = ThemeManager.Surface;
                _navButtons[i].Invalidate();
            }
        }

        // =================================================================
        //  TAB 1: ДАСТУР ҲАҚИДА (About)
        // =================================================================

        private Panel BuildAboutPage()
        {
            var page = new Panel { BackColor = ThemeManager.Background };

            var flow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = false,
                AutoScroll = true,
                Padding = new Padding(ThemeManager.SpaceMD)
            };

            // App name + version
            string version = "1.0.0";
            try
            {
                var asm = Assembly.GetExecutingAssembly();
                var ver = asm.GetName().Version;
                if (ver != null) version = $"{ver.Major}.{ver.Minor}.{ver.Build}";
            }
            catch { }

            flow.Controls.Add(CreateSection("Ўзбек Орфо", 18f, FontStyle.Bold, ThemeManager.Primary));
            flow.Controls.Add(CreateSection($"Версия: {version}", 11f, FontStyle.Regular, ThemeManager.TextSecondary));
            flow.Controls.Add(CreateSpacer(12));
            flow.Controls.Add(CreateParagraph(
                "Ўзбек Орфо — Microsoft Word учун мўлжалланган ўзбек тилида имло ва " +
                "грамматика текширувчи дастур. Дастур сўзларнинг тўғри ёзилишини " +
                "текширади, хатоларни тузатиш учун таклифлар беради ва ўзбек тилига " +
                "хос бўлган қоидаларни қўллаб-қувватлайди."));
            flow.Controls.Add(CreateSpacer(16));
            flow.Controls.Add(CreateSection("Имкониятлар:", 12f, FontStyle.Bold, ThemeManager.TextPrimary));
            flow.Controls.Add(CreateSpacer(4));
            flow.Controls.Add(CreateBulletList(
                "Имло текшируви (бутун ҳужжат ва танланган матн)",
                "Грамматика текшируви (қоидаларга асосланган)",
                "Транслитерация (Лотин ↔ Кирилл)",
                "Луғатни бошқариш (қўшиш, таҳрирлаш, импорт)",
                "Сўз изоҳлари ва маънолари",
                "Махсус белгилар (тутуқ белгиси)",
                "Бўш жойларни тозалаш",
                "Шрифт ўрнатиш (Times New Roman)"));
            flow.Controls.Add(CreateSpacer(16));
            flow.Controls.Add(CreateSection("Тизим талаблари:", 12f, FontStyle.Bold, ThemeManager.TextPrimary));
            flow.Controls.Add(CreateSpacer(4));
            flow.Controls.Add(CreateBulletList(
                "Windows 10 ёки ундан янги",
                "Microsoft Word 2013 / 2016 / 2019 / 365",
                ".NET Framework 4.7.2"));
            flow.Controls.Add(CreateSpacer(16));
            flow.Controls.Add(CreateParagraph("© 2026 Ўзбек Орфо. Барча ҳуқуқлар ҳимояланган."));

            ModernScrollBar.AttachTo(flow, InfoScrollBarWidth);

            page.Controls.Add(flow);
            return page;
        }

        // =================================================================
        //  TAB 2: ҚЎЛЛАНМА (Help / Usage Guide)
        // =================================================================

        private Panel BuildHelpPage()
        {
            var page = new Panel { BackColor = ThemeManager.Background };

            var flow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = false,
                AutoScroll = true,
                Padding = new Padding(ThemeManager.SpaceMD)
            };

            flow.Controls.Add(CreateSection("Фойдаланиш қўлланмаси", 16f, FontStyle.Bold, ThemeManager.Primary));
            flow.Controls.Add(CreateSpacer(12));

            // 1. Текшириш group
            flow.Controls.Add(CreateSection("1. Текшириш", 12f, FontStyle.Bold, ThemeManager.TextPrimary));
            flow.Controls.Add(CreateSpacer(4));
            flow.Controls.Add(CreateParagraph(
                "• 'Матн текшируви' — ҳужжатдаги имло ва грамматик хатоларни текширади. " +
                "Имло хатолари қизил, грамматик хатолар яшил тўлқинли чизиқ билан белгиланади.\n\n" +
                "• 'Хатоларни кўриш' — охирги текширувда топилган хатолар рўйхатини " +
                "диалог ойнасида кўрсатади. Хатоларни тез кўриб чиқиш ва ўчириш учун.\n\n" +
                "• 'Хатолар рўйхати' — топилган барча хатоларни янги ҳужжатда батафсил " +
                "рўйхат шаклида тайёрлайди (сақлаш ва чоп этиш мумкин)."));
            flow.Controls.Add(CreateSpacer(8));

            // 2. Тузатиш group
            flow.Controls.Add(CreateSection("2. Тузатиш", 12f, FontStyle.Bold, ThemeManager.TextPrimary));
            flow.Controls.Add(CreateSpacer(4));
            flow.Controls.Add(CreateParagraph(
                "• 'Вариантлар' — белгиланган хато сўз учун тузатиш вариантларини кўрсатади. " +
                "Хато сўзни тез алмаштириш учун.\n\n" +
                "• 'Барчасини алмаштириш' — танланган хато сўзнинг ҳужжатдаги барча учрашларини " +
                "бирданига тўғри вариантга алмаштиради. Такрорланган хатоларни тез тузатиш учун.\n\n" +
                "• 'Авто тузатиш' — автоматик тузатиш режимини ёқади/ўчиради. " +
                "Текширув пайтида энг яхши вариантни ўзи қўяди."));
            flow.Controls.Add(CreateSpacer(8));

            // 3. Контекст менюси
            flow.Controls.Add(CreateSection("3. Контекст менюси", 12f, FontStyle.Bold, ThemeManager.TextPrimary));
            flow.Controls.Add(CreateSpacer(4));
            flow.Controls.Add(CreateParagraph(
                "Хато ёзилган сўз устида сичқончанинг ўнг тугмасини босганингизда тузатиш таклифлари " +
                "менюда пайдо бўлади. Шу ердан сўзни луғатга қўшиш ёки изоҳини кўриш ҳам мумкин."));
            flow.Controls.Add(CreateSpacer(8));

            // 4. Луғат group
            flow.Controls.Add(CreateSection("4. Луғат", 12f, FontStyle.Bold, ThemeManager.TextPrimary));
            flow.Controls.Add(CreateSpacer(4));
            flow.Controls.Add(CreateParagraph(
                "• 'Луғатга қўшиш' — белгиланган сўзни шахсий луғатга қўшади. " +
                "Кейинги текширувларда хато деб ҳисобланмайди. Тўғри, лекин кам учрайдиган сўзлар учун.\n\n" +
                "• 'Луғатни таҳрирлаш' — шахсий луғатдаги сўзлар рўйхатини кўриш, " +
                "ўчириш ва таҳрирлаш ойнасини очади.\n\n" +
                "• 'Сўз қўшиш' — ташқи файл (.txt, .docx ва ҳ.к.) дан янги сўзларни ўқиб, " +
                "луғатга қўшади. Катта луғат яратиш ёки кенгайтириш учун."));
            flow.Controls.Add(CreateSpacer(8));

            // 5. Транслитерация group
            flow.Controls.Add(CreateSection("5. Транслитерация", 12f, FontStyle.Bold, ThemeManager.TextPrimary));
            flow.Controls.Add(CreateSpacer(4));
            flow.Controls.Add(CreateParagraph(
                "• 'Лотиндан Кириллга' — танланган матн ёки бутун ҳужжатни бир босишда кириллга ўтказади.\n\n" +
                "• 'Кириллдан Лотинга' — танланган матн ёки бутун ҳужжатни бир босишда лотинга ўтказади.\n\n" +
                "• 'Истиснолар' — транслитерация учун истисно сўзлар рўйхатини очади. " +
                "Исмлар, брендлар ва махсус атамаларни қўшиш ёки таҳрирлаш учун."));
            flow.Controls.Add(CreateSpacer(8));

            // 6. Воситалар group
            flow.Controls.Add(CreateSection("6. Воситалар", 12f, FontStyle.Bold, ThemeManager.TextPrimary));
            flow.Controls.Add(CreateSpacer(4));
            flow.Controls.Add(CreateParagraph(
                "• 'Бўшлиқларни тозалаш' — ортиқча бўшлиқлар, икки марта босилган Enter ва " +
                "бошқа формат хатоларини тозалайди. Матнни тартибга солиш учун.\n\n" +
                "• 'Ёзувни алмаштириш' — ҳужжатнинг ҳозирги ёзувини аниқлаб, тескарисига " +
                "(лотин ↔ кирилл) ўтказади. Аралаш ёки номаълум ёзувни тез ўзгартириш учун.\n\n" +
                "• 'TNR ўрнатиш' — танланган матн ёки бутун ҳужжат шрифтини Times New Roman га ўрнатади. " +
                "Имло текширувининг тўғри ишлаши учун тавсия этилади."));
            flow.Controls.Add(CreateSpacer(8));

            // 7. Изоҳлар group
            flow.Controls.Add(CreateSection("7. Изоҳлар", 12f, FontStyle.Bold, ThemeManager.TextPrimary));
            flow.Controls.Add(CreateSpacer(4));
            flow.Controls.Add(CreateParagraph(
                "• 'Изоҳ' — танланган сўзнинг маъноси, имло қоидаси ёки изоҳини кўрсатади. " +
                "Сўзнинг тўғри ёзилишини тушуниш учун.\n\n" +
                "• 'Хатоларни экспорт қилиш' — топилган хатолар рўйхатини янги ҳужжатга ёки файлга сақлайди. " +
                "Ҳисобот тайёрлаш ёки бошқалар билан бўлишиш учун."));
            flow.Controls.Add(CreateSpacer(8));

            // 8. Махсус белгилар group
            flow.Controls.Add(CreateSection("8. Махсус белгилар", 12f, FontStyle.Bold, ThemeManager.TextPrimary));
            flow.Controls.Add(CreateSpacer(4));
            flow.Controls.Add(CreateParagraph(
                "• 'Тутуқ белгиси' [ʻ] — тутуқ белгисини жорий курсор жойига қўяди.\n\n" +
                "• 'Тутуқ белгиси' [ʼ] — тутуқ белгисини жорий курсор жойига қўяди.\n\n" +
                "• 'Тўғри тутуқ белгиси қўйиш' — матндаги нотўғри тутуқ белгиларини тўғри форматга келтиради.\n\n" +
                "• Ёзув тури индикатори — ҳужжатдаги матннинг ҳозирги ёзув турини кўрсатади (Кирилл ёки Лотин)."));
            flow.Controls.Add(CreateSpacer(12));

            // Shortcuts note
            flow.Controls.Add(CreateParagraph(
                "Барча амалиётлар учун тезкор клавиатура тугмалари мавжуд. " +
                "Тўлиқ рўйхат учун чап томондаги 'Тезкор тугмалар' бўлимига ўтинг."));
            flow.Controls.Add(CreateSpacer(8));

            ModernScrollBar.AttachTo(flow, InfoScrollBarWidth);

            page.Controls.Add(flow);
            return page;
        }

        // =================================================================
        //  TAB 3: ТЕЗКОР ТУГМАЛАР (Hotkeys)
        // =================================================================

        private Panel BuildHotkeysPage()
        {
            var page = new Panel { BackColor = ThemeManager.Background };

            var flow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = false,
                AutoScroll = true,
                Padding = new Padding(ThemeManager.SpaceMD)
            };

            flow.Controls.Add(CreateSection("Тезкор тугмалар", 16f, FontStyle.Bold, ThemeManager.Primary));
            flow.Controls.Add(CreateSpacer(8));
            flow.Controls.Add(CreateParagraph(
                "Қуйидаги тугмачалар бирикмасини босиб, тегишли амалиётларни тезроқ бажаришингиз мумкин."));
            flow.Controls.Add(CreateSpacer(16));

            // ── Текшириш ────────────────────────────────────────────────
            flow.Controls.Add(CreateHotkeyGroupHeader("Текшириш"));
            flow.Controls.Add(CreateSpacer(6));
            flow.Controls.Add(CreateHotkeyRow("Ctrl + Alt + Q", "Матн текшируви"));
            flow.Controls.Add(CreateHotkeyRow("Ctrl + Alt + E", "Хатоларни кўриш"));
            flow.Controls.Add(CreateHotkeyRow("Ctrl + Alt + W", "Хатолар рўйхати"));
            flow.Controls.Add(CreateSpacer(12));

            // ── Тузатиш ─────────────────────────────────────────────────
            flow.Controls.Add(CreateHotkeyGroupHeader("Тузатиш"));
            flow.Controls.Add(CreateSpacer(6));
            flow.Controls.Add(CreateHotkeyRow("Ctrl + Alt + J", "Вариантлар"));
            flow.Controls.Add(CreateHotkeyRow("Ctrl + Alt + H", "Барчасини алмаштириш"));
            flow.Controls.Add(CreateHotkeyRow("Ctrl + Alt + A", "Авто тузатиш on/off"));
            flow.Controls.Add(CreateSpacer(12));

            // ── Луғат ───────────────────────────────────────────────────
            flow.Controls.Add(CreateHotkeyGroupHeader("Луғат"));
            flow.Controls.Add(CreateSpacer(6));
            flow.Controls.Add(CreateHotkeyRow("Ctrl + Alt + B", "Луғатга қўшиш"));
            flow.Controls.Add(CreateHotkeyRow("Ctrl + Alt + G", "Луғатни таҳрирлаш"));
            flow.Controls.Add(CreateHotkeyRow("Ctrl + Alt + Y", "Сўз қўшиш (файлдан)"));
            flow.Controls.Add(CreateSpacer(12));

            // ── Транслитерация ──────────────────────────────────────────
            flow.Controls.Add(CreateHotkeyGroupHeader("Транслитерация"));
            flow.Controls.Add(CreateSpacer(6));
            flow.Controls.Add(CreateHotkeyRow("Ctrl + Alt + Shift + C", "Лотиндан Кириллга"));
            flow.Controls.Add(CreateHotkeyRow("Ctrl + Alt + Shift + K", "Кириллдан Лотинга"));
            flow.Controls.Add(CreateHotkeyRow("Ctrl + Alt + X", "Транслитерация истиснолари"));
            flow.Controls.Add(CreateSpacer(12));

            // ── Воситалар ───────────────────────────────────────────────
            flow.Controls.Add(CreateHotkeyGroupHeader("Воситалар"));
            flow.Controls.Add(CreateSpacer(6));
            flow.Controls.Add(CreateHotkeyRow("Ctrl + Alt + Shift + B", "Бўшлиқларни тозалаш"));
            flow.Controls.Add(CreateHotkeyRow("Ctrl + Alt + U", "Ёзувни алмаштириш"));
            flow.Controls.Add(CreateHotkeyRow("Ctrl + Alt + Shift + N", "TNR шрифт ўрнатиш"));
            flow.Controls.Add(CreateSpacer(12));

            // ── Изоҳлар ─────────────────────────────────────────────────
            flow.Controls.Add(CreateHotkeyGroupHeader("Изоҳлар"));
            flow.Controls.Add(CreateSpacer(6));
            flow.Controls.Add(CreateHotkeyRow("Ctrl + Alt + Shift + Z", "Сўз изоҳи"));
            flow.Controls.Add(CreateHotkeyRow("Ctrl + Alt + Shift + O", "Хатоларни экспорт"));
            flow.Controls.Add(CreateSpacer(12));

            // ── Махсус белгилар ─────────────────────────────────────────
            flow.Controls.Add(CreateHotkeyGroupHeader("Махсус белгилар"));
            flow.Controls.Add(CreateSpacer(6));
            flow.Controls.Add(CreateHotkeyRow("Ctrl + Alt + Shift + F", "Тутуқ белгиларини тўғрилаш"));
            flow.Controls.Add(CreateHotkeyRow("Ctrl + Alt + Shift + 1", "Тутуқ белгиси (ʻ) қўйиш"));
            flow.Controls.Add(CreateHotkeyRow("Ctrl + Alt + Shift + 2", "Тутуқ белгиси (ʼ) қўйиш"));
            flow.Controls.Add(CreateSpacer(16));

            ModernScrollBar.AttachTo(flow, InfoScrollBarWidth);

            page.Controls.Add(flow);
            return page;
        }

        private static Control CreateHotkeyGroupHeader(string title)
        {
            var lbl = new Label
            {
                Text = title,
                AutoSize = true,
                Font = new Font("Segoe UI", 11f, FontStyle.Bold),
                ForeColor = ThemeManager.TextSecondary,
                Margin = new Padding(0, 4, 0, 0),
            };
            return lbl;
        }

        private static Control CreateHotkeyRow(string keys, string description)
        {
            var row = new Panel
            {
                Height = 36,
                Width = 720,
                Margin = new Padding(0, 0, 0, 4),
                BackColor = Color.Transparent,
            };

            var badgeFont = new Font("Segoe UI", 10f, FontStyle.Bold);

            // Measure badge width from the key string
            var badgeSize = TextRenderer.MeasureText(keys, badgeFont) + new Size(24, 8);
            int badgeY = (row.Height - badgeSize.Height) / 2;

            // AutoSize=false is critical — otherwise setting Text="" causes AutoSize
            // to collapse the label to zero, hiding the custom-drawn text.
            var badge = new Label
            {
                AutoSize = false,
                Size = badgeSize,
                Location = new Point(0, badgeY),
                Font = badgeFont,
                Text = "",   // text drawn manually in Paint event
                BackColor = Color.Transparent
            };

            badge.Paint += (s, e) =>
            {
                var g = e.Graphics;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                var rect = new Rectangle(0, 0, badge.Width - 1, badge.Height - 1);
                using (var bg = new SolidBrush(ThemeManager.Surface))
                    g.FillRectangle(bg, rect);
                using (var pen = new Pen(ThemeManager.Border, 1f))
                using (var path = CreateRoundedPath(rect, ThemeManager.RadiusSM))
                    g.DrawPath(pen, path);
                using (var br = new SolidBrush(ThemeManager.TextPrimary))
                using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                    g.DrawString(keys, badge.Font, br, new RectangleF(0, 0, badge.Width, badge.Height), sf);
            };

            var descFont = new Font("Segoe UI", 11.5f, FontStyle.Regular);
            var descHeight = TextRenderer.MeasureText(description, descFont).Height;
            var desc = new Label
            {
                AutoSize = false,
                Width = row.Width - badge.Right - 18,
                Height = descHeight + 4,
                Location = new Point(badge.Right + 14, (row.Height - descHeight) / 2),
                Font = descFont,
                Text = description,
                ForeColor = ThemeManager.TextPrimary,
                BackColor = Color.Transparent
            };

            row.Controls.Add(badge);
            row.Controls.Add(desc);
            return row;
        }

        // =================================================================
        //  TAB 4: МАХФИЯЛИК (Privacy Policy)
        // =================================================================

        private Panel BuildPrivacyPage()
        {
            var page = new Panel { BackColor = ThemeManager.Background };

            var flow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = false,
                AutoScroll = true,
                Padding = new Padding(ThemeManager.SpaceMD)
            };

            flow.Controls.Add(CreateSection("Махфиялик сиёсати", 16f, FontStyle.Bold, ThemeManager.Primary));
            flow.Controls.Add(CreateSpacer(12));

            flow.Controls.Add(CreateSection("Маълумотлар йиғилмайди", 12f, FontStyle.Bold, ThemeManager.TextPrimary));
            flow.Controls.Add(CreateSpacer(4));
            flow.Controls.Add(CreateParagraph(
                "Ўзбек Орфо дастури ҳеч қандай шахсий маълумот йиғмайди ва " +
                "ташқи серверларга юбормайди. Дастур тўлиқ офлайн режимда ишлайди — " +
                "интернет уланишини талаб қилмайди."));
            flow.Controls.Add(CreateSpacer(12));

            flow.Controls.Add(CreateSection("Маҳаллий сақлаш", 12f, FontStyle.Bold, ThemeManager.TextPrimary));
            flow.Controls.Add(CreateSpacer(4));
            flow.Controls.Add(CreateParagraph(
                "Барча маълумотлар фақат сизнинг компьютерингизда сақланади:\n\n" +
                "• Қўлланувчи луғати — %AppData%\\UzbekOrfo\\user_dictionary.json\n" +
                "• Журнал файли — %AppData%\\UzbekOrfo\\log.txt\n" +
                "• Созламалар — дастур ичида\n\n" +
                "Бу маълумотлар ҳеч қачон учинчи томонларга узатилмайди."));
            flow.Controls.Add(CreateSpacer(12));

            flow.Controls.Add(CreateSection("Телеметрия йўқ", 12f, FontStyle.Bold, ThemeManager.TextPrimary));
            flow.Controls.Add(CreateSpacer(4));
            flow.Controls.Add(CreateParagraph(
                "Дастурда телеметрия, аналитика ёки фойдаланишни кузатиш " +
                "воситалари мавжуд эмас. Ҳеч қандай маълумот интернет орқали " +
                "юборилмайди."));
            flow.Controls.Add(CreateSpacer(12));

            flow.Controls.Add(CreateSection("Журнал файли", 12f, FontStyle.Bold, ThemeManager.TextPrimary));
            flow.Controls.Add(CreateSpacer(4));
            flow.Controls.Add(CreateParagraph(
                "Дастур ишлаши давомида хатолар ва диагностика маълумотлари " +
                "маҳаллий журнал файлига (log.txt) ёзилади. Бу файл фақат " +
                "сизнинг компьютерингизда сақланади ва истаган вақтда ўчириш мумкин."));
            flow.Controls.Add(CreateSpacer(16));
            flow.Controls.Add(CreateParagraph(
                "Сўнгги янгиланиш: 2026 йил март"));

            ModernScrollBar.AttachTo(flow, InfoScrollBarWidth);

            page.Controls.Add(flow);
            return page;
        }

        // =================================================================
        //  TAB 4: ҚЎЛЛАБ-ҚУВВАТЛАШ (Donate / Support)
        // =================================================================

        private Panel BuildDonatePage()
        {
            var page = new Panel { BackColor = ThemeManager.Background };

            var flow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = false,
                AutoScroll = true,
                Padding = new Padding(ThemeManager.SpaceMD)
            };

            flow.Controls.Add(CreateSection("Дастурни қўллаб-қувватлаш", 16f, FontStyle.Bold, ThemeManager.Primary));
            flow.Controls.Add(CreateSpacer(12));

            flow.Controls.Add(CreateParagraph(
                "Ўзбек Орфо — бепул ва очиқ кодли дастур. Агар дастур сизга " +
                "фойдали бўлса, уни ривожлантиришга ёрдам беришингиз мумкин."));
            flow.Controls.Add(CreateSpacer(16));

            // Tirikchilik.uz
            flow.Controls.Add(CreateSection("Tirikchilik.uz орқали:", 12f, FontStyle.Bold, ThemeManager.TextPrimary));
            flow.Controls.Add(CreateSpacer(4));
            flow.Controls.Add(CreateContactCard(
                "Tirikchilik.uz",
                "tirikchilik.uz",
                "https://tirikchilik.uz",
                ThemeManager.Primary));
            flow.Controls.Add(CreateSpacer(12));

            // Card number
            flow.Controls.Add(CreateSection("Карта рақами орқали:", 12f, FontStyle.Bold, ThemeManager.TextPrimary));
            flow.Controls.Add(CreateSpacer(4));
            flow.Controls.Add(CreateContactCard(
                "9860 1966 0098 2116",
                "Click / Payme орқали тўлаш",
                "https://my.click.uz/auth",
                Color.FromArgb(255, 153, 0)));
            flow.Controls.Add(CreateSpacer(16));

            flow.Controls.Add(CreateSection("Нега қўллаб-қувватлаш керак?", 12f, FontStyle.Bold, ThemeManager.TextPrimary));
            flow.Controls.Add(CreateSpacer(4));
            flow.Controls.Add(CreateBulletList(
                "Янги имкониятлар қўшиш",
                "Луғатни кенгайтириш",
                "Хатоларни тузатиш",
                "Дастурни доимий янгилаб бориш"));
            flow.Controls.Add(CreateSpacer(16));
            flow.Controls.Add(CreateParagraph(
                "Ҳар қандай миқдордаги ёрдам қадрланади! Раҳмат!"));

            ModernScrollBar.AttachTo(flow, InfoScrollBarWidth);

            page.Controls.Add(flow);
            return page;
        }

        // =================================================================
        //  TAB 5: ФИКР БИЛДИРИШ (Feedback)
        // =================================================================

        private Panel BuildFeedbackPage()
        {
            var page = new Panel { BackColor = ThemeManager.Background };

            var flow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = false,
                AutoScroll = true,
                Padding = new Padding(ThemeManager.SpaceMD)
            };

            flow.Controls.Add(CreateSection("Фикр ва таклифлар", 16f, FontStyle.Bold, ThemeManager.Primary));
            flow.Controls.Add(CreateSpacer(12));

            flow.Controls.Add(CreateParagraph(
                "Дастурни яхшилаш учун сизнинг фикрларингиз ва таклифларингиз " +
                "жуда муҳим. Қуйидаги усулларда биз билан боғланишингиз мумкин:"));
            flow.Controls.Add(CreateSpacer(12));

            // Telegram first
            flow.Controls.Add(CreateSection("Telegram:", 12f, FontStyle.Bold, ThemeManager.TextPrimary));
            flow.Controls.Add(CreateSpacer(4));
            flow.Controls.Add(CreateContactCard(
                "@uzbekorfo",
                "t.me/uzbekorfo",
                "https://t.me/uzbekorfo",
                Color.FromArgb(41, 182, 246)));
            flow.Controls.Add(CreateSpacer(12));

            // Email second
            flow.Controls.Add(CreateSection("Электрон почта:", 12f, FontStyle.Bold, ThemeManager.TextPrimary));
            flow.Controls.Add(CreateSpacer(4));
            flow.Controls.Add(CreateContactCard(
                "uzbekorfo@gmail.com",
                "uzbekorfo@gmail.com",
                "mailto:uzbekorfo@gmail.com",
                ThemeManager.Primary));
            flow.Controls.Add(CreateSpacer(16));

            flow.Controls.Add(CreateSection("Хато ҳақида хабар бериш:", 12f, FontStyle.Bold, ThemeManager.TextPrimary));
            flow.Controls.Add(CreateSpacer(4));
            flow.Controls.Add(CreateParagraph(
                "Хато топсангиз, қуйидаги маълумотларни юборинг:\n\n" +
                "1. Windows версиясингиз\n" +
                "2. Microsoft Word версиясингиз\n" +
                "3. Хато қандай юз берганини батафсил ёзинг\n" +
                "4. Имкон бўлса, экранрасмини (скриншотни) юборинг\n\n" +
                "Журнал файли манзили: %AppData%\\UzbekOrfo\\log.txt"));
            flow.Controls.Add(CreateSpacer(16));
            flow.Controls.Add(CreateParagraph(
                "Фикрларингиз учун олдиндан раҳмат! Дастурни биргаликда яхшилаймиз."));

            ModernScrollBar.AttachTo(flow, InfoScrollBarWidth);

            page.Controls.Add(flow);
            return page;
        }

        /// <summary>Like CreateLinkCard but with a "Юбориш" action button instead of "Очиш".</summary>
        private Control CreateContactCard(string title, string subtitle, string url, Color accentColor)
        {
            var card = new Panel
            {
                Width = 580,
                Height = 88,
                Margin = new Padding(0, 0, 0, 4),
                BackColor = ThemeManager.Surface
            };
            card.Paint += (s, e) =>
            {
                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                int r = ThemeManager.RadiusSM;
                var rect = new Rectangle(0, 0, card.Width - 1, card.Height - 1);
                using (var path = CreateRoundedPath(rect, r))
                {
                    using (var bg = new SolidBrush(ThemeManager.Surface))
                        g.FillPath(bg, path);
                    using (var pen = new Pen(ThemeManager.Border))
                        g.DrawPath(pen, path);
                }

                // Accent left bar
                using (var brush = new SolidBrush(accentColor))
                    g.FillRectangle(brush, 0, 6, 4, card.Height - 12);

                // Title
                using (var font = new Font("Segoe UI", 12f, FontStyle.Bold))
                using (var brush = new SolidBrush(ThemeManager.TextPrimary))
                using (var fmt = new StringFormat { Trimming = StringTrimming.EllipsisCharacter, FormatFlags = StringFormatFlags.NoWrap })
                    g.DrawString(title, font, brush, new RectangleF(16, 16, card.Width - 190, 24), fmt);

                // Subtitle
                using (var font = new Font("Segoe UI", 10.5f, FontStyle.Regular))
                using (var brush = new SolidBrush(accentColor))
                using (var fmt = new StringFormat { Trimming = StringTrimming.EllipsisCharacter, FormatFlags = StringFormatFlags.NoWrap })
                    g.DrawString(subtitle, font, brush, new RectangleF(16, 50, card.Width - 190, 22), fmt);
            };

            // Send / open button
            var sendBtn = new ModernButton
            {
                Text = "Юбориш",
                Style = ModernButton.ButtonStyle.Primary,
                Size = new Size(80, 32),
                Font = new Font("Segoe UI", 10f, FontStyle.Regular),
                Location = new Point(card.Width - 92, 28)
            };
            sendBtn.Click += (s, e) =>
            {
                try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); }
                catch { }
            };
            card.Controls.Add(sendBtn);

            // Copy button
            var copyBtn = new ModernButton
            {
                Text = "Нусха",
                Style = ModernButton.ButtonStyle.Ghost,
                Size = new Size(64, 32),
                Font = new Font("Segoe UI", 10f, FontStyle.Regular),
                Location = new Point(card.Width - 162, 28)
            };
            // copy the handle (username or email), not the full URL
            string copyValue = title;
            copyBtn.Click += (s, e) =>
            {
                try
                {
                    Clipboard.SetText(copyValue);
                    copyBtn.Text = "✓";
                    var timer = new Timer { Interval = 1500 };
                    timer.Tick += (t, te) => { copyBtn.Text = "Нусха"; timer.Stop(); timer.Dispose(); };
                    timer.Start();
                }
                catch { }
            };
            card.Controls.Add(copyBtn);

            return card;
        }

        // =================================================================
        //  HELPER METHODS — UI Building Blocks
        // =================================================================

        /// <summary>Creates a styled section header label.</summary>
        private static Control CreateSection(string text, float fontSize, FontStyle style, Color color)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                Font = new Font("Segoe UI", fontSize + 1f, style),
                ForeColor = color,
                Margin = new Padding(0, 0, 0, 0),
                MaximumSize = new Size(720, 0)
            };
        }

        /// <summary>Creates a paragraph label with word wrap.</summary>
        private static Control CreateParagraph(string text)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                Font = new Font("Segoe UI", 11.5f, FontStyle.Regular),
                ForeColor = ThemeManager.TextPrimary,
                Margin = new Padding(0, 0, 0, 4),
                MaximumSize = new Size(720, 0)
            };
        }

        /// <summary>Creates a vertical spacer.</summary>
        private static Control CreateSpacer(int height)
        {
            return new Panel
            {
                Height = height,
                Width = 10,
                BackColor = Color.Transparent,
                Margin = new Padding(0)
            };
        }

        /// <summary>Creates a bullet-point list.</summary>
        private static Control CreateBulletList(params string[] items)
        {
            string text = "";
            foreach (var item in items)
                text += $"  •  {item}\n";

            return new Label
            {
                Text = text.TrimEnd('\n'),
                AutoSize = true,
                Font = new Font("Segoe UI", 11.5f, FontStyle.Regular),
                ForeColor = ThemeManager.TextPrimary,
                Margin = new Padding(4, 0, 0, 4),
                MaximumSize = new Size(710, 0)
            };
        }

        /// <summary>Creates a styled link card with an "Очиш" (Open) button that launches a URL.</summary>
        private Control CreateLinkCard(string title, string subtitle, string url, Color accentColor)
        {
            var card = new Panel
            {
                Width = 580,
                Height = 88,
                Margin = new Padding(0, 0, 0, 4),
                BackColor = ThemeManager.Surface
            };
            card.Paint += (s, e) =>
            {
                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                int r = ThemeManager.RadiusSM;
                var rect = new Rectangle(0, 0, card.Width - 1, card.Height - 1);
                using (var path = CreateRoundedPath(rect, r))
                {
                    using (var bg = new SolidBrush(ThemeManager.Surface))
                        g.FillPath(bg, path);
                    using (var pen = new Pen(ThemeManager.Border))
                        g.DrawPath(pen, path);
                }

                // Accent left bar
                using (var brush = new SolidBrush(accentColor))
                    g.FillRectangle(brush, 0, 6, 4, card.Height - 12);

                // Title
                using (var font = new Font("Segoe UI", 12f, FontStyle.Bold))
                using (var brush = new SolidBrush(ThemeManager.TextPrimary))
                using (var fmt = new StringFormat { Trimming = StringTrimming.EllipsisCharacter, FormatFlags = StringFormatFlags.NoWrap })
                    g.DrawString(title, font, brush, new RectangleF(16, 16, card.Width - 190, 24), fmt);

                // Subtitle
                using (var font = new Font("Segoe UI", 10.5f, FontStyle.Regular))
                using (var brush = new SolidBrush(accentColor))
                using (var fmt = new StringFormat { Trimming = StringTrimming.EllipsisCharacter, FormatFlags = StringFormatFlags.NoWrap })
                    g.DrawString(subtitle, font, brush, new RectangleF(16, 50, card.Width - 190, 22), fmt);
            };

            // Open button
            var openBtn = new ModernButton
            {
                Text = "Очиш",
                Style = ModernButton.ButtonStyle.Primary,
                Size = new Size(72, 32),
                Font = new Font("Segoe UI", 10f, FontStyle.Regular),
                Location = new Point(card.Width - 86, 28)
            };
            openBtn.Click += (s, e) =>
            {
                try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); }
                catch { }
            };
            card.Controls.Add(openBtn);

            // Copy URL button
            var copyBtn = new ModernButton
            {
                Text = "Нусха",
                Style = ModernButton.ButtonStyle.Ghost,
                Size = new Size(64, 32),
                Font = new Font("Segoe UI", 10f, FontStyle.Regular),
                Location = new Point(card.Width - 162, 28)
            };
            copyBtn.Click += (s, e) =>
            {
                try
                {
                    Clipboard.SetText(url);
                    copyBtn.Text = "✓";
                    var timer = new Timer { Interval = 1500 };
                    timer.Tick += (t, te) =>
                    {
                        copyBtn.Text = "Нусха";
                        timer.Stop();
                        timer.Dispose();
                    };
                    timer.Start();
                }
                catch { }
            };
            card.Controls.Add(copyBtn);

            return card;
        }

        /// <summary>Creates a copyable text field (email, telegram, etc).</summary>
        private Control CreateCopyableField(string value)
        {
            var panel = new Panel
            {
                Width = 360,
                Height = 40,
                Margin = new Padding(0, 0, 0, 4),
                BackColor = ThemeManager.Surface
            };
            panel.Paint += (s, e) =>
            {
                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                int r = ThemeManager.RadiusSM;
                var rect = new Rectangle(0, 0, panel.Width - 1, panel.Height - 1);
                using (var path = CreateRoundedPath(rect, r))
                {
                    using (var bg = new SolidBrush(ThemeManager.Surface))
                        g.FillPath(bg, path);
                    using (var pen = new Pen(ThemeManager.Border))
                        g.DrawPath(pen, path);
                }

                using (var font = new Font("Segoe UI", 12f, FontStyle.Regular))
                using (var brush = new SolidBrush(ThemeManager.Primary))
                    g.DrawString(value, font, brush, 12, 9);
            };

            var copyBtn = new ModernButton
            {
                Text = "Нусха",
                Style = ModernButton.ButtonStyle.Ghost,
                Size = new Size(64, 26),
                Font = new Font("Segoe UI", 10f, FontStyle.Regular),
                Location = new Point(panel.Width - 76, 7)
            };
            copyBtn.Click += (s, e) =>
            {
                try
                {
                    Clipboard.SetText(value);
                    copyBtn.Text = "✓";
                    var timer = new Timer { Interval = 1500 };
                    timer.Tick += (t, te) =>
                    {
                        copyBtn.Text = "Нусха";
                        timer.Stop();
                        timer.Dispose();
                    };
                    timer.Start();
                }
                catch { }
            };
            panel.Controls.Add(copyBtn);

            return panel;
        }

        private static GraphicsPath CreateRoundedPath(Rectangle rect, int radius)
        {
            var path = new GraphicsPath();
            int d = radius * 2;
            path.AddArc(rect.X, rect.Y, d, d, 180, 90);
            path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
            path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
            path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        /// <summary>Shows the dialog.</summary>
        public static void ShowInfo()
        {
            using (var form = new AppInfoForm())
            {
                form.ShowDialog();
            }
        }
    }
}
