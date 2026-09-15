using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace UzbekOrfoAddIn
{
    public partial class UzbekOrfoRibbon
    {
        private RibbonToggleButton matnaiEnabled;
        private RibbonToggleButton matnaiLearning;

        private void InitializeMatnAi()
        {
            var tab = Factory.CreateRibbonTab();
            tab.Name = "tabMatnAi"; tab.Label = "MatnAi";
            var suggestions = Factory.CreateRibbonGroup();
            suggestions.Name = "groupMatnAiSuggestions"; suggestions.Label = "So'z takliflari";
            matnaiEnabled = Factory.CreateRibbonToggleButton();
            matnaiEnabled.Name = "toggleMatnAi"; matnaiEnabled.Label = "Yoqish";
            matnaiEnabled.SuperTip = "Offline word completion. No text is inserted until you accept a suggestion.";
            matnaiEnabled.Click += (s, e) => ThisAddIn.Completion?.SetEnabled(matnaiEnabled.Checked);
            suggestions.Items.Add(matnaiEnabled);
            var show = Factory.CreateRibbonButton();
            show.Name = "showMatnAi"; show.Label = "Takliflarni ko'rsatish";
            show.Click += (s, e) => ThisAddIn.Completion?.ShowSuggestions();
            suggestions.Items.Add(show);
            var settings = Factory.CreateRibbonGroup();
            settings.Name = "groupMatnAiSettings"; settings.Label = "Moslashtirish";
            var rebuild = Factory.CreateRibbonButton();
            rebuild.Name = "rebuildMatnAi"; rebuild.Label = "Lug'atni yangilash";
            rebuild.SuperTip = "Rebuild the completion index after dictionary edits.";
            rebuild.Click += (s, e) => ThisAddIn.Completion?.Rebuild();
            settings.Items.Add(rebuild);
            var configure = Factory.CreateRibbonButton();
            configure.Name = "settingsMatnAi"; configure.Label = "Sozlamalar";
            configure.Click += (s, e) =>
            {
                var current = ThisAddIn.Settings;
                if (current == null) return;
                using (var form = new Form { Text = "MatnAi", Width = 360, Height = 205,
                    FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent,
                    MaximizeBox = false, MinimizeBox = false })
                {
                    var minimum = new NumericUpDown { Left = 220, Top = 20, Minimum = 2, Maximum = 15,
                        Value = System.Math.Max(2, current.MinPredictionLength), Width = 75, AccessibleName = "Minimum prefix length" };
                    var count = new NumericUpDown { Left = 220, Top = 60, Minimum = 1, Maximum = 10,
                        Value = current.MaxPredictions, Width = 75, AccessibleName = "Suggestion count" };
                    var save = new Button { Text = "Saqlash", Left = 220, Top = 105, DialogResult = DialogResult.OK };
                    form.Controls.AddRange(new Control[] { minimum, count, save,
                        new Label { Text = "Eng kam harflar", Left = 15, Top = 24, Width = 190 },
                        new Label { Text = "Takliflar soni", Left = 15, Top = 64, Width = 190 } });
                    form.AcceptButton = save;
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        current.MinPredictionLength = (int)minimum.Value;
                        current.MaxPredictions = (int)count.Value;
                        current.Save();
                        ThisAddIn.Completion?.Dismiss();
                    }
                }
            };
            settings.Items.Add(configure);
            matnaiLearning = Factory.CreateRibbonToggleButton();
            matnaiLearning.Name = "learnMatnAi"; matnaiLearning.Label = "Shaxsiy o'rganish";
            matnaiLearning.Click += (s, e) =>
            {
                if (matnaiLearning.Checked && MessageBox.Show(
                    "Save accepted words and their counts locally to improve ranking?\nNo document text is collected or sent.\nTurning this off stops learning and use of counts; Reset clears them.",
                    "MatnAi", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                    matnaiLearning.Checked = false;
                ThisAddIn.Completion?.SetLearningEnabled(matnaiLearning.Checked);
            };
            settings.Items.Add(matnaiLearning);
            var reset = Factory.CreateRibbonButton();
            reset.Name = "resetMatnAi"; reset.Label = "O'rganishni tozalash";
            reset.Click += (s, e) =>
            {
                if (MessageBox.Show("Clear MatnAi acceptance counts? Dictionaries will not be changed. A local .bak recovery copy may remain.",
                    "MatnAi", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    ThisAddIn.Completion?.ResetLearning();
            };
            settings.Items.Add(reset);
            var help = Factory.CreateRibbonGroup();
            help.Name = "groupMatnAiHelp"; help.Label = "Yordam";
            var guide = Factory.CreateRibbonButton();
            guide.Name = "helpMatnAi"; guide.Label = "Qisqa qo'llanma";
            guide.Click += (s, e) => MessageBox.Show(
                "Type at least two letters in ordinary document body text.\n" +
                "Ctrl+Alt+Up/Down selects a suggestion. Ctrl+Alt+Right accepts it; clicking also accepts. Esc dismisses.\n" +
                "Tab and Enter are unchanged. Ctrl+Z undoes acceptance.\n" +
                "Protected content, tables, tracked changes and input composition are excluded.\n" +
                "Offline preview: personal acceptance counts are opt-in; no document collection or sentence prediction.\n" +
                "Dictionary saves refresh the index. Lug'atni yangilash forces a refresh.", "MatnAi");
            help.Items.Add(guide);
            tab.Groups.Add(suggestions); tab.Groups.Add(settings); tab.Groups.Add(help);
            Tabs.Add(tab);
        }
    }
}
