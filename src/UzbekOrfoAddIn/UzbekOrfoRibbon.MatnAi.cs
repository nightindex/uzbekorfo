using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace UzbekOrfoAddIn
{
    public partial class UzbekOrfoRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private void matnaiEnabled_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.Completion?.SetEnabled(matnaiEnabled.Checked);
        }

        private void btnMatnAiRebuild_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.Completion?.Rebuild();
        }

        private void btnMatnAiSettings_Click(object sender, RibbonControlEventArgs e)
        {
                var current = ThisAddIn.Settings;
                if (current == null) return;
                using (var form = new Form { Text = "MatnAI", Width = 360, Height = 205,
                    FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent,
                    MaximizeBox = false, MinimizeBox = false })
                {
                    var minimum = new NumericUpDown { Left = 220, Top = 20, Minimum = 2, Maximum = 15,
                        Value = System.Math.Max(2, current.MinPredictionLength), Width = 75, AccessibleName = "Энг кам ҳарфлар сони" };
                    var count = new NumericUpDown { Left = 220, Top = 60, Minimum = 1, Maximum = 10,
                        Value = current.MaxPredictions, Width = 75, AccessibleName = "Таклифлар сони" };
                    var save = new Button { Text = "Сақлаш", Left = 220, Top = 105, DialogResult = DialogResult.OK };
                    form.Controls.AddRange(new Control[] { minimum, count, save,
                        new Label { Text = "Энг кам ҳарфлар", Left = 15, Top = 24, Width = 190 },
                        new Label { Text = "Таклифлар сони", Left = 15, Top = 64, Width = 190 } });
                    form.AcceptButton = save;
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        current.MinPredictionLength = (int)minimum.Value;
                        current.MaxPredictions = (int)count.Value;
                        current.Save();
                        ThisAddIn.Completion?.Dismiss();
                    }
                }
        }

        private void matnaiLearning_Click(object sender, RibbonControlEventArgs e)
        {
                if (matnaiLearning.Checked && MessageBox.Show(
                    "Қабул қилинган таклифлар ва уларнинг сони рейтингни яхшилаш учун фақат шу компьютерда сақлансинми?\nҲужжат матни йиғилмайди ва юборилмайди.\nЎчирилса, ўрганиш ва бу ҳисоблардан фойдаланиш тўхтайди; «Ўрганишни тозалаш» уларни ўчиради.",
                    "MatnAI", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                    matnaiLearning.Checked = false;
                ThisAddIn.Completion?.SetLearningEnabled(matnaiLearning.Checked);
        }

        private void btnMatnAiReset_Click(object sender, RibbonControlEventArgs e)
        {
                if (MessageBox.Show("MatnAI қабул қилинган таклифлар ҳисобини тозалайсизми? Луғатлар ўзгартирилмайди. Маҳаллий .bak тиклаш нусхаси қолиши мумкин.",
                    "MatnAI", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    ThisAddIn.Completion?.ResetLearning();
        }

        private void btnMatnAiHelp_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show(
                "Оддий ҳужжат матнида камида икки ҳарф ёзинг.\n" +
                "Ctrl+Alt+Up/Down таклифни танлайди. Ctrl+Alt+Right уни қабул қилади; сичқонча билан ҳам қабул қилиш мумкин. Esc ёпади.\n" +
                "Tab ва Enter ўзгармайди. Ctrl+Z қабул қилинган таклифни бир амал билан бекор қилади.\n" +
                "Ҳимояланган матн, жадваллар, ўзгаришларни кузатиш ва киритиш усули фаол бўлганда таклифлар кўрсатилмайди.\n" +
                "Офлайн ишлайди: шахсий қабул қилиш ҳисоблари ихтиёрий; ҳужжат матни йиғилмайди ва гаплар прогноз қилинмайди.\n" +
                "Луғат сақланганда индекс янгиланади. «Луғатни янгилаш» мажбурий янгилайди.", "MatnAI");
        }
    }
}
