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

        private void btnMatnAiSettings_Click(object sender, RibbonControlEventArgs e)
        {
            var current = ThisAddIn.Settings;
            if (current == null) return;

            using (var form = new Form
            {
                Text = "MatnAI созламалари",
                Width = 470,
                Height = 325,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterParent,
                MaximizeBox = false,
                MinimizeBox = false
            })
            {
                var minimum = new NumericUpDown
                {
                    Left = 270,
                    Top = 20,
                    Minimum = 2,
                    Maximum = 15,
                    Value = System.Math.Max(2, current.MinPredictionLength),
                    Width = 75,
                    AccessibleName = "Энг кам ҳарфлар сони"
                };
                var count = new NumericUpDown
                {
                    Left = 270,
                    Top = 60,
                    Minimum = 1,
                    Maximum = 10,
                    Value = current.MaxPredictions,
                    Width = 75,
                    AccessibleName = "Таклифлар сони"
                };
                var learning = new CheckBox
                {
                    Text = "Шахсий ўрганишни ёқиш",
                    Left = 15,
                    Top = 100,
                    Width = 300,
                    Checked = current.MatnAiLearningConsent,
                    AccessibleName = "Шахсий ўрганиш"
                };
                var clearLearning = new Button
                {
                    Text = "Ўрганишни тозалаш",
                    Left = 15,
                    Top = 140,
                    Width = 190,
                    AccessibleDescription = "Сақланган қабул қилиш ҳисобларини ўчиради."
                };
                var refreshIndex = new Button
                {
                    Text = "Индексни ҳозир янгилаш",
                    Left = 220,
                    Top = 140,
                    Width = 210,
                    AccessibleDescription = "MatnAI сўз таклифлари индексини қайта яратади."
                };
                var save = new Button
                {
                    Text = "Сақлаш",
                    Left = 270,
                    Top = 205,
                    DialogResult = DialogResult.OK
                };
                form.Controls.AddRange(new Control[]
                {
                    minimum,
                    count,
                    learning,
                    clearLearning,
                    refreshIndex,
                    save,
                    new Label { Text = "Энг кам ҳарфлар", Left = 15, Top = 24, Width = 230 },
                    new Label { Text = "Таклифлар сони", Left = 15, Top = 64, Width = 230 },
                    new Label
                    {
                        Text = "Қабул қилинган таклифлар сони фақат шу компьютерда сақланади.",
                        Left = 15,
                        Top = 120,
                        Width = 425
                    }
                });
                clearLearning.Click += (s, args) =>
                {
                    if (MessageBox.Show(form,
                        "MatnAI қабул қилинган таклифлар ҳисобини тозалайсизми? Луғатлар ўзгартирилмайди. Маҳаллий .bak тиклаш нусхаси қолиши мумкин.",
                        "MatnAI", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        ThisAddIn.Completion?.ResetLearning();
                };
                refreshIndex.Click += (s, args) =>
                {
                    ThisAddIn.Completion?.Rebuild();
                    MessageBox.Show(form, "Таклифлар индекси янгиланмоқда.", "MatnAI",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                };
                form.AcceptButton = save;

                if (form.ShowDialog() != DialogResult.OK) return;

                bool learningChanged = learning.Checked != current.MatnAiLearningConsent;
                if (learning.Checked && learningChanged && MessageBox.Show(form,
                    "Қабул қилинган таклифлар ва уларнинг сони рейтингни яхшилаш учун фақат шу компьютерда сақлансинми?\nҲужжат матни йиғилмайди ва юборилмайди.\nЎчирилса, ўрганиш ва бу ҳисоблардан фойдаланиш тўхтайди; «Ўрганишни тозалаш» уларни ўчиради.",
                    "MatnAI", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                    return;

                current.MinPredictionLength = (int)minimum.Value;
                current.MaxPredictions = (int)count.Value;
                current.MatnAiLearningConsent = learning.Checked;
                current.Save();
                if (learningChanged) ThisAddIn.Completion?.SetLearningEnabled(learning.Checked);
                else ThisAddIn.Completion?.Dismiss();
            }
        }

        private void btnMatnAiHelp_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show(
                "Оддий ҳужжат матнида камида икки ҳарф ёзинг.\n" +
                "Ctrl+Alt+Up/Down таклифни танлайди. Ctrl+Alt+Right уни қабул қилади; сичқонча билан ҳам қабул қилиш мумкин. Esc ёпади.\n" +
                "Tab ва Enter ўзгармайди. Ctrl+Z қабул қилинган таклифни бир амал билан бекор қилади.\n" +
                "Ҳимояланган матн, жадваллар, ўзгаришларни кузатиш ва киритиш усули фаол бўлганда таклифлар кўрсатилмайди.\n" +
                "Офлайн ишлайди: шахсий қабул қилиш ҳисоблари ихтиёрий; ҳужжат матни йиғилмайди ва гаплар прогноз қилинмайди.\n" +
                "Луғат сақланганда индекс автоматик янгиланади. Зарур бўлса, «Созламалар» ойнасида «Индексни ҳозир янгилаш»ни босинг.",
                "MatnAI");
        }
    }
}
