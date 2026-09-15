using Microsoft.Office.Tools.Ribbon;
using UzbekOrfoAddIn.Forms;

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

            using (var form = new MatnAiSettingsForm(
                current.MinPredictionLength,
                current.MaxPredictions,
                current.MatnAiLearningConsent,
                (minimum, count, learningEnabled) =>
                {
                    bool learningChanged = learningEnabled != current.MatnAiLearningConsent;
                    current.MinPredictionLength = minimum;
                    current.MaxPredictions = count;
                    current.MatnAiLearningConsent = learningEnabled;
                    current.Save();
                    if (learningChanged) ThisAddIn.Completion?.SetLearningEnabled(learningEnabled);
                    else ThisAddIn.Completion?.Dismiss();
                },
                () => ThisAddIn.Completion?.ResetLearning(),
                () => ThisAddIn.Completion?.Rebuild()))
            {
                form.ShowDialog();
            }
        }

        private void btnMatnAiHelp_Click(object sender, RibbonControlEventArgs e)
        {
            using (var form = new MatnAiHelpForm())
                form.ShowDialog();
        }
    }
}
