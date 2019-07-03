using System.Drawing;
using PowerPointLabs.ColorThemes.Extensions;
using PowerPointLabs.HighlightLab.Views;

namespace PowerPointLabs.HighlightLab
{
    internal static class HighlightLabSettings
    {
        public static Color bulletsBackgroundColor = Color.FromArgb(255, 255, 0);

        public static Color bulletsTextHighlightColor = Color.FromArgb(242, 41, 10);
        public static Color bulletsTextDefaultColor = Color.FromArgb(0, 0, 0);

        public static Color textFragmentsBackgroundColor = Color.FromArgb(255, 255, 0);

        public static void ShowSettingsDialog()
        {
            HighlightLabSettingsDialogBox dialog = new HighlightLabSettingsDialogBox(
                bulletsTextHighlightColor, 
                bulletsTextDefaultColor, 
                bulletsBackgroundColor);
            dialog.DialogConfirmedHandler += OnSettingsDialogConfirmed;
            dialog.ShowThematicDialog();
        }

        private static void OnSettingsDialogConfirmed(Color newHighlightColor, Color newDefaultColor, Color newBackgroundColor)
        {
            bulletsTextHighlightColor = newHighlightColor;
            bulletsTextDefaultColor = newDefaultColor;
            bulletsBackgroundColor = newBackgroundColor;
            textFragmentsBackgroundColor = newBackgroundColor;
        }
    }
}
