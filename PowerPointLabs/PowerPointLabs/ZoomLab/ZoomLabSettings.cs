using PowerPointLabs.ColorThemes.Extensions;
using PowerPointLabs.ZoomLab.Views;

namespace PowerPointLabs.ZoomLab
{
    internal static class ZoomLabSettings
    {
        public static bool BackgroundZoomChecked = true;
        public static bool MultiSlideZoomChecked = true;

        public static void ShowSettingsDialog()
        {
            ZoomLabSettingsDialogBox dialog = new ZoomLabSettingsDialogBox(BackgroundZoomChecked, MultiSlideZoomChecked);
            dialog.DialogConfirmedHandler += OnSettingsDialogConfirmed;
            dialog.ShowThematicDialog();
        }

        private static void OnSettingsDialogConfirmed(bool backgroundChecked, bool multiSlideChecked)
        {
            BackgroundZoomChecked = backgroundChecked;
            MultiSlideZoomChecked = multiSlideChecked;
        }
    }
}
