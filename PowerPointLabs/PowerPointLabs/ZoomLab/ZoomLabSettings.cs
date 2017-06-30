using PowerPointLabs.Views;

namespace PowerPointLabs.ZoomLab
{
    internal class ZoomLabSettings
    {
        public static bool BackgroundZoomChecked = true;
        public static bool MultiSlideZoomChecked = true;

        public static void OpenZoomLabSettingsDialog()
        {
            ZoomLabSettingsDialogBox dialog = new ZoomLabSettingsDialogBox(BackgroundZoomChecked, MultiSlideZoomChecked);
            dialog.SettingsHandler += ZoomLabSettingsEdited;
            dialog.ShowDialog();
        }

        public static void ZoomLabSettingsEdited(bool backgroundChecked, bool multiSlideChecked)
        {
            BackgroundZoomChecked = backgroundChecked;
            MultiSlideZoomChecked = multiSlideChecked;
        }
    }
}
