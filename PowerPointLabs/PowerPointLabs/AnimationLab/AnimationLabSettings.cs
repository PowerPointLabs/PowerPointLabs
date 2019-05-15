using PowerPointLabs.AnimationLab.Views;

using PowerPointLabs.Extensions;

namespace PowerPointLabs.AnimationLab
{
    internal static class AnimationLabSettings
    {
        public static float AnimationDuration = 0.5f;
        public static bool IsUseFrameAnimation = false;

        public static void ShowSettingsDialog()
        {
            AnimationLabSettingsDialogBox dialog = new AnimationLabSettingsDialogBox(AnimationDuration, IsUseFrameAnimation);
            dialog.DialogConfirmedHandler += OnSettingsDialogConfirmed;
            Globals.ThisAddIn.ColorThemeChanged += dialog.UpdateColors;
            dialog.ShowDialog();
            Globals.ThisAddIn.ColorThemeChanged -= dialog.UpdateColors;
        }

        private static void OnSettingsDialogConfirmed(float newDuration, bool newFrameChecked)
        {
            AnimationDuration = newDuration;
            IsUseFrameAnimation = newFrameChecked;
        }
    }
}
