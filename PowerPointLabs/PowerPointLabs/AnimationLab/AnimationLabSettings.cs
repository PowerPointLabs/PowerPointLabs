﻿using PowerPointLabs.AnimationLab.Views;

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
            dialog.ShowDialog();
        }

        private static void OnSettingsDialogConfirmed(float newDuration, bool newFrameChecked)
        {
            AnimationDuration = newDuration;
            IsUseFrameAnimation = newFrameChecked;
        }
    }
}
