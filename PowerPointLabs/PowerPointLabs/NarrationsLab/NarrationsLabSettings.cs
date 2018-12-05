using System.Collections.Generic;

using PowerPointLabs.NarrationsLab.Data;
using PowerPointLabs.NarrationsLab.Views;

namespace PowerPointLabs.NarrationsLab
{
    internal static class NarrationsLabSettings
    {
        public static List<string> VoiceNameList = null;
        public static int VoiceSelectedIndex = 0;

        public static bool IsPreviewEnabled = false;
        public static HumanVoice humanVoice;

        public static void ShowSettingsDialog()
        {
            NarrationsLabSettingsDialogBox dialog = NarrationsLabSettingsDialogBox.GetInstance();
            NarrationsLabMainSettingsPage.GetInstance().SetNarrationsLabMainSettings(
                VoiceSelectedIndex,
                humanVoice,
                VoiceNameList,
                NotesToAudio.IsHumanVoiceSelected,
                IsPreviewEnabled);
            NarrationsLabMainSettingsPage.GetInstance().DialogConfirmedHandler += OnSettingsDialogConfirmed;
            dialog.ShowDialog();
        }

        private static void OnSettingsDialogConfirmed(string voiceName, HumanVoice voice, bool isHumanVoiceSelected, bool isPreviewCurrentSlide)
        {
            IsPreviewEnabled = isPreviewCurrentSlide;

            if (!string.IsNullOrWhiteSpace(voiceName))
            {
                NotesToAudio.SetDefaultVoice(voiceName);
                VoiceSelectedIndex = VoiceNameList.IndexOf(voiceName);
            }
            //TODO: This is a simplifying logic. As long as human voice textbox is non-empty, then human voice is always selected
            //To remove human voice, set text box to empty.

            if (voice != null)
            {
                humanVoice = voice;
                NotesToAudio.SetDefaultVoice(humanVoice.voiceName, humanVoice);
                NotesToAudio.IsHumanVoiceSelected = isHumanVoiceSelected;
            }
        }
    }
}
