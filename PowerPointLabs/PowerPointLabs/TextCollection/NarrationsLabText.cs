namespace PowerPointLabs.TextCollection
{
    internal static class NarrationsLabText
    {
        #region Action Framework Tags
        public const string RibbonMenuId = "NarrationsLabMenu";
        public const string AddNarrationsTag = "AddNarrations";
        public const string RecordNarrationsTag = "RecordNarrations";
        public const string RemoveNarrationsTag = "RemoveNarrations";
        public const string SettingsTag = "NarrationsLabSettings";
        #endregion

        #region GUI Text
        public const string RecManagementPanelTitle = "Record Management";

        public const string RibbonMenuLabel = "Narrations";
        public const string AddNarrationsButtonLabel = "Generate Audio Automatically";
        public const string RecordNarrationsButtonLabel = "Record Audio Manually";
        public const string RemoveNarrationsButtonLabel = "Remove Audio";
        public const string SettingsButtonLabel = "Settings";

        public const string RibbonMenuSupertip =
            "Use Narrations Lab to create narration from text in the Notes pane of the selected slides.";
        public const string AddNarrationsButtonSupertip =
            "Create synthesized narration from text in the Notes pane of the selected slides.\n\n" +
            "To perform this action, select the slide(s) with the Speaker Notes, then click this button.";
        public const string RecordNarrationsButtonSupertip =
            "Manually record audio to replace synthesized narration.\n\n" +
            "Click this button to open the Recording interface.";
        public const string RemoveNarrationsButtonSupertip =
            "Remove synthesized audio added using Narrations Lab from the selected slides.\n\n" +
            "To perform this action, select the slide(s) to remove the narrations from, then click this button.";
        public const string SettingsButtonSupertip = "Configure the settings for Narrations Lab.";

        public const string SettingsVoiceSelectionInputTooltip =
            "The voice to be used when generating synthesized audio.\n" +
            "Use [Voice] tags to specify a different voice for a particular section of text.";
        public const string SettingsPreviewCheckboxTooltip =
            "If checked, the current slide's audio and animations will play after the Add Audio button is clicked.";

        #region Recorder
        public const string RecorderInitialTimer = "00:00:00";
        public const string RecorderReadyStatusLabel = "Ready.";
        public const string RecorderRecordingStatusLabel = "Recording...";
        public const string RecorderPlayingStatusLabel = "Playing...";
        public const string RecorderPauseStatusLabel = "Pause";
        public const string RecorderUnrecognizeAudio = "Unrecognized Embedded Audio";
        public const string RecorderScriptStatusNoAudio = "No Audio";
        public const string RecorderWndMessageError = "Fatal error";
        public const string RecorderNoScriptDetail = "No Script Available";
        public const string RecorderNoInputDeviceMsg = "No audio input device was found.\n" +
                                                       "Check that a microphone or other audio input device is attached " +
                                                       "and working.";
        public const string RecorderErrorNoInputDeviceTitle = "Input Device Not Found";
        public const string RecorderErrorSaveRecord = "Do you want to save the recording?";
        public const string RecorderErrorSaveRecordTitle = "Save Recording";
        public const string RecorderErrorReplaceRecordFormat = "Do you want to replace\n{0}\nwith the current recording?";
        public const string RecorderErrorReplaceRecordTitle = "Replace Audio";
        public const string RecorderErrorNoRecordToPlay = "There are no recordings to play.";
        public const string RecorderErrorInvalidOperation = "Invalid Operation";

        public const string InShowControlErrorInvalidRecCommand = "Invalid Recording Command";
        public const string InShowControlRecButtonIdleText = "Stop and Advance";
        public const string InShowControlRecButtonRecText = "Start Recording";
        #endregion
        #endregion
    }
}
