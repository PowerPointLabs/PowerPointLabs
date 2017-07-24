namespace PowerPointLabs.TextCollection
{
    internal static class CaptionsLabText
    {
        #region Action Framework Tags
        public const string RibbonMenuId = "CaptionsLabMenu";
        public const string AddCaptionsTag = "AddCaptions";
        public const string RemoveCaptionsTag = "RemoveCaptions";
        public const string RemoveNotesTag = "RemoveNotes";
        #endregion

        #region GUI Text
        public const string RibbonMenuLabel = "Captions";
        public const string AddCaptionsButtonLabel = "Add Captions";
        public const string RemoveCaptionsButtonLabel = "Remove Captions";
        public const string RemoveAllNotesButtonLabel = "Remove All Notes";

        public const string RibbonMenuSupertip =
            "Use Captions lab to create customizable movie-style subtitles from text in the Notes pane of the selected slides.";
        public const string AddCaptionsButtonSupertip =
            "Create movie-style subtitles from text in the Notes pane for the currently selected slides.\n\n" +
            "To perform this action, select the slide(s) with the notes, then click this button.";
        public const string RemoveCaptionsButtonSupertip =
            "Remove captions added using Captions Lab from the selected slides.\n\n" +
            "To perform this action, select the slide(s) to remove the captions from, then click this button.";
        public const string RemoveAllNotesButtonSupertip =
            "Remove notes from Notes pane of selected slides.\n\n" +
            "To perform this action, select the slide(s) to remove the notes from, then click this button.";

        public const string ErrorDialogTitle = "Unable to execute action";
        public const string ErrorNoSelection = "Select at least one slide to apply captions.";
        public const string ErrorNoNotes = "Captions could not be created because there are no notes entered. Please enter something in the notes and try again.";
        public const string ErrorNoSelectionLog = "No slide in selection.";
        public const string ErrorNoCurrentSlideLog = "No current slide.";
        public const string ErrorNoNotesLog = "No notes on slide.";
        #endregion
    }
}
