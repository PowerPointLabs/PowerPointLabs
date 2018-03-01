namespace PowerPointLabs.TextCollection
{
    internal static class SaveLabText
    {
        #region Action Framework Tags
        public const string RibbonMenuId = "SaveLabMenu";
        public const string SavePresentationsButtonTag = "SavePresentations";
        public const string SaveLabSettingsButtonTag = "SaveLabSettings";
        #endregion

        #region GUI Text
        public const string RibbonMenuLabel = "Save";
        public const string SavePresentationsButtonLabel = "Save Presentations";
        public const string SaveLabSettingsButtonLabel = "Settings";

        public const string RibbonMenuSupertip = "Use Save Lab to save presentations in a separate PowerPoint file.";
        public const string SavePresentationsButtonSupertip =
            "Save presentations to a desired file and location.\n\n" +
            "To perform this action, select the slides to export, then click this button.";
        public const string SaveLabSettingsButtonSupertip = "Configure the default directory for Save Lab for a quick save.";

        public const string ErrorSavingPresentations = "Presentations could not be saved.";
        public const string ErrorZeroSlidesSelected = "No slides were selected to save.";

        public const string FolderDialogDescription = "Select the directory that you want to use as the default.";
        #endregion
    }
}
