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
        public const string SavePresentationsButtonLabel = "Save Selected Slides";
        public const string SaveLabSettingsButtonLabel = "Settings";

        public const string RibbonMenuSupertip = "Use Save Lab to save selected slides as a separate presentation.";
        public const string SavePresentationsButtonSupertip =
            "Save selected slides as a new presentation at a desired file location.\n\n" +
            "To perform this action, select the desired slides to save, then click this button.";
        public const string SaveLabSettingsButtonSupertip = "Configure the default directory for Save Lab for a quick save.";

        public const string ErrorSavingPresentations = "Presentations could not be saved.";
        public const string ErrorZeroSlidesSelected = "No slides were selected to save.";

        public const string FolderDialogDescription = "Select the directory that you want to use as the default.";
        #endregion
    }
}
