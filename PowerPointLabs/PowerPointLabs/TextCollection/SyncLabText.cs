﻿namespace PowerPointLabs.TextCollection
{
    internal static class SyncLabText
    {
        #region Action Framework Tags
        public const string RibbonMenuId = "SyncLabButton";
        public const string RibbonMenuLabel = "Sync";
        public const string RibbonMenuSupertip =
            "Use Sync Lab to make your slides look more consistent.\n\n" +
            "Click this button to open the Sync Lab interface.";
        #endregion

        public const string PaneTag = "SyncLab";
        public const string TaskPanelTitle = "Sync Lab";
        public const string StorageFileName = "Sync Lab - Do not edit";
        public const string DefaultFormatName = "Format";

        public const string ErrorDialogTitle = "Unable to execute action";
        public const string ErrorCopy = "Error: Unable to copy selected item.";
        public const string ErrorSmartArtUnsupported = "Error: SmartArt is currently not supported by SyncLab.";
        public const string ErrorCopySelectionInvalid = "Please select one shape to copy.";
        public const string ErrorPasteSelectionInvalid = "Please select at least one item to apply this format to.";
        public const string ErrorShapeDeleted = "Error in loading shape formats. Removing invalid formats from the list.";
        public const string ErrorSyncPaneNotOpened = "Error: SyncPane not opened.";

        public const string WarningDialogTitle = "Warning";
        public const string WarningSyncPerspectiveShadow =
            "PowerPoint does not allow us to differentiate between perspective and outer shadows with custom settings.\n\n" +
            "The shape will have an outer shadow when synced.";
    }
}
