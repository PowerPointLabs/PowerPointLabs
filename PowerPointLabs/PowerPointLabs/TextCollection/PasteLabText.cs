namespace PowerPointLabs.TextCollection
{
    internal static class PasteLabText
    {
        #region Action Framework Tags
        public const string RibbonMenuId = "PasteLabMenu";
        public const string PasteAtCursorPositionTag = "PasteAtCursorPosition";
        public const string PasteAtOriginalPositionTag = "PasteAtOriginalPosition";
        public const string PasteToFillSlideTag = "PasteToFillSlide";
        public const string PasteToFitSlideTag = "PasteToFitSlide";
        public const string ReplaceWithClipboardTag = "ReplaceWithClipboard";
        public const string PasteIntoGroupTag = "PasteIntoGroup";

        public const string RibbonMenuLabel = "Paste";
        public const string PasteToFillSlideLabel = "Paste To Fill Slide";
        public const string PasteToFitSlideLabel = "Paste To Fit Slide";
        public const string ReplaceWithClipboardLabel = "Replace With Clipboard";
        public const string PasteIntoGroupLabel = "Paste Into Group";
        public const string PasteAtCursorPositionLabel = "Paste At Cursor Position";
        public const string PasteAtOriginalPositionLabel = "Paste At Original Position";
        #endregion

        #region GUI Text
        public const string RibbonMenuSupertip =
            "Use Paste Lab to customize how you paste your copied objects.";
        public const string PasteToFillSlideSupertip =
            "Paste your copied objects to fill the current slide.\n\n" +
            "To perform this action, with object(s) copied, click this button.";
        public const string PasteToFitSlideSupertip =
            "Paste your copied objects to fit the current slide.\n\n" +
            "To perform this action, with object(s) copied, click this button.";
        public const string PasteAtOriginalPositionSupertip =
            "Paste your copied objects at their original positions when they were copied.\n\n" +
            "To perform this action, with object(s) copied, click this button.";
        public const string ReplaceWithClipboardSupertip =
            "Paste your copied objects over your selection while preserving its animations.\n\n" +
            "To perform this action, select an object on the slide, then with object(s) copied, click this button.";
        public const string PasteIntoGroupSupertip =
            "Paste your copied objects into an existing group.\n\n" +
            "To perform this action, select a group on the slide, then with object(s) copied, click this button.";
        #endregion

        public const string ReplaceWithClipboardActionHandlerReminderText = "Please select at least one shape.";
        
        public const string ErrorDialogTitle = "Unable to execute action";
        public const string ErrorEmptyClipboard = "Error: Clipboard is empty.";
        public const string ErrorPaste = "Error: Unable to paste content in clipboard. Try pasting it normally in PowerPoint and then copying the item again.";
        public const string ErrorNoSelection = "Select at least one slide to apply paste actions.";
    }
}
