namespace PowerPointLabs.TextCollection
{
    internal static class ResizeLabText
    {
        public const string PaneTag = "ResizeLab";

        public const string TaskPaneTitle = "Resize Lab";
        public const string RibbonMenuLabel = "Resize";
        public const string RibbonMenuSupertip =
            "Use Resize Lab to accurately resize the objects on your slide.\n\n" +
            "Click this button to open the Resize Lab interface.";

        public const string ErrorInvalidSelection = "You need to select at least {1} {2} before applying '{0}'";
        public const string ErrorNotSameShapes = "You need to select the same type of objects before applying 'Adjust Area Proportionally'";
        public const string ErrorGroupShapeNotSupported = "'Adjust Area Proportionally' does not support grouped objects";
        public const string ErrorUndefined = "Undefined error in Resize Lab";
    }
}
