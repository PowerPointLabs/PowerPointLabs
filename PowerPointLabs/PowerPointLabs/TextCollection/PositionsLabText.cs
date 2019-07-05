namespace PowerPointLabs.TextCollection
{
    internal static class PositionsLabText
    {
        #region Action Framework Tags
        public const string RibbonMenuId = "PositionsLabButton";
        public const string RibbonMenuLabel = "Positions";
        public const string RibbonMenuSupertip =
            "Use Positions Lab to accurately position the objects on your slide.\n\n" +
            "Click this button to open the Positions Lab interface.";
        #endregion

        public const string PaneTag = "PositionsLab";
        public const string TaskPanelTitle = "Positions Lab";

        public const string ErrorUndefined = "Undefined error in Positions Lab.";

        public const string ErrorRepositionFail = "Failed to reposition.";
        public const string ErrorDialogTitle = "Unable to execute action";
        public const string ErrorCorruptedShapesTitle = "Corrupted shapes detected";
        public const string ErrorCorruptedSelection = "Please undo the last operation and try again.";
        public const string ErrorNoSelection = "Please select at least a shape before using this feature.";
        public const string ErrorFewerThanTwoSelection = "Please select at least two shapes before using this feature.";
        public const string ErrorFewerThanThreeSelection = "Please select at least three shapes before using this feature.";
        public const string ErrorFewerThanFourSelection = "Please select at least four shapes before using this feature.";
        public const string ErrorFunctionNotSupportedForWithinShapes = "This function is not supported for Within Corner Most Objects Setting.";
        public const string ErrorFunctionNotSupportedForSlide = "This function is not supported for Within Slide Setting.";
        public const string ErrorFunctionNotSupportedForOverlapRefShapeCenter = "This function is not supported for shapes that overlap the center of the reference shape.";
    }
}
