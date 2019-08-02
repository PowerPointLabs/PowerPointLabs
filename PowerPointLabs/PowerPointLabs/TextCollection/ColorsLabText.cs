namespace PowerPointLabs.TextCollection
{
    internal static class ColorsLabText
    {
        #region Action Framework Tags
        public const string RibbonMenuId = "ColorsLabButton";
        public const string RibbonMenuLabel = "Colors";
        public const string RibbonMenuSupertip =
            "Use Colors Lab to add beautiful colors to your slide.\n\n" +
            "Click this button to open the Colors Lab interface.";
        #endregion

        public const string PaneTag = "ColorsLab";
        public const string TaskPanelTitle = "Colors Lab";

        public const string MainColorBoxTooltip = "Choose the main color: " +
                                   "\r\nDrag the box to favorites palatte, " +
                                   "\r\nor click it to choose one from the Color dialog.";
        public const string FontColorButtonTooltip = "Change the font color of the selected shapes: " +
                                                 "\r\nDrag the button to pick a color.";
        public const string LineColorButtonTooltip = "Change the line color of the selected shapes: " +
                                                 "\r\nDrag the button to pick a color.";
        public const string FillColorButtonTooltip = "Change the fill color of the selected shapes: " +
                                                 "\r\nDrag the button to pick a color.";
        public const string EyeDropperButtonTooltip = "Choose the main color with eyedropper: " +
                                                    "\r\nDrag the button to pick a color.";
        public const string BrightnessSliderTooltip = "Move the slider to adjust the main color’s brightness.";
        public const string SaturationSliderTooltip = "Move the slider to adjust the main color’s saturation.";
        public const string SaveFavoriteColorsButtonTooltip = "Save the favorite color palette.";
        public const string LoadFavoriteColorsButtonTooltip = "Load an existing favorite color palette.";
        public const string ResetFavoriteColorsButtonTooltip = "Reset the current favorite color palette to those last loaded.";
        public const string EmptyFavoriteColorsButtonTooltip = "Empty the favorite color palette.";
        public const string ColorRectangleTooltip = "Click the color to select it as the main color. You can drag-and-drop these colors into the favorites palette.";
        public const string FavoriteColorRectangleTooltip = "Click the color to select it as the main color.";

        public const string ErrorDialogTitle = "Unable to execute action";
        public const string ErrorNoSelection = "To use this feature, select at least one shape.";
    }
}
