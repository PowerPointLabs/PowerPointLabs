namespace PowerPointLabs.TextCollection
{
    internal static class ZoomLabText
    {
        #region Action Framework Tags
        public const string RibbonMenuId = "ZoomLabMenu";
        public const string DrillDownTag = "DrillDown";
        public const string StepBackTag = "StepBack";
        public const string ZoomToAreaTag = "ZoomToArea";
        public const string SettingsTag = "ZoomLabSettings";
        #endregion

        #region GUI Text
        public const string RibbonMenuLabel = "Zoom";
        public const string AddZoomInButtonLabel = "Drill Down";
        public const string AddZoomOutButtonLabel = "Step Back";
        public const string ZoomToAreaButtonLabel = "Zoom To Area";
        public const string SettingsButtonLabel = "Settings";

        public const string RibbonMenuSupertip = "Use Zoom Lab to creating zoom in and out effects for your slides easily.";
        public const string AddZoomInButtonSupertip =
            "Create an animation slide with a zoom-in effect from the currently selected shape to the next slide.\n\n" +
            "To perform this action, select a rectangle shape on the slide to drill down from, then click this button.";
        public const string AddZoomOutButtonSupertip =
            "Create an animation slide with a zoom-out effect from the previous slide to the currently selected shape.\n\n" +
            "To perform this action, select a rectangle shape on the slide to step back to, then click this button.";
        public const string ZoomToAreaButtonSupertip =
            "Zoom into an area of a slide or picture.\n\n" +
            "To perform this action, place a rectangle shape on the portion to magnify, then click this button.\n\n" +
            "This feature works best with high-resolution images.";
        public const string SettingsButtonSupertip = "Configure the settings for Zoom Lab.";

        // Dialog Boxes
        public const string SettingsSlideBackgroundCheckboxTooltip = "Include the slide background while using Zoom Lab.";
        public const string SettingsSeparateSlidesCheckboxTooltip = "Use separate slides for individual animation effects of Zoom To Area.";
        #endregion
    }
}
