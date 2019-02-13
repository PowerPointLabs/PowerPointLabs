namespace PowerPointLabs.TextCollection
{
    internal static class TooltipsLabText
    {
        #region Action Framework Tags
        public const string RibbonMenuId = "TooltipsLabMenu";
        public const string CreateTooltipTag = "CreateTooltip";
        public const string AssignTooltipTag = "AssignTooltip";
        public const string AddTextboxTag = "AddTextbox";
        public const string SettingsTag = "TooltipsLabSettings";
        #endregion

        #region GUI Text
        public const string RibbonMenuLabel = "Tooltips";
        public const string CreateTooltipButtonLabel = "Create Tooltip";
        public const string AssignTooltipButtonLabel = "Assign Tooltip";
        public const string AddTextboxButtonLabel = "Add Textbox";
        public const string SettingsButtonLabel = "Settings";

        public const string RibbonMenuSupertip = "Use Animation Lab to add animations your slides easily.";
        public const string AddAnimationButtonSupertip =
            "Create an animation slide to transition from the currently selected slide to the next slide.\n\n" +
            "To perform this action, duplicate the currently selected slide, move the objects to the desired position, select the original slide, then click this button.";
        public const string InSlideAnimateButtonSupertip =
            "Moves a shape around the slide in multiple steps.\n\n" +
            "To perform this action, copy the shape to locations where it should stop, select the copies in the order they should appear, then click this button.";
        public const string SettingsButtonSupertip = "Configure the settings for Animation Lab.";

        public const string AutoAnimateLoadingText = "Applying auto animation...";
        public const string SettingsDurationInputTooltip = "The duration (in seconds) for the animations in the animation slides to be created.";
        public const string SettingsSmoothAnimationCheckboxTooltip =
            "Use a frame-based approach for smoother resize animations.\n" +
            "This may result in larger file sizes and slower loading times for animated slides.";

        public const string ErrorAutoAnimateDialogTitle = "Unable to execute action";
        public const string ErrorAutoAnimateWrongSlide = "Please select the correct slide.";
        public const string ErrorAutoAnimateNoMatchingShapes = "No matching Shapes were found on the next slide.";
        public const string ErrorAutoAnimateSlideNotAutoAnimate = "The current slide was not added by PowerPointLabs Auto Animate";
        #endregion
    }
}
