namespace PowerPointLabs.TextCollection
{
    internal static class TooltipsLabText
    {
        #region Action Framework Tags
        public const string RibbonMenuId = "TooltipsLabMenu";
        public const string CreateMenuTag = "CreateMenu";
        public const string CreateTooltipTag = "CreateTooltip";
        public const string CreateCalloutTag = "CreateCallout";
        public const string CreateTriggerTag = "CreateTrigger";
        public const string AssignTooltipTag = "AssignTooltip";
        public const string AddTextboxTag = "AddTextbox";
        public const string SettingsTag = "TooltipsLabSettings";
        #endregion

        #region GUI Text
        public const string RibbonMenuLabel = "Tooltips";
        public const string CreateMenuLabel = "Create";
        public const string CreateTooltipButtonLabel = "Tooltip";
        public const string CreateCalloutButtonLabel = "Callout";
        public const string CreateTriggerButtonLabel = "Trigger";
        public const string AssignTooltipButtonLabel = "Assign Tooltip";
        public const string AddTextboxButtonLabel = "Add Textbox";
        public const string SettingsButtonLabel = "Settings";

        public const string RibbonMenuSupertip = "Use Tooltips Lab to add Tooltips easily.";
        public const string CreateTooltipButtonSupertip =
            "Create a trigger shape and a callout shape.";
        public const string CreateCalloutButtonSupertip =
            "Create a callout with a specified trigger shape. \n\n" +
            "To perform this action, select your desired trigger shape. A callout will be generated for you.";
        public const string CreateTriggerButtonSupertip =
            "Create a trigger shape with a specified callout. \n\n" +
            "To perform this action, select your desired callout. A trigger shape will be generated for you.";
        public const string AssignTooltipButtonSupertip =
            "Attach a trigger animation to a group of shapes. \n\n" +
            "To perform this action, select a group of shapes, the first shape selected with the trigger shape.";
        public const string SettingsButtonSupertip = "Configure the settings for Tooltips Lab.";

        public const string ErrorTooltipsDialogTitle = "Unable to execute action";
        public const string ErrorLessThanTwoShapesSelected = "Please select at least two shapes. The first shape will be the trigger shape.";
        #endregion
    }
}