namespace PowerPointLabs.TextCollection
{
    internal static class EffectsLabText
    {
        #region Action Framework Tags
        public const string RibbonMenuId = "EffectsLabMenu";
        public const string MakeTransparentTag = "MakeTransparent";
        public const string MagnifyTag = "Magnify";
        public const string SpotlightMenuId = "SpotlightMenu";
        public const string AddSpotlightTag = "AddSpotlight";
        public const string SpotlightSettingsTag = "SpotlightSettings";
        public const string BlurSelectedMenuId = "BlurSelectedMenu";
        public const string BlurRemainderMenuId = "BlurRemainderMenu";
        public const string BlurBackgroundMenuId = "BlurBackgroundMenu";
        public const string RecolorRemainderMenuId = "RecolorRemainderMenu";
        public const string RecolorBackgroundMenuId = "RecolorBackgroundMenu";
        public const string GrayScaleTag = "GrayScale";
        public const string BlackWhiteTag = "BlackAndWhite";
        public const string GothamTag = "Gotham";
        public const string SepiaTag = "Sepia";
        #endregion

        #region GUI Text
        public const string RibbonMenuLabel = "Effects";
        public const string MakeTransparentButtonLabel = "Make Transparent";
        public const string MagnifyGlassButtonLabel = "Magnifying Glass";
        public const string SpotlightMenuLabel = "Spotlight";
        public const string AddSpotlightButtonLabel = "Create Spotlight";
        public const string SpotlightSettingsButtonLabel = "Settings";
        public const string BlurSelectedButtonLabel = "Blur Selected";
        public const string BlurRemainderButtonLabel = "Blur Remainder";
        public const string BlurBackgroundButtonLabel = "Blur All Except Selected";
        public const string RecolorRemainderButtonLabel = "Recolor Remainder";
        public const string RecolorBackgroundButtonLabel = "Recolor All Except Selected";
        public const string GrayScaleButtonLabel = "Gray Scale";
        public const string BlackWhiteButtonLabel = "Black and White";
        public const string GothamButtonLabel = "Gotham";
        public const string SepiaButtonLabel = "Sepia";

        public const string BlurrinessButtonLabel = "Settings";
        public const string BlurrinessCheckBoxLabel = "Tint ";
        public const string BlurrinessTag = "Blurriness";
        public const string BlurrinessCustom = "CustomPercentage";
        public const string BlurrinessCustomPrefixLabel = "Custom";
        public const string BlurrinessFeatureSelected = "BlurSelected";
        public const string BlurrinessFeatureRemainder = "BlurRemainder";
        public const string BlurrinessFeatureBackground = "BlurBackground";
        public const string RecolorTag = "Recolor";

        public const string RibbonMenuSupertip = "Use Effects Lab to apply elegant effects to your shapes.";
        public const string MakeTransparentSupertip =
            "Adjust the transparency of pictures or shapes.\n\n" +
            "To perform this action, select the shape(s) or picture(s), then click this button.";
        public const string MagnifyGlassSupertip =
            "Magnify a small area or detail on the slide.\n\n" +
            "To perform this action, select the shape over the area to magnify, then click this button.";
        public const string AddSpotlightButtonSupertip =
            "Create a spotlight effect for the slide using selected shapes.\n\n" +
            "To perform this action, draw a shape that the spotlight should outline, select it, then click this button.";
        public const string SpotlightSettingsButtonSupertip =
            "Configure the settings for Spotlight.";
        public const string BlurSelectedSupertip =
            "Blur the area covered by the selected shapes.\n\n" +
            "To perform this action, select the shape(s) over the area to blur, then click this button.";
        public const string BlurRemainderSupertip =
            "Blur evrything in the slide except for the area covered by the selected shapes.\n\n" +
            "To perform this action, select the shape(s) over the area to keep, then click this button.";
        public const string BlurBackgroundSupertip =
            "Blur everything in the slide except for the selected shapes.\n\n" +
            "To perform this action, select the shape(s) or picture(s) to keep, then click this button.";
        public const string RecolorRemainderSupertip =
            "Recolor an area of a slide to attract attention to it.\n\n" +
            "To perform this action, select the shape(s) over the area to keep, then click this button.";
        public const string RecolorBackgroundSupertip =
            "Recolor everything in the slide except for the selected shapes.\n\n" +
            "To perform this action, select the shape(s) or picture(s) to keep, then click this button.";

        public const string ErrorDialogTitle = "Unable to execute action";
        public const string ErrorBlurSelectedNoSelection = "'Blur Selected' requires at least one shape or text box to be selected.";
        public const string ErrorBlurSelectedNonShapeOrTextBox = "'Blur Selected' only supports shape and text box objects.";

        public const string SettingsTintSelectedCheckboxLabel = "Tint Selected";
        public const string SettingsTintRemainderCheckboxLabel = "Tint Remainder";
        public const string SettingsTintBackgroundCheckboxLabel = "Tint All Except Selected";
        public const string SettingsTintCheckboxTooltip = "Adds a tinted effect to your blur";
        public const string SettingsBlurrinessInputTooltip = "The level of blurriness";
        public const string SettingsTransparencyInputTooltip = "The transparency level of the spotlight effect to be created";
        public const string SettingsSoftEdgesSelectionInputTooltip = "The softness of the edges of the spotlight effect to be created";
        #endregion
    }
}
