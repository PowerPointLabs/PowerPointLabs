namespace PowerPointLabs.TextCollection
{
    internal static class CropLabText
    {
        #region Action Framework Tags
        public const string RibbonMenuId = "CropLabMenu";
        public const string CropToShapeTag = "CropToShape";
        public const string CropToSlideTag = "CropToSlide";
        public const string CropToSameDimensionsTag = "CropToSame";
        public const string CropToAspectRatioTag = "CropToAspectRatio";
        public const string CropOutPaddingTag = "CropOutPadding";
        public const string SettingsTag = "CropLabSettings";
        #endregion

        #region GUI Text
        public const string RibbonMenuLabel = "Crop";
        public const string CropToShapeButtonLabel = "Crop To Shape";
        public const string CropOutPaddingButtonLabel = "Crop Out Padding";
        public const string CropToAspectRatioButtonLabel = "Crop To Aspect Ratio";
        public const string CropToSlideButtonLabel = "Crop To Slide";
        public const string CropToSameButtonLabel = "Crop To Same Dimensions";
        public const string SettingsButtonLabel = "Settings";

        public const string RibbonMenuSupertip = "Use Crop Lab to crop shapes and pictures.";
        public const string CropToShapeButtonSupertip =
            "Crop a picture to a custom shape.\n\n" +
            "To perform this action, indicate area-to-keep by drawing one or more shapes upon the picture to crop, select the shape(s), then click this button.";
        public const string CropOutPaddingButtonSupertip =
            "Crop away transparent areas of a picture.\n\n" +
            "To perform this action, select the picture(s) to crop out padding, then click this button.";
        public const string CropToAspectRatioButtonSupertip =
            "Crop a picture to a specific aspect ratio.\n\n" +
            "To perform this action, select the picture(s) to crop to aspect ratio, then click this button.";
        public const string CropToSlideButtonSupertip =
            "Crop a shape or picture to fit within the slide boundaries.\n\n" +
            "To perform this action, select the shape(s) or picture(s) to crop to slide, then click this button.";
        public const string CropToSameButtonSupertip =
            "Crop multiple shapes to the same dimension.\n\n" +
            "To perform this action, select the shape of desired dimensions, then select the other shape(s) to crop, then click this button.";
        public const string SettingsButtonSupertip = "Configure the settings for Crop Lab.";

        public const string ErrorUndefined = "Undefined error in '{0}'.";

        public const string ErrorSelectionIsInvalid = "You need to select at least {1} {2} before applying '{0}'.";
        public const string ErrorSelectionCountZero = "'{0}' requires at least one shape to be selected.";
        public const string ErrorSelectionNonPicture = "'{0}' only supports picture objects.";

        public const string ErrorSelectionMustBeShape = "'{0}' only supports shape objects.";
        public const string ErrorSelectionMustBePicture = "'{0}' only supports picture objects.";
        public const string ErrorSelectionMustBeShapeOrPicture = "'{0}' only supports shape or picture objects.";

        public const string ErrorNoShapeOverBoundary = "All selected objects are inside the slide boundary. No cropping was done.";
        public const string ErrorNoDimensionCropped = "All selected pictures are smaller than reference shape. No cropping was done.";
        public const string ErrorNoPaddingCropped = "All selected pictures have no transparent padding. No cropping was done.";
        public const string ErrorNoAspectRatioCropped = "All selected pictures are already in the given aspect ratio. No cropping was done.";

        public const string ErrorAspectRatioIsInvalid = "The given aspect ratio is invalid. Please enter positive numbers for the width to height ratio.";
        #endregion
    }
}
