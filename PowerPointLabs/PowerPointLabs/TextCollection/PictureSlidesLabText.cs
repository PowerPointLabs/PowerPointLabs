namespace PowerPointLabs.TextCollection
{
    internal static class PictureSlidesLabText
    {
        public const string PaneTag = "PictureSlidesLab";

        public const string RibbonMenuLabel = "Picture Slides";
        public const string RibbonMenuSupertip =
            "Use Picture Slides Lab to create better picture slides using less effort.\n\n" +
            "Click this button to open the Picture Slides Lab interface.";

        /// <summary>
        /// Styles Variation Category Name
        ///
        /// Leave OptionName to be ColorNoEffect to hide color panel & picker
        ///
        /// Color category's name (without spaces) should be equal to corresponding style option's
        /// property name, so that the color picker can work properly
        /// </summary>
        public const string NoEffect = "No Effect";
        public const string ColorHasEffect = "Color";
        public const string TransparencyHasEffect = "Transparency";
        public const string BannerHasEffect = "Banner";
        public const string TextBoxHasEffect = "TextBox";
        public const string VariantCategoryOverlayColor = "Overlay Color";
        public const string VariantCategoryFontColor = "Font Color";
        public const string VariantCategoryTextGlowColor = "Text Glow Color";
        public const string VariantCategoryBannerColor = "Banner Color";
        public const string VariantCategoryTextBoxColor = "TextBox Color";
        public const string VariantCategoryFrostedGlassTextBoxColor = "TextBox Color";
        public const string VariantCategoryFrostedGlassBannerColor = "Banner Color";
        public const string VariantCategoryFrameColor = "Frame Color";
        public const string VariantCategoryCircleColor = "Circle Color";
        public const string VariantCategoryTriangleColor = "Triangle Color";
        public const string VariantCategoryOutlineColor = "Outline Color";
        public const string VariantCategoryTextPosition = "Text Position";
        public const string VariantCategoryFontFamily = "Font";
        public const string VariantCategorySpecialEffects = "Special Effects";
        public const string VariantCategoryBlurriness = "Blurriness";
        public const string VariantCategoryBrightness = "Brightness";
        public const string VariantCategoryFrostedGlassTextBoxTransparency = "TextBox Transparency";
        public const string VariantCategoryFrostedGlassBannerTransparency = "Banner Transparency";
        public const string VariantCategoryFontSizeIncrease = "Font Size";
        public const string VariantCategoryPicture = "Picture";
        public const string VariantCategoryImageReference = "Picture Citation";
        public const string VariantCategoryOverlayTransparency = "Overlay Transparency";
        public const string VariantCategoryBannerTransparency = "Banner Transparency";
        public const string VariantCategoryTextBoxTransparency = "TextBox Transparency";
        public const string VariantCategoryFrameTransparency = "Frame Transparency";
        public const string VariantCategoryCircleTransparency = "Circle Transparency";
        public const string VariantCategoryTriangleTransparency = "Triangle Transparency";
        public const string VariantCategoryTextTransparency = "Text Transparency";

        /// <summary>
        /// Styles Preview Name
        /// </summary>
        public const string StyleNameDirectText = "Direct Text Style";
        public const string StyleNameBlur = "Blur Style";
        public const string StyleNameTextBox = "TextBox Style";
        public const string StyleNameBanner = "Banner Style";
        public const string StyleNameSpecialEffect = "Special Effect Style";
        public const string StyleNameOverlay = "Overlay Style";
        public const string StyleNameOutline = "Outline Style";
        public const string StyleNameFrame = "Frame Style";
        public const string StyleNameCircle = "Circle Style";
        public const string StyleNameTriangle = "Triangle Style";
        public const string StyleNameFrostedGlassTextBox = "Frosted Glass TextBox Style";
        public const string StyleNameFrostedGlassBanner = "Frosted Glass Banner Style";

        /// <summary>
        /// Messages
        /// </summary>
        public const string ErrorImageCorrupted =
            "Failed to load image. The image file is corrupted.";
        public const string ErrorImageDownloadCorrupted =
            "Failed to load image. Please try again.";
        public const string ErrorFailedToLoad =
            "Failed to load image. ";
        public const string ErrorUrlLinkIncorrect =
            "The download link is not in the correct format. Did the link miss out 'http://'?";
        public const string ErrorNoSelectedSlide =
            "Cannot apply styles. Please select a slide first.";
        public const string ErrorFailToInitTempFolder =
            "Failed to initialize Picture Slides Lab. Please verify that sufficient permissions have been granted by Administrator.";
        public const string ErrorNoEmbeddedStyleInfo =
            "No Picture Slides Lab styles are detected for the current slide.";
        public const string ErrorWhenInitialize =
            "Failed to initialize Picture Slides Lab. Some functions may not work.";

        public const string DragAndDropInstruction =
            "Drag and Drop here to get image.";

        public const string InfoPasteNothing = "No picture to paste.";
        public const string InfoPasteThumbnail = "Pasted successfully! But you might have pasted the thumbnail picture.";
        public const string InfoAddPictureCitationSlide = "Added successfully!";
        public const string InfoDeleteAllImage = "Do you want to delete all pictures?";
    }
}
