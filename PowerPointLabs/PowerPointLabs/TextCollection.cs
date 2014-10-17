namespace PowerPointLabs
{
    internal class TextCollection
    {
        # region Common Error
        public const string ErrorNameTooLong = "The name's length cannot be more than 255 characters.";
        public const string ErrorInvalidCharacter = "The name cannot be empty, or contain the following characters:'<', '>', ':', '\"', '/', '\\', '|', '?', or '*'.";
        public const string ErrorFileNameExist = "A file already exists with that name.";
        # endregion

        # region URLs
        public const string FeedbackUrl = "http://powerpointlabs.info/contact.html";
        public const string HelpDocumentUrl = "http://powerpointlabs.info/docs.html";
        public const string PowerPointLabsWebsiteUrl = "http://PowerPointLabs.info";
        # endregion

        # region Ribbon XML
        # region Supertips
        public const string AddAnimationButtonSupertip =
            "Creates an animation slide to transition from the currently selected slide to the next slide.";
        public const string ReloadButtonSupertip =
            "Recreates an existing animation slide with new animations.\n\n" +
            "To activate, select the original slide or the animation slide then click this button.";
        public const string InSlideAnimateButtonSupertip =
            "Moves a shape around the slide in multiple steps.\n\n" +
            "To activate, copy the shape to locations where it should stop, select the copies in the order they should appear, then click this button";
        
        public const string AddZoomInButtonSupertip =
            "Creates an animation slide with a zoom-in effect from the currently selected shape to the next slide.\n\n" +
            "To activate, select a rectangle shape on the slide to drill down from, then click this button.";
        public const string AddZoomOutButtonSupertip =
            "Creates an animation slide with a zoom-out effect from the previous slide to the currently selected shape.\n\n" +
            "To activate, select a rectangle shape on the slide to step back to, then click this button.";
        public const string ZoomToAreaButtonSupertip =
            "Zoom into an area of a slide or image.\n\nTo activate, place a rectangle shape on the portion to magnify, then click this button.\n\n" +
            "This feature works best with high-resolution images.";
        
        public const string MoveCropShapeButtonSupertip =
            "Crop a picture to a custom shape.\n\n" +
            "To activate, draw one or more shapes upon the picture to crop, select the shape(s), then click this button.";
        
        public const string AddSpotlightButtonSupertip =
            "Creates a spotlight effect for a selected shape.\n\n" +
            "To activate, draw a shape that the spotlight should outline, select it, then click this button.";
        public const string ReloadSpotlightButtonSupertip =
            "Adjusts the transparency and edges of an existing spotlight.\n\n" +
            "To activate, set the transparency level and soft edges width, select the existing spotlight shape, then click this button.";
        
        public const string AddAudioButtonSupertip =
            "Creates synthesized narration from text in the Speaker Notes pane of the selected slides.";
        public const string GenerateRecordButtonSupertip =
            "Creates synthesized narration from text in a slide's Speaker Notes pane.";
        public const string AddRecordButtonSupertip =
            "Manually record audio to replace synthesized narration.";
        public const string RemoveAudioButtonSupertip =
            "Removes synthesized audio added using Auto Narrate from the selected slides."; 
        
        public const string AddCaptionsButtonSupertip =
            "Creates movie-style subtitles from text in the Speaker Notes pane, and adds it to the selected slides.";
        public const string RemoveCaptionsButtonSupertip =
            "Removes captions added using Auto Captions from the selected slides.";
        public const string RemoveAllNotesButtonSupertip = "Remove notes from note pane of selected slides.";
        
        public const string HighlightBulletsTextButtonSupertip =
            "Highlights selected bullet points by changing the text's color.\n\n" +
            "To activate, select the bullet points to highlight, then click this button.";
        public const string HighlightBulletsBackgroundButtonSupertip =
            "Highlights selected bullet points by changing the text's background color.\n\n" +
            "To activate, select the bullet points to highlight, then click this button.";
        public const string HighlightTextFragmentsButtonSupertip =
            "Highlights the selected text fragments.\n\n" +
            "To activate, select the text to highlight, then click this button.";
        
        public const string ColorPickerButtonSupertip = @"Opens Custom Color Picker";
        
        public const string CustomeShapeButtonSupertip = @"Manage your custom shapes."; // Custome -> Custom?
        
        public const string HelpButtonSupertip = @"Click this to visit PowerPointLabs help page in our website.";
        public const string FeedbackButtonSupertip = @"Click this to email us problem reports or other feedback. ";
        public const string AboutButtonSupertip = @"Information about the PowerPointLabs plugin.";
        # endregion

        # region Tab Labels
        public const string PowerPointLabsAddInsTabLabel = "PowerPointLabs";
        # endregion

        # region Button Labels
        public const string CombineShapesLabel = "Combine Shapes";

        public const string AutoAnimateGroupLabel = "Auto Animate";
        public const string AddAnimationButtonLabel = "Add Animation Slide";
        public const string AddAnimationReloadButtonLabel = "Recreate Animation";
        public const string AddAnimationInSlideAnimateButtonLabel = "Animate In Slide";

        public const string AutoZoomGroupLabel = "Auto Zoom";
        public const string AddZoomInButtonLabel = "Drill Down";
        public const string AddZoomOutButtonLabel = "Step Back";
        public const string ZoomToAreaButtonLabel = "Zoom To Area";

        public const string AutoCropGroupLabel = "Auto Crop";
        public const string MoveCropShapeButtonLabel = "Crop To Shape";

        public const string SpotLightGroupLabel = "Spotlight";
        public const string AddSpotlightButtonLabel = "Create Spotlight";
        public const string ReloadSpotlightButtonLabel = "Recreate Spotlight";

        public const string EmbedAudioGroupLabel = "Auto Narrate";
        public const string AddAudioButtonLabel = "Add Audio";
        public const string GenerateRecordButtonLabel = "Generate Audio Automatically";
        public const string AddRecordButtonLabel = "Record Audio Manually";
        public const string RemoveAudioButtonLabel = "Remove Audio";

        public const string EmbedCaptionGroupLabel = "Auto Captions";
        public const string AddCaptionsButtonLabel = "Add Captions";
        public const string RemoveCaptionsButtonLabel = "Remove Captions";
        public const string RemoveAllNotesButtonLabel = "Remove All Notes";

        public const string HighlightBulletsGroupLabel = "Highlight Bullets";
        public const string HighlightBulletsTextButtonLabel = "Highlight Points";
        public const string HighlightBulletsBackgroundButtonLabel = "Highlight Background";
        public const string HighlightTextFragmentsButtonLabel = "Highlight Text";

        public const string LabsGroupLabel = "Labs";
        public const string ColorPickerButtonLabel = "Colors Lab";
        public const string CustomeShapeButtonLabel = "Shapes Lab";

        public const string PPTLabsHelpGroupLabel = "Help";
        public const string HelpButtonLabel = "Help";
        public const string FeedbackButtonLabel = "Report Issues/ Send Feedback";
        public const string AboutButtonLabel = "About";
        # endregion

        # region Context Menu Labels
        public const string NameEditShapeLabel = "Edit Name";
        public const string SpotlightShapeLabel = "Add Spotlight";
        public const string ZoomInContextMenuLabel = "Drill Down";
        public const string ZoomOutContextMenuLabel = "Step Back";
        public const string ZoomToAreaContextMenuLabel = "Zoom To Area";
        public const string HighlightBulletsMenuShapeLabel = "Highlight Bullets";
        public const string HighlightBulletsTextShapeLabel = "Highlight Text";
        public const string HighlightBulletsBackgroundShapeLabel = "Highlight Background";
        public const string ConvertToPictureShapeLabel = "Convert to Picture";
        public const string AddCustomShapeShapeLabel = "Add to Shapes Lab";
        public const string CutOutShapeShapeLabel = "Crop To Shape";
        public const string FitToWidthShapeLabel = "Fit To Width";
        public const string FitToHeightShapeLabel = "Fit To Height";
        public const string InSlideAnimateGroupLabel = "Animate In-Slide";
        public const string ApplyAutoMotionThumbnailLabel = "Add Animation Slide";
        public const string ContextSpeakSelectedTextLabel = "Speak Selected Text";
        public const string ContextAddCurrentSlideLabel = "Add Audio (Current Slide)";
        public const string ContextReplaceAudioLabel = "Replace Audio";
        # endregion
        # endregion

        # region Ribbon
        public static readonly string AboutInfo =
            "          PowerPointLabs Plugin Version " + Properties.Settings.Default.Version + " [Release date: " + Properties.Settings.Default.ReleaseDate + "]\n     Developed at School of Computing, National University of Singapore.\n        For more information, visit our website " + PowerPointLabsWebsiteUrl;

        public const string AboutInfoTitle = "About PowerPointLabs";
        # endregion

        # region ThisAddIn
        public const string AccessTempFolderErrorMsg = "Error when accessing temp folder";
        public const string CreatTempFolderErrorMsg = "Error when creating temp folder";
        public const string ExtraErrorMsg = "Error when extracting";
        public const string PrepareMediaErrorMsg = "Error when preparing media files";
        public const string VersionNotCompatibleErrorMsg =
            "Some features of PowerPointLabs do not work with presentations saved in " +
            "the .ppt format. To use them, please resave the " +
            "presentation with the .pptx format.";
        public const string OnlinePresentationNotCompatibleErrorMsg =
        	"Some features of PowerPointLabs do not work with online presentations. " +
        	"To use them, please save the file locally.";
        public const string ShapeGalleryInitErrorMsg =
            "Could not connect to shape database from your default location.\n\n" +
            "To check your default location, right click on the panel and select 'Settings' option.";
        public const string TabActivateErrorTitle = "Unable to activate 'Double Click to Open Property' feature";
        public const string TabActivateErrorDescription =
            "To activate 'Double Click to Open Property' feature, you need to enable 'Home' tab " +
            "in Options -> Customize Ribbon -> Main Tabs -> tick the checkbox of 'Home' -> click OK but" +
            "ton to save.";
        public const string ShapesLabTaskPanelTitle = "Shapes Lab";
        public const string ColorsLabTaskPanelTitle = "Colors Lab";
        public const string RecManagementPanelTitle = "Record Management";
        # endregion 

        # region ShapeGalleryPresentation
        public const string ShapeCorruptedError =
            "Some shapes in the Shapes Lab were corrupted, but some of the them are recovered.";
        # endregion

        # region CropToShape

        public class CropToShapeText
        {
            //------------ Msg -------------
            public const string ErrorMessageForSelectionCountZero = "'Crop To Shape' requires at least one shape to be selected.";
            public const string ErrorMessageForSelectionNonShape = "'Crop To Shape' only supports shape objects.";
            public const string ErrorMessageForExceedSlideBound = "The selected shape needs to be within the slide's boundaries.";
            public const string ErrorMessageForRotationNonZero = "'Crop To Shape' does not currently support rotated shapes.";
            public const string ErrorMessageForUndefined = "Undefined error in 'Crop To Shape'.";
        }

        # endregion

        # region ConvertToPicture
        public const string ErrorTypeNotSupported = "Convert to Picture only supports Shapes and Charts.";
        public const string ErrorWindowTitle = "Convert to Picture: Unsupported Object";
        # endregion

        # region Task Pane - Recorder
        public const string RecorderInitialTimer = "00:00:00";
        public const string RecorderReadyStatusLabel = "Ready.";
        public const string RecorderRecordingStatusLabel = "Recording...";
        public const string RecorderPlayingStatusLabel = "Playing...";
        public const string RecorderPauseStatusLabel = "Pause";
        public const string RecorderUnrecognizeAudio = "Unrecognized Embedded Audio";
        public const string RecorderScriptStatusNoAudio = "No Audio";
        public const string RecorderWndMessageError = "Fatal error";
        public const string RecorderNoScriptDetail = "No Script Available";
        public const string RecorderNoInputDeviceMsg = "No audio input device was found.\n" +
                                                       "Check that a microphone or other audio input device is attached " +
                                                       "and working.";
        public const string RecorderNoInputDeviceMsgBoxTitle = "Input Device Not Found";
        public const string RecorderSaveRecordMsg = "Do you want to save the recording?";
        public const string RecorderSaveRecordMsgBoxTitle = "Save Recording";
        public const string RecorderReplaceRecordMsgFormat = "Do you want to replace\n{0}\nwith the current recording?";
        public const string RecorderReplaceRecordMsgBoxTitle = "Replace Audio";
        public const string RecorderNoRecordToPlayError = "There are no recordings to play.";
        public const string RecorderInvalidOperation = "Invalid Operation";
        # endregion

        # region Task Pane - Colors Lab

        public class ColorsLabText
        {
            //----------- Tooltips -----------
            public const string MainColorBoxTooltips = "Choose the main color: " +
                                       "\r\nDrag the button to pick a color, " +
                                       "\r\nor click it to choose one from the Color dialog.";
            public const string FontColorButtonTooltips = "Change the font color of the selected shapes: " +
                                                     "\r\nDrag the button to pick a color, " +
                                                     "\r\nor click it to choose one from the Color dialog.";
            public const string LineColorButtonTooltips = "Change the line color of the selected shapes: " +
                                                     "\r\nDrag the button to pick a color, " +
                                                     "\r\nor click it to choose one from the Color dialog.";
            public const string FillColorButtonTooltips = "Change the fill color of the selected shapes: " +
                                                     "\r\nDrag the button to pick a color, " +
                                                     "\r\nor click it to choose one from the Color dialog.";
            public const string BrightnessSliderTooltips = "Move the slider to adjust the main color’s brightness.";
            public const string SaturationSliderTooltips = "Move the slider to adjust the main color’s saturation.";
            public const string SaveFavoriteColorsButtonTooltips = "Save the favorite color palette.";
            public const string LoadFavoriteColorsButtonTooltips = "Load an existing favorite color palette.";
            public const string ResetFavoriteColorsButtonTooltips = "Reset the current favorite color palette to those last loaded.";
            public const string EmptyFavoriteColorsButtonTooltips = "Empty the favorite color palette.";
            public const string ColorRectangleTooltips = "Click the color to select it as the main color. You can drag-and-drop these colors into the favorites palette.";
            public const string ThemeColorRectangleTooltips = "Click the color to select it as the main color.";

            //------------ Msg ------------
            public const string InfoHowToActivateFeature = "To use this feature, select at least one shape.";
        }
        
        # endregion

        # region Task Pane - Shapes Lab
        public const string CustomShapeFileNameInvalid = "Invalid shape name.";
        public const string CustomShapeNoShapeTextFirstLine = "No shapes saved yet.";
        public const string CustomShapeNoShapeTextSecondLine = "Right-click any object on a slide to save it in this panel.";
        public const string CustomShapeNoPanelSelectedError = "No shape selected.";
        public const string CustomShapeViewTypeNotSupported = "Shapes Lab does not support the current view type.";
        public const string CustomeShapeSaveLocationChangedSuccessFormat =
            "Default saving path has been changed to \n{0}\nAll shapes have been moved to the new location.";
        public const string CustomeShapeSetAsDefaultCategorySuccessFormat = "{0} has been set as default category.";
        public const string CustomShapeSaveLocationChangedSuccessTitle = "Success";
        public const string CustomShapeMigrationError =
            "The folder cannot be migrated entirely. Please check if your destination loaction forbids this action.";
        public const string CustomShapeOriginalFolderDeletionError =
            "The original folder could not be deleted because some of the files in folder is still in use. You could " +
            "try to delete this folder manually when those files are closed.";
        public const string CustomShapeMigratingDialogTitle = "Migrating...";
        public const string CustomShapeMigratingDialogContent = "Shapes are being migrated, please wait...";
        public const string CustomShapeRemoveLastCategoryError = "Removing the last category is not allowed.";
        public const string CustomShapeDuplicateCategoryNameError = "The name has already been used.";
        public const string CustomShapeRemoveDefaultCategoryMessage =
            "You are removing your default category. After removing this category, the first category will be made " +
            "as default category. Continue?";
        public const string CustomShapeRemoveDefaultCategoryCaption = "Removing Default Category";
        public const string CustomShapeImportFileError = "Import File could not be opened.";
        public const string CustomShapeImportSuccess = "Successfully imported";
        
        public const string CustomShapeShapeContextStripAddToSlide = "Add To Slide";
        public const string CustomShapeShapeContextStripEditName = "Edit Name";
        public const string CustomShapeShapeContextStripMoveShape = "Move Shape To";
        public const string CustomShapeShapeContextStripRemoveShape = "Remove Shape";
        public const string CustomShapeShapeContextStripCopyShape = "Copy Shape To";

        public const string CustomShapeCategoryContextStripAddCategory = "Add Category";
        public const string CustomShapeCategoryContextStripRemoveCategory = "Remove Category";
        public const string CustomShapeCategoryContextStripRenameCategory = "Rename Category";
        public const string CustomShapeCategoryContextStripImportCategory = "Import Category";
        public const string CustomShapeCategoryContextStripSetAsDefaultCategory = "Set as Default";
        public const string CustomShapeCategoryContextStripCategorySettings = "Settings";
        # endregion

        # region Control - ShapesLabSetting
        public const string FolderDialogDescription = "Select the directory that you want to use as the default.";
        public const string FolderNonEmptyErrorMsg = "Please select an empty folder as default saving folder.";
        # endregion

        # region Control - SlideShow Recorder Control
        public const string InShowControlInvalidRecCommandError = "Invalid Recording Command";
        public const string InShowControlRecButtonIdleText = "Stop and Advance";
        public const string InShowControlRecButtonRecText = "Start Recording";
        # endregion

        #region Control - Loading Dialog
        public const string LoadingDialogDefaultTitle = "Loading...";
        public const string LoadingDialogDefaultContent = "Loading, please wait...";
        # endregion

        # region Error Dialog
        public const string UserFeedBack = " Help us fix the problem by emailing ";
        public const string Email = @"pptlabs@comp.nus.edu.sg";
        # endregion

        # region Install and Update related

        public const string QuickTutorialFileName = "Tutorial.pptx";
        public const string VstoName = "PowerPointLabsInstaller.vsto";
        public const string InstallerName = "data.zip";

        # endregion
    }
}
