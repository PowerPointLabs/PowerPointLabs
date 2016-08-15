namespace PowerPointLabs
{
    public class TextCollection
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
        public const string SingleShapeDownloadUrl = "http://www.comp.nus.edu.sg/~pptlabs/gallery.html";
        # endregion

        # region Ribbon XML
        # region Supertips
        # region Auto Animation
        public const string AddAnimationButtonSupertip =
            "Creates an animation slide to transition from the currently selected slide to the next slide.";
        public const string ReloadButtonSupertip =
            "Recreates an existing animation slide with new animations.\n\n" +
            "To activate, select the original slide or the animation slide then click this button.";
        public const string InSlideAnimateButtonSupertip =
            "Moves a shape around the slide in multiple steps.\n\n" +
            "To activate, copy the shape to locations where it should stop, select the copies in the order they should appear, then click this button";
        # endregion

        # region Auto Zoom
        public const string AddZoomInButtonSupertip =
            "Creates an animation slide with a zoom-in effect from the currently selected shape to the next slide.\n\n" +
            "To activate, select a rectangle shape on the slide to drill down from, then click this button.";
        public const string AddZoomOutButtonSupertip =
            "Creates an animation slide with a zoom-out effect from the previous slide to the currently selected shape.\n\n" +
            "To activate, select a rectangle shape on the slide to step back to, then click this button.";
        public const string ZoomToAreaButtonSupertip =
            "Zoom into an area of a slide or image.\n\nTo activate, place a rectangle shape on the portion to magnify, then click this button.\n\n" +
            "This feature works best with high-resolution images.";
        # endregion

        # region Auto Crop
        public const string MoveCropShapeButtonSupertip =
            "Crop a picture to a custom shape.\n\n" +
            "To activate, draw one or more shapes upon the picture to crop, select the shape(s), then click this button.";
        # endregion

        # region Spotlight
        public const string AddSpotlightButtonSupertip =
            "Creates a spotlight effect for a selected shape.\n\n" +
            "To activate, draw a shape that the spotlight should outline, select it, then click this button.";
        # endregion

        # region Auto Narrate
        public const string AddAudioButtonSupertip =
            "Creates synthesized narration from text in the Speaker Notes pane of the selected slides.";
        public const string GenerateRecordButtonSupertip =
            "Creates synthesized narration from text in a slide's Speaker Notes pane.";
        public const string AddRecordButtonSupertip =
            "Manually record audio to replace synthesized narration.";
        public const string RemoveAudioButtonSupertip =
            "Removes synthesized audio added using Auto Narrate from the selected slides.";
        # endregion

        # region Auto Caption
        public const string AddCaptionsButtonSupertip =
            "Creates movie-style subtitles from text in the Speaker Notes pane, and adds it to the selected slides.";
        public const string RemoveCaptionsButtonSupertip =
            "Removes captions added using Auto Captions from the selected slides.";
        public const string RemoveAllNotesButtonSupertip = "Remove notes from note pane of selected slides.";
        # endregion

        # region Highlight Points
        public const string HighlightBulletsTextButtonSupertip =
            "Highlights selected bullet points by changing the text's color.\n\n" +
            "To activate, select the bullet points to highlight, then click this button.";
        public const string HighlightBulletsBackgroundButtonSupertip =
            "Highlights selected bullet points by changing the text's background color.\n\n" +
            "To activate, select the bullet points to highlight, then click this button.";
        public const string HighlightTextFragmentsButtonSupertip =
            "Highlights the selected text fragments.\n\n" +
            "To activate, select the text to highlight, then click this button.";
        # endregion

        # region Labs
        # region Colors Lab
        public const string ColorPickerButtonSupertip = @"Opens Custom Color Picker";
        # endregion

        # region Shapes Lab
        public const string CustomeShapeButtonSupertip = @"Manage your custom shapes.";
        # endregion

        # region Effects Lab
        public const string EffectsLabMenuSupertip = @"Apply elegant effects to shapes.";
        public const string EffectsLabMakeTransparentSupertip = @"Adjust the transparency of pictures or shapes.";
        public const string EffectsLabMagnifyGlassSupertip = @"Magnify a small area or detail on the slide.";
        public const string EffectsLabBlurSelectedSupertip = @"Blur the selected shapes.";
        public const string EffectsLabBlurRemainderSupertip = @"Draw attention to an area of the slide by blurring everything else.";
        public const string EffectsLabColorizeRemainderSupertip = @"Recolor an area of a slide to attract attention to it.";
        public const string EffectsLabBlurBackgroundSupertip = @"Blur everything in the slide except for the selected shapes.";
        public const string EffectsLabColorizeBackgroundSupertip = @"Recolor everything in the slide except for the selected shapes.";
        # endregion

        # region Agenda Lab
        public const string AgendaLabSupertip = "Generate professional-look agenda automatically.\n\n To use this feature, you need " +
                                                "to group up your into appropriate sections. Each section will be used as one item in " +
                                                "the agenda.";
        public const string AgendaLabBulletPointSupertip = "Generate an agenda in bullet point style.";
        public const string AgendaLabVisualAgendaSupertip = "Generate an agenda in visual style.";
        public const string AgendaLabBeamAgendaSupertip = "Generate agenda side bar for selected slides.";
        public const string AgendaLabUpdateAgendaSupertip = "Synchronize agenda's layout and format with the first slide.";
        public const string AgendaLabRemoveAgendaSupertip = "Remove agenda generated by PowerPointLabs.";
        public const string AgendaLabAgendaSettingsSupertip = "Configure agenda settings.";
        public const string AgendaLabBulletAgendaSettingsSupertip = "Set color scheme for Bullet Agenda.";
        # endregion

        # region Drawing Lab
        public const string DrawingsLabButtonSupertip = @"Opens the Drawing Lab Interface";
        #endregion

        # region Drawing Lab
        public const string PositionsLabSupertip = "Open Positions Lab window.";
        #endregion

        #region Resize Lab
        public const string ResizeLabButtonSupertip = "Opens the Resize Lab Interface";
        #endregion

        #endregion

        #region Help
        public const string HelpButtonSupertip = @"Click this to visit PowerPointLabs help page in our website.";
        public const string FeedbackButtonSupertip = @"Click this to email us problem reports or other feedback. ";
        public const string AboutButtonSupertip = @"Information about the PowerPointLabs plugin.";
        # endregion
        # endregion

        # region Tab Labels
        public const string PowerPointLabsAddInsTabLabel = "PowerPointLabs";
        # endregion

        # region Button Labels
        public const string CombineShapesLabel = "Combine Shapes";

        # region Auto Animation
        public const string AutoAnimateGroupLabel = "Animation Lab";
        public const string AddAnimationButtonLabel = "Add Animation Slide";
        public const string AddAnimationReloadButtonLabel = "Recreate Animation";
        public const string AddAnimationInSlideAnimateButtonLabel = "Animate In Slide";
        # endregion

        # region Auto Zoom
        public const string AutoZoomGroupLabel = "Zoom Lab";
        public const string AddZoomInButtonLabel = "Drill Down";
        public const string AddZoomOutButtonLabel = "Step Back";
        public const string ZoomToAreaButtonLabel = "Zoom To Area";
        # endregion

        # region Auto Crop
        public const string AutoCropGroupLabel = "Crop Lab";
        public const string MoveCropShapeButtonLabel = "Crop To Shape";
        # endregion

        # region Spotlight
        public const string SpotLightGroupLabel = "Spotlight Lab";
        public const string AddSpotlightButtonLabel = "Create Spotlight";
        public const string ReloadSpotlightButtonLabel = "Recreate Spotlight";
        # endregion

        # region Auto Narration
        public const string EmbedAudioGroupLabel = "Narrations Lab";
        public const string AddAudioButtonLabel = "Add Audio";
        public const string GenerateRecordButtonLabel = "Generate Audio Automatically";
        public const string AddRecordButtonLabel = "Record Audio Manually";
        public const string RemoveAudioButtonLabel = "Remove Audio";
        # endregion

        # region Auto Caption
        public const string EmbedCaptionGroupLabel = "Captions Lab";
        public const string AddCaptionsButtonLabel = "Add Captions";
        public const string RemoveCaptionsButtonLabel = "Remove Captions";
        public const string RemoveAllNotesButtonLabel = "Remove All Notes";
        # endregion

        # region Highlight Points
        public const string HighlightBulletsGroupLabel = "Highlight Bullets Lab";
        public const string HighlightBulletsTextButtonLabel = "Highlight Points";
        public const string HighlightBulletsBackgroundButtonLabel = "Highlight Background";
        public const string HighlightTextFragmentsButtonLabel = "Highlight Text";
        # endregion

        # region Labs
        public const string LabsGroupLabel = "Labs";

        # region Colors Lab
        public const string ColorPickerButtonLabel = "Colors Lab";
        # endregion

        # region Shapes Lab
        public const string CustomeShapeButtonLabel = "Shapes Lab";
        # endregion

        # region Effects Lab
        public const string EffectsLabButtonLabel = "Effects Lab";
        public const string EffectsLabMakeTransparentButtonLabel = "Make Transparent";
        public const string EffectsLabMagnifyGlassButtonLabel = "Magnifying Glass";
        public const string EffectsLabBlurSelectedButtonLabel = "Blur Selected";
        public const string EffectsLabBlurRemainderButtonLabel = "Blur Remainder";
        public const string EffectsLabBlurBackgroundButtonLabel = "Blur All Except Selected";
        public const string EffectsLabRecolorRemainderButtonLabel = "Recolor Remainder";
        public const string EffectsLabRecolorBackgroundButtonLabel = "Recolor All Except Selected";
        # endregion

        # region Agenda Lab
        public const string AgendaLabButtonLabel = "Agenda Lab";
        public const string AgendaLabBulletPointButtonLabel = "Create Text Agenda";
        public const string AgendaLabVisualAgendaButtonLabel = "Create Visual Agenda";
        public const string AgendaLabBeamAgendaButtonLabel = "Create Sidebar Agenda";
        public const string AgendaLabUpdateAgendaButtonLabel = "Synchronize Agenda";
        public const string AgendaLabRemoveAgendaButtonLabel = "Remove Agenda";
        public const string AgendaLabAgendaSettingsButtonLabel = "Agenda Settings";
        public const string AgendaLabBulletAgendaSettingsButtonLabel = "Bullet Agenda Settings";
        # endregion

        # region Drawing Lab
        public const string DrawingsLabButtonLabel = "Drawing Lab";
        # endregion

        # region Positions Lab
        public const string PositionsLabButtonLabel = "Positions Lab";
        # endregion

        # region Resize Lab
        public const string ResizeLabButtonLabel = "Resize Lab";
        # endregion
        # endregion

        # region Help
        public const string PPTLabsHelpGroupLabel = "Help";
        public const string HelpButtonLabel = "Help";
        public const string FeedbackButtonLabel = "Report Issues/ Send Feedback";
        public const string AboutButtonLabel = "About";
        # endregion
        # endregion

        # region Context Menu Labels

        public const string PowerPointLabsMenuLabel = "PowerPointLabs";
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
        public const string HideSelectedShapeLabel = "Hide the Shape";
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

        # region PowerPointSlide

        public const string NotesPageStorageText = "This notes page is used to store data - Do not edit the notes. ";

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
            "To check your default location, right click on the Shapes Lab's panel and select 'Settings' option.";
        public const string TabActivateErrorTitle = "Unable to activate 'Double Click to Open Property' feature";
        public const string TabActivateErrorDescription =
            "To activate 'Double Click to Open Property' feature, you need to enable 'Home' tab " +
            "in Options -> Customize Ribbon -> Main Tabs -> tick the checkbox of 'Home' -> click OK but" +
            "ton to save.";
        public const string ShapesLabTaskPanelTitle = "Shapes Lab";
        public const string ColorsLabTaskPanelTitle = "Colors Lab";
        public const string DrawingsLabTaskPanelTitle = "Drawing Lab";
        public const string RecManagementPanelTitle = "Record Management";
        public const string PositionsLabTaskPanelTitle = "Positions Lab";
        public const string ResizeLabsTaskPaneTitle = "Resize Lab";
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
            public const string ErrorMessageForUndefined = "Undefined error in 'Crop To Shape'.";
        }

        # endregion

        # region ConvertToPicture
        public const string ErrorTypeNotSupported = "Convert to Picture only supports Shapes and Charts.";
        public const string ErrorWindowTitle = "Convert to Picture: Unsupported Object";
        # endregion

        public class PictureSlidesLabText
        {
            public const string PictureSlidesLabSupertip = "Open Picture Slides Lab window.";

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

        #region Agenda Lab
        // Errors
        public const string AgendaLabErrorDialogTitle = "Unable to execute action";
        public const string AgendaLabNoSectionError = "Please group the slides into sections before generating agenda.";
        public const string AgendaLabSingleSectionError = "Please divide the slides into two or more sections.";
        public const string AgendaLabEmptySectionError = "Presentation contains empty section(s). Please fill them up or remove them.";
        public const string AgendaLabAgendaExistError = "Agenda already exists. The previous agenda will be removed and regenerated. Do you want to proceed?";
        public const string AgendaLabAgendaExistErrorCaption = "Confirm Update";
        public const string AgendaLabNoAgendaError = "There is no generated agenda.";
        public const string AgendaLabNoReferenceSlideError = "No reference slide could be found. Either replace the reference slide or regenerate the agenda.";
        public const string AgendaLabInvalidReferenceSlideError = "The current reference slide is invalid. Either replace the reference slide or regenerate the agenda.";
        public static string AgendaLabSectionNameTooLongError = "One of the section names exceeds the maximum size allowed by Agenda Lab. Please rename the section accordingly.";

        // Dialog Boxes
        public const string AgendaLabGeneratingDialogTitle = "Generating...";
        public const string AgendaLabGeneratingDialogContent = "Agenda is generating, please wait...";
        public const string AgendaLabSynchronizingDialogTitle = "Synchronizing...";
        public const string AgendaLabSynchronizingDialogContent = "Agenda is being synchronized, please wait...";

        public const string AgendaLabReorganiseSidebarTitle = "Reorganise Sidebar";
        public const string AgendaLabReorganiseSidebarContent = "The sections have been changed. Do you wish to reorganise the items in the sidebar?";

        public const string AgendaLabBeamGenerateSingleSlideDialogTitle = "Generate on all slides";
        public const string AgendaLabBeamGenerateSingleSlideDialogContent = "Only one slide is selected. Would you like to generate the sidebar on all slides instead?";

        // Agenda Content
        public const string AgendaLabTitleContent = "Agenda";

        public const string AgendaLabBulletVisitedContent = "Visited bullet format";
        public const string AgendaLabBulletHighlightedContent = "Highlighted bullet format";
        public const string AgendaLabBulletUnvisitedContent = "Unvisited bullet format";
        public const string AgendaLabBeamHighlightedText = "Highlighted";

        public const string AgendaLabTemplateSlideInstructions =
                            "This slide is used as a ‘Template' for generating agenda slides. Please do not delete this slide.\r" +
                            "Adjust the design of this slide and click the 'Sync Agenda' (in Agenda Lab) to replicate the design in the other slides.";
        # endregion

        # region Drawing Lab

        public const string DrawingsLabSelectExactlyOneShape = "Please select a single shape.";
        public const string DrawingsLabSelectAtLeastOneShape = "Please select at least one shape.";
        public const string DrawingsLabSelectExactlyTwoShapes = "Please select two shapes.";
        public const string DrawingsLabSelectAtLeastTwoShapes = "Please select at least two shapes.";
        public const string DrawingsLabSelectTwoSetsOfShapes = "Please select two sets of shapes.";
        public const string DrawingsLabSelectStartAndEndShape = "Please select a start shape and an end shape";

        public const string DrawingsLabErrorCannotGroup = "These shapes cannot be grouped.";
        public const string DrawingsLabErrorNothingUngrouped = "Please select shapes that have been grouped.";

        public const string DrawingsLabMultiCloneDialogText = "Number of Extra Copies";
        public const string DrawingsLabMultiCloneDialogHeader = "Multi-Clone";

        public const string DrawingsLabSetTextDialogHeader = "Set Text";

        # endregion

        # region Effects Lab

        public const string EffectsLabBlurSelectedErrorNoSelection = "'Blur Selected'  requires at least one shape or text box to be selected.";
        public const string EffectsLabBlurSelectedErrorNonShapeOrTextBox = "'Blur Selected' only supports shape and text box objects.";

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
        public const string CustomShapeDefaultShapeName = "My Shape Untitled";

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
        public const string CustomShapeImportNoSlideError = "Import File is empty.";
        public const string CustomShapeImportAppendCategoryError = "Your computer does not support this feature.";
        public const string CustomShapeImportSingleCategoryErrorFormat =
            "{0} contains multiple categories. Try \"Import Category\" instead.";
        public const string CustomShapeImportSuccess = "Successfully imported";

        public const string CustomShapeImportShapeFileDialogTitle = "Import Shapes";
        public const string CustomShapeImportLibraryFileDialogTitle = "Import Library";
        
        public const string CustomShapeShapeContextStripAddToSlide = "Add To Slide";
        public const string CustomShapeShapeContextStripEditName = "Edit Name";
        public const string CustomShapeShapeContextStripMoveShape = "Move Shape To";
        public const string CustomShapeShapeContextStripRemoveShape = "Remove Shape";
        public const string CustomShapeShapeContextStripCopyShape = "Copy Shape To";

        public const string CustomShapeCategoryContextStripAddCategory = "Add Category";
        public const string CustomShapeCategoryContextStripRemoveCategory = "Remove Category";
        public const string CustomShapeCategoryContextStripRenameCategory = "Rename Category";
        public const string CustomShapeCategoryContextStripImportCategory = "Import Library";
        public const string CustomShapeCategoryContextStripImportShapes = "Import Shapes";
        public const string CustomShapeCategoryContextStripSetAsDefaultCategory = "Set as Default Category";
        public const string CustomShapeCategoryContextStripCategorySettings = "Shapes Lab Settings";
        #endregion

        #region Positions Lab
        public class PositionsLabText
        {
            public const string ErrorNoSelection = "Please select at least a shape before using this feature";
            public const string ErrorFewerThanTwoSelection = "Please select at least two shapes before using this feature";
            public const string ErrorFewerThanThreeSelection = "Please select at least three shapes before using this feature";
            public const string ErrorFunctionNotSupportedForWithinShapes = "This function is not supported for Within Corner Most Objects Setting.";
            public const string ErrorFunctionNotSupportedForSlide = "This function is not supported for Within Slide Setting.";
            public const string ErrorUndefined = "'Undefined error in Resize Lab'";
        }
        #endregion

        #region Task Pane - Resize Lab

        public class ResizeLabText
        {
            public const string ErrorInvalidSelection = "You need to select at least {1} {2} before applying '{0}'";
            public const string ErrorUndefined = "'Undefined error in Resize Lab'";
        }

        #endregion

        #region Control - ShapesLabSetting
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
