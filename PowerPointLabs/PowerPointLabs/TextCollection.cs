namespace PowerPointLabs
{
    public class TextCollection1
    {
        # region Common Error
        public const string ErrorNameTooLong = "The name's length cannot be more than 255 characters.";
        public const string ErrorInvalidCharacter = "The name cannot be empty, or contain the following characters:'<', '>', ':', '\"', '/', '\\', '|', '?', or '*'.";
        public const string ErrorFileNameExist = "A file already exists with that name.";
        # endregion

        # region URLs
        public const string FeedbackUrl = "http://www.comp.nus.edu.sg/~pptlabs/contact.html";
        public const string HelpDocumentUrl = "http://www.comp.nus.edu.sg/~pptlabs/docs/";
        public const string PowerPointLabsWebsiteUrl = "http://PowerPointLabs.info";
        public const string SingleShapeDownloadUrl = "http://www.comp.nus.edu.sg/~pptlabs/gallery.html";
        # endregion

        # region Ribbon XML
        # region Supertips


        # region Crop Lab
        public const string CropLabMenuSupertip =
            "Use Crop Lab to crop shapes and pictures.";
        public const string CropLabSettingsSupertip =
            "Configure the settings for Crop Lab.";
        public const string MoveCropShapeButtonSupertip =
            "Crop a picture to a custom shape.\n\n" +
            "To perform this action, indicate area-to-keep by drawing one or more shapes upon the picture to crop, select the shape(s), then click this button.";
        public const string CropOutPaddingSupertip =
            "Crop away transparent areas of a picture.\n\n" +
            "To perform this action, select the picture(s) to crop out padding, then click this button.";
        public const string CropToAspectRatioSupertip =
            "Crop a picture to a specific aspect ratio.\n\n" +
            "To perform this action, select the picture(s) to crop to aspect ratio, then click this button.";
        public const string CropToSlideButtonSupertip =
            "Crop a shape or picture to fit within the slide boundaries.\n\n" +
            "To perform this action, select the shape(s) or picture(s) to crop to slide, then click this button.";
        public const string CropToSameButtonSupertip =
            "Crop multiple shapes to the same dimension.\n\n" +
            "To perform this action, select the shape of desired dimensions, then select the other shape(s) to crop, then click this button.";
        #endregion

        #region Narrations Lab
        public const string NarrationsLabMenuSupertip =
            "Use Narrations Lab to create narration from text in the Notes pane of the selected slides.";
        public const string AddNarrationsButtonSupertip =
            "Create synthesized narration from text in the Notes pane of the selected slides.\n\n" +
            "To perform this action, select the slide(s) with the Speaker Notes, then click this button.";
        public const string RecordNarrationsButtonSupertip =
            "Manually record audio to replace synthesized narration.\n\n" +
            "Click this button to open the Recording interface.";
        public const string RemoveNarrationsButtonSupertip =
            "Remove synthesized audio added using Narrations Lab from the selected slides.\n\n" +
            "To perform this action, select the slide(s) to remove the narrations from, then click this button.";
        public const string NarrationsLabSettingsSupertip =
            "Configure the settings for Narrations Lab.";
        #endregion

        #region Picture Slides Lab
        public const string PictureSlidesLabMenuLabel = "Picture Slides";
        public const string PictureSlidesLabMenuSupertip =
            "Use Picture Slides Lab to create better picture slides using less effort.\n\n" +
            "Click this button to open the Picture Slides Lab interface.";
        #endregion

        #region Captions Lab
        public const string CaptionsLabMenuSupertip =
            "Use Captions lab to create customizable movie-style subtitles from text in the Notes pane of the selected slides.";
        public const string AddCaptionsButtonSupertip =
            "Create movie-style subtitles from text in the Notes pane for the currently selected slides.\n\n" +
            "To perform this action, select the slide(s) with the notes, then click this button.";
        public const string RemoveCaptionsButtonSupertip =
            "Remove captions added using Captions Lab from the selected slides.\n\n" +
            "To perform this action, select the slide(s) to remove the captions from, then click this button.";
        public const string RemoveAllNotesButtonSupertip =
            "Remove notes from Notes pane of selected slides.\n\n" +
            "To perform this action, select the slide(s) to remove the notes from, then click this button.";
        # endregion

        # region Highlight Lab
        public const string HighlightLabMenuSupertip =
            "Use Highlight Lab to highlight bullet points and text.";
        public const string AddSpotlightButtonSupertip =
            "Create a spotlight effect for the slide using selected shapes.\n\n" +
            "To perform this action, draw a shape that the spotlight should outline, select it, then click this button.";
        public const string SpotlightPropertiesButtonSupertip =
            "Configure the settings for Spotlight.";
        public const string HighlightBulletsTextButtonSupertip =
            "Highlight selected bullet points by changing the font color.\n\n" +
            "To perform this action, select the bullet points to highlight, then click this button.";
        public const string HighlightBulletsBackgroundButtonSupertip =
            "Highlight selected bullet points by changing the text background color.\n\n" +
            "To perform this action, select the bullet points to highlight, then click this button.";
        public const string HighlightTextFragmentsButtonSupertip =
            "Highlight the selected text fragments.\n\n" +
            "To perform this action, select the text fragment(s) to highlight, then click this button.";
        public const string RemoveHighlightButtonSupertip =
            "Remove all Highlighting from the current slide.\n\n" +
            "To perform this action, click this button.";
        public const string HighlightLabSettingsSupertip =
            "Configure the settings for Highlight Lab.";
        #endregion

        #region More Labs
        #region Colors Lab
        public const string ColorsLabMenuSupertip =
            "Use Colors Lab to add beautiful colors to your slide.\n\n" +
            "Click this button to open the Colors Lab interface.";
        # endregion

        # region Shapes Lab
        public const string CustomShapeButtonSupertip =
            "Use Shapes Lab to manage your custom shapes.\n\n" +
            "Click this button to open the Shapes Lab interface.";
        # endregion

        # region Effects Lab
        public const string EffectsLabMenuSupertip =
            "Use Effects Lab to apply elegant effects to your shapes.";
        public const string EffectsLabMakeTransparentSupertip =
            "Adjust the transparency of pictures or shapes.\n\n" +
            "To perform this action, select the shape(s) or picture(s), then click this button.";
        public const string EffectsLabMagnifyGlassSupertip =
            "Magnify a small area or detail on the slide.\n\n" +
            "To perform this action, select the shape over the area to magnify, then click this button.";
        public const string EffectsLabBlurSelectedSupertip =
            "Blur the area covered by the selected shapes.\n\n" +
            "To perform this action, select the shape(s) over the area to blur, then click this button.";
        public const string EffectsLabBlurRemainderSupertip =
            "Blur evrything in the slide except for the area covered by the selected shapes.\n\n" +
            "To perform this action, select the shape(s) over the area to keep, then click this button.";
        public const string EffectsLabBlurBackgroundSupertip =
            "Blur everything in the slide except for the selected shapes.\n\n" +
            "To perform this action, select the shape(s) or picture(s) to keep, then click this button.";
        public const string EffectsLabRecolorRemainderSupertip =
            "Recolor an area of a slide to attract attention to it.\n\n" +
            "To perform this action, select the shape(s) over the area to keep, then click this button.";
        public const string EffectsLabRecolorBackgroundSupertip =
            "Recolor everything in the slide except for the selected shapes.\n\n" +
            "To perform this action, select the shape(s) or picture(s) to keep, then click this button.";
        # endregion

        # region Agenda Lab
        public const string AgendaLabSupertip =
            "Use Agenda Lab to generate professional-looking agendas automatically.\n\n" +
            "To use this feature, you need to group up your into appropriate sections. " +
            "Each section will be used as one item in the agenda.";
        public const string AgendaLabBulletPointSupertip =
            "Generate an agenda in bullet point style.\n\n" +
            "To perform this action, group your slides into sections, then click this button.";
        public const string AgendaLabVisualAgendaSupertip =
            "Generate an agenda in visual style.\n\n" +
            "To perform this action, group your slides into sections, then click this button.";
        public const string AgendaLabBeamAgendaSupertip =
            "Generate agenda side bar for selected slides.\n\n" +
            "To perform this action, group your slides into sections, then click this button.";
        public const string AgendaLabUpdateAgendaSupertip =
            "Synchronize agenda's layout and format with the first (template) slide.\n\n" +
            "To perform this action, make the changes you want on the first (template) slide, then click this button.";
        public const string AgendaLabRemoveAgendaSupertip =
            "Remove agenda generated by PowerPointLabs.\n\n" +
            "To perform this action, click this button";
        # endregion

        # region Positions Lab
        public const string PositionsLabMenuSupertip =
            "Use Positions Lab to accurately position the objects on your slide.\n\n" +
            "Click this button to open the Positions Lab interface.";
        #endregion

        #region Resize Lab
        public const string ResizeLabMenuSupertip =
            "Use Resize Lab to accurately resize the objects on your slide.\n\n" +
            "Click this button to open the Resize Lab interface.";
        #endregion

        #region Timer Lab
        public const string TimerLabMenuSupertip =
            "Use Timer Lab to create beautiful timers for your slides.\n\n" +
            "Click this button to open the Timer Lab interface.";
        #endregion

        #region Sync Lab
        public const string SyncLabMenuSupertip =
            "Use Sync Lab to make your slides look more consistent.\n\n" +
            "Click this button to open the Sync Lab interface.";
        #endregion

        #endregion

        #region Help
        public const string HelpMenuSupertip =
            "More information about PowerPointLabs.";
        public const string UserGuideButtonSupertip = "Click this button to visit PowerPointLabs help page in our website.";
        public const string TutorialButtonSupertip = "Click this button to open the tutorial for PowerPointLabs.";
        public const string FeedbackButtonSupertip = "Click this button to email us problem reports or other feedback. ";
        public const string AboutButtonSupertip = "Click this button for information about the PowerPointLabs add-in.";
        # endregion
        # endregion

        # region Tab Labels
        public const string PowerPointLabsAddInsTabLabel = "PowerPointLabs";
        #endregion

        #region Group Labels
        public const string AnimationsGroupLabel = "Animations";
        public const string AudioGroupLabel = "Audio";
        public const string EffectsGroupLabel = "Effects";
        public const string FormattingGroupLabel = "Formatting";
        public const string MoreLabsGroupLabel = "More Labs";
        #endregion

        #region Button Labels
        public const string CombineShapesLabel = "Combine Shapes";

        # region Crop Lab
        public const string CropLabMenuLabel = "Crop";
        public const string MoveCropShapeButtonLabel = "Crop To Shape";
        public const string CropOutPaddingLabel = "Crop Out Padding";
        public const string CropToAspectRatioLabel = "Crop To Aspect Ratio";
        public const string CropToSlideButtonLabel = "Crop To Slide";
        public const string CropToSameButtonLabel = "Crop To Same Dimensions";
        public const string CropLabSettingsButtonLabel = "Settings";
        # endregion

        # region Narrations Lab
        public const string NarrationsLabMenuLabel = "Narrations";
        public const string AddNarrationsButtonLabel = "Generate Audio Automatically";
        public const string RecordNarrationsButtonLabel = "Record Audio Manually";
        public const string RemoveNarrationsButtonLabel = "Remove Audio";
        public const string NarrationsLabSettingsButtonLabel = "Settings";
        # endregion

        # region Captions Lab
        public const string CaptionsLabMenuLabel = "Captions";
        public const string AddCaptionsButtonLabel = "Add Captions";
        public const string RemoveCaptionsButtonLabel = "Remove Captions";
        public const string RemoveAllNotesButtonLabel = "Remove All Notes";
        # endregion

        # region Highlight Lab
        public const string HighlightLabMenuLabel = "Highlight";
        public const string SpotlightPropertiesButtonLabel = "Settings";
        public const string SpotlightMenuLabel = "Spotlight";
        public const string AddSpotlightButtonLabel = "Create Spotlight";
        public const string ReloadSpotlightButtonLabel = "Recreate Spotlight";
        public const string HighlightBulletsTextButtonLabel = "Highlight Points";
        public const string HighlightBulletsBackgroundButtonLabel = "Highlight Background";
        public const string HighlightTextFragmentsButtonLabel = "Highlight Text";
        public const string RemoveHighlightButtonLabel = "Remove Highlighting";
        public const string HighlightLabSettingsButtonLabel = "Settings";
        #endregion

        #region Colors Lab
        public const string ColorPickerButtonLabel = "Colors";
        # endregion

        # region Shapes Lab
        public const string CustomeShapeButtonLabel = "Shapes";
        # endregion

        # region Effects Lab
        public const string EffectsLabButtonLabel = "Effects";
        public const string EffectsLabMakeTransparentButtonLabel = "Make Transparent";
        public const string EffectsLabMagnifyGlassButtonLabel = "Magnifying Glass";
        public const string EffectsLabBlurSelectedButtonLabel = "Blur Selected";
        public const string EffectsLabBlurRemainderButtonLabel = "Blur Remainder";
        public const string EffectsLabBlurBackgroundButtonLabel = "Blur All Except Selected";
        public const string EffectsLabRecolorRemainderButtonLabel = "Recolor Remainder";
        public const string EffectsLabRecolorBackgroundButtonLabel = "Recolor All Except Selected";
        public const string EffectsLabGrayScaleButtonLabel = "Gray Scale";
        public const string EffectsLabBlackWhiteButtonLabel = "Black and White";
        public const string EffectsLabGothamButtonLabel = "Gotham";
        public const string EffectsLabSepiaButtonLabel = "Sepia";
        # endregion

        # region Agenda Lab
        public const string AgendaLabButtonLabel = "Agenda";
        public const string AgendaLabBulletPointButtonLabel = "Create Text Agenda";
        public const string AgendaLabVisualAgendaButtonLabel = "Create Visual Agenda";
        public const string AgendaLabBeamAgendaButtonLabel = "Create Sidebar Agenda";
        public const string AgendaLabUpdateAgendaButtonLabel = "Synchronize Agenda";
        public const string AgendaLabRemoveAgendaButtonLabel = "Remove Agenda";
        public const string AgendaLabAgendaSettingsButtonLabel = "Agenda Settings";
        public const string AgendaLabBulletAgendaSettingsButtonLabel = "Bullet Agenda Settings";
        # endregion

        # region Positions Lab
        public const string PositionsLabButtonLabel = "Positions";
        # endregion

        # region Resize Lab
        public const string ResizeLabButtonLabel = "Resize";
        #endregion

        #region Sync Lab
        public const string SyncLabButtonLabel = "Sync";
        # endregion

        # region Timer Lab
        public const string TimerLabButtonLabel = "Timer";
        # endregion
        
        # region Help Menu
        public const string HelpMenuLabel = "Help";
        public const string UserGuideButtonLabel = "User Guide";
        public const string TutorialButtonLabel = "Tutorial";
        public const string FeedbackButtonLabel = "Report Issues/ Send Feedback";
        public const string AboutButtonLabel = "About";
        #endregion
        #endregion

        # region Dynamic Menu Labels

        public const string DynamicMenuId = "DynamicMenu";
        public const string DynamicMenuButtonId = "Button";
        public const string DynamicMenuCheckBoxId = "CheckBox";
        public const string DynamicMenuOptionId = "Option";
        public const string DynamicMenuSeparatorId = "Separator";
        public const string DynamicMenuXmlButton = "<button id=\"{0}\" tag=\"{1}\" getLabel=\"GetLabel\" onAction=\"OnAction\"/>";
        public const string DynamicMenuXmlImageButton = "<button id=\"{0}\" tag=\"{1}\" getLabel=\"GetLabel\" getImage=\"GetImage\" getEnabled=\"GetEnabled\" onAction=\"OnAction\"/>";
        public const string DynamicMenuXmlCheckBox = "<checkBox id=\"{0}\" tag=\"{1}\" getLabel=\"GetLabel\" getPressed=\"GetPressed\" onAction=\"OnCheckBoxAction\"/>";
        public const string DynamicMenuXmlMenu = "<menu xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">{0}</menu>";
        public const string DynamicMenuXmlMenuSeparator = "<menuSeparator id=\"{0}Separator\"/>";
        public const string DynamicMenuXmlTitleMenuSeparator = "<menuSeparator id=\"{0}Separator\" title=\"{1}\"/>";
        public const string EffectsLabBlurrinessButtonLabel = "Settings";
        public const string EffectsLabBlurrinessCheckBoxLabel = "Tint ";
        public const string EffectsLabBlurrinessTag = "Blurriness";
        public const string EffectsLabBlurrinessCustom = "CustomPercentage";
        public const string EffectsLabBlurrinessCustomPrefixLabel = "Custom";
        public const string EffectsLabBlurrinessFeatureSelected = "BlurSelected";
        public const string EffectsLabBlurrinessFeatureRemainder = "BlurRemainder";
        public const string EffectsLabBlurrinessFeatureBackground = "BlurBackground";
        public const string EffectsLabRecolorTag = "Recolor";

        # endregion

        #region Context Menu Labels

        public const string PowerPointLabsMenuLabel = "PowerPointLabs";
        public const string PasteLabMenuLabel = "Paste Lab";
        public const string ShortcutsLabMenuLabel = "Shortcuts";

        public const string EditNameShapeLabel = "Edit Name";
        public const string SpotlightShapeLabel = "Add Spotlight";
        public const string ZoomInContextMenuLabel = "Drill Down";
        public const string ZoomOutContextMenuLabel = "Step Back";
        public const string ZoomToAreaContextMenuLabel = "Zoom To Area";
        public const string HighlightBulletsMenuShapeLabel = "Highlight Bullets";
        public const string HighlightBulletsTextShapeLabel = "Highlight Text";
        public const string HighlightBulletsBackgroundShapeLabel = "Highlight Background";
        public const string ConvertToPictureShapeLabel = "Convert To Picture";
        public const string AddIntoGroup = "Add Into Group";
        public const string AddCustomShapeShapeLabel = "Add To Shapes Lab";
        public const string HideSelectedShapeLabel = "Hide Shape";
        public const string CutOutShapeShapeLabel = "Crop To Shape";
        public const string InSlideAnimateGroupLabel = "Animate In-Slide";
        public const string ApplyAutoMotionThumbnailLabel = "Add Animation Slide";
        public const string ContextSpeakSelectedTextLabel = "Speak Selected Text";
        public const string ContextAddCurrentSlideLabel = "Add Audio (Current Slide)";
        public const string ContextReplaceAudioLabel = "Replace Audio";
        #endregion

        #region Ribbon Content
        public const string RibbonButton = "Button";
        public const string RibbonMenu = "Menu";

        public const string AnimationsGroupId = "AnimationsGroup";
        public const string AnimationLabMenuId = "AnimationLabMenu";
        public const string AddAnimationSlideTag = "AddAnimationSlide";
        public const string AnimateInSlideTag = "AnimateInSlide";
        public const string AnimationLabSettingsTag = "AnimationLabSettings";

        public const string ZoomLabMenuId = "ZoomLabMenu";
        public const string DrillDownTag = "DrillDown";
        public const string StepBackTag = "StepBack";
        public const string ZoomToAreaTag = "ZoomToArea";
        public const string ZoomLabSettingsTag = "ZoomLabSettings";

        public const string AudioGroupId = "AudioGroup";
        public const string NarrationsLabMenuId = "NarrationsLabMenu";
        public const string AddNarrationsTag = "AddNarrations";
        public const string RecordNarrationsTag = "RecordNarrations";
        public const string RemoveNarrationsTag = "RemoveNarrations";
        public const string NarrationsLabSettingsTag = "NarrationsLabSettings";

        public const string CaptionsLabMenuId = "CaptionsLabMenu";
        public const string AddCaptionsTag = "AddCaptions";
        public const string RemoveCaptionsTag = "RemoveCaptions";
        public const string RemoveNotesTag = "RemoveNotes";
        public const string CaptionsLabSettingsTag = "CaptionsLabSettings";

        public const string EffectsGroupId = "EffectsGroup";
        public const string HighlightLabMenuId = "HighlightLabMenu";
        public const string HighlightPointsTag = "HighlightPoints";
        public const string HighlightBackgroundTag = "HighlightBackground";
        public const string HighlightTextTag = "HighlightText";
        public const string RemoveHighlightTag = "RemoveHighlight";
        public const string HighlightLabSettingsTag = "HighlightLabSettings";

        public const string EffectsLabMenuId = "EffectsLabMenu";
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

        public const string FormattingGroupId = "FormattingGroup";
        public const string PositionsLabTag = "PositionsLab";

        public const string ResizeLabTag = "ResizeLab";

        public const string ColorsLabTag = "ColorsLab";

        public const string SyncLabTag = "SyncLab";

        public const string ShapesLabTag = "ShapesLab";

        public const string CropLabMenuId = "CropLabMenu";
        public const string CropToShapeTag = "CropToShape";
        public const string CropToSlideTag = "CropToSlide";
        public const string CropToSameDimensionsTag = "CropToSame";
        public const string CropToAspectRatioTag = "CropToAspectRatio";
        public const string CropOutPaddingTag = "CropOutPadding";
        public const string CropLabSettingsTag = "CropLabSettings";

        public const string PasteLabMenuId = "PasteLabMenu";
        public const string PasteAtCursorPositionTag = "PasteAtCursorPosition";
        public const string PasteAtOriginalPositionTag = "PasteAtOriginalPosition";
        public const string PasteToFillSlideTag = "PasteToFillSlide";
        public const string ReplaceWithClipboardTag = "ReplaceWithClipboard";
        public const string PasteIntoGroupTag = "PasteIntoGroup";

        public const string MoreLabsGroupId = "MoreLabsGroup";
        public const string TimerLabTag = "TimerLab";

        public const string AgendaLabMenuId = "AgendaLabMenu";
        public const string TextAgendaTag = "TextAgenda";
        public const string VisualAgendaTag = "VisualAgenda";
        public const string BeamAgendaTag = "BeamAgenda";
        public const string RemoveAgendaTag = "RemoveAgenda";
        public const string UpdateAgendaTag = "UpdateAgenda";

        public const string PictureSlidesLabTag = "PictureSlidesLab";

        public const string HelpMenuId = "HelpMenu";
        public const string HelpTag = "Help";
        public const string TutorialTag = "Tutorial";
        public const string FeedbackTag = "Feedback";
        public const string AboutTag = "About";
        #endregion

        #region Context Menu Content

        public const string MenuShape = "MenuShape";
        public const string MenuLine = "MenuLine";
        public const string MenuFreeform = "MenuFreeform";
        public const string MenuPicture = "MenuPicture";
        public const string MenuGroup = "MenuGroup";
        public const string MenuInk = "MenuInk";
        public const string MenuVideo = "MenuVideo";
        public const string MenuTextEdit = "MenuTextEdit";
        public const string MenuChart = "MenuChart";
        public const string MenuTable = "MenuTable";
        public const string MenuTableCell = "MenuTableWhole";
        public const string MenuSlide = "MenuFrame";
        public const string MenuSmartArt = "MenuSmartArtBackground";
        public const string MenuEditSmartArtBase = "MenuSmartArtEdit";
        public const string MenuEditSmartArt = MenuEditSmartArtBase + "SmartArt";
        public const string MenuEditSmartArtText = MenuEditSmartArtBase + "Text";
        public const string MenuNotes = "MenuNotes";

        public const string MenuSeparator = "MenuSeparator";

        public const string EditNameTag = "EditName";
        public const string ConvertToPictureTag = "ConvertToPicture";
        public const string HideShapeTag = "HideShape";
        public const string AddCustomShapeTag = "AddShape";
        public const string AddIntoGroupTag = "AddIntoGroup";
        public const string SpeakSelectedTag = "SpeakSelected";
        #endregion
        #endregion

        #region PowerPointSlide

        public const string NotesPageStorageText = "This notes page is used to store data - Do not edit the notes. ";

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
        public const string RecManagementPanelTitle = "Record Management";
        public const string PositionsLabTaskPanelTitle = "Positions Lab";
        public const string ResizeLabsTaskPaneTitle = "Resize Lab";
        public const string TimerLabTaskPaneTitle = "Timer Lab";
        public const string SyncLabTaskPanelTitle = "Sync Lab";
        #endregion

        #region Graphics
        public const string TemporaryImageStorageFileName = "temp.png";
        #endregion

        #region ShapeGalleryPresentation
        public const string ShapeCorruptedError =
            "Some shapes in the Shapes Lab were corrupted, but some of the them are recovered.";
        # endregion

        #region ConvertToPicture
        public const string ErrorTypeNotSupported = "Convert to Picture only supports Shapes and Charts.";
        public const string ErrorWindowTitle = "Convert to Picture: Unsupported Object";
        # endregion

        #region Task Pane - Recorder
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

        #region Narrations Lab
        // Dialog Boxes
        public const string NarrationsLabSettingsVoiceSelectionInputTooltip = 
            "The voice to be used when generating synthesized audio.\n" +
            "Use [Voice] tags to specify a different voice for a particular section of text.";
        public const string NarrationsLabSettingsPreviewCheckboxTooltip =
            "If checked, the current slide's audio and animations will play after the Add Audio button is clicked.";
        #endregion

        #region Zoom Lab
        // Dialog Boxes
        public const string ZoomLabSettingsSlideBackgroundCheckboxTooltip = "Include the slide background while using Zoom Lab.";
        public const string ZoomLabSettingsSeparateSlidesCheckboxTooltip = "Use separate slides for individual animation effects of Zoom To Area.";
        #endregion

        #region Task Pane - Resize Lab
        public class ResizeLabText
        {
            public const string ErrorInvalidSelection = "You need to select at least {1} {2} before applying '{0}'";
            public const string ErrorNotSameShapes = "You need to select the same type of objects before applying 'Adjust Area Proportionally'";
            public const string ErrorGroupShapeNotSupported = "'Adjust Area Proportionally' does not support grouped objects";
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
        public const string ReportIssueEmail = @"pptlabs@comp.nus.edu.sg";
        # endregion

        # region Install and Update related

        public const string QuickTutorialFileName = "Tutorial.pptx";
        public const string VstoName = "PowerPointLabsInstaller.vsto";
        public const string InstallerName = "data.zip";

        # endregion
    }
}
