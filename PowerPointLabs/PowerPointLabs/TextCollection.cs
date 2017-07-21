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
