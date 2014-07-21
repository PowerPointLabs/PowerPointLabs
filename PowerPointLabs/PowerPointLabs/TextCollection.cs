namespace PowerPointLabs
{
    internal class TextCollection
    {
        # region URLs
        public const string FeedbackUrl = "http://powerpointlabs.info/contact.html";
        public const string HelpDocumentUrl = "http://powerpointlabs.info/docs.html";
        public const string PowerPointLabsWebsiteUrl = "http://PowerPointLabs.info";
        # endregion

        # region Ribbon XML
        public const string AddAnimationButtonSupertip =
            "Adds a slide with animations to create an animated transition from the currently selected slide to the next slide.";
        public const string ReloadButtonSupertip =
            "Recreates an existing animation slide with new animations.\n\n" +
            "To activate, select the original slide or the animation slide (added previously) and click this button.";
        public const string InSlideAnimateButtonSupertip =
            "Moves a shape around the slide in multiple steps.\n\n" +
            "To activate, copy the shape to locations where you want it to stop, select the copies in the order you want them to appear, and click this button";
        public const string AddZoomInButtonSupertip =
            "Adds a slide with animations to create a zoom-in effect from the currently selected shape to the next slide.\n\n" +
            "To activate, select a rectangle shape on the slide that you want to drill down from and click on this button.";
        public const string AddZoomOutButtonSupertip =
            "Adds a slide with animations to create a zoom-out effect from the previous slide to the currently selected shape.\n\n" +
            "To activate, select a rectangle shape on the slide that you want to step back to and click on this button.";
        public const string ZoomToAreaButtonSupertip =
            "Zoom into an area of a slide or image.\n\nTo activate, place a rectangle shape on the portion that you want to magnify and click on this button.\n\n" +
            "This feature works best with high-resolution images.";
        public const string MoveCropShapeButtonSupertip =
            "Crop a picture according to your selected shape(s).\n\n" +
            "To activate, draw a shape upon the picture you want to crop, select the shape, and then click on this button.";
        public const string AddSpotlightButtonSupertip =
            "Creates a spotlight effect for a selected shape.\n\n" +
            "To activate, draw a shape to match the spotlight you want, select the shape, and click this button.";
        public const string ReloadSpotlightButtonSupertip =
            "Adjusts the transparency and edges of an existing spotlight.\n\n" +
            "To activate, set the transparency level and soft edges width, select the existing spotlight shape, and click this button.";
        public const string AddAudioButtonSupertip =
            "Generates synthesized audio from text in the Speaker Notes pane, and adds it to the slide(s).\n\n" +
            "To add audios for multiple slides, select the slides first (e.g. Ctrl+A to select all slides, Ctrl+click to select only some slides) before clicking the 'Add Audio' button.";
        public const string GenerateRecordButtonSupertip =
            "Generates synthesized audio from text in the Speaker Notes pane, and adds it to the slide(s).";
        public const string AddRecordButtonSupertip =
            "Use this option to record audio manually to replace the generated audio.";
        public const string RemoveAudioButtonSupertip =
            "Removes any generated audio that was added using Auto Narrate from the slide(s).\n\n" +
            "To remove audios for multiple slides, select the slides first (e.g. Ctrl+A to select all slides, Ctrl+click to select only some slides) before clicking the 'Remove Audio' button.";
        public const string AddCaptionsButtonSupertip =
            "Generates movie-style subtitles from text in the Speaker Notes pane, and adds it to the slide(s).\n\n" +
            "To add captions for multiple slides, select the slides first (e.g. Ctrl+A to select all slides, Ctrl+click to select only some slides) before clicking the 'Add Captions' button.";
        public const string RemoveCaptionsButtonSupertip =
            "Removes any captions that were added using Auto Captions from the slide(s).\n\n" +
            "To remove captions for multiple slides, select the slides first (e.g. Ctrl+A to select all slides, Ctrl+click to select only some slides) before clicking the 'Remove Captions' button.";
        public const string HighlightBulletsTextButtonSupertip =
            "Highlights all the bullet points on the specified slide by changing the text color.";
        public const string HighlightBulletsBackgroundButtonSupertip =
            "Highlights all the bullet points on the specified slide by highlighting the text background";
        public const string ColorPickerButtonSupertip = @"Opens Custom Color Picker";
        public const string CustomeShapeButtonSupertip = @"Manage your custom shapes on the fly.";
        public const string HelpButtonSupertip = @"Click this to visit PowerPointLabs help page in our website.";
        public const string FeedbackButtonSupertip = @"Click this to email us problem reports or other feedback. ";
        public const string AboutButtonSupertip = @"Information about the PowerPointLabs plugin.";
        # endregion

        # region Ribbon
        public const string AboutInfo =
            "          PowerPointLabs Plugin Version 1.8.10 [Release date: 18 Jul 2014]\n     Developed at School of Computing, National University of Singapore.\n        For more information, visit our website " + PowerPointLabsWebsiteUrl;
        public const string AboutInfoTitle = "About PowerPointLabs";
        # endregion

        # region ThisAddIn
        public const string AccessTempFolderErrorMsg = "Error when accessing temp folder";
        public const string CreatTempFolderErrorMsg = "Error when creating temp folder";
        public const string ExtraErrorMsg = "Error when extracting";
        public const string PrepareMediaErrorMsg = "Error when preparing media files";
        public const string VersionNotCompatibleMsg =
            "This file is not fully compatible with some features of PowerPointLabs because it is " +
            "in the outdated .ppt format used by PowerPoint 2007 (and older). If you wish to use the " +
            "full power of PowerPointLabs to enhance this file, please save in the .pptx format used " +
            "by PowerPoint 2010 and newer.";
        public const string TabActivateErrorTitle = "Unable to activate 'Double Click to Open Property' feature";
        public const string TabActivateErrorDescription =
            "To activate 'Double Click to Open Property' feature, you need to enable 'Home' tab " +
            "in Options -> Customize Ribbon -> Main Tabs -> tick the checkbox of 'Home' -> click OK but" +
            "ton to save.";
        public const string ShapesLabTaskPanelTitle = "Shapes Lab";
        public const string ColorsLabTaskPanelTitle = "Colors Lab";
        public const string RecManagementPanelTitle = "Record Management";
        # endregion 

        # region ConvertToPicture
        public const string ErrorTypeNotSupported = "Convert to Picture only supports Shapes and Charts.";
        public const string ErrorWindowTitle = "Unable to Convert to Picture";
        # endregion

        # region Task Pane - Recorder
        public const string RecorderInitialTimer = "00:00:00";
        public const string RecorderReadyStatusLabel = "Ready.";
        public const string RecorderRecordingStatusLabel = "Recording...";
        public const string RecorderPlayingStatusLabel = "Playing...";
        public const string RecorderPauseStatusLabel = "Pause";
        public const string RecorderUnrecognizeAudio = "Unrecognize Embedded Audio";
        public const string RecorderScriptStatusNoAudio = "No Audio";
        public const string RecorderWndMessageError = "Fatal error";
        public const string RecorderNoScriptDetail = "No Script Available";
        public const string RecorderNoInputDeviceMsg = "No Input Device suitable for the recording.\n" +
                                                       "Make sure your computer has a built-in voice picker and has been enabled, " +
                                                       "or an external voice input device has been connected.";
        public const string RecorderNoInputDeviceMsgBoxTitle = "Input Device Not Found";
        public const string RecorderSaveRecordMsg = "Do you want to save the record?";
        public const string RecorderSaveRecordMsgBoxTitle = "Save Record";
        public const string RecorderReplaceRecordMsgFormat = "Do you want to replace\n{0}\nwith current record?";
        public const string RecorderReplaceRecordMsgBoxTitle = "Replacement";
        public const string RecorderNoRecordToPlayError = "No record to play back. Please record first.";
        public const string RecorderInvalidOperation = "Invalid Operation";
        # endregion

        # region Task Pane - Custom Shape
        public const string CustomShapeFileNameInvalid = "Invalid shape name encountered";
        public const string CustomShapeNoShapeTextFirstLine = "No shapes saved yet.";
        public const string CustomShapeNoShapeTextSecondLine = "Right-click any object in the slides to save it in this panel.";
        public const string CustomShapeNoPanelSelectedError = "No shape selected";
        public const string CustomShapeViewTypeNotSupported = "Shapes Lab does not support current view type.";
        # endregion

        # region Control - Labeled Thumbnail 
        public const string LabeledThumbnailTooLongNameError = "Name's length cannot exceed 255";
        public const string LabeledThumbnailInvalidCharacterError = "Empty name, '<', '>', ':', '\"', '/', '\\', '|', '?', and '*' are not allowed for the name";
        public const string LabeledThumbnailFileNameExistError = "File name is already used";
        # endregion

        # region Error Dialog
        public const string UserFeedBack = " Help us fix the problem by emailing ";
        public const string Email = @"pptlabs@comp.nus.edu.sg";
        # endregion
    }
}
