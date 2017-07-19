﻿namespace PowerPointLabs
{
    public class TextCollection
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
        # region Animation Lab
        public const string AnimationLabMenuSupertip =
            "Use Animation Lab to add animations your slides easily.";
        public const string AddAnimationButtonSupertip =
            "Create an animation slide to transition from the currently selected slide to the next slide.\n\n" +
            "To perform this action, duplicate the currently selected slide, move the objects to the desired position, select the original slide, then click this button.";
        public const string InSlideAnimateButtonSupertip =
            "Moves a shape around the slide in multiple steps.\n\n" +
            "To perform this action, copy the shape to locations where it should stop, select the copies in the order they should appear, then click this button.";
        public const string AnimationLabSettingsSupertip =
            "Configure the settings for Animation Lab.";
        # endregion

        # region Zoom Lab
        public const string ZoomLabMenuSupertip =
            "Use Zoom Lab to creating zoom in and out effects for your slides easily.";
        public const string AddZoomInButtonSupertip =
            "Create an animation slide with a zoom-in effect from the currently selected shape to the next slide.\n\n" +
            "To perform this action, select a rectangle shape on the slide to drill down from, then click this button.";
        public const string AddZoomOutButtonSupertip =
            "Create an animation slide with a zoom-out effect from the previous slide to the currently selected shape.\n\n" +
            "To perform this action, select a rectangle shape on the slide to step back to, then click this button.";
        public const string ZoomToAreaButtonSupertip =
            "Zoom into an area of a slide or picture.\n\n" +
            "To perform this action, place a rectangle shape on the portion to magnify, then click this button.\n\n" +
            "This feature works best with high-resolution images.";
        public const string ZoomLabSettingsSupertip =
            "Configure the settings for Zoom Lab.";
        # endregion

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

        #region Paste Lab
        public const string PasteLabMenuSupertip =
            "Use Paste Lab to customize how you paste your copied objects.";
        public const string PasteToFillSlideSupertip =
            "Paste your copied objects to fill the current slide.\n\n" +
            "To perform this action, with object(s) copied, click this button.";
        public const string PasteAtOriginalPositionSupertip =
            "Paste your copied objects at their original positions when they were copied.\n\n" +
            "To perform this action, with object(s) copied, click this button.";
        public const string ReplaceWithClipboardSupertip =
            "Paste your copied objects over your selection while preserving its animations.\n\n" +
            "To perform this action, select an object on the slide, then with object(s) copied, click this button.";
        public const string PasteIntoGroupSupertip =
            "Paste your copied objects into an existing group.\n\n" +
            "To perform this action, select a group on the slide, then with object(s) copied, click this button.";
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
        public const string EffectsLabColorizeRemainderSupertip =
            "Recolor an area of a slide to attract attention to it.\n\n" +
            "To perform this action, select the shape(s) over the area to keep, then click this button.";
        public const string EffectsLabColorizeBackgroundSupertip =
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

        # region Drawing Lab
        public const string DrawingsLabButtonSupertip = "Open the Drawing Lab Interface";
        #endregion

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
        # endregion

        # region Button Labels
        public const string CombineShapesLabel = "Combine Shapes";

        # region Animation Lab
        public const string AnimationLabMenuLabel = "Animation";
        public const string AddAnimationButtonLabel = "Add Animation Slide";
        public const string AddAnimationReloadButtonLabel = "Recreate Animation";
        public const string AddAnimationInSlideAnimateButtonLabel = "Animate In Slide";
        public const string AnimationLabSettingsButtonLabel = "Settings";
        # endregion

        # region Zoom Lab
        public const string ZoomLabMenuLabel = "Zoom";
        public const string AddZoomInButtonLabel = "Drill Down";
        public const string AddZoomOutButtonLabel = "Step Back";
        public const string ZoomToAreaButtonLabel = "Zoom To Area";
        public const string ZoomLabSettingsButtonLabel = "Settings";
        # endregion

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

        # region Drawing Lab
        public const string DrawingsLabButtonLabel = "Drawing";
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
        public const string EffectsLabBlurrinessFeatureSelected = "EffectsLabBlurSelected";
        public const string EffectsLabBlurrinessFeatureRemainder = "EffectsLabBlurRemainder";
        public const string EffectsLabBlurrinessFeatureBackground = "EffectsLabBlurBackground";

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

        public const string AnimationLabMenuId = "AnimationLabMenu";
        public const string AddAnimationSlideTag = "AddAnimationSlide";
        public const string AnimateInSlideTag = "AnimateInSlide";
        public const string AnimationLabSettingsTag = "AnimationLabSettings";

        public const string ZoomLabMenuId = "ZoomLabMenu";
        public const string DrillDownTag = "DrillDown";
        public const string StepBackTag = "StepBack";
        public const string ZoomToAreaTag = "ZoomToArea";
        public const string ZoomLabSettingsTag = "ZoomLabSettings";

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

        public const string HighlightLabMenuId = "HighlightLabMenu";
        public const string HighlightPointsTag = "HighlightPoints";
        public const string HighlightBackgroundTag = "HighlightBackground";
        public const string HighlightTextTag = "HighlightText";
        public const string RemoveHighlightTag = "RemoveHighlight";
        public const string HighlightLabSettingsTag = "HighlightLabSettings";

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
        public const string HideSelectedShapeTag = "HideShape";
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
        public const string DrawingsLabTaskPanelTitle = "Drawing Lab";
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

        # region CropToShape

        public class CropToShapeText
        {
            //------------ Msg -------------
            public const string ErrorMessageForSelectionCountZero = "'Crop To Shape' requires at least one shape to be selected.";
            public const string ErrorMessageForSelectionNonShape = "'Crop To Shape' only supports shape objects.";
            public const string ErrorMessageForUndefined = "Undefined error in 'Crop To Shape'.";
        }

        #endregion

        #region CropToSlide

        public class CropToSlideText
        {
            //------------ Msg -------------
            public const string ErrorMessageForSelectionCountZero = "'Crop To Slide' requires at least one shape to be selected.";
            public const string ErrorMessageForSelectionNonPicture = "'Crop To Slide' only supports picture objects.";
            public const string ErrorMessageForUndefined = "Undefined error in 'Crop To Slide'.";
        }

        #endregion

        #region CropLab

        public class CropLabText
        {
            public const string ErrorSelectionIsInvalid = "You need to select at least {1} {2} before applying '{0}'.";
            public const string ErrorSelectionMustBeShape = "'{0}' only supports shape objects.";
            public const string ErrorSelectionMustBePicture = "'{0}' only supports picture objects.";
            public const string ErrorSelectionMustBeShapeOrPicture = "'{0}' only supports shape or picture objects.";
            public const string ErrorAspectRatioIsInvalid = "The given aspect ratio is invalid. Please enter positive numbers for the width to height ratio.";
            public const string ErrorUndefined = "'Undefined error in Crop Lab'.";
            public const string ErrorMessageNoShapeOverBoundary = "All selected objects are inside the slide boundary. No cropping was done.";
            public const string ErrorMessageNoDimensionCropped = "All selected pictures are smaller than reference shape. No cropping was done.";
            public const string ErrorMessageNoPaddingCropped = "All selected pictures have no transparent padding. No cropping was done.";
            public const string ErrorMessageNoAspectRatioCropped = "All selected pictures are already in the given aspect ratio. No cropping was done.";
        }

        #endregion

        #region ConvertToPicture
        public const string ErrorTypeNotSupported = "Convert to Picture only supports Shapes and Charts.";
        public const string ErrorWindowTitle = "Convert to Picture: Unsupported Object";
        # endregion

        #region PictureSlidesLab
        public class PictureSlidesLabText
        {
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
        #endregion

        #region Animation Lab
        // Errors
        public const string AnimationLabAutoAnimateErrorDialogTitle = "Unable to perform action";
        public const string AnimationLabAutoAnimateErrorWrongSlide = "Please select the correct slide.";
        public const string AnimationLabAutoAnimateErrorNoMatchingShapes = "No matching Shapes were found on the next slide.";
        public const string AnimationLabAutoAnimateErrorSlideNotAutoAnimate = "The current slide was not added by PowerPointLabs Auto Animate";

        // Dialog Boxes
        public const string AnimationLabAutoAnimateLoadingText = "Applying auto animation...";
        public const string AnimationLabSettingsDurationInputTooltip = "The duration (in seconds) for the animations in the animation slides to be created.";
        public const string AnimationLabSettingsSmoothAnimationCheckboxTooltip = 
            "Use a frame-based approach for smoother resize animations.\n" +
            "This may result in larger file sizes and slower loading times for animated slides.";
        #endregion

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
        public const string AgendaLabSectionNameTooLongError = "One of the section names exceeds the maximum size allowed by Agenda Lab. Please rename the section accordingly.";

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
        // Errors
        public const string EffectsLabBlurSelectedErrorNoSelection = "'Blur Selected' requires at least one shape or text box to be selected.";
        public const string EffectsLabBlurSelectedErrorNonShapeOrTextBox = "'Blur Selected' only supports shape and text box objects.";

        // Dialog Boxes
        public const string EffectsLabSettingsTintCheckboxForTintSelected = "Tint Selected";
        public const string EffectsLabSettingsTintCheckboxForTintRemainder = "Tint Remainder";
        public const string EffectsLabSettingsTintCheckboxForTintBackground = "Tint All Except Selected";
        public const string EffectsLabSettingsTintCheckboxTooltip = "Adds a tinted effect to your blur.";
        public const string EffectsLabSettingsBlurrinessInputTooltip = "The level of blurriness.";
        public const string SpotlightSettingsTransparencyInputTooltip = "The transparency level of the spotlight effect to be created.";
        public const string SpotlightSettingsSoftEdgesSelectionInputTooltip = "The softness of the edges of the spotlight effect to be created.";
        #endregion

        #region Captions Lab

        public const string CaptionsLabErrorDialogTitle = "Unable to perform action";
        public const string CaptionsLabErrorNoSelection = "Select at least one slide to apply captions.";
        public const string CaptionsLabErrorNoNotes = "Captions could not be created because there are no notes entered. Please enter something in the notes and try again.";
        public const string CaptionsLabErrorNoSelectionLog = "No slide in selection";
        public const string CaptionsLabErrorNoCurrentSlideLog = "No current slide";
        public const string CaptionsLabErrorNoNotesLog = "No notes on slide";

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
            "The folder cannot be migrated entirely. Please check if your destination location forbids this action.";
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

        #region Narrations Lab
        // Dialog Boxes
        public const string NarrationsLabSettingsVoiceSelectionInputTooltip = 
            "The voice to be used when generating synthesized audio.\n" +
            "Use [Voice] tags to specify a different voice for a particular section of text.";
        public const string NarrationsLabSettingsPreviewCheckboxTooltip =
            "If checked, the current slide's audio and animations will play after the Add Audio button is clicked.";
        #endregion

        #region Positions Lab
        public class PositionsLabText
        {
            public const string ErrorNoSelection = "Please select at least a shape before using this feature";
            public const string ErrorFewerThanTwoSelection = "Please select at least two shapes before using this feature";
            public const string ErrorFewerThanThreeSelection = "Please select at least three shapes before using this feature";
            public const string ErrorFewerThanFourSelection = "Please select at least four shapes before using this feature";
            public const string ErrorFunctionNotSupportedForWithinShapes = "This function is not supported for Within Corner Most Objects Setting.";
            public const string ErrorFunctionNotSupportedForSlide = "This function is not supported for Within Slide Setting.";
            public const string ErrorFunctionNotSupportedForOverlapRefShapeCenter = "This function is not supported for shapes that overlap the center of the reference shape.";
            public const string ErrorUndefined = "'Undefined error in Resize Lab'";
        }
        #endregion

        #region Paste Lab
        public class PasteLabText
        {
            public const string PasteLabMenu = "Paste";
            public const string PasteToFillSlide = "Paste To Fill Slide";
            public const string ReplaceWithClipboard = "Replace With Clipboard";
            public const string PasteIntoGroup = "Paste Into Group";
            public const string PasteAtCursorPosition = "Paste At Cursor Position";
            public const string PasteAtOriginalPosition = "Paste At Original Position";
        }
        #endregion

        #region Sync Lab
        public const string SyncLabErrorDialogTitle = "Unable to execute action";
        public const string SyncLabCopySelectError = "Please select one shape to copy.";
        public const string SyncLabPasteSelectError = "Please select at least one item to apply.";
        public const string SyncLabShapeDeletedError = "Error in loading shape formats. Removing invalid formats from the list.";
        public const string SyncLabCopyError = "Error: Unable to copy selected item.";
        public const string SyncLabStorageFileName = "Sync Lab - Do not edit";
        public const string SyncLabDefaultFormatName = "Format";
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
        public const string Email = @"pptlabs@comp.nus.edu.sg";
        # endregion

        # region Install and Update related

        public const string QuickTutorialFileName = "Tutorial.pptx";
        public const string VstoName = "PowerPointLabsInstaller.vsto";
        public const string InstallerName = "data.zip";

        # endregion
    }
}
