namespace PowerPointLabs.TextCollection
{
    public class CommonText
    {
        # region Common Error
        public const string ErrorTitle = "Error";
        public const string ErrorSlideSelectionTitle = "Slide selection error";
        public const string ErrorDuringSetup = "Error During Setup.";

        public const string ErrorNameTooLong = "The name's length cannot be more than 255 characters.";
        public const string ErrorInvalidCharacter = "The name cannot be empty, contain the following characters:'<', '>', ':', '\"', '/', '\\', '|', '?', or '*', or be a Windows reserved file name: CON, PRN, AUX, NUL, COM1, COM2, COM3, COM4, COM5, COM6, COM7, COM8, COM9, LPT1, LPT2, LPT3, LPT4, LPT5, LPT6, LPT7, LPT8, or LPT9.";
        public const string ErrorFileNameExist = "A file already exists with that name.";   
        # endregion

        #region URLs
        public const string FeedbackUrl = "http://www.comp.nus.edu.sg/~pptlabs/contact.html";
        public const string HelpDocumentUrl = "http://www.comp.nus.edu.sg/~pptlabs/docs/";
        public const string PowerPointLabsWebsiteUrl = "http://PowerPointLabs.info";
        public const string SingleShapeDownloadUrl = "http://www.comp.nus.edu.sg/~pptlabs/gallery.html";
        # endregion

        # region Tab Labels
        public const string PowerPointLabsAddInsTabLabel = "PowerPointLabs";
        public const string PowerPointLabsMenuLabel = "PowerPointLabs";
        public const string CombineShapesLabel = "Combine Shapes";
        #endregion

        #region Ribbon Groups
        public const string AnimationsGroupLabel = "Animations";
        public const string AudioGroupLabel = "Audio";
        public const string EffectsGroupLabel = "Effects";
        public const string FormattingGroupLabel = "Formatting";
        public const string MoreLabsGroupLabel = "More Labs";

        public const string RibbonMenu = "Menu";
        public const string AnimationsGroupId = "AnimationsGroup";
        public const string AudioGroupId = "AudioGroup";
        public const string EffectsGroupId = "EffectsGroup";
        public const string FormattingGroupId = "FormattingGroup";
        public const string MoreLabsGroupId = "MoreLabsGroup";
        #endregion

        #region Dynamic Menu Labels
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
        # endregion

        #region PowerPointSlide
        public const string NotesPageStorageText = "This notes page is used to store data - Do not edit the notes. ";
        # endregion

        # region ThisAddIn Error Messages
        public const string ErrorAccessTempFolder = "Error when accessing temp folder";
        public const string ErrorCreateTempFolder = "Error when creating temp folder";
        public const string ErrorExtract = "Error when extracting";
        public const string ErrorPrepareMedia = "Error when preparing media files";
        public const string ErrorVersionNotCompatible =
            "Some features of PowerPointLabs do not work with presentations saved in " +
            "the .ppt format. To use them, please resave the " +
            "presentation with the .pptx format.";
        public const string ErrorOnlinePresentationNotCompatible =
            "Some features of PowerPointLabs do not work with online presentations. " +
            "To use them, please save the file locally.";
        public const string ErrorShapeGalleryInit =
            "Could not connect to shape database from your default location.\n\n" +
            "To check your default location, right click on the Shapes Lab's panel and select 'Settings' option.";
        public const string ErrorTabActivateTitle = "Unable to activate 'Double Click to Open Property' feature";
        public const string ErrorTabActivate =
            "To activate 'Double Click to Open Property' feature, you need to enable 'Home' tab " +
            "in Options -> Customize Ribbon -> Main Tabs -> tick the checkbox of 'Home' -> click OK but" +
            "ton to save.";
        #endregion

        #region Graphics
        public const string TemporaryImageStorageFileName = "temp.png";
        public const string TemporaryCompressedImageStorageFileName = "temp.jpeg";
        #endregion

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
