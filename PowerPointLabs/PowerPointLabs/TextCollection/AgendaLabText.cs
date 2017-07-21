namespace PowerPointLabs.TextCollection
{
    internal static class AgendaLabText
    {
        public const string ErrorDialogTitle = "Unable to execute action";
        public const string NoSectionError = "Please group the slides into sections before generating agenda.";
        public const string SingleSectionError = "Please divide the slides into two or more sections.";
        public const string EmptySectionError = "Presentation contains empty section(s). Please fill them up or remove them.";
        public const string AgendaExistError = "Agenda already exists. The previous agenda will be removed and regenerated. Do you want to proceed?";
        public const string AgendaExistErrorCaption = "Confirm Update";
        public const string NoAgendaError = "There is no generated agenda.";
        public const string NoReferenceSlideError = "No reference slide could be found. Either replace the reference slide or regenerate the agenda.";
        public const string InvalidReferenceSlideError = "The current reference slide is invalid. Either replace the reference slide or regenerate the agenda.";
        public const string SectionNameTooLongError = "One of the section names exceeds the maximum size allowed by Agenda Lab. Please rename the section accordingly.";

        // Dialog Boxes
        public const string GeneratingDialogTitle = "Generating...";
        public const string GeneratingDialogContent = "Agenda is generating, please wait...";
        public const string SynchronizingDialogTitle = "Synchronizing...";
        public const string SynchronizingDialogContent = "Agenda is being synchronized, please wait...";

        public const string ReorganiseSidebarTitle = "Reorganise Sidebar";
        public const string ReorganiseSidebarContent = "The sections have been changed. Do you wish to reorganise the items in the sidebar?";

        public const string BeamGenerateSingleSlideDialogTitle = "Generate on all slides";
        public const string BeamGenerateSingleSlideDialogContent = "Only one slide is selected. Would you like to generate the sidebar on all slides instead?";

        // Agenda Content
        public const string TitleContent = "Agenda";

        public const string BulletVisitedContent = "Visited bullet format";
        public const string BulletHighlightedContent = "Highlighted bullet format";
        public const string BulletUnvisitedContent = "Unvisited bullet format";
        public const string BeamHighlightedText = "Highlighted";

        public const string TemplateSlideInstructions =
                            "This slide is used as a ‘Template' for generating agenda slides. Please do not delete this slide.\r" +
                            "Adjust the design of this slide and click the 'Sync Agenda' (in Agenda Lab) to replicate the design in the other slides.";
    }
}
