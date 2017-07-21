namespace PowerPointLabs.TextCollection
{
    internal static class ShapesLabText
    {
        public const string RibbonMenuLabel = "Shapes";
        public const string RibbonMenuSupertip =
           "Use Shapes Lab to manage your custom shapes.\n\n" +
           "Click this button to open the Shapes Lab interface.";

        public const string FileNameInvalid = "Invalid shape name.";
        public const string NoShapeTextFirstLine = "No shapes saved yet.";
        public const string NoShapeTextSecondLine = "Right-click any object on a slide to save it in this panel.";
        public const string NoPanelSelectedError = "No shape selected.";
        public const string ViewTypeNotSupported = "Shapes Lab does not support the current view type.";
        public const string SaveLocationChangedSuccessFormat =
            "Default saving path has been changed to \n{0}\nAll shapes have been moved to the new location.";
        public const string SetAsDefaultCategorySuccessFormat = "{0} has been set as default category.";
        public const string SaveLocationChangedSuccessTitle = "Success";
        public const string MigrationError =
            "The folder cannot be migrated entirely. Please check if your destination location forbids this action.";
        public const string OriginalFolderDeletionError =
            "The original folder could not be deleted because some of the files in folder is still in use. You could " +
            "try to delete this folder manually when those files are closed.";
        public const string MigratingDialogTitle = "Migrating...";
        public const string MigratingDialogContent = "Shapes are being migrated, please wait...";
        public const string RemoveLastCategoryError = "Removing the last category is not allowed.";
        public const string DuplicateCategoryNameError = "The name has already been used.";
        public const string RemoveDefaultCategoryMessage =
            "You are removing your default category. After removing this category, the first category will be made " +
            "as default category. Continue?";
        public const string RemoveDefaultCategoryCaption = "Removing Default Category";
        public const string ImportFileError = "Import File could not be opened.";
        public const string ImportNoSlideError = "Import File is empty.";
        public const string ImportAppendCategoryError = "Your computer does not support this feature.";
        public const string ImportSingleCategoryErrorFormat =
            "{0} contains multiple categories. Try \"Import Category\" instead.";
        public const string ImportSuccess = "Successfully imported";

        public const string ImportShapeFileDialogTitle = "Import Shapes";
        public const string ImportLibraryFileDialogTitle = "Import Library";

        public const string ShapeContextStripAddToSlide = "Add To Slide";
        public const string ShapeContextStripEditName = "Edit Name";
        public const string ShapeContextStripMoveShape = "Move Shape To";
        public const string ShapeContextStripRemoveShape = "Remove Shape";
        public const string ShapeContextStripCopyShape = "Copy Shape To";

        public const string CategoryContextStripAddCategory = "Add Category";
        public const string CategoryContextStripRemoveCategory = "Remove Category";
        public const string CategoryContextStripRenameCategory = "Rename Category";
        public const string CategoryContextStripImportCategory = "Import Library";
        public const string CategoryContextStripImportShapes = "Import Shapes";
        public const string CategoryContextStripSetAsDefaultCategory = "Set as Default Category";
        public const string CategoryContextStripCategorySettings = "Shapes Lab Settings";
    }
}
