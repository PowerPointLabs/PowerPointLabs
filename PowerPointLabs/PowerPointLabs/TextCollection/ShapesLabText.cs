namespace PowerPointLabs.TextCollection
{
    internal static class ShapesLabText
    {
        #region Action Framework Tags
        public const string RibbonMenuId = "ShapesLabMenu";
        public const string RibbonMenuSupertip = "Use Shapes Lab to manage your custom shapes.";
        public const string RibbonMenuLabel = "Shapes";

        public const string ShapesLabButtonLabel = "Shapes Lab";
        public const string ShapesLabButtonSupertip = "Click this button to open the Shapes Lab interface.";
        #endregion

        public const string PaneTag = "ShapesLab";
        public const string TaskPanelTitle = "Shapes Lab";

        public const string ErrorDialogTitle = "Unable to execute action";
        public const string ErrorAddSelectionInvalid = "Please select one shape to add.";

        public const string ErrorFileNameInvalid = "Invalid shape name.";
        public const string ErrorNoShapeTextFirstLine = "No shapes saved yet.";
        public const string ErrorNoShapeTextSecondLine = "Right-click any object on a slide to save it in this panel.";
        public const string ErrorNoPanelSelected = "No shape selected.";
        public const string ErrorViewTypeNotSupported = "Shapes Lab does not support the current view type.";
        public const string SuccessSaveLocationChanged =
            "Default saving path has been changed to \n{0}\nAll shapes have been moved to the new location.";
        public const string SuccessSaveLocationChangedTitle = "Success";
        public const string SuccessSetAsDefaultCategory = "{0} has been set as default category.";
        public const string ErrorMigration =
            "The folder cannot be migrated entirely. Please check if your destination location forbids this action.";
        public const string ErrorOriginalFolderDeletion =
            "The original folder could not be deleted because some of the files in folder is still in use. You could " +
            "try to delete this folder manually when those files are closed.";
        public const string MigratingDialogTitle = "Migrating...";
        public const string MigratingDialogContent = "Shapes are being migrated, please wait...";
        public const string ErrorRemoveLastCategory = "Removing the last category is not allowed.";
        public const string ErrorDuplicateCategoryName = "The name has already been used.";
        public const string RemoveDefaultCategoryMessage =
            "You are removing your default category. After removing this category, the first category will be made " +
            "as default category. Continue?";
        public const string RemoveDefaultCategoryCaption = "Removing Default Category";
        public const string ErrorImportFile = "Import File could not be opened.";
        public const string ErrorImportNoSlide = "Import File is empty.";
        public const string ErrorImportAppendCategory = "Your computer does not support this feature.";
        public const string ErrorImportSingleCategory =
            "{0} contains multiple categories. Try \"Import Category\" instead.";
        public const string SuccessImport = "Successfully imported.";

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

        public const string ErrorSameShapeNameInDestination = "{0} exists in {1}. Please rename your shape before moving.";
        public const string ErrorShapeCorrupted = "Some shapes in the Shapes Lab were corrupted, but some of the them are recovered.";

        public const string FolderDialogDescription = "Select the directory that you want to use as the default.";
        public const string ErrorFolderNonEmpty = "Please select an empty folder as default saving folder.";

        public const string AddShapeToolTip = "Adds a shape to Shapes Lab.";
        public const string DisabledAddShapeToolTip = AddShapeToolTip + "\nStart by selecting a shape.";
    }
}
