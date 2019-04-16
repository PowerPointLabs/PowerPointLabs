using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ShapesLab
{
    class ShapesLabUtils
    {
        public static void SyncShapeAdd(ThisAddIn addIn, string shapeName, string shapeFullName, string category)
        {
            if (addIn == null)
            {
                return;
            }
            foreach (PowerPoint.DocumentWindow window in Globals.ThisAddIn.Application.Windows)
            {
                if (window == addIn.Application.ActiveWindow)
                {
                    continue;
                }
                CustomShapePane shapePaneControl = (addIn.GetControlFromWindow(typeof(CustomShapePane), window)
                    as CustomShapePane);
                if (shapePaneControl?.CurrentCategory == category)
                {
                    shapePaneControl.AddCustomShape(shapeName, shapeFullName, false);
                }
            }
        }

        public static void SyncShapeRemove(ThisAddIn addIn, string shapeName, string category)
        {
            if (addIn == null)
            {
                return;
            }
            foreach (PowerPoint.DocumentWindow window in Globals.ThisAddIn.Application.Windows)
            {
                if (window == addIn.Application.ActiveWindow)
                {
                    continue;
                }
                CustomShapePane shapePaneControl = (addIn.GetControlFromWindow(typeof(CustomShapePane), window)
                    as CustomShapePane);
                if (shapePaneControl?.CurrentCategory == category)
                {
                    shapePaneControl.RemoveCustomShape(shapeName);
                }
            }
        }

        public static void SyncShapeRename(ThisAddIn addIn, string shapeOldName, string shapeNewName, string category)
        {
            if (addIn == null)
            {
                return;
            }
            foreach (PowerPoint.DocumentWindow window in Globals.ThisAddIn.Application.Windows)
            {
                if (window == addIn.Application.ActiveWindow)
                {
                    continue;
                }
                CustomShapePane shapePaneControl = (addIn.GetControlFromWindow(typeof(CustomShapePane), window)
                    as CustomShapePane);
                if (shapePaneControl?.CurrentCategory == category)
                {
                    shapePaneControl.RenameCustomShape(shapeOldName, shapeNewName);
                }
            }
        }

        public static void SyncAddCategory(ThisAddIn addIn, string newCategoryName)
        {
            if (addIn == null)
            {
                return;
            }
            foreach (PowerPoint.DocumentWindow window in Globals.ThisAddIn.Application.Windows)
            {
                if (window == addIn.Application.ActiveWindow)
                {
                    continue;
                }
                CustomShapePane shapePaneControl = (addIn.GetControlFromWindow(typeof(CustomShapePane), window)
                    as CustomShapePane);
                shapePaneControl?.AddCategory(newCategoryName);
            }
        }

        public static void SyncRemoveCategory(ThisAddIn addIn, int removedCategoryIndex)
        {
            if (addIn == null)
            {
                return;
            }
            foreach (PowerPoint.DocumentWindow window in Globals.ThisAddIn.Application.Windows)
            {
                if (window == addIn.Application.ActiveWindow)
                {
                    continue;
                }
                CustomShapePane shapePaneControl = (addIn.GetControlFromWindow(typeof(CustomShapePane), window)
                    as CustomShapePane);
                shapePaneControl?.RemoveCategory(removedCategoryIndex);
            }
        }

        public static void SyncRenameCategory(ThisAddIn addIn, int renameCategoryIndex, string newCategoryName)
        {
            if (addIn == null)
            {
                return;
            }
            foreach (PowerPoint.DocumentWindow window in Globals.ThisAddIn.Application.Windows)
            {
                if (window == addIn.Application.ActiveWindow)
                {
                    continue;
                }
                CustomShapePane shapePaneControl = (addIn.GetControlFromWindow(typeof(CustomShapePane), window)
                    as CustomShapePane);
                shapePaneControl?.RenameCategory(renameCategoryIndex, newCategoryName);
            }
        }
    }
}
