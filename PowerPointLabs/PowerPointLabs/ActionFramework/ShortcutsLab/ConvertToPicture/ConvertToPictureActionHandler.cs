using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Models;
using PowerPointLabs.ShortcutsLab;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ShortcutsLab
{
    [ExportActionRibbonId(ShortcutsLabText.ConvertToPictureTag)]
    class ConvertToPictureActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();
            PowerPointPresentation pres = this.GetCurrentPresentation();
            PowerPointSlide slide = this.GetCurrentSlide();
            Selection selection = GetSelection();
            if (selection == null)
            {
                MessageBox.Show("The shape selection has changed.", "Please try again");
                return;
            }
            ConvertToPicture.Convert(pres, slide, selection);
        }

        private Selection GetSelection()
        {
            Selection selection = this.GetCurrentSelection();
            if (selection.HasChildShapeRange)
            {
                return DuplicateChildAsSelection(selection);
            }
            return selection;
        }

        private Selection DuplicateChildAsSelection(Selection selection)
        {
            ShapeRange shapeRange = null;
            int numTries = 0;

            while (shapeRange == null && numTries < 50)
            {
                shapeRange = RegroupChildShapeRange(selection);
                numTries++;
            }

            if (shapeRange == null)
            {
                return null;
            }

            selection.Unselect();
            shapeRange.Select();
            selection = this.GetCurrentSelection();
            return selection;
        }

        private static ShapeRange RegroupChildShapeRange(Selection selection)
        {
            ShapeRange childShapeRange = selection.ChildShapeRange;
            ShapeRange oldShapeRange = null;
            try
            {
                oldShapeRange = selection.ShapeRange.Ungroup();
                //childShapeRange[1].ParentGroup.Ungroup();
            }
            catch
            {
                // when an undo is performed and
                // selecting a different number of shapes in a group, will throw exception
                return null;
            }
            ShapeRange shapeRange = childShapeRange.Duplicate();
            if (oldShapeRange != null)
            {
                oldShapeRange.Regroup();
            }
            return shapeRange;
        }

        private static int GetRepresentativeShape(ShapeRange shapeRange)
        {
            return shapeRange.Count > 1 ? shapeRange.Count - 1 : 1;
        }

    }
}
