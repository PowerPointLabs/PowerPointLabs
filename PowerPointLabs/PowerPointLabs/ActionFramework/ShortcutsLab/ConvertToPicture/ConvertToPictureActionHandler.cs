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
            ShapeRange shapeRange = RegroupChildShapeRange(selection);

            selection.Unselect();
            shapeRange.Select();
            selection = this.GetCurrentSelection();
            return selection;
        }

        private static ShapeRange RegroupChildShapeRange(Selection selection)
        {
            ShapeRange childShapeRange = selection.ChildShapeRange;
            ShapeRange oldShapeRange = selection.ShapeRange.Ungroup();

            ShapeRange shapeRange = childShapeRange.Duplicate();
            return shapeRange;
        }

        private static int GetRepresentativeShape(ShapeRange shapeRange)
        {
            return shapeRange.Count > 1 ? shapeRange.Count - 1 : 1;
        }

    }
}
