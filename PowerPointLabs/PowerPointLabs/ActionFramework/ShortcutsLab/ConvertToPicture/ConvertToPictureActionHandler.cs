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
                ShapeRange childShapeRange = selection.ChildShapeRange;
                ShapeRange oldShapeRange = childShapeRange[GetLastShape(childShapeRange)].ParentGroup.Ungroup();
                ShapeRange shapeRange = childShapeRange.Duplicate();
                oldShapeRange.Regroup();

                selection.Unselect();
                shapeRange.Select();
                selection = this.GetCurrentSelection();
            }
            return selection;
        }

        private static int GetLastShape(ShapeRange shapeRange)
        {
            return shapeRange.Count > 1 ? shapeRange.Count - 1 : 1;
        }
    }
}
