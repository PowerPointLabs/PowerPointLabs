using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.Models;

namespace PowerPointLabs.ActionFramework.Action.PasteLab
{
    [ExportActionRibbonId("PasteAtOriginalPosition")]
    class PasteAtOriginalPositionActionHandler : PasteLabActionHandler
    {
        ShapeRange pastedShapes = null;
        
        protected override void ExecutePasteAction(string ribbonId, PowerPointPresentation presentation, PowerPointSlide slide,
                                                    Selection selection, ShapeRange pastingShapes)
        {
            pastingShapes.Delete();

            PowerPointSlide tempSlide = presentation.AddSlide(index: slide.Index);
            pastingShapes = tempSlide.Shapes.Paste();
            pastedShapes = slide.CopyShapesToSlide(pastingShapes);
            tempSlide.Delete();
        }

        protected override void CleanupPasteAction()
        {
            pastedShapes.Copy();
        }
    }
}