using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.PasteLab
{
    [ExportActionRibbonId(PasteLabText.PasteAtOriginalPositionTag)]
    class PasteAtOriginalPositionActionHandler : PasteLabActionHandler
    {
        protected override ShapeRange ExecutePasteAction(string ribbonId, PowerPointPresentation presentation, PowerPointSlide slide,
                                                        ShapeRange selectedShapes, ShapeRange selectedChildShapes)
        {
            PowerPointSlide tempSlide = presentation.AddSlide(index: slide.Index);
            ShapeRange tempPastingShapes = PasteShapesFromClipboard(tempSlide);
            if (tempPastingShapes == null)
            {
                tempSlide.Delete();
                return PasteShapesFromClipboard(slide);
            }

            ShapeRange pastingShapes = slide.CopyShapesToSlide(tempPastingShapes);
            tempSlide.Delete();

            return pastingShapes;
        }
    }
}