using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.Models;
using PowerPointLabs.PasteLab;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.PasteLab
{
    [ExportActionRibbonId(PasteLabText.PasteToFillSlideTag)]
    class PasteToFillSlideActionHandler : PasteLabActionHandler
    {
        protected override ShapeRange ExecutePasteAction(string ribbonId, PowerPointPresentation presentation, PowerPointSlide slide,
                                                        ShapeRange selectedShapes, ShapeRange selectedChildShapes)
        {
            ShapeRange pastingShapes = PasteShapesFromClipboard(slide);
            if (pastingShapes == null)
            {
                return null;
            }

            PasteToFillSlide.Execute(slide, pastingShapes, presentation.SlideWidth, presentation.SlideHeight);
            return null;
        }
    }
}
