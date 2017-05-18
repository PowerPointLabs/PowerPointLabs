using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.Models;

namespace PowerPointLabs.ActionFramework.Action.PasteLab
{
    [ExportActionRibbonId("PasteToFillSlide")]
    class PasteToFillSlideActionHandler : PasteLabActionHandler
    {
        protected override void ExecutePasteAction(string ribbonId, PowerPointPresentation presentation, PowerPointSlide slide,
                                                    Selection selection, ShapeRange pastingShapes)
        {
            PowerPointLabs.PasteLab.PasteToFillSlide.Execute(slide, pastingShapes, presentation.SlideWidth, presentation.SlideHeight);
        }
    }
}
