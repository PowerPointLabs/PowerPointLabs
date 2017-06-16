using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.Models;
using PowerPointLabs.PasteLab;

namespace PowerPointLabs.ActionFramework.Action.PasteLab
{
    [ExportActionRibbonId(
        TextCollection.PasteToFillSlideId + TextCollection.MenuShape,
        TextCollection.PasteToFillSlideId + TextCollection.MenuLine,
        TextCollection.PasteToFillSlideId + TextCollection.MenuFreeform,
        TextCollection.PasteToFillSlideId + TextCollection.MenuPicture,
        TextCollection.PasteToFillSlideId + TextCollection.MenuGroup,
        TextCollection.PasteToFillSlideId + TextCollection.MenuInk,
        TextCollection.PasteToFillSlideId + TextCollection.MenuVideo,
        TextCollection.PasteToFillSlideId + TextCollection.MenuTextEdit,
        TextCollection.PasteToFillSlideId + TextCollection.MenuChart,
        TextCollection.PasteToFillSlideId + TextCollection.MenuTable,
        TextCollection.PasteToFillSlideId + TextCollection.MenuTableCell,
        TextCollection.PasteToFillSlideId + TextCollection.MenuSlide,
        TextCollection.PasteToFillSlideId + TextCollection.MenuSmartArt,
        TextCollection.PasteToFillSlideId + TextCollection.MenuEditSmartArt,
        TextCollection.PasteToFillSlideId + TextCollection.MenuEditSmartArtText,
        TextCollection.PasteToFillSlideId + TextCollection.RibbonButton)]
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
