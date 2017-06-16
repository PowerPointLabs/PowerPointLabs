using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.Models;
using PowerPointLabs.PasteLab;

namespace PowerPointLabs.ActionFramework.Action.PasteLab
{
    [ExportActionRibbonId(
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuShape,
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuLine,
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuFreeform,
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuPicture,
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuGroup,
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuInk,
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuVideo,
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuTextEdit,
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuChart,
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuTable,
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuTableCell,
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuSlide,
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuSmartArt,
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuEditSmartArt,
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuEditSmartArtText,
        TextCollection.PasteToFillSlideMenuId + TextCollection.RibbonButton)]
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
