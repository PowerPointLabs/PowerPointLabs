using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.Models;

namespace PowerPointLabs.ActionFramework.Action.PasteLab
{
    [ExportActionRibbonId(
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuShape,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuLine,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuFreeform,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuPicture,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuGroup,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuInk,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuVideo,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuTextEdit,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuChart,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuTable,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuTableCell,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuSlide,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuSmartArt,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuEditSmartArt,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuEditSmartArtText,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.RibbonButton)]
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