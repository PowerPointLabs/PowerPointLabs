using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.Models;

namespace PowerPointLabs.ActionFramework.Action.PasteLab
{
    [ExportActionRibbonId(
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuShape,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuLine,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuFreeform,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuPicture,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuGroup,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuInk,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuVideo,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuTextEdit,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuChart,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuTable,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuTableCell,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuSlide,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuSmartArt,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuEditSmartArt,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuEditSmartArtText,
        TextCollection.PasteAtOriginalPositionId + TextCollection.RibbonButton)]
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