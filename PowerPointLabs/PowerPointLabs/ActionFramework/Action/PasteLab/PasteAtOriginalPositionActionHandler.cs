using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.Models;
using PowerPointLabs.PasteLab;

namespace PowerPointLabs.ActionFramework.Action.PasteLab
{
    [ExportActionRibbonId(
        "PasteAtOriginalPositionMenuShape", "PasteAtOriginalPositionMenuLine", "PasteAtOriginalPositionMenuFreeform",
        "PasteAtOriginalPositionMenuPicture", "PasteAtOriginalPositionMenuGroup", "PasteAtOriginalPositionMenuInk",
        "PasteAtOriginalPositionMenuVideo", "PasteAtOriginalPositionMenuTextEdit", "PasteAtOriginalPositionMenuChart",
        "PasteAtOriginalPositionMenuTable", "PasteAtOriginalPositionMenuTableWhole", "PasteAtOriginalPositionMenuFrame",
        "PasteAtOriginalPositionMenuSmartArtBackground", "PasteAtOriginalPositionMenuSmartArtEditSmartArt",
        "PasteAtOriginalPositionMenuSmartArtEditText", "PasteAtOriginalPositionButton")]
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
                return null;
            }

            ShapeRange pastingShapes = slide.CopyShapesToSlide(tempPastingShapes);
            tempSlide.Delete();

            return pastingShapes;
        }
    }
}