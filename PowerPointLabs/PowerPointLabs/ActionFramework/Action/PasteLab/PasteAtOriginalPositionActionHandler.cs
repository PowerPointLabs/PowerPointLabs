using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.Models;
using PowerPointLabs.PasteLab;

namespace PowerPointLabs.ActionFramework.Action.PasteLab
{
    [ExportActionRibbonId(
        "PasteAtOriginalPositionMenuFrame",
        "PasteAtOriginalPositionMenuShape",
        "PasteAtOriginalPositionMenuLine",
        "PasteAtOriginalPositionMenuFreeform",
        "PasteAtOriginalPositionMenuPicture",
        "PasteAtOriginalPositionMenuGroup",
        "PasteAtOriginalPositionMenuInk",
        "PasteAtOriginalPositionMenuVideo",
        "PasteAtOriginalPositionMenuTextEdit",
        "PasteAtOriginalPositionMenuChart",
        "PasteAtOriginalPositionMenuTable",
        "PasteAtOriginalPositionMenuTableWhole",
        "PasteAtOriginalPositionMenuSmartArtBackground",
        "PasteAtOriginalPositionMenuSmartArtEditSmartArt",
        "PasteAtOriginalPositionMenuSmartArtEditText")]
    class PasteAtOriginalPositionActionHandler : PasteLabActionHandler
    {
        protected override ShapeRange ExecutePasteAction(string ribbonId, PowerPointPresentation presentation, PowerPointSlide slide,
                                                        ShapeRange selectedShapes, ShapeRange selectedChildShapes)
        {
            PowerPointSlide tempSlide = presentation.AddSlide(index: slide.Index);
            ShapeRange tempPastingShapes = tempSlide.Shapes.Paste();
            ShapeRange pastingShapes = slide.CopyShapesToSlide(tempPastingShapes);
            tempSlide.Delete();
            return pastingShapes;
        }
    }
}