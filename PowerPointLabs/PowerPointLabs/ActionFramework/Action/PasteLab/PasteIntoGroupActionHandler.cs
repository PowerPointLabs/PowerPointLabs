using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using PowerPointLabs.PasteLab;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.Action.PasteLab
{
    [ExportActionRibbonId("PasteIntoGroupMenuGroup", "PasteIntoGroupButton")]
    class PasteIntoGroupActionHandler : PasteLabActionHandler
    {
        protected override ShapeRange ExecutePasteAction(string ribbonId, PowerPointPresentation presentation, PowerPointSlide slide,
                                                        ShapeRange selectedShapes, ShapeRange selectedChildShapes)
        {
            if (selectedShapes.Count <= 0)
            {
                Logger.Log("PasteIntoGroup failed. No valid shape is selected.");
                return null;
            }

            if (selectedShapes.Count == 1 && !Graphics.IsAGroup(selectedShapes[1]))
            {
                Logger.Log("PasteIntoGroup failed. Selection is only a single shape.");
                return null;
            }

            ShapeRange pastingShapes = slide.Shapes.Paste();
            return PasteIntoGroup.Execute(presentation, slide, selectedShapes, pastingShapes);
        }
    }
}