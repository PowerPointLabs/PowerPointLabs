using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using PowerPointLabs.PasteLab;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.PasteLab
{
    [ExportActionRibbonId(PasteLabText.PasteIntoGroupTag)]
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

            if (selectedShapes.Count == 1 && !ShapeUtil.IsAGroup(selectedShapes[1]))
            {
                Logger.Log("PasteIntoGroup failed. Selection is only a single shape.");
                return null;
            }

            ShapeRange pastingShapes = PasteShapesFromClipboard(slide);
            if (pastingShapes == null)
            {
                return null;
            }

            return PasteIntoGroup.Execute(presentation, slide, selectedShapes, pastingShapes);
        }
    }
}