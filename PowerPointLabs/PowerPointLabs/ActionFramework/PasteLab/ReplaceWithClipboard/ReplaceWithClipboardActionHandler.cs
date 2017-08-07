using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.Models;
using PowerPointLabs.PasteLab;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.PasteLab
{
    [ExportActionRibbonId(PasteLabText.ReplaceWithClipboardTag)]
    class ReplaceWithClipboardActionHandler : PasteLabActionHandler
    {
        protected override ShapeRange ExecutePasteAction(string ribbonId, PowerPointPresentation presentation, PowerPointSlide slide,
                                                        ShapeRange selectedShapes, ShapeRange selectedChildShapes)
        {
            if (selectedShapes.Count <= 0)
            {
                MessageBox.Show("Please select at least one shape.", "Error");
                return null;
            }

            ShapeRange pastingShapes = PasteShapesFromClipboard(slide);
            if (pastingShapes == null)
            {
                return null;
            }

            return ReplaceWithClipboard.Execute(presentation, slide, selectedShapes, selectedChildShapes, pastingShapes);
        }
    }
}