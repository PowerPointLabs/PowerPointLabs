using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.Models;
using PowerPointLabs.PasteLab;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

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
                MessageBox.Show(TextCollection.PasteLabText.ReplaceWithClipboardActionHandlerReminderText, TextCollection.CommonText.ErrorTitle);
                return null;
            }

            ShapeRange pastingShapes = ClipboardUtil.PasteShapesFromClipboard(presentation, slide);
            if (pastingShapes == null)
            {
                return null;
            }

            return ReplaceWithClipboard.Execute(presentation, slide, selectedShapes, selectedChildShapes, pastingShapes);
        }
    }
}