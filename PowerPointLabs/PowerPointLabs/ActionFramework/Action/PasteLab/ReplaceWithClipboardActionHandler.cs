using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.Models;
using PowerPointLabs.PasteLab;

namespace PowerPointLabs.ActionFramework.Action.PasteLab
{
    [ExportActionRibbonId("ReplaceWithClipboard", "ReplaceWithClipboardFreeform", "ReplaceWithClipboardPicture")]
    class ReplaceWithClipboardActionHandler : PasteLabActionHandler
    {
        protected override ShapeRange ExecutePasteAction(string ribbonId, PowerPointPresentation presentation, PowerPointSlide slide,
                                                        Selection selection, ShapeRange pastingShapes)
        {
            if (!IsSelectionShapes(selection))
            {
                MessageBox.Show("Please select at least one shape.", "Error");
                pastingShapes.Delete();
                return null;
            }

            return ReplaceWithClipboard.Execute(presentation, slide, selection, pastingShapes);
        }
    }
}