using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.Models;
using PowerPointLabs.PasteLab;

namespace PowerPointLabs.ActionFramework.Action.PasteLab
{
    [ExportActionRibbonId(
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuShape,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuLine,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuFreeform,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuPicture,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuGroup,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuInk,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuVideo,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuChart,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuTable,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuTableCell,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuSmartArt,
        TextCollection.ReplaceWithClipboardId + TextCollection.RibbonButton)]
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