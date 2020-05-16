using System.Windows;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
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

            if (selectedShapes.Count == 1 && !selectedShapes[1].IsAGroup())
            {
                Logger.Log("PasteIntoGroup failed. Selection is only a single shape.");
                return null;
            }

            this.StartNewUndoEntry();

            ShapeRange pastingShapes = ClipboardUtil.PasteShapesFromClipboard(presentation, slide);
            if (pastingShapes == null)
            {
                Logger.Log("PasteLab: Could not paste clipboard contents.");
                MessageBox.Show(PasteLabText.ErrorPaste, PasteLabText.ErrorDialogTitle);
                return null;
            }

            return PasteIntoGroup.Execute(presentation, slide, selectedShapes, pastingShapes);
        }
    }
}