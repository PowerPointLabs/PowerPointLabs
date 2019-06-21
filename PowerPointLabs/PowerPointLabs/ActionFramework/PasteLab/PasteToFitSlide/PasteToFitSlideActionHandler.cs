using System.Windows;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using PowerPointLabs.PasteLab;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Windows;

namespace PowerPointLabs.ActionFramework.PasteLab
{
    [ExportActionRibbonId(PasteLabText.PasteToFitSlideTag)]
    class PasteToFitSlideActionHandler : PasteLabActionHandler
    {
        protected override ShapeRange ExecutePasteAction(string ribbonId, PowerPointPresentation presentation, PowerPointSlide slide,
                                                        ShapeRange selectedShapes, ShapeRange selectedChildShapes)
        {
            ShapeRange pastingShapes = ClipboardUtil.PasteShapesFromClipboard(presentation, slide);
            if (pastingShapes == null)
            {
                Logger.Log("PasteLab: Could not paste clipboard contents.");
                MessageBoxUtil.Show(PasteLabText.ErrorPaste, PasteLabText.ErrorDialogTitle);
                return pastingShapes;
            }

            PasteToFitSlide.Execute(presentation, slide, pastingShapes, presentation.SlideWidth, presentation.SlideHeight);
            return null;
        }
    }
}
