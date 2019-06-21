using System.Windows;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Windows;

namespace PowerPointLabs.ActionFramework.PasteLab
{
    [ExportActionRibbonId(PasteLabText.PasteAtOriginalPositionTag)]
    class PasteAtOriginalPositionActionHandler : PasteLabActionHandler
    {
        protected override ShapeRange ExecutePasteAction(string ribbonId, PowerPointPresentation presentation, PowerPointSlide slide,
                                                        ShapeRange selectedShapes, ShapeRange selectedChildShapes)
        {
            PowerPointSlide tempSlide = presentation.AddSlide(index: slide.Index);
            ShapeRange tempPastingShapes = ClipboardUtil.PasteShapesFromClipboard(presentation, tempSlide);
            if (tempPastingShapes == null)
            {
                tempSlide.Delete();
                ShapeRange shapes = ClipboardUtil.PasteShapesFromClipboard(presentation, slide);
                if (shapes == null) 
                {
                    Logger.Log("PasteLab: Could not paste clipboard contents.");
                    MessageBoxUtil.Show(PasteLabText.ErrorPaste, PasteLabText.ErrorDialogTitle);
                }
                return shapes;
            }

            ShapeRange pastingShapes = slide.CopyShapesToSlide(tempPastingShapes);
            tempSlide.Delete();

            return pastingShapes;
        }
    }
}