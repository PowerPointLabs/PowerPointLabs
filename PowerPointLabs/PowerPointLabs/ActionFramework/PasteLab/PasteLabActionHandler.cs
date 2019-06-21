using System.Collections.Generic;
using System.Windows;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Windows;

namespace PowerPointLabs.ActionFramework.PasteLab
{
    abstract class PasteLabActionHandler : ActionHandler
    {
        private static readonly string SelectOrderTagName = "PasteLabSelectOrder";
        private static readonly string SelectChildOrderTagName = "PasteLabSelectChildOrder";

        // Sealed method: Subclasses should override ExecutePasteAction instead
        protected sealed override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();
            
            PowerPointPresentation presentation = this.GetCurrentPresentation();
            PowerPointSlide slide = this.GetCurrentSlide();
            Selection selection = this.GetCurrentSelection();

            if (ClipboardUtil.IsClipboardEmpty())
            {
                Logger.Log(ribbonId + " failed. Clipboard is empty.");
                MessageBoxUtil.Show(PasteLabText.ErrorEmptyClipboard, PasteLabText.ErrorDialogTitle);
                return;
            }

            if (slide == null)
            {
                Logger.Log(ribbonId + " failed. Selection is empty.");
                MessageBoxUtil.Show(PasteLabText.ErrorNoSelection, PasteLabText.ErrorDialogTitle);
                return;
            }

            ShapeRange passedSelectedShapes = null;
            ShapeRange passedSelectedChildShapes = null;

            if (ShapeUtil.IsSelectionShape(selection) && !IsSelectionIgnored(ribbonId))
            {
                // When pasting some objects, the selection may change to the pasted object (e.g. jpg from desktop).
                // Therefore we must capture the selection first.
                ShapeRange selectedShapes = selection.ShapeRange;

                // Preserve selection by tagging them
                for (int i = 1; i <= selectedShapes.Count; i++)
                {
                    selectedShapes[i].Tags.Add(SelectOrderTagName, i.ToString());
                }

                ShapeRange selectedChildShapes = null;
                if (selection.HasChildShapeRange)
                {
                    selectedChildShapes = selection.ChildShapeRange;
                    for (int i = 1; i <= selectedChildShapes.Count; i++)
                    {
                        selectedChildShapes[i].Tags.Add(SelectChildOrderTagName, i.ToString());
                    }
                }

                // Corruption correction
                ShapeRange correctedShapes = ShapeUtil.CorruptionCorrection(selectedShapes, slide);

                // Reselect the preserved selections
                List<Shape> correctedShapeList = new List<Shape>();
                List<Shape> correctedChildShapeList = new List<Shape>();
                foreach (Shape shape in correctedShapes)
                {
                    correctedShapeList.Add(shape);
                    correctedChildShapeList.AddRange(ShapeUtil.GetChildrenWithNonEmptyTag(shape, SelectChildOrderTagName));
                }
                correctedShapeList.Sort((sh1, sh2) => int.Parse(sh1.Tags[SelectOrderTagName]) - int.Parse(sh2.Tags[SelectOrderTagName]));
                correctedChildShapeList.Sort((sh1, sh2) => int.Parse(sh1.Tags[SelectChildOrderTagName]) - int.Parse(sh2.Tags[SelectChildOrderTagName]));
                passedSelectedShapes = slide.ToShapeRange(correctedShapeList);
                passedSelectedChildShapes = slide.ToShapeRange(correctedChildShapeList);

                // Remove shape tags after they have been used
                ShapeUtil.DeleteTagFromShapes(passedSelectedShapes, SelectOrderTagName);
                ShapeUtil.DeleteTagFromShapes(passedSelectedChildShapes, SelectChildOrderTagName);
            }

            ShapeRange result = ExecutePasteAction(ribbonId, presentation, slide, passedSelectedShapes, passedSelectedChildShapes);
            if (result != null)
            {
                result.Select();
            }
        }

        protected abstract ShapeRange ExecutePasteAction(string ribbonId, PowerPointPresentation presentation, PowerPointSlide slide,
                                                        ShapeRange selectedShapes, ShapeRange selectedChildShapes);

        private bool IsSelectionIgnored(string ribbonId)
        {
            return ribbonId.StartsWith("PasteAtCursorPosition") ||
                ribbonId.StartsWith("PasteAtOriginalPosition") ||
                ribbonId.StartsWith("PasteToFillSlide") ||
                ribbonId.StartsWith("PasteToFitSlide");
        }
    }
}
