﻿using System.Collections.Generic;
using System.Runtime.InteropServices;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ActionFramework.Util;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.PasteLab
{
    abstract class PasteLabActionHandler : BaseUtilActionHandler
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

            if (Graphics.IsClipboardEmpty())
            {
                Logger.Log(ribbonId + " failed. Clipboard is empty.");
                return;
            }

            ShapeRange passedSelectedShapes = null;
            ShapeRange passedSelectedChildShapes = null;

            if (IsSelectionShapes(selection) && !IsSelectionIgnored(ribbonId))
            {
                // Save clipboard onto a temp slide, because CorruptionCorrrection uses Copy-Paste
                PowerPointSlide tempClipboardSlide = presentation.AddSlide(index: slide.Index);
                ShapeRange tempClipboardShapes = PasteShapesFromClipboard(tempClipboardSlide);

                // Nothing is pasted, stop now
                if (tempClipboardShapes == null)
                {
                    tempClipboardSlide.Delete();
                    return;
                }

                // Preserve selection using tags
                ShapeRange selectedShapes = selection.ShapeRange;
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
                ShapeRange correctedShapes = Graphics.CorruptionCorrection(selectedShapes, slide);

                // Reselect the preserved selections
                List<Shape> correctedShapeList = new List<Shape>();
                List<Shape> correctedChildShapeList = new List<Shape>();
                foreach (Shape shape in correctedShapes)
                {
                    correctedShapeList.Add(shape);
                    correctedChildShapeList.AddRange(Graphics.GetChildrenWithNonEmptyTag(shape, SelectChildOrderTagName));
                }
                correctedShapeList.Sort((sh1, sh2) => int.Parse(sh1.Tags[SelectOrderTagName]) - int.Parse(sh2.Tags[SelectOrderTagName]));
                correctedChildShapeList.Sort((sh1, sh2) => int.Parse(sh1.Tags[SelectChildOrderTagName]) - int.Parse(sh2.Tags[SelectChildOrderTagName]));
                passedSelectedShapes = slide.ToShapeRange(correctedShapeList);
                passedSelectedChildShapes = slide.ToShapeRange(correctedChildShapeList);

                // Remove the tags after they have been used
                Graphics.DeleteTagFromShapes(passedSelectedShapes, SelectOrderTagName);
                Graphics.DeleteTagFromShapes(passedSelectedChildShapes, SelectChildOrderTagName);

                // Revert clipboard
                tempClipboardShapes.Copy();
                tempClipboardSlide.Delete();
            }

            ShapeRange result = ExecutePasteAction(ribbonId, presentation, slide, passedSelectedShapes, passedSelectedChildShapes);
            if (result != null)
            {
                result.Select();
            }
        }

        protected abstract ShapeRange ExecutePasteAction(string ribbonId, PowerPointPresentation presentation, PowerPointSlide slide,
                                                        ShapeRange selectedShapes, ShapeRange selectedChildShapes);

        protected ShapeRange PasteShapesFromClipboard(PowerPointSlide slide)
        {
            try
            {
                return slide.Shapes.Paste();
            }
            catch (COMException e)
            {
                // May be thrown if there is placeholder shape in clipboard
                Logger.LogException(e, "PasteShapeFromClipboard");
                return null;
            }
        }

        private bool IsSelectionIgnored(string ribbonId)
        {
            return ribbonId.StartsWith("PasteAtCursorPosition") ||
                ribbonId.StartsWith("PasteAtOriginalPosition") ||
                ribbonId.StartsWith("PasteToFillSlide");
        }
    }
}
