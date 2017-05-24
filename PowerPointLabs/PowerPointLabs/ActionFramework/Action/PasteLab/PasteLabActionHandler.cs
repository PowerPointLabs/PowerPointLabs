using System.Collections.Generic;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ActionFramework.Util;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.Action.PasteLab
{
    abstract class PasteLabActionHandler : BaseUtilActionHandler
    {
        private static readonly string SelectOrderTagName = "SelectOrderTagName";
        private static readonly string SelectChildOrderTagName = "SelectChildOrderTagName";

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

            if (IsSelectionShapes(selection))
            {
                // Save clipboard onto a temp slide, because CorruptionCorrrection uses Copy-Paste
                PowerPointSlide tempClipboardSlide = presentation.AddSlide(index: slide.Index);
                ShapeRange tempClipboardShapes = tempClipboardSlide.Shapes.Paste();

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
                    if (Graphics.IsAGroup(shape))
                    {
                        for (int i = 1; i <= shape.GroupItems.Count; i++)
                        {
                            Shape child = shape.GroupItems.Range(i)[1];
                            if (!child.Tags[SelectChildOrderTagName].Equals(""))
                            {
                                correctedChildShapeList.Add(child);
                            }
                        }
                    }
                }
                correctedShapeList.Sort((sh1, sh2) => int.Parse(sh2.Tags[SelectOrderTagName]) - int.Parse(sh1.Tags[SelectOrderTagName]));
                correctedChildShapeList.Sort((sh1, sh2) => int.Parse(sh2.Tags[SelectChildOrderTagName]) - int.Parse(sh1.Tags[SelectChildOrderTagName]));
                passedSelectedShapes = slide.ToShapeRange(correctedShapeList);
                passedSelectedChildShapes = slide.ToShapeRange(correctedChildShapeList);

                // Remove the tags after they have been used
                foreach (Shape shape in passedSelectedShapes)
                {
                    shape.Tags.Delete(SelectOrderTagName);
                }
                foreach (Shape shape in passedSelectedChildShapes)
                {
                    shape.Tags.Delete(SelectChildOrderTagName);
                }

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
    }
}
