using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Models;

namespace PowerPointLabs.PasteLab
{
    static internal class ReplaceWithClipboard
    {
        public static void Execute(PowerPointPresentation presentation, PowerPointSlide slide, Selection selection, ShapeRange pastingShapes)
        {
            Shape selectedShape = selection.ShapeRange[1];

            if (selection.HasChildShapeRange)
            {
                selectedShape = selection.ChildShapeRange[1];
                float posLeft = selectedShape.Left;
                float posTop = selectedShape.Top;
                selectedShape.Delete();

                PasteIntoGroup.Execute(presentation, slide, selection.ShapeRange, pastingShapes, posLeft, posTop);

                return;
            }

            Shape pastingShape = pastingShapes[1];
            if (pastingShapes.Count > 1)
            {
                pastingShape = pastingShapes.Group();
            }
            pastingShape.Left = selectedShape.Left;
            pastingShape.Top = selectedShape.Top;

            foreach (Effect eff in slide.TimeLine.MainSequence)
            {
                if (eff.Shape == selectedShape)
                {
                    Effect newEff = slide.TimeLine.MainSequence.Clone(eff);
                    newEff.Shape = pastingShape;
                    eff.Delete();
                }
            }
            
            selectedShape.Delete();
        }
    }
}
