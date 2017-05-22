using System.Collections.Generic;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Models;

namespace PowerPointLabs.PasteLab
{
    static internal class ReplaceWithClipboard
    {
        public static ShapeRange Execute(PowerPointPresentation presentation, PowerPointSlide slide, Selection selection, ShapeRange pastingShapes)
        {
            Shape selectedShape = selection.ShapeRange[1];

            if (selection.HasChildShapeRange)
            {
                selectedShape = selection.ChildShapeRange[1];
                float posLeft = selectedShape.Left;
                float posTop = selectedShape.Top;

                Shape selectedGroup = selectedShape.ParentGroup;
                Shape tempSelectedGroup = slide.CopyShapeToSlide(selectedGroup);
                slide.DeleteShapeAnimations(tempSelectedGroup);
                slide.TransferAnimation(selectedGroup, tempSelectedGroup);

                List<Shape> selectedGroupShapeList = new List<Shape>();
                int selectedGroupCount = selectedGroup.GroupItems.Count;
                for (int i = 1; i <= selectedGroupCount; i++)
                {
                    Shape shape = selectedGroup.GroupItems.Range(i)[1];
                    if (shape.Name.Equals(selectedShape.Name))
                    {
                        continue;
                    }
                    selectedGroupShapeList.Add(shape);
                }
                
                ShapeRange shapesToGroup = slide.ToShapeRange(selectedGroupShapeList);
                shapesToGroup = slide.CopyShapesToSlide(shapesToGroup);
                selectedGroup.Delete();
                
                ShapeRange result = PasteIntoGroup.Execute(presentation, slide, shapesToGroup, pastingShapes, posLeft, posTop);
                slide.TransferAnimation(tempSelectedGroup, result[1]);
                
                tempSelectedGroup.Delete();
                return result;
            }

            Shape pastingShape = pastingShapes[1];
            if (pastingShapes.Count > 1)
            {
                pastingShape = pastingShapes.Group();
            }
            pastingShape.Left = selectedShape.Left;
            pastingShape.Top = selectedShape.Top;

            slide.DeleteShapeAnimations(pastingShape);
            slide.TransferAnimation(selectedShape, pastingShape);
            selectedShape.Delete();
            
            return slide.ToShapeRange(pastingShape);
        }
    }
}
