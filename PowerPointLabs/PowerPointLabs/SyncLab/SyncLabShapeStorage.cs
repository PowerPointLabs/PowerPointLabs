using System;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Models;

namespace PowerPointLabs.SyncLab
{
    public class SyncLabShapeStorage : PowerPointPresentation
    {
#pragma warning disable 0618

        int nextKey = 0;

        public SyncLabShapeStorage() : base()
        {
            Path = Globals.ThisAddIn.PrepareTempFolder(
                                         PowerPointPresentation.Current.Presentation);
            Name = string.Format(TextCollection.SyncLabStorageTemplateName,
                                 DateTime.Now.ToString("yyyyMMddHHmmssffff"));
            OpenInBackground();
            ClearShapes();
        }

        // Saves shape in storage
        // Returns a key to find the shape by,
        // or null if the shape cannot be copied
        public string CopyShape(Shape shape)
        {
            // copies a shape, and returns a shape name
            shape.Copy();
            Shape copiedShape = null;
            try
            {
                copiedShape = Slides[0].Shapes.Paste()[1];
            }
            catch
            {
                return null;
            }
            string shapeKey = nextKey.ToString();
            nextKey++;
            copiedShape.Name = shapeKey;
            ForceSave();
            return shapeKey;
        }

        public Shape GetShape(string shapeKey)
        {
            var shapes = Slides[0].Shapes;
            for (int i = 1; i <= shapes.Count; i++)
            {
                if (shapes[i].Name.Equals(shapeKey))
                {
                    return shapes[i];
                }
            }
            return null;
        }

        public void RemoveShape(string shapeKey)
        {
            var index = 1;
            var shapes = Slides[0].Shapes;
            while (index <= shapes.Count)
            {
                if (shapes[index].Name.Equals(shapeKey))
                {
                    shapes[index].Delete();
                }
                else
                {
                    index++;
                }
            }
        }

        public void ForceSave()
        {
            Save();
            Close();
            OpenInBackground();
        }

        public void ClearShapes()
        {
            while (SlideCount > 0)
            {
                GetSlide(1).Delete();
            }
            AddSlide();
            Slides[0].DeleteAllShapes();
        }
    }
}
