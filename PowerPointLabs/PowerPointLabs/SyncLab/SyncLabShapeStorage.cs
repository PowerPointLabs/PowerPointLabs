using System;
using Microsoft.Office.Core;
using PowerPointLabs.Models;
using PowerPointLabs.SyncLab.Views;
using PowerPointLabs.TextCollection;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace PowerPointLabs.SyncLab
{
    /// <summary>
    /// Saves shapes into a PowerPointPresentation that exists in the background.
    /// The exact saved shapes may change in type but style will be retained.
    /// Eg: PlaceHolders are saved as Textboxes
    /// </summary>
    public sealed class SyncLabShapeStorage : PowerPointPresentation
    {

        public const int FormatStorageSlide = 0;

        private int nextKey = 0;

        private static readonly Lazy<SyncLabShapeStorage> StorageInstance =
            new Lazy<SyncLabShapeStorage>(() => new SyncLabShapeStorage());

        public static SyncLabShapeStorage Instance
        {
            get { return StorageInstance.Value; }
        }

        private SyncLabShapeStorage() : base()
        {
            Path = System.IO.Path.GetTempPath();
            Name = SyncLabText.StorageFileName;
            OpenInBackground();
            ClearShapes();
        }

        /// <summary>
        /// Saves shape in storage
        /// Returns a key to find the shape by,
        /// or null if the shape cannot be copied
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="formats">Required for msoPlaceholder</param>
        /// <returns></returns>
        public string CopyShape(Shape shape, FormatTreeNode[] formats)
        {
            // copies a shape, and returns a shape name
            Shape copiedShape = null;
            try
            {
                if (shape.Type == MsoShapeType.msoPlaceholder)
                {
                    copiedShape = CopyMsoPlaceHolder(formats, shape);
                }
                else
                {
                    shape.Copy();
                    copiedShape = Slides[0].Shapes.Paste()[1];
                }
                    
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
            Shapes shapes = Slides[0].Shapes;
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
            int index = 1;
            Shapes shapes = Slides[0].Shapes;
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
            Slides[FormatStorageSlide].DeleteAllShapes();
        }

        /// <summary>
        /// Fake a copy by creating a textbox with the same formats
        /// Copy/Pasting MsoPlaceHolder doesn't work.
        /// Note: Shapes.AddPlaceholder(..) is not applicable.
        /// It restores a deleted placeholder to the slide, not create a shape
        /// </summary>
        /// <param name="formats"></param>
        /// <param name="msoPlaceHolder"></param>
        /// <returns></returns>
        private Shape CopyMsoPlaceHolder(FormatTreeNode[] formats, Shape msoPlaceHolder)
        {
            Shape savedShape = Slides[0].Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                msoPlaceHolder.Left,
                msoPlaceHolder.Top,
                msoPlaceHolder.Width,
                msoPlaceHolder.Height);
            
            SyncFormatUtil.ApplyFormats(formats, msoPlaceHolder, savedShape);
            return savedShape;
        }
    }
}
