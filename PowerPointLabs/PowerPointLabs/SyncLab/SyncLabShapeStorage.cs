using System;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
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
        /// <returns>identifier of copied shape</returns>
        public string CopyShape(Shape shape, FormatTreeNode[] formats)
        {
            Shape copiedShape = null;
            if (shape.Type == MsoShapeType.msoPlaceholder)
            {
                copiedShape = SyncFormatUtil.CopyMsoPlaceHolder(formats, shape);
            }
            else
            {
                try
                {
                    shape.Copy();
                    copiedShape = Slides[0].Shapes.Paste()[1];
                }
                catch
                {
                    copiedShape = null;
                }
            }

            if (copiedShape == null)
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
    }
}
