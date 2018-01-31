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
                copiedShape = CopyMsoPlaceHolder(formats, shape);
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

        /// <summary>
        /// Fake a copy by creating a similar object with the same formats
        /// Copy/Pasting MsoPlaceHolder doesn't work.
        /// Note: Shapes.AddPlaceholder(..) does not work as well.
        /// It restores a deleted placeholder to the slide, not create a shape
        /// </summary>
        /// <param name="formats"></param>
        /// <param name="msoPlaceHolder"></param>
        /// <returns>returns null if input placeholder is not supported</returns>
        private Shape CopyMsoPlaceHolder(FormatTreeNode[] formats, Shape msoPlaceHolder)
        {
            PpPlaceholderType realType = msoPlaceHolder.PlaceholderFormat.Type;
            
            // charts, tables, pictures & smart shapes may return a general type,
            // ppPlaceHolderObject or ppPlaceHolderVerticalObject
            bool isGeneralType = realType == PpPlaceholderType.ppPlaceholderObject ||
                                 realType == PpPlaceholderType.ppPlaceholderVerticalObject;
            if (isGeneralType)
            {
                realType = GetSpecificPlaceholderType(msoPlaceHolder);
            }
            
            // create an appropriate shape, based on placeholder type
            Shape shapeTemplate = null;
            switch (realType)
            {
                case PpPlaceholderType.ppPlaceholderBody:
                case PpPlaceholderType.ppPlaceholderCenterTitle:
                case PpPlaceholderType.ppPlaceholderSubtitle:
                    // unable to support FarEast text orientation, API does not give us enough information
                    shapeTemplate = Slides[0].Shapes.AddTextbox(
                        MsoTextOrientation.msoTextOrientationHorizontal,
                        msoPlaceHolder.Left,
                        msoPlaceHolder.Top,
                        msoPlaceHolder.Width,
                        msoPlaceHolder.Height);
                    break;
                case PpPlaceholderType.ppPlaceholderVerticalBody:
                case PpPlaceholderType.ppPlaceholderVerticalTitle:
                    // unable to support FarEast text orientation, API does not give us enough information
                    shapeTemplate = Slides[0].Shapes.AddTextbox(
                        MsoTextOrientation.msoTextOrientationVertical,
                        msoPlaceHolder.Left,
                        msoPlaceHolder.Top,
                        msoPlaceHolder.Width,
                        msoPlaceHolder.Height);
                    break;
                case PpPlaceholderType.ppPlaceholderChart:
                case PpPlaceholderType.ppPlaceholderOrgChart:
                    // charts are not yet supported by Synclab
                    break;
                case PpPlaceholderType.ppPlaceholderTable:
                    // tables are not yet supported by Synclab
                    break;
                case PpPlaceholderType.ppPlaceholderPicture:
                case PpPlaceholderType.ppPlaceholderBitmap:
                    // TODO: images are not yet supported by SyncLab, see PictureFormat for things to copy
                    // do nothing for now
                    break;
                case PpPlaceholderType.ppPlaceholderVerticalObject:
                case PpPlaceholderType.ppPlaceholderObject:
                    // already narrowed down the type 
                    // should only perform actions valid for all placeholder objects here 
                    // do nothing for now
                    break;
                default:
                    // types not listed above are types that do not make sense to be supported by Synclab
                    // eg. footer, header, date
                    break;
            }
            
            if (shapeTemplate == null)
            {
                // placeholder type is not supported, no copy made
                return null;
            }
            
            SyncFormatUtil.ApplyFormats(formats, msoPlaceHolder, shapeTemplate);
            return shapeTemplate;
        }

        /// <summary>
        /// Targets only msoPlaceHolderObject & msoPlaceHolderVerticalObject
        /// Attempt to return a more specific placeholder type, if we can determine it
        /// </summary>
        /// <param name="placeHolder"></param>
        /// <returns>a specific type, or the shape's original type</returns>
        private PpPlaceholderType GetSpecificPlaceholderType(Shape placeHolder)
        {
            bool isPicture = IsPlaceHolderPicture(placeHolder);
            if (isPicture)
            {
                return PpPlaceholderType.ppPlaceholderPicture;
            }
            else
            {
                return placeHolder.PlaceholderFormat.Type;
            }
        }

        /// <summary>
        /// Checks if a placeholder is a Picture
        /// shape.PlaceHolderFormat.Type is insufficient, sometimes returning the more general "ppPlaceHolderObject".
        /// </summary>
        /// <param name="placeHolder"></param>
        /// <returns></returns>
        private bool IsPlaceHolderPicture(Shape placeHolder)
        {
            try
            {
                // attempt to access PictureFormat properties, an exception will be thrown
                // if shape is not a Picture.
                float attemptToAccess = placeHolder.PictureFormat.CropTop;
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
