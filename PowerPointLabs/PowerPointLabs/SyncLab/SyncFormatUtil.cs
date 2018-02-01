using System;
using System.Drawing;
using System.Drawing.Text;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.SyncLab.ObjectFormats;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace PowerPointLabs.SyncLab
{
    public class SyncFormatUtil
    {
        #region Display Image Utils

        public static Shapes GetTemplateShapes()
        {
            SyncLabShapeStorage shapeStorage = SyncLabShapeStorage.Instance;
            return shapeStorage.Slides[SyncLabShapeStorage.FormatStorageSlide].Shapes;
        }
        
        public static Bitmap GetTextDisplay(string text, System.Drawing.Font font, Size imageSize)
        {
            Bitmap image = new Bitmap(imageSize.Width, imageSize.Height);
            Graphics g = Graphics.FromImage(image);
            g.TextRenderingHint = TextRenderingHint.AntiAlias;
            SizeF textSize = g.MeasureString(text, font);
            if (textSize.Width == 0 || textSize.Height == 0)
            {
                // nothing to print
                return image;
            }
            if (textSize.Width > imageSize.Width || textSize.Height > imageSize.Height)
            {
                double scale = Math.Min(imageSize.Width / textSize.Width, imageSize.Height / textSize.Height);
                font = new System.Drawing.Font(font.FontFamily, Convert.ToSingle(font.Size * scale),
                                                            font.Style, font.Unit, font.GdiCharSet, font.GdiVerticalFont);
                textSize = g.MeasureString(text, font);
            }
            float xPos = Convert.ToSingle((imageSize.Width - textSize.Width) / 2);
            float yPos = Convert.ToSingle((imageSize.Height - textSize.Height) / 2);
            g.DrawString(text, font, Brushes.Black, xPos, yPos);
            g.Dispose();
            return image;
        }

        #endregion

        #region Shape Name Utils
        public static bool IsValidFormatName(string name)
        {
            name = name.Trim();
            return name.Length > 0;
        }

        #endregion

        #region Sync Shape Format utils

        /// <summary>
        /// Applies the specified formats from one shape to multiple shapes
        /// </summary>
        /// <param name="formats">Formats to apply</param>
        /// <param name="formatShape">source shape</param>
        /// <param name="newShapes">destination shape</param>
        public static void ApplyFormats(Format[] formats, Shape formatShape, ShapeRange newShapes)
        {
            foreach (Shape newShape in newShapes)
            {
                ApplyFormats(formats, formatShape, newShape);
            }
        }

        public static void ApplyFormats(Format[] formats, Shape formatShape, Shape newShape)
        {
            foreach (Format format in formats)
            {
                format.SyncFormat(formatShape, newShape);
            }
        }
        
        #endregion

        #region PlaceHolder utils

        public static bool CanCopyMsoPlaceHolder(Shape placeholder)
        {
            var emptyArray = new Format[0];
            Shape copyAttempt = CopyMsoPlaceHolder(emptyArray, placeholder);
            
            if (copyAttempt == null)
            {
                return false;
            }
            
            copyAttempt.Delete();
            return true;
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
        public static Shape CopyMsoPlaceHolder(Format[] formats, Shape msoPlaceHolder)
        {
            Shapes templateShapes = GetTemplateShapes();
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
                // the type never seems to be anything other than subtitle, center title, title, body or object.
                // still, place the rest here to be safe.
                case PpPlaceholderType.ppPlaceholderBody:
                case PpPlaceholderType.ppPlaceholderCenterTitle:
                case PpPlaceholderType.ppPlaceholderTitle:
                case PpPlaceholderType.ppPlaceholderSubtitle:
                case PpPlaceholderType.ppPlaceholderVerticalBody:
                case PpPlaceholderType.ppPlaceholderVerticalTitle:
                    shapeTemplate = templateShapes.AddTextbox(
                        msoPlaceHolder.TextFrame.Orientation,
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
            
            ApplyFormats(formats, msoPlaceHolder, shapeTemplate);
            return shapeTemplate;
        }

        /// <summary>
        /// Targets only msoPlaceHolderObject & msoPlaceHolderVerticalObject
        /// Attempt to return a more specific placeholder type, if we can determine it
        /// </summary>
        /// <param name="placeHolder"></param>
        /// <returns>a specific type, or the shape's original type</returns>
        private static PpPlaceholderType GetSpecificPlaceholderType(Shape placeHolder)
        {
            bool isPicture = IsPlaceHolderPicture(placeHolder);
            bool isBody = IsPlaceHolderBody(placeHolder);
            if (isPicture)
            {
                return PpPlaceholderType.ppPlaceholderPicture;
            }
            else if (isBody)
            {
                return PpPlaceholderType.ppPlaceholderBody;
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
        private static bool IsPlaceHolderPicture(Shape placeHolder)
        {
            try
            {
                // attempt to access PictureFormat properties, an exception will be thrown
                // if shape is not a Picture.
                float unused = placeHolder.PictureFormat.CropTop;
                return true;
            }
            catch
            {
                return false;
            }
        }
        
        /// <summary>
        /// Checks if a placeholder is a Body
        /// shape.PlaceHolderFormat.Type is insufficient, sometimes returning the more general "ppPlaceHolderObject".
        /// </summary>
        /// <param name="placeHolder"></param>
        /// <returns></returns>
        private static bool IsPlaceHolderBody(Shape placeHolder)
        {
            return placeHolder.HasTextFrame == MsoTriState.msoTrue;
        }
        #endregion
    }
}
