using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using PowerPointLabs.Models;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs
{
    public class CropToShape
    {
        private const int ErrorCodeForSelectionCountZero = 0;
        private const int ErrorCodeForSelectionNonShape = 1;
        private const int ErrorCodeForExceedSlideBound = 2;
        private const int ErrorCodeForRotationNonZero = 3;

        private const string ErrorMessageForSelectionCountZero = TextCollection.CropToShapeText.ErrorMessageForSelectionCountZero;
        private const string ErrorMessageForSelectionNonShape = TextCollection.CropToShapeText.ErrorMessageForSelectionNonShape;
        private const string ErrorMessageForExceedSlideBound = TextCollection.CropToShapeText.ErrorMessageForExceedSlideBound;
        private const string ErrorMessageForRotationNonZero = TextCollection.CropToShapeText.ErrorMessageForRotationNonZero;
        private const string ErrorMessageForUndefined = TextCollection.CropToShapeText.ErrorMessageForUndefined;

        private const string MessageBoxTitle = "Unable to crop";

        private static readonly string SlidePicture = Path.GetTempPath() + @"\slide.png";
        private static readonly string FillInBackgroundPicture = Path.GetTempPath() + @"\currentFillInBg.png";

        public static PowerPoint.Shape Crop(PowerPoint.Selection selection, double magnifyRatio = 1.0,
                                            bool handleError = true)
        {
            try
            {
                VerifyIsSelectionValid(selection);
            }
            catch (Exception e)
            {
                if (handleError)
                {
                    MessageBox.Show(GetErrorMessageForErrorCode(e.Message), MessageBoxTitle);
                    return null;
                }

                throw;
            }

            return Crop(selection.ShapeRange, handleError: handleError);
        }

        public static PowerPoint.Shape Crop(PowerPoint.ShapeRange shapeRange, double magnifyRatio = 1.0,
                                            bool handleError = true)
        {
            try
            {
                if (!VerifyIsShapeRangeValid(shapeRange, handleError)) return null;

                var shape = GetShapeForSelection(shapeRange);
                TakeScreenshotProxy(shape);
                var filledShape = FillInShapeWithScreenshot(shape, magnifyRatio);

                return filledShape;
            }
            catch (Exception e)
            {
                if (handleError)
                {
                    MessageBox.Show(GetErrorMessageForErrorCode(e.Message), MessageBoxTitle);
                    return null;
                }

                throw;
            }
        }

        private static bool VerifyIsShapeRangeValid(PowerPoint.ShapeRange shapeRange, bool handleError)
        {
            try
            {
                if (shapeRange.Count < 1)
                {
                    ThrowErrorCode(ErrorCodeForSelectionCountZero);
                }

                if (!IsShapeForSelection(shapeRange))
                {
                    ThrowErrorCode(ErrorCodeForSelectionNonShape);
                }

                return true;
            }
            catch (Exception e)
            {
                if (handleError)
                {
                    MessageBox.Show(GetErrorMessageForErrorCode(e.Message), MessageBoxTitle);
                    return false;
                }

                throw;
            }
        }

        private static void VerifyIsSelectionValid(PowerPoint.Selection selection)
        {
            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                ThrowErrorCode(ErrorCodeForSelectionCountZero);
            }
        }

        private static PowerPoint.Shape GetShapeForSelection(PowerPoint.ShapeRange shapeRange)
        {
            var rangeOriginal = shapeRange;
            //some shapes in the selection cannot be used due to 
            //Powerpoint's 'Delete-Undo' issue: when a shape got deleted or cut programmatically, and users undo,
            //then we can only read the shape's name/width/height/left/top.. for others, it'll throw an exception
            //'Cut-Paste' is a common workaround method for this issue
            rangeOriginal.Cut();
            rangeOriginal = PowerPointCurrentPresentationInfo.CurrentSlide.Shapes.Paste();

            var rangeCopy = MakeCopyForShapeRange(rangeOriginal);
            var ungroupedRangeCopy = UngroupAllForShapeRange(rangeCopy);

            var mergedShape = ungroupedRangeCopy[1];
            if (ungroupedRangeCopy.Count > 1)
            {
                mergedShape = ungroupedRangeCopy.Group();
            }

            rangeOriginal.Delete();

            return mergedShape;
        }

        private static PowerPoint.Shape FillInShapeWithScreenshot(PowerPoint.Shape shape, double magnifyRatio = 1.0)
        {
            if (shape.Type != Office.MsoShapeType.msoGroup)
            {
                CreateFillInBackgroundForShape(shape, magnifyRatio);
                shape.Fill.UserPicture(FillInBackgroundPicture);
            }
            else
            {
                using (var slideImage = (Bitmap)Image.FromFile(SlidePicture))
                {
                    foreach (var shapeGroupItem in (from PowerPoint.Shape sh in shape.GroupItems select sh))
                    {
                        CreateFillInBackground(shapeGroupItem, slideImage);
                        shapeGroupItem.Fill.UserPicture(FillInBackgroundPicture);
                    }
                }
            }
            shape.Line.Visible = Office.MsoTriState.msoFalse;
            shape.Copy();
            var shapeToReturn = PowerPointCurrentPresentationInfo.CurrentSlide.Shapes.Paste()[1];
            shape.Delete();
            return shapeToReturn;
        }

        private static void CreateFillInBackgroundForShape(PowerPoint.Shape shape, double magnifyRatio = 1.0)
        {
            using (var slideImage = (Bitmap)Image.FromFile(SlidePicture))
            {
                CreateFillInBackground(shape, slideImage, magnifyRatio);
            }
        }

        private static void CreateFillInBackground(PowerPoint.Shape shape, Bitmap slideImage, double magnifyRatio = 1.0)
        {
            var croppedImage = KiCut(slideImage,
                shape.Left * Utils.Graphics.PictureExportingRatio,
                shape.Top * Utils.Graphics.PictureExportingRatio,
                shape.Width * Utils.Graphics.PictureExportingRatio,
                shape.Height * Utils.Graphics.PictureExportingRatio,
                magnifyRatio);
            croppedImage.Save(FillInBackgroundPicture, ImageFormat.Png);
        }

        private static Bitmap KiCut(Bitmap original, float startX, float startY, float width, float height,
                                    double magnifyRatio = 1.0)
        {
            if (original == null) return null;
            if (startX >= original.Width || startY >= original.Height) return null;
            try
            {
                var outputImage = new Bitmap((int)width, (int)height, PixelFormat.Format32bppArgb);
                
                var inverseRatio = 1 / magnifyRatio;
                
                var newWidth = width * inverseRatio;
                var newHeight = height * inverseRatio;
                var newY = startY + (1 - inverseRatio) / 2 * width;
                var newX = startX + (1 - inverseRatio) / 2 * width;

                var inputGraphics = Graphics.FromImage(outputImage);
                inputGraphics.DrawImage(original,
                    new Rectangle(0, 0, (int)width, (int)height),
                    new Rectangle((int)newX, (int)newY, (int)newWidth, (int)newHeight),
                    GraphicsUnit.Pixel);
                inputGraphics.Dispose();

                return outputImage;
            }
            catch
            {
                return null;
            }
        }

        private static void TakeScreenshotProxy(PowerPoint.Shape shape)
        {
            shape.Visible = Office.MsoTriState.msoFalse;
            Utils.Graphics.ExportSlide(PowerPointCurrentPresentationInfo.CurrentSlide, SlidePicture);
            shape.Visible = Office.MsoTriState.msoTrue;
        }

        private static PowerPoint.ShapeRange MakeCopyForShapeRange(PowerPoint.ShapeRange rangeOriginal)
        {
            //Change shape's name in rangeOriginal, so that shape's name in rangeCopy is the same.
            //This is a naming mechanism in office:
            //When shape's name is the default one, its copy's name will be different (e.g. index got changed).
            //When shape's name is not the default one, its copy's name will be the same as the original shape's
            //use Guid here to ensure that name is unique
            var appendString = Guid.NewGuid().ToString() + "temp";
            ModifyNameForShapeRange(rangeOriginal, appendString);

            rangeOriginal.Copy();
            var rangeCopy = PowerPointCurrentPresentationInfo.CurrentSlide.Shapes.Paste();
            AdjustSamePositionForShapeRange(rangeOriginal, rangeCopy);

            appendString = "_Copy";
            ModifyNameForShapeRange(rangeCopy, appendString);
            return rangeCopy;
        }

        /// <summary>
        /// Assumption: 2 ranges have the same names
        /// </summary>
        /// <param name="rangeReference"></param>
        /// <param name="rangeCopy"></param>
        private static void AdjustSamePositionForShapeRange(IEnumerable rangeReference, IEnumerable rangeCopy)
        {
            var nameMap = (from PowerPoint.Shape shape in rangeReference select shape)
                .ToDictionary(shape => shape.Name, shape => new Tuple<float, float>(shape.Left, shape.Top));
            foreach (var shape in (from PowerPoint.Shape sh in rangeCopy select sh))
            {
                shape.Left = nameMap[shape.Name].Item1;
                shape.Top = nameMap[shape.Name].Item2;
            }
        }

        private static void ModifyNameForShapeRange(IEnumerable range, string appendString)
        {
            foreach (var sh in range)
            {
                ((PowerPoint.Shape) sh).Name += appendString;
            }
        }

        private static bool IsWithinSlide(PowerPoint.Shape shape)
        {
            //-1 and +1 for better user experience
            bool cond1 = shape.Left >= -1;
            bool cond2 = shape.Top >= -1;
            bool cond3 = shape.Left + shape.Width <= PowerPointPresentation.Current.SlideWidth + 1;
            bool cond4 = shape.Top + shape.Height <= PowerPointPresentation.Current.SlideHeight + 1;
            return cond1 && cond2 && cond3 && cond4;
        }

        public static PowerPoint.ShapeRange UngroupAllForShapeRange(IEnumerable range, bool remove = true)
        {
            var ungroupedShapeNames = new List<string>();
            var queue = new Queue<PowerPoint.Shape>();

            foreach (var item in range)
            {
                queue.Enqueue(item as PowerPoint.Shape);
            }
            while (queue.Count != 0)
            {
                var shape = queue.Dequeue();
                if (shape.Type == Office.MsoShapeType.msoGroup)
                {
                    var subRange = shape.Ungroup();
                    foreach (var item in subRange)
                    {
                        queue.Enqueue(item as PowerPoint.Shape);
                    }
                }
                else if ((int)shape.Rotation != 0)
                {
                    if (remove)
                    {
                        RemoveShapesForUngroupAll(shape, ungroupedShapeNames, queue);
                    }

                    ThrowErrorCode(ErrorCodeForRotationNonZero);
                }
                else if (!IsShape(shape))
                {
                    if (remove)
                    {
                        RemoveShapesForUngroupAll(shape, ungroupedShapeNames, queue);
                    }

                    ThrowErrorCode(ErrorCodeForSelectionNonShape);
                }
                else
                {
                    shape.Name += Guid.NewGuid().ToString();
                    ungroupedShapeNames.Add(shape.Name);
                }
            }
            return PowerPointCurrentPresentationInfo.CurrentSlide.Shapes.Range(ungroupedShapeNames.ToArray());
        }

        private static void RemoveShapesForUngroupAll(PowerPoint.Shape shape, List<string> ungroupedShapes, Queue<PowerPoint.Shape> queue)
        {
            shape.Delete();
            if (ungroupedShapes.Count > 0)
            {
                PowerPointCurrentPresentationInfo.CurrentSlide.Shapes.Range(ungroupedShapes.ToArray()).Delete();
            }
            while (queue.Count != 0)
            {
                queue.Dequeue().Delete();
            }
        }

        private static bool IsShapeForSelection(PowerPoint.ShapeRange shapeRange)
        {
            return (from PowerPoint.Shape shape in shapeRange select shape).All(IsShape);
        }

        private static bool IsShape(PowerPoint.Shape shape)
        {
            return shape.Type == Office.MsoShapeType.msoAutoShape 
                || shape.Type == Office.MsoShapeType.msoFreeform 
                || shape.Type == Office.MsoShapeType.msoGroup;
        }

        private static void ThrowErrorCode(int typeOfError)
        {
            throw new Exception(typeOfError.ToString(CultureInfo.InvariantCulture));
        }

        private static void IgnoreExceptionThrown(){}

        public static string GetErrorMessageForErrorCode(string errorCode)
        {
            var errorCodeInteger = -1;
            try
            {
                errorCodeInteger = Int32.Parse(errorCode);
            }
            catch
            {
                IgnoreExceptionThrown();
            }
            switch (errorCodeInteger)
            {
                case ErrorCodeForSelectionCountZero:
                    return ErrorMessageForSelectionCountZero;
                case ErrorCodeForSelectionNonShape:
                    return ErrorMessageForSelectionNonShape;
                case ErrorCodeForExceedSlideBound:
                    return ErrorMessageForExceedSlideBound;
                case ErrorCodeForRotationNonZero:
                    return ErrorMessageForRotationNonZero;
                default:
                    return ErrorMessageForUndefined;
            }
        }

        public static Bitmap GetCutOutShapeMenuImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.CutOutShapeMenu);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetCutOutShapeMenuImage");
                throw;
            }
        }
    }
}
