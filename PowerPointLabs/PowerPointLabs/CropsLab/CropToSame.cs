using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Threading;

namespace PowerPointLabs
{
    public class CropToSame
    {
#pragma warning disable 0618
        private const int ErrorCodeForSelectionCountZero = 0;
        private const int ErrorCodeForSelectionNonPicture = 1;

        private const string ErrorMessageForSelectionCountZero = TextCollection.CropToSlideText.ErrorMessageForSelectionCountZero;
        private const string ErrorMessageForSelectionNonPicture = TextCollection.CropToSlideText.ErrorMessageForSelectionNonPicture;
        private const string ErrorMessageForUndefined = TextCollection.CropToSlideText.ErrorMessageForUndefined;

        private const string MessageBoxTitle = "Unable to crop";

        private static readonly string ShapePicture = Path.GetTempPath() + @"\shape.png";

        private static DispatcherTimer cropTimer = new DispatcherTimer();

        private static PowerPoint.ShapeRange shapes;

        public static void StartCropToSame(PowerPoint.Selection selection, Office.IRibbonControl control, bool handleError = true)
        {
            try
            {
                VerifyIsSelectionValid(selection);
                if (!VerifyIsShapeRangeValid(selection.ShapeRange, handleError)) return;
                cropTimer.Tick -= CropHandler;
                cropTimer.Tick += CropHandler;

                shapes = selection.ShapeRange;
                cropTimer.Start();
            }
            catch (Exception e)
            {
                if (handleError)
                {
                    ProcessErrorMessage(e);
                    return;
                }

                throw;
            }
            
        }

        private static bool IsFirstShapeSelected()
        {
            foreach (PowerPoint.Shape shape in PowerPointCurrentPresentationInfo.CurrentSelection.ShapeRange)
            {
                MessageBox.Show(shapes[1].Id + " " + shape.Id);
                if (shapes[1].Id == shape.Id)
                {
                    return true;
                }
            }
            return false;
        }
        private static void CropHandler(object sender, EventArgs e)
        {
            if (!IsFirstShapeSelected())
            {
                cropTimer.Stop();
            }
            /*
            for (int i = 2; i < shapes.Count; i++)
            {

                shapes[i].PictureFormat.CropTop = shapes[1].PictureFormat.CropTop;
                shapes[i].PictureFormat.CropLeft = shapes[1].PictureFormat.CropLeft;
                shapes[i].PictureFormat.CropRight = shapes[1].PictureFormat.CropRight;
                shapes[i].PictureFormat.CropBottom = shapes[1].PictureFormat.CropBottom;
            }
            */
        }

        private static bool VerifyIsShapeRangeValid(PowerPoint.ShapeRange shapeRange, bool handleError)
        {
            try
            {
                if (shapeRange.Count < 1)
                {
                    ThrowErrorCode(ErrorCodeForSelectionCountZero);
                }

                if (!IsPictureForSelection(shapeRange))
                {
                    ThrowErrorCode(ErrorCodeForSelectionNonPicture);
                }

                return true;
            }
            catch (Exception e)
            {
                if (handleError)
                {
                    ProcessErrorMessage(e);
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

        private static bool IsPictureForSelection(PowerPoint.ShapeRange shapeRange)
        {
            return (from PowerPoint.Shape shape in shapeRange select shape).All(IsPicture);
        }

        private static bool IsPicture(PowerPoint.Shape shape)
        {
            return shape.Type == Office.MsoShapeType.msoPicture ||
                   shape.Type == Office.MsoShapeType.msoLinkedPicture;
        }

        private static void ThrowErrorCode(int typeOfError)
        {
            throw new Exception(typeOfError.ToString(CultureInfo.InvariantCulture));
        }

        private static void IgnoreExceptionThrown() { }

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
                case ErrorCodeForSelectionNonPicture:
                    return ErrorMessageForSelectionNonPicture;
                default:
                    return ErrorMessageForUndefined;
            }
        }

        private static void ProcessErrorMessage(Exception e)
        {
            //This method prompts the error message to user. If it has an unrecognised error code,
            //an alternative message window with erro trace stack pops up and prompts the user to
            //send the trace stack to the developer team.
            var errMessage = GetErrorMessageForErrorCode(e.Message);
            if (!string.Equals(errMessage, ErrorMessageForUndefined, StringComparison.Ordinal))
            {
                MessageBox.Show(errMessage, MessageBoxTitle);
            }
            else
            {
                Views.ErrorDialogWrapper.ShowDialog(MessageBoxTitle, e.Message, e);
            }
        }

    }
}
