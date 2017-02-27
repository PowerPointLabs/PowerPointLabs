using System;
using System.Drawing;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.CropLab;
using Office = Microsoft.Office.Core;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId("CropToSameButton")]
    class CropToSameActionHandler : CropLabActionHandler
    {
        private static readonly string ShapePicture = Path.GetTempPath() + @"\shape.png";
        private static readonly string FeatureName = "Crop To Same Dimensions";

        protected override void ExecuteAction(string ribbonId)
        {
            CropLabErrorHandler errorHandler = CropLabErrorHandler.InitializeErrorHandler(CropLabUIControl.GetSharedInstance());
            if (!VerifyIsSelectionValid(this.GetCurrentSelection()))
            {
                HandleInvalidSelectionError(CropLabErrorHandler.ErrorCodeSelectionIsInvalid, FeatureName, CropLabErrorHandler.SelectionTypePicture, 2, errorHandler);
                return;
            }
            ShapeRange shapeRange = this.GetCurrentSelection().ShapeRange;
            if (shapeRange.Count < 2)
            {
                HandleInvalidSelectionError(CropLabErrorHandler.ErrorCodeSelectionIsInvalid, FeatureName, CropLabErrorHandler.SelectionTypePicture, 2, errorHandler);
                return;
            }
            if (!IsPictureForSelection(shapeRange))
            {
                HandleErrorCodeIfRequired(CropLabErrorHandler.ErrorCodeSelectionMustBePicture, FeatureName, errorHandler);
                return;
            }
            var refShape = shapeRange[1];
            float refScaleWidth = PowerPointLabs.Utils.Graphics.GetScaleWidth(refShape);
            float refScaleHeight = PowerPointLabs.Utils.Graphics.GetScaleHeight(refShape);
            //MessageBox.Show(shapes[1].PictureFormat.CropTop.ToString() + " " + shapes[1].PictureFormat.CropBottom.ToString() + " " + shapes[1].Height.ToString() + " " + );
            //refShape.ScaleHeight(0.5F, Microsoft.Office.Core.MsoTriState.msoFalse);
            float epsilon = 0.001F;
            for (int i = 2; i <= shapeRange.Count; i++)
            {

                float scaleWidth = PowerPointLabs.Utils.Graphics.GetScaleWidth(shapeRange[i]);
                float scaleHeight = PowerPointLabs.Utils.Graphics.GetScaleHeight(shapeRange[i]);
                float heightToCrop = shapeRange[i].Height - refShape.Height;
                float widthToCrop = shapeRange[i].Width - refShape.Width;

                float cropTop = Math.Max(shapeRange[1].PictureFormat.CropTop, epsilon);
                float cropBottom = Math.Max(shapeRange[1].PictureFormat.CropBottom, epsilon);
                float cropLeft = Math.Max(shapeRange[1].PictureFormat.CropLeft, epsilon);
                float cropRight = Math.Max(shapeRange[1].PictureFormat.CropRight, epsilon);

                float refShapeCroppedHeight = cropTop + cropBottom;
                float refShapeCroppedWidth = cropLeft + cropRight;

                shapeRange[i].PictureFormat.CropTop = Math.Max(0, heightToCrop * cropTop / refShapeCroppedHeight / scaleHeight);
                shapeRange[i].PictureFormat.CropLeft = Math.Max(0, widthToCrop * cropLeft / refShapeCroppedWidth / scaleWidth);
                shapeRange[i].PictureFormat.CropRight = Math.Max(0, widthToCrop * cropRight / refShapeCroppedWidth / scaleWidth);
                shapeRange[i].PictureFormat.CropBottom = Math.Max(0, heightToCrop * cropBottom / refShapeCroppedHeight / scaleHeight);
            }
        }
    }
}
