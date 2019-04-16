using System;

using Microsoft.Office.Core;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.PictureSlidesLab.Service.Effect;
using PowerPointLabs.PictureSlidesLab.Util;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.PictureSlidesLab.Service
{
    partial class EffectsDesigner
    {
        public const int BlurDegreeForFrostedGlassEffect = 95;

        private PowerPoint.Shape AddPicture(string imageFile, EffectName effectName)
        {
            PowerPoint.Shape imageShape = Shapes.AddPicture(imageFile,
                MsoTriState.msoFalse, MsoTriState.msoTrue, 0,
                0);
            ChangeName(imageShape, effectName);
            return imageShape;
        }

        private int CalculateTextBoxMargin(int fontSizeToIncrease)
        {
            if (fontSizeToIncrease == -1 || fontSizeToIncrease == 0)
            {
                return 10;
            }

            return (int) (fontSizeToIncrease * 0.25 + 10);
        }

        private void CropPicture(PowerPoint.Shape picShape)
        {
            try
            {
                if (picShape.Left < 0)
                {
                    picShape.PictureFormat.Crop.ShapeLeft = 0;
                }
                if (picShape.Top < 0)
                {
                    picShape.PictureFormat.Crop.ShapeTop = 0;
                }
                if (picShape.Left + picShape.Width > SlideWidth)
                {
                    picShape.PictureFormat.Crop.ShapeWidth = SlideWidth - picShape.Left;
                }
                if (picShape.Top + picShape.Height > SlideHeight)
                {
                    picShape.PictureFormat.Crop.ShapeHeight = SlideHeight - picShape.Top;
                }
            }
            catch (Exception e)
            {
                // some kind of picture cannot be cropped
                Logger.LogException(e, "CropPicture");
            }
        }

        private void CropPicture(PowerPoint.Shape picShape, float targetLeft, float targetTop, float targetWidth, float targetHeight)
        {
            try
            {
                picShape.PictureFormat.Crop.ShapeLeft = targetLeft;
                picShape.PictureFormat.Crop.ShapeTop = targetTop;
                picShape.PictureFormat.Crop.ShapeWidth = targetWidth;
                picShape.PictureFormat.Crop.ShapeHeight = targetHeight;
            }
            catch (Exception e)
            {
                // some kind of picture cannot be cropped
                Logger.LogException(e, "CropPicture");
            }
        }

        /// <summary>
        /// change the shape name, so that they can be managed (eg delete) by name easily
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="effectName"></param>
        private static void ChangeName(PowerPoint.Shape shape, EffectName effectName)
        {
            ShapeUtil.ChangeName(shape, effectName, ShapeNamePrefix);
        }

        private static void AddTag(PowerPoint.Shape shape, string tagName, String value)
        {
            ShapeUtil.AddTag(shape, tagName, value);
        }
    }
}
