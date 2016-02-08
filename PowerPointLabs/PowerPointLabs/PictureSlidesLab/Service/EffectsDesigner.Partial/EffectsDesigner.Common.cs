using System;
using Microsoft.Office.Core;
using PowerPointLabs.PictureSlidesLab.Service.Effect;
using PowerPointLabs.PictureSlidesLab.Util;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.PictureSlidesLab.Service
{
    partial class EffectsDesigner
    {
        private PowerPoint.Shape AddPicture(string imageFile, EffectName effectName)
        {
            var imageShape = Shapes.AddPicture(imageFile,
                MsoTriState.msoFalse, MsoTriState.msoTrue, 0,
                0);
            ChangeName(imageShape, effectName);
            return imageShape;
        }

        private void CropPicture(PowerPoint.Shape picShape)
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
