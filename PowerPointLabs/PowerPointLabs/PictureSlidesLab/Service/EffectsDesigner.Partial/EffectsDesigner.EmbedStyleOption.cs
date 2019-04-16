using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows;

using Microsoft.Office.Core;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service.Effect;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.PictureSlidesLab.Service
{
    partial class EffectsDesigner
    {
        /// <summary>
        /// embed style information into the image shapes,
        /// and then return a list of shape in which
        /// index-0 is the original image
        /// index-1 is the cropped image
        /// </summary>
        /// <param name="originalImageFile"></param>
        /// <param name="croppedImageFile"></param>
        /// <param name="imageContext"></param>
        /// <param name="imageSource"></param>
        /// <param name="rect"></param>
        /// <param name="opt"></param>
        /// <returns></returns>
        public List<PowerPoint.Shape> EmbedStyleOptionsInformation(string originalImageFile, string croppedImageFile,
            string imageContext, string imageSource, Rect rect, StyleOption opt)
        {
            if (originalImageFile == null)
            {
                return new List<PowerPoint.Shape>();
            }

            PowerPoint.Shape originalImage = AddPicture(originalImageFile, EffectName.Original_DO_NOT_REMOVE);
            float slideWidth = SlideWidth;
            float slideHeight = SlideHeight;
            FitToSlide.AutoFit(originalImage, slideWidth, slideHeight);
            originalImage.Visible = MsoTriState.msoFalse;

            PowerPoint.Shape croppedImage = AddPicture(croppedImageFile, EffectName.Cropped_DO_NOT_REMOVE);
            FitToSlide.AutoFit(croppedImage, slideWidth, slideHeight);
            croppedImage.Visible = MsoTriState.msoFalse;

            List<PowerPoint.Shape> result = new List<PowerPoint.Shape>();
            result.Add(originalImage);
            result.Add(croppedImage);

            // store source image info
            AddTag(originalImage, Tag.ReloadOriginImg, originalImageFile);
            AddTag(originalImage, Tag.ReloadCroppedImg, croppedImageFile);
            AddTag(originalImage, Tag.ReloadImgContext, imageContext);
            AddTag(originalImage, Tag.ReloadImgSource, imageSource);
            AddTag(originalImage, Tag.ReloadRectX, rect.X.ToString(CultureInfo.InvariantCulture));
            AddTag(originalImage, Tag.ReloadRectY, rect.Y.ToString(CultureInfo.InvariantCulture));
            AddTag(originalImage, Tag.ReloadRectWidth, rect.Width.ToString(CultureInfo.InvariantCulture));
            AddTag(originalImage, Tag.ReloadRectHeight, rect.Height.ToString(CultureInfo.InvariantCulture));

            // store style info
            Type type = opt.GetType();
            System.Reflection.PropertyInfo[] props = type.GetProperties();
            foreach (System.Reflection.PropertyInfo propertyInfo in props)
            {
                try
                {
                    AddTag(originalImage, Tag.ReloadPrefix + propertyInfo.Name,
                        propertyInfo.GetValue(opt, null).ToString());
                }
                catch (Exception e)
                {
                    Logger.LogException(e, "EmbedStyleOptionsInformation");
                }
            }
            return result;
        }
    }
}
