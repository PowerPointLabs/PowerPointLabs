using System.Drawing;
using ImageProcessor;
using ImageProcessor.Imaging.Filters;
using PowerPointLabs.PictureSlidesLab.Service.Effect;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.PictureSlidesLab.Service
{
    partial class EffectsDesigner
    {
        public PowerPoint.Shape ApplySpecialEffectEffect(IMatrixFilter effectFilter, bool isActualSize)
        {
            Source.SpecialEffectImageFile = SpecialEffectImage(effectFilter, Source.FullSizeImageFile ?? Source.ImageFile, isActualSize);
            var specialEffectImageShape = AddPicture(Source.SpecialEffectImageFile, EffectName.SpecialEffect);
            var slideWidth = SlideWidth;
            var slideHeight = SlideHeight;
            FitToSlide.AutoFit(specialEffectImageShape, slideWidth, slideHeight);
            CropPicture(specialEffectImageShape);
            return specialEffectImageShape;
        }

        private static string SpecialEffectImage(IMatrixFilter effectFilter, string imageFilePath,
            bool isActualSize)
        {
            var specialEffectImageFile = Util.TempPath.GetPath("fullsize_specialeffect");
            using (var imageFactory = new ImageFactory())
            {
                var image = imageFactory
                        .Load(imageFilePath)
                        .Image;
                var ratio = (float)image.Width / image.Height;
                if (isActualSize)
                {
                    image = imageFactory
                        .Resize(new Size((int)(768 * ratio), 768))
                        .Filter(effectFilter)
                        .Image;
                }
                else
                {
                    image = imageFactory
                        .Resize(new Size((int)(300 * ratio), 300))
                        .Filter(effectFilter)
                        .Image;
                }
                image.Save(specialEffectImageFile);
            }
            return specialEffectImageFile;
        }
    }
}
