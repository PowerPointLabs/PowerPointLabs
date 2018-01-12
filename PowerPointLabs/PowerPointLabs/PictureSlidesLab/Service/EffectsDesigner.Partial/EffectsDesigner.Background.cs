using Microsoft.Office.Core;
using PowerPointLabs.PictureSlidesLab.Service.Effect;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.PictureSlidesLab.Service
{
    partial class EffectsDesigner
    {
        public PowerPoint.Shape ApplyBackgroundEffect()
        {
            var imageShape = AddPicture(Source.FullSizeImageFile ?? Source.ImageFile, EffectName.BackGround);
            imageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            var slideWidth = SlideWidth;
            var slideHeight = SlideHeight;
            FitToSlide.AutoFit(imageShape, slideWidth, slideHeight);

            CropPicture(imageShape);
            return imageShape;
        }
    }
}
