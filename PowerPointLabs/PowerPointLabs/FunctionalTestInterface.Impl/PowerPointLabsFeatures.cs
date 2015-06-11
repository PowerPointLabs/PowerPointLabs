using System;
using FunctionalTestInterface;
using PowerPointLabs.Models;

namespace PowerPointLabs.FunctionalTestInterface.Impl
{
    [Serializable]
    class PowerPointLabsFeatures : MarshalByRefObject, IPowerPointLabsFeatures
    {
        public void AutoCrop()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            CropToShape.Crop(selection);
        }

        public void AutoAnimate()
        {
            PowerPointLabs.AutoAnimate.AddAutoAnimation();
        }

        public void AnimateInSlide()
        {
            PowerPointLabs.AnimateInSlide.isHighlightBullets = false;
            PowerPointLabs.AnimateInSlide.AddAnimationInSlide();
        }

        public void AutoCaptions()
        {
            NotesToCaptions.EmbedCaptionsOnSelectedSlides();
        }

        public void Spotlight()
        {
            PowerPointLabs.Spotlight.AddSpotlightEffect();
        }

        public void FitToWidth()
        {
            var selectedShape = PowerPointCurrentPresentationInfo.CurrentSelection.ShapeRange[1];
            FitToSlide.FitToWidth(selectedShape);
        }

        public void FitToHeight()
        {
            var selectedShape = PowerPointCurrentPresentationInfo.CurrentSelection.ShapeRange[1];
            FitToSlide.FitToHeight(selectedShape);
        }

        public void ConvertToPic()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            ConvertToPicture.Convert(selection);
        }
    }
}
