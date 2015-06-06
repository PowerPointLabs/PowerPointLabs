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
    }
}
