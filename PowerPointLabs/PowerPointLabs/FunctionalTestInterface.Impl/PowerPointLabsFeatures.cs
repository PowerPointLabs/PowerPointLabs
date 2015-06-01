using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
    }
}
