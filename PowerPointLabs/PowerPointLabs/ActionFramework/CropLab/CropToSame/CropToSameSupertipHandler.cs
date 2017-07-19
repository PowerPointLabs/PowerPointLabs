﻿using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.CropLab
{
    [ExportSupertipRibbonId(TextCollection.CropToSameDimensionsTag)]
    class CropToSameSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.CropToSameButtonSupertip;
        }
    }
}
