﻿using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.CropLab
{
    [ExportLabelRibbonId(TextCollection.CropToAspectRatioTag)]
    class CropToAspectRatioLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.CropToAspectRatioLabel;
        }
    }
}
