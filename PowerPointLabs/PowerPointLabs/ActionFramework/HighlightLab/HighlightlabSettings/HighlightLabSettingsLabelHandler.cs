﻿using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label.HighlightLab
{
    [ExportLabelRibbonId(TextCollection.HighlightLabSettingsTag)]
    class HighlightLabSettingsLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.HighlightLabSettingsButtonLabel;
        }
    }
}
