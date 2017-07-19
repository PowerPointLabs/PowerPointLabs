﻿using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.HighlightLab
{
    [ExportEnabledRibbonId(TextCollection.HighlightPointsTag)]
    class HighlightPointsEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return this.GetRibbonUi().HighlightBulletsEnabled;
        }
    }
}