﻿using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Supertip.AnimationLab
{
    [ExportSupertipRibbonId("AnimationLabSettingsButton")]
    class AnimationLabSettingsSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.AnimationLabSettingsSupertip;
        }
    }
}
