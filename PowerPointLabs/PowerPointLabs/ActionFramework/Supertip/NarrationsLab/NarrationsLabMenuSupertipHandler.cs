﻿using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Supertip.NarrationsLab
{
    [ExportSupertipRibbonId("NarrationsLabMenu")]
    class NarrationsLabMenuSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.NarrationsLabMenuSupertip;
        }
    }
}
