﻿using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.Enabled.PasteLab
{
    [ExportEnabledRibbonId(
        "PasteAtOriginalPosition",
        "PasteAtOriginalPositionShape",
        "PasteAtOriginalPositionFreeform",
        "PasteAtOriginalPositionPicture",
        "PasteAtOriginalPositionGroup")]
    class PasteAtOriginalPositionEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return !Graphics.IsClipboardEmpty();
        }
    }
}