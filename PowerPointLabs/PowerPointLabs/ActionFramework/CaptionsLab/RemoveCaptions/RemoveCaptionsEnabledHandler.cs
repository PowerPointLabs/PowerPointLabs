using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.CaptionsLab
{
    [ExportEnabledRibbonId(CaptionsLabText.RemoveCaptionsTag)]
    class RemoveCaptionsEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return this.GetRibbonUi().RemoveCaptionsEnabled;
        }
    }
}