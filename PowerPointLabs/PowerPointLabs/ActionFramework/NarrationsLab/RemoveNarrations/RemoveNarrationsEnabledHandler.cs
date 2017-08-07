using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.NarrationsLab
{
    [ExportEnabledRibbonId(NarrationsLabText.RemoveNarrationsTag)]
    class RemoveNarrationsEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return this.GetRibbonUi().RemoveAudioEnabled;
        }
    }
}