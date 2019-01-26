using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;
using PowerPointLabs.TimerLab;

namespace PowerPointLabs.ActionFramework.TimerLab
{
    [ExportEnabledRibbonId(TimerLabText.PaneTag)]
    class TimerLabEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return PowerPointLabs.TimerLab.TimerLab.IsTimerEnabled;
        }
    }
}
