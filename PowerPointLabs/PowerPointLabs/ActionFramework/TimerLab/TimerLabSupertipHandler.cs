using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.TimerLab
{
    [ExportSupertipRibbonId(TimerLabText.PaneTag)]
    class TimerLabSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TimerLabText.RibbonMenuSupertip;
        }
    }
}
