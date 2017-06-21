using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.TimerLab
{
    [ExportSupertipRibbonId(TextCollection.TimerLabTag)]
    class TimerLabSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.TimerLabMenuSupertip;
        }
    }
}
