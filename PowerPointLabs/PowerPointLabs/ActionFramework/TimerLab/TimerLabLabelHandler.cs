using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.TimerLab
{
    [ExportLabelRibbonId(TextCollection1.TimerLabTag)]
    class TimerLabLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TimerLabText.RibbonMenuLabel;
        }
    }
}
