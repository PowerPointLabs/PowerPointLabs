using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.TimerLab
{
    [ExportLabelRibbonId(TextCollection1.TimerLabTag)]
    class TimerLabLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection1.TimerLabButtonLabel;
        }
    }
}
