using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.AgendaLab
{
    [ExportLabelRibbonId(AgendaLabText.VisualAgendaTag)]
    class VisualAgendaLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return AgendaLabText.VisualAgendaButtonLabel;
        }
    }
}
