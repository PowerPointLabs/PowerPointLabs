using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.AgendaLab
{
    [ExportLabelRibbonId(AgendaLabText.RemoveAgendaTag)]
    class RemoveAgendaLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return AgendaLabText.RemoveAgendaButtonLabel;
        }
    }
}
