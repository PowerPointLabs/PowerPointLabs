using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.AgendaLab
{
    [ExportSupertipRibbonId(AgendaLabText.VisualAgendaTag)]
    class VisualAgendaSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return AgendaLabText.VisualAgendaSupertip;
        }
    }
}
