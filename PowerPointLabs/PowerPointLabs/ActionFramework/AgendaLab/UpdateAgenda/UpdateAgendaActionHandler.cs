using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.AgendaLab;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.AgendaLab
{
    [ExportActionRibbonId(AgendaLabText.UpdateAgendaTag)]
    class UpdateAgendaActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            AgendaLabMain.SynchronizeAgenda();
        }
    }
}
