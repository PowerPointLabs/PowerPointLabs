using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.AgendaLab;

namespace PowerPointLabs.ActionFramework.AgendaLab
{
    [ExportActionRibbonId(TextCollection.UpdateAgendaTag)]
    class UpdateAgendaActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            AgendaLabMain.SynchroniseAgenda();
        }
    }
}
