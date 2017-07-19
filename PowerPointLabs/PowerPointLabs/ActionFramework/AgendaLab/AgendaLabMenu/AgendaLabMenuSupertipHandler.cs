using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.AgendaLab
{
    [ExportSupertipRibbonId(TextCollection.AgendaLabMenuId)]
    class AgendaLabMenuSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.AgendaLabSupertip;
        }
    }
}
