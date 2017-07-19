using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.AgendaLab
{
    [ExportSupertipRibbonId(TextCollection.TextAgendaTag)]
    class TextAgendaSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.AgendaLabBulletPointSupertip;
        }
    }
}
