using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.CaptionsLab
{
    [ExportSupertipRibbonId(TextCollection.RemoveNotesTag)]
    class RemoveNotesSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.RemoveAllNotesButtonSupertip;
        }
    }
}
