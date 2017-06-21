using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.PasteLab
{
    [ExportSupertipRibbonId(TextCollection.ReplaceWithClipboardTag)]
    class ReplaceWithClipboardButtonSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.ReplaceWithClipboardSupertip;
        }
    }
}
