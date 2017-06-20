using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.ResizeLab
{
    [ExportSupertipRibbonId(TextCollection.ResizeLabTag)]
    class ResizeLabSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.ResizeLabMenuSupertip;
        }
    }
}
