using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ResizeLab
{
    [ExportLabelRibbonId(ResizeLabText.PaneTag)]
    class ResizeLabLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return ResizeLabText.RibbonMenuLabel;
        }
    }
}
