using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.ResizeLab
{
    [ExportLabelRibbonId(TextCollection.ResizeLabTag)]
    class ResizeLabLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.ResizeLabButtonLabel;
        }
    }
}
