using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.SyncLab
{
    [ExportLabelRibbonId(TextCollection1.SyncLabTag)]
    class SyncLabLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection1.SyncLabButtonLabel;
        }
    }
}
