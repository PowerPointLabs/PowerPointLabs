using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.SyncLab
{
    [ExportLabelRibbonId(SyncLabText.PaneTag)]
    class SyncLabLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return SyncLabText.RibbonMenuLabel;
        }
    }
}
