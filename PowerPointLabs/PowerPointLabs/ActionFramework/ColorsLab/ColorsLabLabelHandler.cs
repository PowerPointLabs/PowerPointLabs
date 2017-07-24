using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ColorsLab
{
    [ExportLabelRibbonId(ColorsLabText.PaneTag)]
    class ColorsLabLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return ColorsLabText.RibbonMenuLabel;
        }
    }
}
