using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportLabelRibbonId(EffectsLabText.RecolorRemainderMenuId)]
    class RecolorRemainderLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return EffectsLabText.RecolorRemainderButtonLabel;
        }
    }
}
