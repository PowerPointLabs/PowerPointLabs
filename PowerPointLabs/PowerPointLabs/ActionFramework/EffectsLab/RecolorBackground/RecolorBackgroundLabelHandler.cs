using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportLabelRibbonId(TextCollection.RecolorBackgroundMenuId)]
    class RecolorBackgroundLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.EffectsLabRecolorBackgroundButtonLabel;
        }
    }
}
