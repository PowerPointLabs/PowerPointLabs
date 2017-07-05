using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportLabelRibbonId(TextCollection.MakeTransparentTag)]
    class MakeTransparentLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.EffectsLabMakeTransparentButtonLabel;
        }
    }
}
