using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportLabelRibbonId(TextCollection.BlurRemainderMenuId)]
    class BlurRemainderLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.EffectsLabBlurRemainderButtonLabel;
        }
    }
}
