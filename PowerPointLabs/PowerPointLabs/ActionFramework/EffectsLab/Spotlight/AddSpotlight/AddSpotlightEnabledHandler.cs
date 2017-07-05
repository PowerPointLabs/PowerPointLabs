using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportEnabledRibbonId(TextCollection.AddSpotlightTag)]
    class AddSpotlightEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return IsSelectionAllShapeWithArea();
        }
    }
}