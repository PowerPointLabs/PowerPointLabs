
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.TooltipsLab
{
    [ExportEnabledRibbonId(TooltipsLabText.CreateTooltipTag)]
    class CreateTooltipEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            PowerPointSlide currentSlide = this.GetCurrentSlide();
            return currentSlide != null;
        }
    }
}