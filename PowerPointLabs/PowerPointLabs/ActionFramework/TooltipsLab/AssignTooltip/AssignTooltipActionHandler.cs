using System;
using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.TooltipsLab;
using PowerPointLabs.Utils;
using PowerPointLabs.Views;

namespace PowerPointLabs.ActionFramework.TooltipsLab
{
    [ExportActionRibbonId(TooltipsLabText.AssignTooltipTag)]
    class AssignTooltipActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            Selection selection = this.GetCurrentSelection();
            PowerPointSlide currentSlide = this.GetCurrentSlide();

            if (currentSlide == null)
            {
                return;
            }

            if (!ShapeUtil.IsSelectionShape(selection))
            {
                return;
            }

            try
            {
                bool success = AssignTooltip.AddTriggerAnimation(currentSlide, selection);
                if (success)
                {
                    // Open the animation custom pane if it is not opened
                    if (!this.GetApplication().CommandBars.GetPressedMso("AnimationCustom"))
                    {
                        this.GetApplication().CommandBars.ExecuteMso("AnimationCustom");
                    }
                }
            }
            catch (Exception e)
            {
                ErrorDialogBox.ShowDialog("PowerPointLabs", e.Message, e);
            }
        }
    }
}
