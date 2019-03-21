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
    [ExportActionRibbonId(TooltipsLabText.ConvertToTooltipTag)]
    class ConvertToTooltipActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();
            PowerPointSlide currentSlide = this.GetCurrentSlide();
            Selection selection = this.GetCurrentSelection();

            if (currentSlide == null || !ShapeUtil.IsSelectionShape(selection))
            {
                return;
            }

            try
            {
                bool success = ConvertToTooltip.AddTriggerAnimation(currentSlide, selection);
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
