using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.TooltipsLab;

namespace PowerPointLabs.ActionFramework.TooltipsLab
{
    [ExportActionRibbonId(TooltipsLabText.AddTextboxTag)]
    class AddTextboxActionHandler : ActionHandler
    {
        public static Selection GetNewSelection(Shape shape1, Shape shape2)
        {
            Globals.ThisAddIn.Application.ActiveWindow.Selection.Unselect();
            shape1.Select();
            shape2.Select(Microsoft.Office.Core.MsoTriState.msoFalse);
            return Globals.ThisAddIn.Application.ActiveWindow.Selection;
        }

        protected override void ExecuteAction(string ribbonId)
        {
            Selection selection = this.GetCurrentSelection();
            PowerPointSlide currentSlide = this.GetCurrentSlide();
            this.StartNewUndoEntry();

            AddTextbox.AddTextboxToCallout(currentSlide, selection);
        }
    }
}
