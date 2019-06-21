using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Models;
using PowerPointLabs.ShortcutsLab;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Windows;

namespace PowerPointLabs.ActionFramework.ShortcutsLab
{
    [ExportActionRibbonId(ShortcutsLabText.AddIntoGroupTag)]
    class AddIntoGroupActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            PowerPointPresentation presentation = this.GetCurrentPresentation();
            PowerPointSlide slide = this.GetCurrentSlide();
            Selection selection = this.GetCurrentSelection();

            if (!ShapeUtil.IsSelectionShape(selection) || selection.ShapeRange.Count < 2)
            {
                MessageBoxUtil.Show(TextCollection.ShortcutsLabText.AddIntoGroupActionHandlerReminderText, TextCollection.CommonText.ErrorTitle);
                return;
            }
            
            ShapeRange result = AddIntoGroup.Execute(presentation, slide, selection);
            result.Select();
        }
    }
}