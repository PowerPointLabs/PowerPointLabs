using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Handlers;
using PowerPointLabs.Models;
using PowerPointLabs.ShortcutsLab;

namespace PowerPointLabs.ActionFramework.ShortcutsLab
{
    [ExportActionRibbonId(TextCollection1.AddIntoGroupTag)]
    class AddIntoGroupActionHandler : BaseUtilActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            PowerPointPresentation presentation = this.GetCurrentPresentation();
            PowerPointSlide slide = this.GetCurrentSlide();
            Selection selection = this.GetCurrentSelection();

            if (!IsSelectionShapes(selection) || selection.ShapeRange.Count < 2)
            {
                MessageBox.Show("Please select more than one shape.", "Error");
                return;
            }
            
            ShapeRange result = AddIntoGroup.Execute(presentation, slide, selection);
            result.Select();
        }
    }
}