using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Util;
using PowerPointLabs.Models;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId("MergeIntoGroup")]
    class MergeIntoGroupActionHandler : BaseUtilActionHandler
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

            MiscFeatures.MergeIntoGroup.Execute(presentation, slide, selection);
        }
    }
}