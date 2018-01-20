using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.HighlightLab
{
    [ExportEnabledRibbonId(HighlightLabText.HighlightTextTag)]
    class HighlightTextEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            try
            {
                if (this.GetAddIn().Application.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionText && 
                    this.GetAddIn().Application.ActiveWindow.Selection.TextRange2.TrimText().Length > 0)
                {
                    return this.GetRibbonUi().HighlightTextFragmentsEnabled;
                }
                else
                {
                    return false;
                }
            }
            catch (System.Runtime.InteropServices.COMException) // If this exception is caught, it means nothing has been selected yet
            {
                return false;
            }
        }
    }
}