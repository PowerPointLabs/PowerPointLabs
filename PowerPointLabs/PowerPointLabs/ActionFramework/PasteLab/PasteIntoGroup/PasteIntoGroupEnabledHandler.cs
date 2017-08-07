using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.PasteLab
{
    [ExportEnabledRibbonId(PasteLabText.PasteIntoGroupTag)]
    class PasteIntoGroupEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            Selection currentSelection = this.GetCurrentSelection();
            return !GraphicsUtil.IsClipboardEmpty() &&
                ShapeUtil.IsSelectionMultipleOrGroup(currentSelection) &&
                !ShapeUtil.HasPlaceholderInSelection(currentSelection);
        }
    }
}