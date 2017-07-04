﻿using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportActionRibbonId(TextCollection.AddSpotlightTag)]
    class AddSpotlightActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            if (this.GetAddIn().Application.ActiveWindow.Selection.Type !=
                Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }

            this.StartNewUndoEntry();

            Spotlight.AddSpotlightEffect();
        }
    }
}
