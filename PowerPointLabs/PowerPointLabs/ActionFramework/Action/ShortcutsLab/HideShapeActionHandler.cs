﻿using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId(
        "HideShapeMenuShape", "HideShapeMenuLine", "HideShapeMenuFreeform",
        "HideShapeMenuPicture", "HideShapeMenuGroup", "HideShapeMenuInk",
        "HideShapeMenuVideo", "HideShapeMenuTextEdit", "HideShapeMenuChart",
        "HideShapeMenuTable", "HideShapeMenuTableWhole", "HideShapeMenuSmartArtBackground",
        "HideShapeMenuSmartArtEditSmartArt", "HideShapeMenuSmartArtEditText")]
    class HideActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            var selectedShapes = this.GetCurrentSelection().ShapeRange;
            selectedShapes.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
        }
    }
}
