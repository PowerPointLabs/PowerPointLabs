using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.Enabled.PasteLab
{
    [ExportEnabledRibbonId(
        "EditNameMenuShape", "EditNameMenuLine", "EditNameMenuFreeform",
        "EditNameMenuPicture", "EditNameMenuGroup", "EditNameMenuInk",
        "EditNameMenuVideo", "EditNameMenuTextEdit", "EditNameMenuChart",
        "EditNameMenuTable", "EditNameMenuTableWhole", "EditNameMenuSmartArtBackground",
        "EditNameMenuSmartArtEditSmartArt", "EditNameMenuSmartArtEditText")]
    class EditNameEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return IsSelectionSingleShape();
        }
    }
}