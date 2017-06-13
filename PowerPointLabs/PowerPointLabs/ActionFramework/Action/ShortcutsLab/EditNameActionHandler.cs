using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Views;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId(
        "EditNameMenuShape",
        "EditNameMenuLine",
        "EditNameMenuFreeform",
        "EditNameMenuPicture",
        "EditNameMenuGroup",
        "EditNameMenuInk",
        "EditNameMenuVideo",
        "EditNameMenuChart",
        "EditNameMenuTable",
        "EditNameMenuTableWhole",
        "EditNameMenuSmartArtBackground",
        "EditNameMenuSmartArtEditSmartArt",
        "EditNameMenuSmartArtEditText")]
    class EditNameActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            var selectedShape = this.GetCurrentSelection().ShapeRange[1];
            var editForm = new Form1(this.GetRibbonUi(), selectedShape.Name);
            editForm.ShowDialog();
        }
    }
}
