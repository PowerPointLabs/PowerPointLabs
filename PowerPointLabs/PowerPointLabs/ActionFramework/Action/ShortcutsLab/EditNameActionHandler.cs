using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Views;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId(
        "EditNameMenuShape", "EditNameMenuLine", "EditNameMenuFreeform",
        "EditNameMenuPicture", "EditNameMenuGroup", "EditNameMenuInk",
        "EditNameMenuVideo", "EditNameMenuTextEdit", "EditNameMenuChart",
        "EditNameMenuTable", "EditNameMenuTableWhole",  "EditNameMenuSmartArtBackground",
        "EditNameMenuSmartArtEditSmartArt", "EditNameMenuSmartArtEditText")]
    class EditNameActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            var selection = this.GetCurrentSelection();
            var selectedShape = selection.ShapeRange[1];
            if (selection.HasChildShapeRange)
            {
                selectedShape = selection.ChildShapeRange[1];
            }
            
            var editForm = new Form1(this.GetRibbonUi(), selectedShape);
            editForm.ShowDialog();
        }
    }
}
