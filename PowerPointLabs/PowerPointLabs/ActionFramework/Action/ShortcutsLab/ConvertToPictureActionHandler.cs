using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId(
        "ConvertToPictureMenuShape",
        "ConvertToPictureMenuLine",
        "ConvertToPictureMenuFreeform",
        "ConvertToPictureMenuGroup",
        "ConvertToPictureMenuInk",
        "ConvertToPictureMenuVideo",
        "ConvertToPictureMenuTextEdit",
        "ConvertToPictureMenuChart",
        "ConvertToPictureMenuTable",
        "ConvertToPictureMenuTableWhole",
        "ConvertToPictureMenuSmartArtBackground",
        "ConvertToPictureMenuSmartArtEditSmartArt",
        "ConvertToPictureMenuSmartArtEditText")]
    class ConvertToPictureActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            var selection = this.GetCurrentSelection();
            ConvertToPicture.Convert(selection);
        }
    }
}
