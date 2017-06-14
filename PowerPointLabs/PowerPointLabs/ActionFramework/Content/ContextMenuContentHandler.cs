using System.Collections.Generic;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Content
{
    [ExportContentRibbonId("MenuShape", "MenuLine", "MenuFreeform", "MenuPicture", "MenuFrame", "MenuGroup",
                           "MenuInk", "MenuVideo", "MenuTextEdit", "MenuChart", "MenuTable", "MenuTableWhole",
                           "MenuSmartArtBackground", "MenuSmartArtEditSmartArt", "MenuSmartArtEditText")]

    class ContextMenuContentHandler : ContentHandler
    {
        protected override string GetContent(string ribbonId)
        {
            var contextMenuGroups = GetContextMenuGroups(ribbonId);
            var xmlString = new System.Text.StringBuilder();

            foreach (ContextMenuGroup group in contextMenuGroups)
            {
                string id = group.Name.Replace(" ", "");
                xmlString.Append(string.Format(TextCollection.DynamicMenuXmlTitleMenuSeparator, id, group.Name));

                foreach (string groupItem in group.Items)
                {
                    string item = groupItem.Replace(" ", "");
                    xmlString.Append(string.Format(TextCollection.DynamicMenuXmlImageButton, item + ribbonId));
                }
            }

            return string.Format(TextCollection.DynamicMenuXmlMenu, xmlString);
        }

        private List<ContextMenuGroup> GetContextMenuGroups(string ribbonId)
        {
            List<ContextMenuGroup> contextMenuGroups = new List<ContextMenuGroup>();
            ContextMenuGroup pasteLab = new ContextMenuGroup("Paste Lab", new List<string>());
            ContextMenuGroup shortcuts = new ContextMenuGroup("Shortcuts", new List<string>());
            contextMenuGroups.Add(pasteLab);

            // All context menus will have these buttons
            pasteLab.Items.Add(TextCollection.PasteLabText.PasteAtCursorPosition);
            pasteLab.Items.Add(TextCollection.PasteLabText.PasteAtOriginalPosition);
            pasteLab.Items.Add(TextCollection.PasteLabText.PasteToFillSlide);

            // Context menus other than slide will have these buttons
            if (ribbonId != "MenuFrame")
            {
                // We only add shortcuts group if context menu is not for slide
                contextMenuGroups.Add(shortcuts);

                if (!ribbonId.Contains("MenuSmartArtEdit") && ribbonId != "MenuTextEdit" && ribbonId != "MenuTable")
                {
                    pasteLab.Items.Add(TextCollection.PasteLabText.ReplaceWithClipboard);
                }

                shortcuts.Items.Add(TextCollection.EditNameShapeLabel);

                // Context menus other than picture will have these buttons
                if (ribbonId != "MenuPicture")
                {
                    shortcuts.Items.Add(TextCollection.ConvertToPictureShapeLabel);
                }

                shortcuts.Items.Add(TextCollection.HideSelectedShapeLabel);
                shortcuts.Items.Add(TextCollection.AddCustomShapeShapeLabel);

                // Context menu group will have these buttons
                if (ribbonId == "MenuGroup")
                {
                    pasteLab.Items.Add(TextCollection.PasteLabText.PasteIntoGroup);
                    shortcuts.Items.Add(TextCollection.AddIntoGroup);
                }
            }
            return contextMenuGroups;
        }

        public class ContextMenuGroup
        {
            public string Name { get; private set; }
            public List<string> Items { get; private set; }

            public ContextMenuGroup(string groupName, List<string> groupItems)
            {
                Name = groupName;
                Items = groupItems;
            }
        }
    }
}
