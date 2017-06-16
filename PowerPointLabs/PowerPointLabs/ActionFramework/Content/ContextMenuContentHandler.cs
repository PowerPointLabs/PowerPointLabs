using System.Collections.Generic;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Content
{
    [ExportContentRibbonId(
        TextCollection.MenuShape, 
        TextCollection.MenuLine,
        TextCollection.MenuFreeform,
        TextCollection.MenuPicture, 
        TextCollection.MenuSlide,
        TextCollection.MenuGroup,
        TextCollection.MenuInk,
        TextCollection.MenuVideo,
        TextCollection.MenuTextEdit,
        TextCollection.MenuChart,
        TextCollection.MenuTable,
        TextCollection.MenuTableCell,
        TextCollection.MenuSmartArt,
        TextCollection.MenuEditSmartArt,
        TextCollection.MenuEditSmartArtText)]

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
                    xmlString.Append(string.Format(TextCollection.DynamicMenuXmlImageButton, groupItem + ribbonId, groupItem));
                }
            }

            return string.Format(TextCollection.DynamicMenuXmlMenu, xmlString);
        }

        private List<ContextMenuGroup> GetContextMenuGroups(string ribbonId)
        {
            List<ContextMenuGroup> contextMenuGroups = new List<ContextMenuGroup>();
            ContextMenuGroup pasteLab = new ContextMenuGroup(TextCollection.PasteLabMenuLabel, new List<string>());
            ContextMenuGroup shortcuts = new ContextMenuGroup(TextCollection.ShortcutsLabMenuLabel, new List<string>());
            contextMenuGroups.Add(pasteLab);

            // All context menus will have these buttons
            pasteLab.Items.Add(TextCollection.PasteAtCursorPositionTag);
            pasteLab.Items.Add(TextCollection.PasteAtOriginalPositionTag);
            pasteLab.Items.Add(TextCollection.PasteToFillSlideTag);

            // Context menus other than slide will have these buttons
            if (ribbonId != TextCollection.MenuSlide)
            {
                // We only add shortcuts group if context menu is not for slide
                contextMenuGroups.Add(shortcuts);

                if (!ribbonId.Contains(TextCollection.MenuEditSmartArtBase) &&
                    ribbonId != TextCollection.MenuTextEdit &&
                    ribbonId != TextCollection.MenuTable)
                {
                    pasteLab.Items.Add(TextCollection.ReplaceWithClipboardTag);
                }

                shortcuts.Items.Add(TextCollection.EditNameTag);

                // Context menus other than picture will have these buttons
                if (ribbonId != TextCollection.MenuPicture)
                {
                    shortcuts.Items.Add(TextCollection.ConvertToPictureTag);
                }

                shortcuts.Items.Add(TextCollection.HideSelectedShapeTag);
                shortcuts.Items.Add(TextCollection.AddCustomShapeTag);

                // Context menu group will have these buttons
                if (ribbonId == TextCollection.MenuGroup)
                {
                    pasteLab.Items.Add(TextCollection.PasteIntoGroupTag);
                    shortcuts.Items.Add(TextCollection.AddIntoGroupTag);
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
