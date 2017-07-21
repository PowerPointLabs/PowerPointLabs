using System.Collections.Generic;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.ShortcutsLab
{
    [ExportContentRibbonId(
        TextCollection1.MenuShape, TextCollection1.MenuLine, TextCollection1.MenuFreeform,
        TextCollection1.MenuPicture, TextCollection1.MenuSlide, TextCollection1.MenuGroup,
        TextCollection1.MenuInk, TextCollection1.MenuVideo, TextCollection1.MenuTextEdit,
        TextCollection1.MenuChart, TextCollection1.MenuTable, TextCollection1.MenuTableCell,
        TextCollection1.MenuSmartArt, TextCollection1.MenuEditSmartArt, TextCollection1.MenuEditSmartArtText)]

    class ContextMenuContentHandler : ContentHandler
    {
        protected override string GetContent(string ribbonId)
        {
            var contextMenuGroups = GetContextMenuGroups(ribbonId);
            var xmlString = new System.Text.StringBuilder();

            foreach (ContextMenuGroup group in contextMenuGroups)
            {
                string id = group.Name.Replace(" ", "");
                xmlString.Append(string.Format(TextCollection1.DynamicMenuXmlTitleMenuSeparator,
                    id + TextCollection1.MenuSeparator, group.Name));

                foreach (string groupItem in group.Items)
                {
                    xmlString.Append(string.Format(TextCollection1.DynamicMenuXmlImageButton, groupItem + ribbonId, groupItem));
                }
            }

            return string.Format(TextCollection1.DynamicMenuXmlMenu, xmlString);
        }

        private List<ContextMenuGroup> GetContextMenuGroups(string ribbonId)
        {
            List<ContextMenuGroup> contextMenuGroups = new List<ContextMenuGroup>();
            ContextMenuGroup pasteLab = new ContextMenuGroup(TextCollection1.PasteLabMenuLabel, new List<string>());
            ContextMenuGroup shortcuts = new ContextMenuGroup(TextCollection1.ShortcutsLabMenuLabel, new List<string>());
            contextMenuGroups.Add(pasteLab);

            // All context menus will have these buttons
            pasteLab.Items.Add(TextCollection1.PasteAtCursorPositionTag);
            pasteLab.Items.Add(TextCollection1.PasteAtOriginalPositionTag);
            pasteLab.Items.Add(TextCollection1.PasteToFillSlideTag);

            // Context menus other than slide will have these buttons
            if (ribbonId != TextCollection1.MenuSlide)
            {
                // We only add shortcuts group if context menu is not for slide
                contextMenuGroups.Add(shortcuts);

                if (!ribbonId.Contains(TextCollection1.MenuEditSmartArtBase) &&
                    ribbonId != TextCollection1.MenuTextEdit &&
                    ribbonId != TextCollection1.MenuTable)
                {
                    pasteLab.Items.Add(TextCollection1.ReplaceWithClipboardTag);
                }

                shortcuts.Items.Add(TextCollection1.EditNameTag);

                // Context menus other than picture will have these buttons
                if (ribbonId != TextCollection1.MenuPicture)
                {
                    shortcuts.Items.Add(TextCollection1.ConvertToPictureTag);
                }

                shortcuts.Items.Add(TextCollection1.HideShapeTag);
                shortcuts.Items.Add(TextCollection1.AddCustomShapeTag);

                // Context menu group will have these buttons
                if (ribbonId == TextCollection1.MenuGroup)
                {
                    pasteLab.Items.Add(TextCollection1.PasteIntoGroupTag);
                    shortcuts.Items.Add(TextCollection1.AddIntoGroupTag);
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
