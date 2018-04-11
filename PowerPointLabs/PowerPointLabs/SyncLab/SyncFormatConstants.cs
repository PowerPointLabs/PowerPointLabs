using System.Collections.Generic;
using System.Drawing;
using PowerPointLabs.SyncLab.Views;
using Font = System.Drawing.Font;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    public class SyncFormatConstants
    {

        public static readonly Size DisplayImageSize = new Size(30, 30);

        public static readonly string DisplayFontString = "Text";
        public static readonly int DisplayImageFontSize = 12;
        public static readonly Font DisplayImageFont = new Font("Arial", DisplayImageFontSize);

        public static readonly int ColorBlack = 0;
        public static readonly int DisplayLineWeight = 5;

        public static FormatTreeNode[] FormatCategories => CreateFormatCategories();

        public static List<Format> Formats
        {
            get
            {
                List<Format> list = new List<Format>();
                list.AddRange(GetFormatsFromFormatTreeNode(FormatCategories));
                return list;
            }
        }
        
        /// <Summary>
        /// Collect all format objects from an array of FormatTreeNodes
        /// </Summary>
        /// <param name="nodes"></param>
        /// <returns>Collected formats</returns>
        private static Format[] GetFormatsFromFormatTreeNode(FormatTreeNode[] nodes)
        {
            List<Format> list = new List<Format>();
            foreach (FormatTreeNode node in nodes)
            {
                if (node.IsFormatNode)
                {
                    list.Add(node.Format);
                }
                else if (node.ChildrenNodes != null)
                {
                    list.AddRange(GetFormatsFromFormatTreeNode(node.ChildrenNodes));
                }
            }

            return list.ToArray();
        }

        private static FormatTreeNode[] CreateFormatCategories()
        {
            FormatTreeNode[] formats =
                {
                    new FormatTreeNode(
                            "Text",
                            new FormatTreeNode("Font", new FontFormat()),
                            new FormatTreeNode("Font Size", new FontSizeFormat()),
                            new FormatTreeNode("Font Color", new FontColorFormat()),
                            new FormatTreeNode("Style", new FontStyleFormat())),
                    new FormatTreeNode(
                            "Fill",
                            new FormatTreeNode("Color", new FillFormat()),
                            new FormatTreeNode("Transparency", new FillTransparencyFormat())),
                    new FormatTreeNode(
                            "Line",
                            new FormatTreeNode("Width", new LineWeightFormat()),
                            new FormatTreeNode("Compound Type", new LineCompoundTypeFormat()),
                            new FormatTreeNode("Dash Type", new LineDashTypeFormat()),
                            new FormatTreeNode("Arrow", new LineArrowFormat()),
                            new FormatTreeNode("Color", new LineFillFormat()),
                            new FormatTreeNode("Transparency", new LineTransparencyFormat())),
                    new FormatTreeNode(
                            "Visual Effects",
                            new FormatTreeNode("Artistic Effect", new PictureEffectsFormat()),
                            new FormatTreeNode("Glow", 
                                new FormatTreeNode("Color", new GlowColorFormat()),
                                new FormatTreeNode("Size", new GlowSizeFormat()),
                                new FormatTreeNode("Transparency", new GlowTransparencyFormat())),
                            new FormatTreeNode("Reflection", new ReflectionEffectFormat())),
                    new FormatTreeNode(
                            "Size/Position",
                            new FormatTreeNode("Width", new PositionWidthFormat()),
                            new FormatTreeNode("Height", new PositionHeightFormat()),
                            new FormatTreeNode("X", new PositionXFormat()),
                            new FormatTreeNode("Y", new PositionYFormat()))
                };
            return formats;
        }
    }
}
