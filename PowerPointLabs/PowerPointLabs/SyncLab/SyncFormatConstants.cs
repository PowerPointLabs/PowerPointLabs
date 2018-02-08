using System.Collections.Generic;
using System.Drawing;
using System.Windows.Documents;
using PowerPointLabs.SyncLab.Views;

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

        private static FormatTreeNode[] formatCategories = InitFormatCategories();

        public static FormatTreeNode[] FormatCategories
        {
            get
            {
                FormatTreeNode[] result = new FormatTreeNode[formatCategories.Length];
                for (int i = 0; i < result.Length; i++)
                {
                    result[i] = formatCategories[i].Clone();
                }
                return result;
            }
        }

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
        /// <Summary>
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

        private static FormatTreeNode[] InitFormatCategories()
        {
            FormatTreeNode[] formats =
                {
                    new FormatTreeNode(
                            "Text",
                            new FormatTreeNode("Font", new Format(typeof(FontFormat))),
                            new FormatTreeNode("Font Size", new Format(typeof(FontSizeFormat))),
                            new FormatTreeNode("Font Color", new Format(typeof(FontColorFormat))),
                            new FormatTreeNode("Style", new Format(typeof(FontStyleFormat)))),
                    new FormatTreeNode(
                            "Fill",
                            new FormatTreeNode("Color", new Format(typeof(FillFormat))),
                            new FormatTreeNode("Transparency", new Format(typeof(FillTransparencyFormat)))),
                    new FormatTreeNode(
                            "Line",
                            new FormatTreeNode("Width", new Format(typeof(LineWeightFormat))),
                            new FormatTreeNode("Compound Type", new Format(typeof(LineCompoundTypeFormat))),
                            new FormatTreeNode("Dash Type", new Format(typeof(LineDashTypeFormat))),
                            new FormatTreeNode("Arrow", new Format(typeof(LineArrowFormat))),
                            new FormatTreeNode("Color", new Format(typeof(LineFillFormat))),
                            new FormatTreeNode("Transparency", new Format(typeof(LineTransparencyFormat)))),
                    new FormatTreeNode(
                            "Size/Position",
                            new FormatTreeNode("Width", new Format(typeof(PositionWidthFormat))),
                            new FormatTreeNode("Height", new Format(typeof(PositionHeightFormat))),
                            new FormatTreeNode("X", new Format(typeof(PositionXFormat))),
                            new FormatTreeNode("Y", new Format(typeof(PositionYFormat))))
                };
            return formats;
        }
    }
}
