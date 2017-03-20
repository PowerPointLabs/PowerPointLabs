using System.Drawing;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    public class SyncFormatConstants
    {

        public static readonly Size DisplayImageSize = new Size(30, 30);

        public static readonly int DisplayImageFontSize = DisplayImageSize.Height;
        public static readonly Font DisplayImageFont = new Font("Arial", DisplayImageFontSize);

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

        private static FormatTreeNode[] InitFormatCategories()
        {
            FormatTreeNode[] formats = new FormatTreeNode[]
                {
                    new FormatTreeNode(
                            "Text",
                            new FormatTreeNode("Font", new Format(typeof(FontFormat))),
                            new FormatTreeNode("Font Size", new Format(typeof(FontSizeFormat))),
                            new FormatTreeNode("Font Color", new Format(typeof(FontColorFormat))),
                            new FormatTreeNode("Style", new Format(typeof(FontStyleFormat)))
                            /*new FormatTreeNode("Shadow", new Format(typeof(FillFormat))),
                            new FormatTreeNode("Strikethrough", new Format(typeof(FillFormat))),
                            new FormatTreeNode("Character Spacing", new Format(typeof(FillFormat))),
                            new FormatTreeNode("Line Spacing", new Format(typeof(FillFormat))),
                            new FormatTreeNode("Alignment", new Format(typeof(FillFormat)))*/
                        ),
                    new FormatTreeNode(
                            "Fill",
                            new FormatTreeNode("Fill", new Format(typeof(FillFormat)))
                        ),
                    new FormatTreeNode(
                            "Line",
                            //missing dash style
                            new FormatTreeNode("Arrow", new Format(typeof(LineArrowFormat))),
                            new FormatTreeNode("Weight", new Format(typeof(LineWeightFormat))),
                            new FormatTreeNode("Compound Type", new Format(typeof(LineCompoundTypeFormat))),
                            new FormatTreeNode("Fill", new Format(typeof(LineFillFormat)))
                        ),
                    //not easy
                    /*new FormatTreeNode(
                            "Effect",
                            new FormatTreeNode("Shadow", new Format(typeof(FillFormat))),
                            new FormatTreeNode("Reflection", new Format(typeof(FillFormat))),
                            new FormatTreeNode("Glow", new Format(typeof(FillFormat))),
                            new FormatTreeNode("Soft Edge", new Format(typeof(FillFormat))),
                            new FormatTreeNode("Bevel", new Format(typeof(FillFormat))),
                            new FormatTreeNode("3D Rotation", new Format(typeof(FillFormat)))
                        ),*/
                    new FormatTreeNode(
                            "Size/Position",
                            new FormatTreeNode("Width", new Format(typeof(PositionWidthFormat))),
                            new FormatTreeNode("Height", new Format(typeof(PositionHeightFormat))),
                            new FormatTreeNode("X", new Format(typeof(PositionXFormat))),
                            new FormatTreeNode("Y", new Format(typeof(PositionYFormat)))
                        )
                };
            return formats;
        }
    }
}
