using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class SyncFormatConstants
    {

        public static FormatTreeNode[] FormatCategories
        {
            get
            {
                return new FormatTreeNode[]
                {
                    new FormatTreeNode(
                            "Text",
                            new FormatTreeNode("Font", new Format(typeof(FillFormat))),
                            new FormatTreeNode("Font Size", new Format(typeof(FillFormat))),
                            new FormatTreeNode("Font Color", new Format(typeof(FillFormat))),
                            new FormatTreeNode("Bold", new Format(typeof(FillFormat))),
                            new FormatTreeNode("Italics", new Format(typeof(FillFormat))),
                            new FormatTreeNode("Underline", new Format(typeof(FillFormat))),
                            new FormatTreeNode("Shadow", new Format(typeof(FillFormat))),
                            new FormatTreeNode("Strikethrough", new Format(typeof(FillFormat))),
                            new FormatTreeNode("Character Spacing", new Format(typeof(FillFormat))),
                            new FormatTreeNode("Line Spacing", new Format(typeof(FillFormat))),
                            new FormatTreeNode("Alignment", new Format(typeof(FillFormat)))
                        ),
                    new FormatTreeNode(
                            "Fill",
                            new FormatTreeNode("Fill", new Format(typeof(FillFormat)))
                        ),
                    new FormatTreeNode(
                            "Line",
                            new FormatTreeNode("Start Arrow", new Format(typeof(FillFormat))),
                            new FormatTreeNode("End Arrow", new Format(typeof(FillFormat))),
                            new FormatTreeNode("Weight", new Format(typeof(FillFormat))),
                            new FormatTreeNode("Style", new Format(typeof(FillFormat))),
                            new FormatTreeNode("Fill", new Format(typeof(FillFormat)))
                        ),
                    new FormatTreeNode(
                            "Effect",
                            new FormatTreeNode("Shadow", new Format(typeof(FillFormat))),
                            new FormatTreeNode("Reflection", new Format(typeof(FillFormat))),
                            new FormatTreeNode("Glow", new Format(typeof(FillFormat))),
                            new FormatTreeNode("Soft Edge", new Format(typeof(FillFormat))),
                            new FormatTreeNode("Bevel", new Format(typeof(FillFormat))),
                            new FormatTreeNode("3D Rotation", new Format(typeof(FillFormat)))
                        ),
                    new FormatTreeNode(
                            "Size/Position",
                            new FormatTreeNode("Width", new Format(typeof(FillFormat))),
                            new FormatTreeNode("Height", new Format(typeof(FillFormat))),
                            new FormatTreeNode("X", new Format(typeof(FillFormat))),
                            new FormatTreeNode("Y", new Format(typeof(FillFormat)))
                        )
                };
                /*
                KeyValuePair<String, Type>[] types = new KeyValuePair<String, Type>[]
                {
                    new KeyValuePair<string, Type>("Fill\\Fill Style", typeof(FillFormat)),
                    new KeyValuePair<string, Type>("Fill\\Fill Style 2", typeof(FillFormat)),
                    new KeyValuePair<string, Type>("Fill\\Fill Style 3", typeof(FillFormat)),
                    new KeyValuePair<string, Type>("Fill\\Fill Style 4", typeof(FillFormat)),
                    new KeyValuePair<string, Type>("Line\\Line Fill", typeof(LineFormat))
                };
                KeyValuePair<String, Format>[] categories =
                        new KeyValuePair<String, Format>[types.Length];
                HashSet<String> seenCategories = new HashSet<String>();
                for (int i = 0; i < categories.Length; i++)
                {
                    string category = types[i].Key;
                    Type formatType = types[i].Value;
                    Debug.Assert(!seenCategories.Contains(types[i].Key), "Duplicate key");
                    seenCategories.Add(types[i].Key);
                    categories[i] = new KeyValuePair<string, Format>(
                            types[i].Key,
                            new Format(types[i].Value)
                        );
                }
                return categories;
                */
                /*
                FormatCategory[] categories = new FormatCategory[]
                {
                    new FormatCategory(
                        "Text",
                        new Type[] {
                        }
                    ),
                    new FormatCategory(
                        "Fill",
                        new Type[] {
                            typeof(FillFormat)
                        }
                    ),
                    new FormatCategory(
                        "Line",
                        new Type[] {
                            typeof(LineFormat)
                        }
                    ),
                    new FormatCategory(
                        "Effect",
                        new Type[] {
                        }
                    ),
                    new FormatCategory(
                        "Size/Position",
                        new Type[] {
                        }
                    )
                };
                return categories;
                */
            }
        }
    }
}
