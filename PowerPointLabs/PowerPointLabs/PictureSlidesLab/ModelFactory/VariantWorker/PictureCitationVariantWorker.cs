using System.Collections.Generic;
using System.ComponentModel.Composition;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker
{
    [Export("GeneralVariantWorker", typeof(IVariantWorker))]
    [ExportMetadata("GeneralVariantWorkerOrder", 5)]
    class PictureCitationVariantWorker : IVariantWorker
    {
        public string GetVariantName()
        {
            return TextCollection.PictureSlidesLabText.VariantCategoryImageReference;
        }

        public List<StyleVariant> GetVariants()
        {
            return new List<StyleVariant>
            {
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "No Effect"},
                    {"IsInsertReference", false},
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Bottom Right"},
                    {"IsInsertReference", true},
                    {"ImageReferenceAlignment", 3},
                    {"CitationFontSize", 14},
                    {"ImageReferenceTextBoxColor", ""}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Bottom"},
                    {"IsInsertReference", true},
                    {"ImageReferenceAlignment", 2},
                    {"CitationFontSize", 14},
                    {"ImageReferenceTextBoxColor", ""}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Bottom Left"},
                    {"IsInsertReference", true},
                    {"ImageReferenceAlignment", 1},
                    {"CitationFontSize", 14},
                    {"ImageReferenceTextBoxColor", ""}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Bottom Right (Small Font)"},
                    {"IsInsertReference", true},
                    {"ImageReferenceAlignment", 3},
                    {"CitationFontSize", 10},
                    {"ImageReferenceTextBoxColor", ""}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Bottom (Small Font)"},
                    {"IsInsertReference", true},
                    {"ImageReferenceAlignment", 2},
                    {"CitationFontSize", 10},
                    {"ImageReferenceTextBoxColor", ""}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Bottom Left (Small Font)"},
                    {"IsInsertReference", true},
                    {"ImageReferenceAlignment", 1},
                    {"CitationFontSize", 10},
                    {"ImageReferenceTextBoxColor", ""}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Bottom With Banner"},
                    {"IsInsertReference", true},
                    {"ImageReferenceAlignment", 2},
                    {"CitationFontSize", 12},
                    {"ImageReferenceTextBoxColor", "#000000"}
                })
            };
        }
    }
}
