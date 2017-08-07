using System.Collections.Generic;
using System.ComponentModel.Composition;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker
{
    [Export("GeneralVariantWorker", typeof(IVariantWorker))]
    [ExportMetadata("GeneralVariantWorkerOrder", 3)]
    class PictureVariantWorker : IVariantWorker
    {
        public string GetVariantName()
        {
            return PictureSlidesLabText.VariantCategoryPicture;
        }

        public List<StyleVariant> GetVariants()
        {
            return new List<StyleVariant>
            {
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Picture 0"},
                    {"PictureIndex", 0}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Picture 1"},
                    {"PictureIndex", 1}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Picture 2"},
                    {"PictureIndex", 2}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Picture 3"},
                    {"PictureIndex", 3}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Picture 4"},
                    {"PictureIndex", 4}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Picture 5"},
                    {"PictureIndex", 5}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Picture 6"},
                    {"PictureIndex", 6}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Picture 7"},
                    {"PictureIndex", 7}
                })
            };
        }
    }
}
