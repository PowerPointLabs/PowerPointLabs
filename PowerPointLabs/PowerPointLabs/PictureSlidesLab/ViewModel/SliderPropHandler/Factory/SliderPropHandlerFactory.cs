using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using PowerPointLabs.PictureSlidesLab.ViewModel.SliderPropHandler.Interface;

namespace PowerPointLabs.PictureSlidesLab.ViewModel.SliderPropHandler.Factory
{
    [Export(typeof(SliderPropHandlerFactory))]
    class SliderPropHandlerFactory
    {
        [ImportMany(typeof(ISliderPropHandler))]
        private IEnumerable<Lazy<ISliderPropHandler, IPropHandlerNameMetadata>> ImportedSliderPropHandlers { get; set; }

        public ISliderPropHandler GetSliderPropHandler(string propHandlerName)
        {
            if (propHandlerName.Contains(TextCollection.PictureSlidesLabText.TransparencyHasEffect))
            {
                var transparencyPropHandler = (TransparencySliderPropHandler)ImportedSliderPropHandlers
                    .Where(propHandler => propHandler.Metadata.PropHandlerName == TextCollection.PictureSlidesLabText.TransparencyHasEffect)
                    .Select(propHandler => propHandler.Value)
                    .FirstOrDefault();
                transparencyPropHandler.PropName = propHandlerName;
                return transparencyPropHandler;
            }
            else
            {
                return ImportedSliderPropHandlers
                    .Where(propHandler => propHandler.Metadata.PropHandlerName == propHandlerName)
                    .Select(propHandler => propHandler.Value)
                    .FirstOrDefault();
            }
        }

        public struct SliderProperties
        {
            public int Value;
            public int Maximum;
            public int TickFrequency;
        }
    }
}
