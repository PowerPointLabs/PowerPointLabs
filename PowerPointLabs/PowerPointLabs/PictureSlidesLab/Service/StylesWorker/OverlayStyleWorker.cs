using System.Collections.Generic;
using System.ComponentModel.Composition;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service.StylesWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.Service.StylesWorker
{
    [Export(typeof(IStyleWorker))]
    [ExportMetadata("WorkerOrder", 3)]
    class OverlayStyleWorker : IStyleWorker
    {
        public IList<Shape> Execute(StyleOption option, EffectsDesigner designer, ImageItem source, Shape imageShape, Settings settings)
        {
            var result = new List<Shape>();
            if (option.IsUseOverlayStyle)
            {
                var backgroundOverlayShape = designer.ApplyOverlayEffect(option.OverlayColor, option.OverlayTransparency);
                result.Add(backgroundOverlayShape);
            }
            return result;
        }
    }
}
