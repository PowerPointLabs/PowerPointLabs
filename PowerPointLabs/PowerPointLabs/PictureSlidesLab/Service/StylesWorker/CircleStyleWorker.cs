using System.Collections.Generic;
using System.ComponentModel.Composition;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service.StylesWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.Service.StylesWorker
{
    [Export(typeof(IStyleWorker))]
    [ExportMetadata("WorkerOrder", 8)]
    class CircleStyleWorker : IStyleWorker
    {
        public IList<Shape> Execute(StyleOption option, EffectsDesigner designer, ImageItem source, Shape imageShape, Settings settings)
        {
            List<Shape> result = new List<Shape>();
            if (option.IsUseCircleStyle)
            {
                Shape circleOverlayShape = designer.ApplyCircleRingsEffect(option.CircleColor, option.CircleTransparency);
                result.Add(circleOverlayShape);
            }
            return result;
        }
    }
}
