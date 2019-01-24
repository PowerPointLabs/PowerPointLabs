using System.Collections.Generic;
using System.ComponentModel.Composition;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service.StylesWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.Service.StylesWorker
{
    [Export(typeof(IStyleWorker))]
    [ExportMetadata("WorkerOrder", 7)]
    class FrameStyleWorker : IStyleWorker
    {
        public IList<Shape> Execute(StyleOption option, EffectsDesigner designer, ImageItem source, Shape imageShape, Settings settings)
        {
            List<Shape> result = new List<Shape>();
            if (option.IsUseFrameStyle)
            {
                Shape frameOverlayShape = designer.ApplyAlbumFrameEffect(option.FrameColor, option.FrameTransparency);
                result.Add(frameOverlayShape);
            }
            return result;
        }
    }
}
