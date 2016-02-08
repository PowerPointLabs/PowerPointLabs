using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.PictureSlidesLab.Model;

namespace PowerPointLabs.PictureSlidesLab.Service.StylesWorker
{
    class FrameStyleWorker : IStyleWorker
    {
        public IList<Shape> Execute(StyleOption option, EffectsDesigner designer, ImageItem source, Shape imageShape)
        {
            var result = new List<Shape>();
            if (option.IsUseFrameStyle)
            {
                var frameOverlayShape = designer.ApplyAlbumFrameEffect(option.FrameColor, option.FrameTransparency);
                result.Add(frameOverlayShape);
            }
            return result;
        }
    }
}
