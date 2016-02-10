using System.Collections.Generic;
using System.ComponentModel.Composition;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service.StylesWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.Service.StylesWorker
{
    [Export(typeof(IStyleWorker))]
    [ExportMetadata("WorkerOrder", 6)]
    class OutlineStyleWorker : IStyleWorker
    {
        public IList<Shape> Execute(StyleOption option, EffectsDesigner designer, ImageItem source, Shape imageShape)
        {
            var result = new List<Shape>();
            if (option.IsUseOutlineStyle)
            {
                var outlineOverlayShape = designer.ApplyRectOutlineEffect(imageShape, option.OutlineColor, 0);
                result.Add(outlineOverlayShape);
            }
            return result;
        }
    }
}
