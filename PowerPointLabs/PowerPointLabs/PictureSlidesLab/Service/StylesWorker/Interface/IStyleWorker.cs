using System.Collections.Generic;

using PowerPointLabs.PictureSlidesLab.Model;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.PictureSlidesLab.Service.StylesWorker.Interface
{
    public interface IStyleWorker
    {
        /// <summary>
        /// Apply the style using given effects designer and style option.
        /// </summary>
        /// <param name="option"></param>
        /// <param name="designer"></param>
        /// <param name="source"></param>
        /// <param name="imageShape"></param>
        /// <param name="settings"></param>
        /// <returns>Return empty list if nothing to return</returns>
        IList<PowerPoint.Shape> Execute(StyleOption option, EffectsDesigner designer, ImageItem source, 
            PowerPoint.Shape imageShape, Settings settings);
    }
}
