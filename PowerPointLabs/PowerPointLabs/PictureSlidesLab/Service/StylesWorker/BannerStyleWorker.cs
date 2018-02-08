﻿using System.Collections.Generic;
using System.ComponentModel.Composition;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service.StylesWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.Service.StylesWorker
{
    [Export(typeof(IStyleWorker))]
    [ExportMetadata("WorkerOrder", 4)]
    class BannerStyleWorker : IStyleWorker
    {
        public IList<Shape> Execute(StyleOption option, EffectsDesigner designer, ImageItem source, Shape imageShape, Settings settings)
        {
            List<Shape> result = new List<Shape>();
            if (option.IsUseBannerStyle)
            {
                Shape bannerOverlayShape = ApplyBannerStyle(option, designer, imageShape);
                result.Add(bannerOverlayShape);
            }
            return result;
        }

        private Shape ApplyBannerStyle(StyleOption option, EffectsDesigner effectsDesigner, Shape imageShape)
        {
            return effectsDesigner.ApplyRectBannerEffect(option.GetBannerDirection(), option.GetTextBoxPosition(),
                        imageShape, option.BannerColor, option.BannerTransparency);
        }
    }
}
