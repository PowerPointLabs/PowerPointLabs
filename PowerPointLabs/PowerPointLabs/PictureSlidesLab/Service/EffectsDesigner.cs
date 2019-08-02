using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Util;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.PictureSlidesLab.Service
{
    /// <summary>
    /// EffectsDesigner provides APIs to generate effects (the elements of a style).
    /// 
    /// To support any new effect, create a new partial class of EffectsDesigner under 
    /// folder `EffectsDesigner.Partial`.
    /// </summary>
    public partial class EffectsDesigner : PowerPointSlide
    {
        public const string ShapeNamePrefix = "pptPictureSlidesLab";

        // the picture to apply/preview
        private ImageItem Source { get; set; }

        private float SlideWidth { get; set; }

        private float SlideHeight { get; set; }

        # region APIs

        /// <summary>
        /// For `preview`
        /// </summary>
        /// <param name="slide">the slide at the background</param>
        public EffectsDesigner(PowerPoint.Slide slide)
            : base(slide)
        {
            
        }

        public void PreparePreviewing(PowerPoint.Slide contentSlide, float slideWidth, float slideHeight, ImageItem source)
        {
            Logger.Log("PreparePreviewing begins");
            InitLayoutAndDesign(contentSlide);
            DeleteAllShapes();
            CopyShapes(contentSlide);
            Setup(slideWidth, slideHeight, source);
            Logger.Log("PreparePreviewing done");
        }

        /// <summary>
        /// For `apply`
        /// </summary>
        /// <param name="slide">the slide to apply the style</param>
        /// <param name="slideWidth"></param>
        /// <param name="slideHeight"></param>
        /// <param name="source"></param>
        public EffectsDesigner(PowerPoint.Slide slide, float slideWidth, float slideHeight, ImageItem source)
            : base(slide)
        {
            Setup(slideWidth, slideHeight, source);
        }

        # endregion

        # region Helper Funcs

        private void Setup(float slideWidth, float slideHeight, ImageItem source)
        {
            SlideWidth = slideWidth;
            SlideHeight = slideHeight;
            Source = source;
            DeleteShapesWithPrefix(ShapeNamePrefix);
        }

        private void InitLayoutAndDesign(PowerPoint.Slide contentSlide)
        {
            if (contentSlide.Layout == PowerPoint.PpSlideLayout.ppLayoutCustom)
            {
                _slide.CustomLayout = contentSlide.CustomLayout;
                // remove target textbox from the layout
                Shape shape = ShapeUtil.GetTextShapeToProcess(Shapes);
                if (shape != null)
                {
                    shape.SafeDelete();
                }
            }
            else
            {
                _slide.Layout = contentSlide.Layout;
            }
            _slide.Design = contentSlide.Design;
        }

        private void CopyShapes(PowerPoint.Slide contentSlide)
        {
            try
            {
                // copy shapes from content slide to preview slide
                _slide.Shapes.SafeCopy(contentSlide.Shapes.Range());
            }
            catch
            {
                // nothing to copy-paste
            }
        }

        #endregion
    }
}
