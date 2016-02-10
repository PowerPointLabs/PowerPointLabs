using PowerPointLabs.Models;
using PowerPointLabs.PictureSlidesLab.Model;
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

        // the slide that contains text (e.g. current slide)
        private PowerPoint.Slide ContentSlide { get; }

        # region APIs

        public static EffectsDesigner CreateEffectsDesignerForApply(PowerPoint.Slide slide,
            float slideWidth, float slideHeight, ImageItem source)
        {
            return new EffectsDesigner(slide, slideWidth, slideHeight, source);
        }

        public static EffectsDesigner CreateEffectsDesignerForPreview(PowerPoint.Slide slide,
            PowerPoint.Slide contentSlide, float slideWidth, float slideHeight, ImageItem source)
        {
            return new EffectsDesigner(slide, contentSlide, slideWidth, slideHeight, source);
        }

        /// <summary>
        /// For `apply`
        /// </summary>
        /// <param name="slide">the slide to apply the style</param>
        /// <param name="slideWidth"></param>
        /// <param name="slideHeight"></param>
        /// <param name="source"></param>
        private EffectsDesigner(PowerPoint.Slide slide, float slideWidth, float slideHeight, ImageItem source)
            : base(slide)
        {
            Setup(slideWidth, slideHeight, source);
        }

        /// <summary>
        /// For `preview`
        /// </summary>
        /// <param name="slide">the temp slide to produce preview image</param>
        /// <param name="contentSlide">the slide that contains content</param>
        /// <param name="slideWidth"></param>
        /// <param name="slideHeight"></param>
        /// <param name="source"></param>
        private EffectsDesigner(PowerPoint.Slide slide, PowerPoint.Slide contentSlide, 
            float slideWidth, float slideHeight, ImageItem source)
            : base(slide)
        {
            ContentSlide = contentSlide;
            Setup(slideWidth, slideHeight, source);
        }

        # endregion

        # region Helper Funcs

        private void Setup(float slideWidth, float slideHeight, ImageItem source)
        {
            SlideWidth = slideWidth;
            SlideHeight = slideHeight;
            Source = source;
            PrepareShapesForPreview();
        }

        private void PrepareShapesForPreview()
        {
            try
            {
                if (ContentSlide != null && _slide != ContentSlide)
                {
                    // copy shapes from content slide to preview slide
                    DeleteAllShapes();
                    ContentSlide.Shapes.Range().Copy();
                    _slide.Shapes.Paste();
                }
                DeleteShapesWithPrefix(ShapeNamePrefix);
            }
            catch
            {
                // nothing to copy-paste
            }
        }

        #endregion
    }
}
