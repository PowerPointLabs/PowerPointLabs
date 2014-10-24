using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ImageProcessor;
using ImageProcessor.Imaging.Filters;
using Core = Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Models
{
    class PowerPointBgEffectSlide : PowerPointSlide
    {
        # region Constructor
        private PowerPointBgEffectSlide(Slide slide) : base(slide)
        {
            AddPowerPointLabsIndicator().ZOrder(Core.MsoZOrderCmd.msoBringToFront);
        }

        public new static PowerPointSlide FromSlideFactory(Slide refSlide)
        {
            if (refSlide == null)
            {
                return null;
            }

            // here we cut-paste the shape to get a reference of those shapes
            var oriShapeRange = refSlide.Shapes.Paste();

            if (!CropToShape.VerifyIsShapeRangeValid(oriShapeRange))
            {
                return null;
            }

            // TODO: make use of PowerPointLabs.Presentation Model!!!
            var curPresentation = PowerPointCurrentPresentationInfo.CurrentPresentation;
            var curSlideIndex = PowerPointCurrentPresentationInfo.CurrentSlide.Index;
            var customLayout = curPresentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutText];
            var rawSlide = curPresentation.Slides.AddSlide(curSlideIndex + 1, customLayout);
            var newSlide = PowerPointSlide.FromSlideFactory(rawSlide);

            newSlide.DeleteAllShapes();

            // get a copy of original cover shapes
            var copyShapeRange = newSlide.Shapes.Paste();
            
            // crop in the original slide and put into clipboard
            var croppedShape = MakeFrontImage(oriShapeRange);

            if (croppedShape == null) return null;

            croppedShape.Cut();

            // swap the uncropped shapes and cropped shapes
            var pastedCrop = newSlide.Shapes.Paste();
            
            // calibrate pasted shapes
            pastedCrop.Left -= 12;
            pastedCrop.Top -= 12;

            copyShapeRange.Cut();
            refSlide.Shapes.Paste().Select();

            return new PowerPointBgEffectSlide(rawSlide);
        }
        # endregion

        # region API
        public void BlurBackground()
        {
            MakeBackgroundImage(null);
        }

        public void GreyScaleBackground()
        {
            MakeBackgroundImage(MatrixFilters.GreyScale);
        }

        public void BlackWhiteBackground()
        {
            MakeBackgroundImage(MatrixFilters.BlackWhite);
        }

        public void SepiaBackground()
        {
            MakeBackgroundImage(MatrixFilters.Sepia);
        }

        public void GothamBackground()
        {
            MakeBackgroundImage(MatrixFilters.Gotham);
        }
        # endregion

        # region Helper Functions
        private static Shape MakeFrontImage(ShapeRange shapeRange)
        {
            // soften cropped shape's edge
            shapeRange.SoftEdge.Type = Core.MsoSoftEdgeType.msoSoftEdgeType5;

            return CropToShape.Crop(shapeRange);
        }

        private void MakeBackgroundImage(IMatrixFilter filter)
        {
            var bgPicSaveTempPath = Path.Combine(Path.GetTempPath(), "slide.png");

            using (var imageFactory = new ImageFactory())
            {
                var image = imageFactory.Load(bgPicSaveTempPath);

                image = filter == null ? image.GaussianBlur(20) : image.Filter(filter);

                image.Save(bgPicSaveTempPath);
            }

            var newBackground = Shapes.AddPicture(bgPicSaveTempPath, Core.MsoTriState.msoFalse, Core.MsoTriState.msoTrue,
                                                  0, 0,
                                                  PowerPointCurrentPresentationInfo.SlideWidth,
                                                  PowerPointCurrentPresentationInfo.SlideHeight);

            newBackground.ZOrder(Core.MsoZOrderCmd.msoSendToBack);
        }
        # endregion
    }
}
