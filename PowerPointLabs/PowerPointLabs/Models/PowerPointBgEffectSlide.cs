using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using ImageProcessor;
using ImageProcessor.Imaging.Filters;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.CropLab;
using PowerPointLabs.Utils;

using Core = Microsoft.Office.Core;

namespace PowerPointLabs.Models
{
    class PowerPointBgEffectSlide : PowerPointSlide
    {
#pragma warning disable 0618
        private static readonly string AnimatedBackgroundPath = Path.Combine(Path.GetTempPath(), "animatedSlide.png");

        # region Constructor
        private PowerPointBgEffectSlide(Slide slide) : base(slide)
        {
            AddPowerPointLabsIndicator().ZOrder(Core.MsoZOrderCmd.msoBringToFront);
        }

        public static PowerPointBgEffectSlide BgEffectFactory(Slide refSlide, bool coverShape = true)
        {
            if (refSlide == null)
            {
                return null;
            }

            // here we cut-paste the shape to get a reference of those shapes
            ShapeRange oriShapeRange = refSlide.Shapes.Paste();

            // preprocess the shapes, eliminate animations for shapes
            foreach (Shape shape in oriShapeRange)
            {
                FromSlideFactory(refSlide).RemoveAnimationsForShape(shape);
            }

            // TODO: make use of PowerPointLabs.Presentation Model!!!
            // cut the original shape cover again and duplicate the slide
            // here the slide will be duplicated without the original shape cover
            oriShapeRange.Cut();
            PowerPointSlide newSlide = FromSlideFactory(refSlide.Duplicate()[1]);
            
            // get a copy of original cover shapes
            ShapeRange copyShapeRange = newSlide.Shapes.Paste();
            // paste the original shape cover back
            oriShapeRange = refSlide.Shapes.Paste();
            
            // make the range invisible before animated the slide
            copyShapeRange.Visible = Core.MsoTriState.msoFalse;

            MakeAnimatedBackground(newSlide);

            copyShapeRange.Visible = Core.MsoTriState.msoCTrue;
            oriShapeRange.Visible = Core.MsoTriState.msoCTrue;

            newSlide.Transition.EntryEffect = PpEntryEffect.ppEffectFadeSmoothly;
            newSlide.Transition.Duration = 0.5f;

            PowerPointBgEffectSlide bgEffectSlide = new PowerPointBgEffectSlide(newSlide.GetNativeSlide());

            try
            {
                if (coverShape)
                {
                    bgEffectSlide = PrepareForeground(oriShapeRange, copyShapeRange, refSlide, newSlide);
                }
            }
            catch (InvalidOperationException e)
            {
                refSlide.Delete();
                throw new InvalidOperationException(e.Message);
            }

            return bgEffectSlide;
        }
        # endregion

        # region API
        public void BlurAllBackground(int percentage)
        {
            AddBackgroundImage(null, percentage: percentage);
        }

        public void BlurBackground(int percentage, bool isTint)
        {
            AddBackgroundImage(null, percentage: percentage, isTint: isTint);
        }

        public void GrayScaleBackground()
        {
            AddBackgroundImage(MatrixFilters.GreyScale);
        }

        public void BlackWhiteBackground()
        {
            AddBackgroundImage(MatrixFilters.BlackWhite);
        }

        public void SepiaBackground()
        {
            AddBackgroundImage(MatrixFilters.Sepia);
        }

        public void GothamBackground()
        {
            AddBackgroundImage(MatrixFilters.Gotham);
        }
        # endregion

        # region Helper Functions
        private void AddBackgroundImage(IMatrixFilter filter, int percentage = 0, bool isTint = false)
        {
            if (filter == null)
            {
                EffectsLab.EffectsLabBlur.BlurImage(AnimatedBackgroundPath, percentage);
            }
            else
            {
                using (ImageFactory imageFactory = new ImageFactory())
                {
                    ImageFactory image = imageFactory.Load(AnimatedBackgroundPath);

                    image = image.Filter(filter);

                    image.Save(AnimatedBackgroundPath);
                }
            }

            Shape newBackground = Shapes.AddPicture(AnimatedBackgroundPath, Core.MsoTriState.msoFalse,
                                                  Core.MsoTriState.msoTrue,
                                                  0, 0,
                                                  PowerPointPresentation.Current.SlideWidth,
                                                  PowerPointPresentation.Current.SlideHeight);

            newBackground.ZOrder(Core.MsoZOrderCmd.msoSendToBack);

            if (filter == null && isTint)
            {
                EffectsLab.EffectsLabBlur.GenerateOverlayShape(this, newBackground);
            }
        }

        private static Shape MakeFrontImage(Slide refSlide, ShapeRange shapeRange)
        {
            foreach (Shape shape in shapeRange)
            {
                float softEdgeRadius = Math.Min(Math.Min(shape.Width, shape.Height) * 0.15f, 10f);
                if (ShapeUtil.IsAGroup(shape))
                {
                    for (int i = 1; i <= shape.GroupItems.Count; i++)
                    {
                        Shape child = shape.GroupItems.Range(i)[1];
                        child.SoftEdge.Radius = softEdgeRadius;
                    }
                }
                else
                {
                    shape.SoftEdge.Radius = softEdgeRadius;
                }
            }

            Shape croppedShape = CropToShape.Crop(FromSlideFactory(refSlide), shapeRange, handleError: false);

            return croppedShape;
        }

        private static void MakeAnimatedBackground(PowerPointSlide curSlide)
        {
            foreach (Shape shape in curSlide.Shapes.Cast<Shape>().Where(curSlide.HasExitAnimation))
            {
                shape.Delete();
            }

            curSlide.MoveMotionAnimation();

            Utils.GraphicsUtil.ExportSlide(curSlide, AnimatedBackgroundPath);

            List<Shape> visibleShape = curSlide.Shapes.Cast<Shape>().Where(x => x.Visible == Core.MsoTriState.msoTrue).ToList();
            
            foreach (Shape shape in visibleShape)
            {
                shape.Delete();
            }

            List<Shape> placeHolders =
                curSlide.Shapes.Cast<Shape>().Where(x => x.Type == Core.MsoShapeType.msoPlaceholder).ToList();

            foreach (Shape placeHolder in placeHolders)
            {
                placeHolder.Delete();
            }
        }

        private static PowerPointBgEffectSlide PrepareForeground(ShapeRange oriShapeRange, ShapeRange copyShapeRange,
                                                                 Slide refSlide, PowerPointSlide newSlide)
        {
            try
            {
                // crop in the original slide and put into clipboard
                Shape croppedShape = MakeFrontImage(refSlide, oriShapeRange);

                croppedShape.Cut();

                // swap the uncropped shapes and cropped shapes
                ShapeRange pastedCrop = newSlide.Shapes.Paste();

                // calibrate pasted shapes
                pastedCrop.Left -= 12;
                pastedCrop.Top -= 12;

                // ungroup front image if necessary
                if (pastedCrop[1].Type == Core.MsoShapeType.msoGroup)
                {
                    pastedCrop[1].Ungroup();
                }

                copyShapeRange.Cut();
                oriShapeRange = refSlide.Shapes.Paste();

                oriShapeRange.Fill.ForeColor.RGB = 0xaaaaaa;
                oriShapeRange.Fill.Transparency = 0.7f;
                oriShapeRange.Line.Visible = Core.MsoTriState.msoTrue;
                oriShapeRange.Line.ForeColor.RGB = 0x000000;

                ShapeUtil.MakeShapeViewTimeInvisible(oriShapeRange, refSlide);

                oriShapeRange.Select();

                // finally add transition to the new slide
                newSlide.Transition.EntryEffect = PpEntryEffect.ppEffectFadeSmoothly;
                newSlide.Transition.Duration = 0.5f;

                return new PowerPointBgEffectSlide(newSlide.GetNativeSlide());
            }
            catch (Exception e)
            {
                string errorMessage = e.Message;
                errorMessage = errorMessage.Replace("Crop To Shape", "Blur/Recolor Remainder");

                newSlide.Delete();

                throw new InvalidOperationException(errorMessage);
            }
        }
        # endregion
    }
}
