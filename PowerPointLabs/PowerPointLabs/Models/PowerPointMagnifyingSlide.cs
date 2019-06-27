using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.AnimationLab;
using PowerPointLabs.CropLab;
using PowerPointLabs.Utils;
using PowerPointLabs.ZoomLab;

using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Models
{
    class PowerPointMagnifyingSlide : PowerPointSlide
    {
#pragma warning disable 0618
        private PowerPoint.Shape indicatorShape = null;
        private PowerPoint.Shape zoomSlideCroppedShapes = null;

        private PowerPointMagnifyingSlide(PowerPoint.Slide slide) : base(slide)
        {
            _slide.Name = "PPTLabsMagnifyingSlide" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
            DeleteHiddenShapes();
        }

        public static PowerPointSlide FromSlideFactory(PowerPoint.Slide slide)
        {
            if (slide == null)
            {
                return null;
            }

            return new PowerPointMagnifyingSlide(slide);
        }

        public void AddZoomToAreaAnimation(PowerPoint.Shape zoomShape)
        {
            PrepareForZoomToArea(zoomShape);
            PowerPoint.Shape shapeToZoom = null, referenceShape = null;
            if (!ZoomLabSettings.BackgroundZoomChecked)
            {
                shapeToZoom = GetShapeToZoom(zoomShape);
                referenceShape = GetReferenceShape(shapeToZoom);
                DefaultMotionAnimation.AddDefaultMotionAnimation(this, shapeToZoom, referenceShape, 0.5f, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
            }
            else
            {
                shapeToZoom = zoomSlideCroppedShapes.Duplicate()[1];
                DeleteShapeAnimations(shapeToZoom);
                LegacyShapeUtil.CopyShapePosition(zoomSlideCroppedShapes, ref shapeToZoom);

                referenceShape = GetReferenceShape(zoomShape);
                DefaultMotionAnimation.AddZoomToAreaMotionAnimation(this, shapeToZoom, zoomShape, referenceShape, 0.5f, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
            }

            shapeToZoom.Name = "PPTLabsMagnifyAreaSlide" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
            referenceShape.Delete();
            zoomSlideCroppedShapes.Visible = Office.MsoTriState.msoFalse;
            indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
        }

        private void PrepareForZoomToArea(PowerPoint.Shape zoomShape)
        {
            MoveMotionAnimation();

            //Delete zoom shapes and shapes with exit animations
            List<PowerPoint.Shape> shapes = _slide.Shapes.Cast<PowerPoint.Shape>().ToList();
            IEnumerable<PowerPoint.Shape> matchingShapes = shapes.Where(current => (HasExitAnimation(current) || current.Equals(zoomShape)));
            foreach (PowerPoint.Shape s in matchingShapes)
            {
                s.Delete();
            }

            float magnifyRatio = PowerPointPresentation.Current.SlideWidth / zoomShape.Width;
            AddZoomSlideCroppedPicture(magnifyRatio);

            DeleteSlideNotes();
            DeleteSlideMedia();
            ManageSlideTransitions();
            indicatorShape = AddPowerPointLabsIndicator();

            //Add fade out effect for non-zoom shapes
            shapes = _slide.Shapes.Cast<PowerPoint.Shape>().ToList();
            matchingShapes = shapes.Where(current => (!(current.Equals(indicatorShape) || current.Equals(zoomSlideCroppedShapes))));
            foreach (PowerPoint.Shape s in matchingShapes)
            {
                DeleteShapeAnimations(s);
                if (!ZoomLabSettings.BackgroundZoomChecked)
                {
                    PowerPoint.Effect effectFade = _slide.TimeLine.MainSequence.AddEffect(s, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                    effectFade.Exit = Office.MsoTriState.msoTrue;
                    effectFade.Timing.Duration = 0.25f;
                }
                else
                {
                    s.Visible = Office.MsoTriState.msoFalse;
                }
            }
        }

        //Return shape which helps to calculate the amount of zoom-in animation
        private PowerPoint.Shape GetReferenceShape(PowerPoint.Shape shapeToZoom)
        {
            PowerPoint.Shape referenceShape = shapeToZoom.Duplicate()[1];
            referenceShape.LockAspectRatio = Office.MsoTriState.msoTrue;
            if (referenceShape.Width > referenceShape.Height)
            {
                referenceShape.Width = PowerPointPresentation.Current.SlideWidth;
            }
            else
            {
                referenceShape.Height = PowerPointPresentation.Current.SlideHeight;
            }

            referenceShape.Left = (PowerPointPresentation.Current.SlideWidth / 2) - (referenceShape.Width / 2);
            referenceShape.Top = (PowerPointPresentation.Current.SlideHeight / 2) - (referenceShape.Height / 2);

            return referenceShape;
        }

        //Return cropped version of slide picture
        private PowerPoint.Shape GetShapeToZoom(PowerPoint.Shape zoomShape)
        {
            PowerPoint.Shape shapeToZoom = zoomSlideCroppedShapes.Duplicate()[1];
            DeleteShapeAnimations(shapeToZoom);
            LegacyShapeUtil.CopyShapePosition(zoomSlideCroppedShapes, ref shapeToZoom);

            shapeToZoom.PictureFormat.CropLeft += zoomShape.Left;
            shapeToZoom.PictureFormat.CropTop += zoomShape.Top;
            shapeToZoom.PictureFormat.CropRight += (PowerPointPresentation.Current.SlideWidth - (zoomShape.Left + zoomShape.Width));
            shapeToZoom.PictureFormat.CropBottom += (PowerPointPresentation.Current.SlideHeight - (zoomShape.Top + zoomShape.Height));

            LegacyShapeUtil.CopyShapePosition(zoomShape, ref shapeToZoom);
            return shapeToZoom;
        }

        //Stores slide-size crop of the current slide as a global variable
        private void AddZoomSlideCroppedPicture(float magnifyRatio = 1.0f)
        {
            PowerPointSlide zoomSlideCopy = this.Duplicate();
            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(zoomSlideCopy.Index);

            Shape cropShape = zoomSlideCopy.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, PowerPointPresentation.Current.SlideWidth - 0.01f, PowerPointPresentation.Current.SlideHeight - 0.01f);
            cropShape.Select();
            Selection sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            Shape croppedShape = CropToShape.Crop(zoomSlideCopy, sel, magnifyRatio: magnifyRatio);
            zoomSlideCroppedShapes = GraphicsUtil.CutAndPaste(croppedShape, _slide);

            zoomSlideCroppedShapes.Name = "PPTLabsMagnifyAreaGroup" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
            ShapeUtil.FitShapeToSlide(ref zoomSlideCroppedShapes);
            zoomSlideCopy.Delete();
        }

        private void ManageSlideTransitions()
        {
            RemoveSlideTransitions();
            _slide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoTrue;
            _slide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoFalse;
            _slide.SlideShowTransition.AdvanceTime = 0;
        }
    }
}
