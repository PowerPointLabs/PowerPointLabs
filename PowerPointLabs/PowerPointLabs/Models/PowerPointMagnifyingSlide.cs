using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Models
{
    class PowerPointMagnifyingSlide : PowerPointSlide
    {
        private PowerPoint.Shape indicatorShape = null;
        private PowerPoint.Shape zoomSlideCroppedShapes = null;

        private PowerPointMagnifyingSlide(PowerPoint.Slide slide) : base(slide)
        {
            _slide.Name = "PPTLabsMagnifyingSlide" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
        }

        new public static PowerPointSlide FromSlideFactory(PowerPoint.Slide slide)
        {
            if (slide == null)
                return null;

            return new PowerPointMagnifyingSlide(slide);
        }

        public void AddZoomToAreaAnimation(PowerPoint.Shape zoomShape)
        {
            PrepareForZoomToArea(zoomShape);
            PowerPoint.Shape shapeToZoom = null, referenceShape = null;
            if (!ZoomToArea.backgroundZoomChecked)
            {
                shapeToZoom = GetShapeToZoom(zoomShape);
                referenceShape = GetReferenceShape(shapeToZoom);
                DefaultMotionAnimation.AddDefaultMotionAnimation(this, shapeToZoom, referenceShape, 0.5f, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
            }
            else
            {
                shapeToZoom = zoomSlideCroppedShapes.Duplicate()[1];
                DeleteShapeAnimations(shapeToZoom);
                CopyShapePosition(zoomSlideCroppedShapes, ref shapeToZoom);

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

            List<PowerPoint.Shape> shapes = _slide.Shapes.Cast<PowerPoint.Shape>().ToList();
            var matchingShapes = shapes.Where(current => (HasExitAnimation(current) || current.Equals(zoomShape)));
            foreach (PowerPoint.Shape s in matchingShapes)
                s.Delete();

            AddZoomSlideCroppedPicture();

            DeleteSlideNotes();
            DeleteSlideMedia();
            ManageSlideTransitions();
            indicatorShape = AddPowerPointLabsIndicator();

            shapes = _slide.Shapes.Cast<PowerPoint.Shape>().ToList();
            matchingShapes = shapes.Where(current => (!(current.Equals(indicatorShape) || current.Equals(zoomSlideCroppedShapes))));
            foreach (PowerPoint.Shape s in matchingShapes)
            {
                DeleteShapeAnimations(s);
                if (!ZoomToArea.backgroundZoomChecked)
                {
                    PowerPoint.Effect effectFade = _slide.TimeLine.MainSequence.AddEffect(s, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                    effectFade.Exit = Office.MsoTriState.msoTrue;
                    effectFade.Timing.Duration = 0.25f;
                }
                else
                    s.Visible = Office.MsoTriState.msoFalse;
            }
        }

        private PowerPoint.Shape GetReferenceShape(PowerPoint.Shape shapeToZoom)
        {
            shapeToZoom.Copy();

            PowerPoint.Shape referenceShape = _slide.Shapes.Paste()[1];
            referenceShape.LockAspectRatio = Office.MsoTriState.msoTrue;
            if (referenceShape.Width > referenceShape.Height)
                referenceShape.Width = PowerPointPresentation.SlideWidth;
            else
                referenceShape.Height = PowerPointPresentation.SlideHeight;

            referenceShape.Left = (PowerPointPresentation.SlideWidth / 2) - (referenceShape.Width / 2);
            referenceShape.Top = (PowerPointPresentation.SlideHeight / 2) - (referenceShape.Height / 2);

            return referenceShape;
        }

        private PowerPoint.Shape GetShapeToZoom(PowerPoint.Shape zoomShape)
        {
            PowerPoint.Shape shapeToZoom = zoomSlideCroppedShapes.Duplicate()[1];
            DeleteShapeAnimations(shapeToZoom);
            CopyShapePosition(zoomSlideCroppedShapes, ref shapeToZoom);

            shapeToZoom.PictureFormat.CropLeft += zoomShape.Left;
            shapeToZoom.PictureFormat.CropTop += zoomShape.Top;
            shapeToZoom.PictureFormat.CropRight += (PowerPointPresentation.SlideWidth - (zoomShape.Left + zoomShape.Width));
            shapeToZoom.PictureFormat.CropBottom += (PowerPointPresentation.SlideHeight - (zoomShape.Top + zoomShape.Height));

            CopyShapePosition(zoomShape, ref shapeToZoom);
            return shapeToZoom;
        }

        private void AddZoomSlideCroppedPicture()
        {
            PowerPointSlide zoomSlideCopy = this.Duplicate();
            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(zoomSlideCopy.Index);

            PowerPoint.Shape cropShape = zoomSlideCopy.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, PowerPointPresentation.SlideWidth - 0.01f, PowerPointPresentation.SlideHeight - 0.01f);
            cropShape.Select();
            PowerPoint.Selection sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            PowerPoint.Shape croppedShape = Globals.ThisAddIn.ribbon.CropShapeToSlide(ref sel);
            croppedShape.Cut();

            zoomSlideCroppedShapes = _slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            zoomSlideCroppedShapes.Name = "PPTLabsMagnifyAreaGroup" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
            FitShapeToSlide(ref zoomSlideCroppedShapes);
            zoomSlideCopy.Delete();
        }

        private void ManageSlideTransitions()
        {
            base.RemoveSlideTransitions();
            _slide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoTrue;
            _slide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoFalse;
            _slide.SlideShowTransition.AdvanceTime = 0;
        }

        private void FitShapeToSlide(ref PowerPoint.Shape shapeToMove)
        {
            shapeToMove.LockAspectRatio = Office.MsoTriState.msoFalse;
            shapeToMove.Left = 0;
            shapeToMove.Top = 0;
            shapeToMove.Width = PowerPointPresentation.SlideWidth;
            shapeToMove.Height = PowerPointPresentation.SlideHeight;
        }

        private static void CopyShapePosition(PowerPoint.Shape shapeToCopy, ref PowerPoint.Shape shapeToMove)
        {
            shapeToMove.Left = shapeToCopy.Left + (shapeToCopy.Width / 2) - (shapeToMove.Width / 2);
            shapeToMove.Top = shapeToCopy.Top + (shapeToCopy.Height / 2) - (shapeToMove.Height / 2);
        }
    }
}
