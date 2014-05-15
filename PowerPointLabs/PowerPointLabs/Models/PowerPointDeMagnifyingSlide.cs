using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Models
{
    class PowerPointDeMagnifyingSlide : PowerPointSlide
    {
        private PowerPoint.Shape indicatorShape = null;
        private PowerPoint.Shape zoomSlideCroppedShapes = null;
        private PowerPointDeMagnifyingSlide(PowerPoint.Slide slide) : base(slide)
        {
            _slide.Name = "PPTLabsDeMagnifyingSlide" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
        }

        new public static PowerPointSlide FromSlideFactory(PowerPoint.Slide slide)
        {
            if (slide == null)
                return null;

            return new PowerPointDeMagnifyingSlide(slide);
        }

        public void AddZoomToAreaAnimation(PowerPoint.Shape zoomShape)
        {
            PrepareForZoomToArea(zoomShape);

            PowerPoint.Effect lastEffect = null;
            if (!ZoomToArea.backgroundZoomChecked)
            {
                zoomSlideCroppedShapes.LockAspectRatio = Office.MsoTriState.msoTrue;
                if (zoomSlideCroppedShapes.Width > zoomSlideCroppedShapes.Height)
                    zoomSlideCroppedShapes.Width = PowerPointPresentation.SlideWidth;
                else
                    zoomSlideCroppedShapes.Height = PowerPointPresentation.SlideHeight;

                zoomSlideCroppedShapes.Left = (PowerPointPresentation.SlideWidth / 2) - (zoomSlideCroppedShapes.Width / 2);
                zoomSlideCroppedShapes.Top = (PowerPointPresentation.SlideHeight / 2) - (zoomSlideCroppedShapes.Height / 2);

                DefaultMotionAnimation.AddDefaultMotionAnimation(this, zoomSlideCroppedShapes, zoomShape, 0.5f, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);

                bool isFirst = true;
                PowerPoint.Effect effectFade = null;
                foreach (PowerPoint.Shape tmp in _slide.Shapes)
                {
                    if (!(tmp.Equals(zoomSlideCroppedShapes) || tmp.Equals(indicatorShape)))
                    {
                        if (isFirst)
                            effectFade = _slide.TimeLine.MainSequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                        else
                            effectFade = _slide.TimeLine.MainSequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);

                        effectFade.Timing.Duration = 0.25f;
                        isFirst = false;
                    }
                }

                effectFade = _slide.TimeLine.MainSequence.AddEffect(zoomSlideCroppedShapes, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                effectFade.Exit = Office.MsoTriState.msoTrue;
                effectFade.Timing.Duration = 0.25f;
            }
            else
            {
                GetShapeToZoomWithBackground(zoomShape);
                FrameMotionAnimation.animationType = FrameMotionAnimation.FrameMotionAnimationType.kZoomToAreaDeMagnify;
                FrameMotionAnimation.AddStepBackFrameMotionAnimation(this, zoomSlideCroppedShapes);
                lastEffect = _slide.TimeLine.MainSequence[_slide.TimeLine.MainSequence.Count];

                bool isFirst = true;
                PowerPoint.Effect effectFade = null;
                foreach (PowerPoint.Shape tmp in _slide.Shapes)
                {
                    if (!(tmp.Equals(zoomSlideCroppedShapes) || tmp.Equals(indicatorShape)) && !(tmp.Name.Contains("PPTLabsMagnifyShape")) && !(tmp.Name.Contains("PPTLabsMagnifyArea")))
                    {
                        tmp.Visible = Office.MsoTriState.msoTrue;
                        if (isFirst)
                            effectFade = _slide.TimeLine.MainSequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                        else
                            effectFade = _slide.TimeLine.MainSequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);

                        effectFade.Timing.Duration = 0.01f;
                        isFirst = false;
                    }
                }

                lastEffect.MoveTo(_slide.TimeLine.MainSequence.Count);
                lastEffect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                lastEffect.Timing.TriggerDelayTime = 0.0f;
                lastEffect.Timing.Duration = 0.01f;
            }

            indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
        }

        private void PrepareForZoomToArea(PowerPoint.Shape zoomShape)
        {
            RemoveAnimationsForShapes(_slide.Shapes.Cast<PowerPoint.Shape>().ToList());
            GetShapesWithPrefix("PPTLabsIndicator")[0].Delete();
            DeleteShapesWithPrefix("PPTLabsMagnifyAreaSlide");

            AddZoomSlideCroppedPicture(zoomShape);

            DeleteSlideNotes();
            DeleteSlideMedia();
            ManageSlideTransitions();
            indicatorShape = AddPowerPointLabsIndicator();
        }

        private void GetShapeToZoomWithBackground(PowerPoint.Shape zoomShape)
        {
            PowerPoint.Shape referenceShape = GetReferenceShape(zoomShape);

            float finalWidthMagnify = referenceShape.Width;
            float initialWidthMagnify = zoomShape.Width;
            float finalHeightMagnify = referenceShape.Height;
            float initialHeightMagnify = zoomShape.Height;

            zoomShape.Copy();
            PowerPoint.Shape zoomShapeCopy = _slide.Shapes.Paste()[1];
            CopyShapeAttributes(zoomShape, ref zoomShapeCopy);

            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(_slide.SlideIndex);
            zoomSlideCroppedShapes.Select();
            zoomShapeCopy.Select(Office.MsoTriState.msoFalse);
            PowerPoint.ShapeRange selection = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
            PowerPoint.Shape groupShape = selection.Group();

            groupShape.Width *= (finalWidthMagnify / initialWidthMagnify);
            groupShape.Height *= (finalHeightMagnify / initialHeightMagnify);
            groupShape.Ungroup();
            zoomSlideCroppedShapes.Left += (referenceShape.Left - zoomShapeCopy.Left);
            zoomSlideCroppedShapes.Top += (referenceShape.Top - zoomShapeCopy.Top);
            zoomShapeCopy.Delete();
            referenceShape.Delete();
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

        private void AddZoomSlideCroppedPicture(PowerPoint.Shape zoomShape)
        {
            zoomSlideCroppedShapes = GetShapesWithPrefix("PPTLabsMagnifyAreaGroup")[0];
            zoomSlideCroppedShapes.Visible = Office.MsoTriState.msoTrue;
            DeleteShapeAnimations(zoomSlideCroppedShapes);

            if (!ZoomToArea.backgroundZoomChecked)
            {
                zoomSlideCroppedShapes.PictureFormat.CropLeft += zoomShape.Left;
                zoomSlideCroppedShapes.PictureFormat.CropTop += zoomShape.Top;
                zoomSlideCroppedShapes.PictureFormat.CropRight += (PowerPointPresentation.SlideWidth - (zoomShape.Left + zoomShape.Width));
                zoomSlideCroppedShapes.PictureFormat.CropBottom += (PowerPointPresentation.SlideHeight - (zoomShape.Top + zoomShape.Height));

                CopyShapePosition(zoomShape, ref zoomSlideCroppedShapes);
            }
        }

        private void ManageSlideTransitions()
        {
            base.RemoveSlideTransitions();
            _slide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoFalse;
            _slide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoTrue;
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

        private static void CopyShapeSize(PowerPoint.Shape shapeToCopy, ref PowerPoint.Shape shapeToMove)
        {
            shapeToMove.LockAspectRatio = Office.MsoTriState.msoFalse;
            shapeToMove.Width = shapeToCopy.Width;
            shapeToMove.Height = shapeToCopy.Height;
        }

        private static void CopyShapeAttributes(PowerPoint.Shape shapeToCopy, ref PowerPoint.Shape shapeToMove)
        {
            CopyShapeSize(shapeToCopy, ref shapeToMove);
            CopyShapePosition(shapeToCopy, ref shapeToMove);
        }
    }
}
