using System;
using System.Drawing;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.FunctionalTestInterface.Impl.Controller;
using TestInterface;

namespace PowerPointLabs.FunctionalTestInterface.Impl
{
    [Serializable]
    class PowerPointLabsFeatures : MarshalByRefObject, IPowerPointLabsFeatures
    {
        public IColorsLabController ColorsLab
        {
            get { return ColorsLabController.Instance; }
        }

        public IShapesLabController ShapesLab
        {
            get { return ShapesLabController.Instance; }
        }

        public IPositionsLabController PositionsLab
        {
            get { return PositionsLabController.Instance; }
        }

        public ISyncLabController SyncLab
        {
            get { return SyncLabController.Instance; }
        }

        public ITimerLabController TimerLab
        {
            get { return TimerLabController.Instance; }
        }

        private Ribbon1 Ribbon
        {
            get { return FunctionalTestExtensions.GetRibbonUi(); }
        }

        public void AutoCrop()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl(TextCollection.CropToShapeTag);
                Ribbon.OnAction(control);
            });
        }

        public void CropOutPadding()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl(TextCollection.CropOutPaddingTag);
                Ribbon.OnAction(control);
            });
        }

        public void CropToAspectRatioW1H10()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl("CropToAspectRatioOption1_10");
                control.Tag = TextCollection.CropToAspectRatioTag;
                Ribbon.OnAction(control);
            });
        }

        public void CropToSlide()
        {
            var control = new RibbonControl(TextCollection.CropToSlideTag);
            Ribbon.OnAction(control);
        }

        public void CropToSame()
        {
            var control = new RibbonControl(TextCollection.CropToSameDimensionsTag);
            Ribbon.OnAction(control);
        }

        public void AutoAnimate()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl(TextCollection.AddAnimationSlideTag);
                Ribbon.OnAction(control);
            });
        }

        public void AnimateInSlide()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl(TextCollection.AnimateInSlideTag);
                Ribbon.OnAction(control);
            });
        }

        public void AutoCaptions()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.AddCaptionsTag));
            });
        }

        public void Spotlight()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.SpotlightBtnClick(new RibbonControl("Spotlight"));
            });
        }

        public void SetSpotlightProperties(float newTransparency, float newSoftEdge, Color newColor)
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.SpotlightPropertiesEdited(newTransparency, newSoftEdge, newColor);
            });
        }

        public void OpenSpotlightDialog()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.SpotlightDialogButtonPressed(new RibbonControl("OpenSpotlightDialog"));
            });
        }

        public void ConvertToPic()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.ConvertToPictureTag));
            });
        }

        public void DrillDown()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.DrillDownTag));
            });
        }

        public void StepBack()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.StepBackTag));
            });
        }

        public void AddZoomToArea()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.ZoomToAreaTag));
            });
        }

        public void SetZoomProperties(bool backgroundChecked, bool multiSlideChecked)
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.ZoomPropertiesEdited(backgroundChecked, multiSlideChecked);
            });
        }

        public void HighlightPoints()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.HighlightPointsTag));
            });
        }

        public void HighlightBackground()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.HighlightBackgroundTag));
            });
        }

        public void HighlightFragments()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.HighlightTextTag));
            });
        }

        public void RemoveHighlight()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.RemoveHighlightTag));
            });
        }

        public void AutoNarrate()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.AddNarrationsTag));
            });
        }

        public void GenerateTextAgenda()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.TextAgendaTag));
            });
        }

        public void GenerateVisualAgenda()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.VisualAgendaTag));
            });
        }

        public void GenerateBeamAgenda()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.BeamAgendaTag));
            });
        }

        public void RemoveAgenda()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.RemoveAgendaTag));
            });
        }

        public void SynchronizeAgenda()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.UpdateAgendaTag));
            });
        }

        public void TransparentEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.TransparentEffectClick(new RibbonControl("TransparentEffect"));
            });
        }

        public void MagnifyingGlassEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.MagnifyGlassEffectClick(new RibbonControl("MagnifyingGlassEffect"));
            });
        }
        
        public void BlurrinessOverlay(string feature, bool pressed)
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl(feature + TextCollection.DynamicMenuCheckBoxId);
                control.Tag = "Blurriness";
                Ribbon.OnCheckBoxAction(control, pressed);
            });
        }

        public void BlurSelectedEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl("EffectsLabBlurSelectedOption90");
                control.Tag = "Blurriness";
                Ribbon.OnAction(control);
            });
        }

        public void BlurRemainderEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl("EffectsLabBlurRemainderOption90");
                control.Tag = "Blurriness";
                Ribbon.OnAction(control);
            });
        }

        public void GreyScaleRemainderEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.GreyScaleRemainderEffectClick(new RibbonControl("GreyScaleEffect"));
            });
        }


        public void GothamRemainderEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.GothamRemainderEffectClick(new RibbonControl("GothamEffect"));
            });
        }

        public void SepiaRemainderEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.SepiaRemainderEffectClick(new RibbonControl("SepiaEffect"));
            });
        }


        public void BlurBackgroundEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl("EffectsLabBlurBackgroundOption90");
                control.Tag = "Blurriness";
                Ribbon.OnAction(control);
            });
        }

        public void BlackAndWhiteBackgroundEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.BlackWhiteBackgroundEffectClick(new RibbonControl("BlackAndWhiteEffect"));
            });
        }

        public void SepiaBackgroundEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.SepiaBackgroundEffectClick(new RibbonControl("SepiaEffect"));
            });
        }

        public void PasteToFillSlide()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl(TextCollection.PasteToFillSlideTag);
                Ribbon.OnAction(control);
            });
        }

        public void PasteAtCursorPosition()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl(TextCollection.PasteAtCursorPositionTag);
                Ribbon.OnAction(control);
            });
        }

        public void PasteAtOriginalPosition()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl(TextCollection.PasteAtOriginalPositionTag);
                Ribbon.OnAction(control);
            });
        }

        public void PasteIntoGroup()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl(TextCollection.PasteIntoGroupTag);
                Ribbon.OnAction(control);
            });
        }

        public void ReplaceWithClipboard()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl(TextCollection.ReplaceWithClipboardTag);
                Ribbon.OnAction(control);
            });
        }
    }
}
