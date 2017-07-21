using System;
using System.Drawing;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.EffectsLab;
using PowerPointLabs.FunctionalTestInterface.Impl.Controller;
using PowerPointLabs.ZoomLab;
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
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection1.CropToShapeTag));
            }));
        }

        public void CropOutPadding()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection1.CropOutPaddingTag));
            }));
        }

        public void CropToAspectRatioW1H10()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                var control = new RibbonControl("CropToAspectRatioOption1_10");
                control.Tag = TextCollection1.CropToAspectRatioTag;
                Ribbon.OnAction(control);
            }));
        }

        public void CropToSlide()
        {
            Ribbon.OnAction(new RibbonControl(TextCollection1.CropToSlideTag));
        }

        public void CropToSame()
        {
            Ribbon.OnAction(new RibbonControl(TextCollection1.CropToSameDimensionsTag));
        }

        public void AutoAnimate()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection1.AddAnimationSlideTag));
            }));
        }

        public void AnimateInSlide()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection1.AnimateInSlideTag));
            }));
        }

        public void AutoCaptions()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection1.AddCaptionsTag));
            }));
        }

        public void Spotlight()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl("AddSpotlight"));
            });
        }

        public void SetSpotlightProperties(float newTransparency, float newSoftEdge, Color newColor)
        {
            UIThreadExecutor.Execute(() =>
            {
                EffectsLabSettings.SpotlightTransparency = newTransparency;
                EffectsLabSettings.SpotlightSoftEdges = newSoftEdge;
                EffectsLabSettings.SpotlightColor = newColor;
            });
        }

        public void OpenSpotlightDialog()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl("SpotlightSettings"));
            });
        }

        public void ConvertToPic()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection1.ConvertToPictureTag));
            }));
        }

        public void DrillDown()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection1.DrillDownTag));
            }));
        }

        public void StepBack()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection1.StepBackTag));
            }));
        }

        public void AddZoomToArea()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection1.ZoomToAreaTag));
            }));
        }

        public void SetZoomProperties(bool backgroundChecked, bool multiSlideChecked)
        {
            UIThreadExecutor.Execute(() =>
            {
                ZoomLabSettings.BackgroundZoomChecked = backgroundChecked;
                ZoomLabSettings.MultiSlideZoomChecked = multiSlideChecked;
            });
        }

        public void HighlightPoints()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection1.HighlightPointsTag));
            }));
        }

        public void HighlightBackground()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection1.HighlightBackgroundTag));
            }));
        }

        public void HighlightFragments()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection1.HighlightTextTag));
            }));
        }

        public void RemoveHighlight()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection1.RemoveHighlightTag));
            }));
        }

        public void AutoNarrate()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection1.AddNarrationsTag));
            }));
        }

        public void GenerateTextAgenda()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection1.TextAgendaTag));
            }));
        }

        public void GenerateVisualAgenda()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection1.VisualAgendaTag));
            }));
        }

        public void GenerateBeamAgenda()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection1.BeamAgendaTag));
            }));
        }

        public void RemoveAgenda()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection1.RemoveAgendaTag));
            }));
        }

        public void SynchronizeAgenda()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection1.UpdateAgendaTag));
            }));
        }

        public void TransparentEffect()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection1.MakeTransparentTag));
            }));
        }

        public void MagnifyingGlassEffect()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection1.MagnifyTag));
            }));
        }

        public void SetTintForBlurSelected(bool isTinted)
        {
            EffectsLabSettings.IsTintSelected = isTinted;
        }

        public void SetTintForBlurRemainder(bool isTinted)
        {
            EffectsLabSettings.IsTintRemainder = isTinted;
        }

        public void SetTintForBlurBackground(bool isTinted)
        {
            EffectsLabSettings.IsTintBackground = isTinted;
        }

        public void BlurSelectedEffect()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                var control = new RibbonControl("BlurSelectedOption90");
                control.Tag = TextCollection1.EffectsLabBlurrinessTag;
                Ribbon.OnAction(control);
            }));
        }

        public void BlurRemainderEffect()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                var control = new RibbonControl("BlurRemainderOption90");
                control.Tag = TextCollection1.EffectsLabBlurrinessTag;
                Ribbon.OnAction(control);
            }));
        }

        public void BlurBackgroundEffect()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                var control = new RibbonControl("BlurBackgroundOption90");
                control.Tag = TextCollection1.EffectsLabBlurrinessTag;
                Ribbon.OnAction(control);
            }));
        }

        public void GreyScaleRemainderEffect()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                var control = new RibbonControl("GrayScaleRecolorRemainderMenu");
                control.Tag = TextCollection1.EffectsLabRecolorTag;
                Ribbon.OnAction(control);
            }));
        }

        public void BlackAndWhiteRemainderEffect()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                var control = new RibbonControl("BlackAndWhiteRecolorRemainderMenu");
                control.Tag = TextCollection1.EffectsLabRecolorTag;
                Ribbon.OnAction(control);
            }));
        }

        public void GothamRemainderEffect()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                var control = new RibbonControl("GothamRecolorRemainderMenu");
                control.Tag = TextCollection1.EffectsLabRecolorTag;
                Ribbon.OnAction(control);
            }));
        }

        public void SepiaRemainderEffect()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                var control = new RibbonControl("SepiaRecolorRemainderMenu");
                control.Tag = TextCollection1.EffectsLabRecolorTag;
                Ribbon.OnAction(control);
            }));
        }

        public void GreyScaleBackgroundEffect()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                var control = new RibbonControl("GrayScaleRecolorBackgroundMenu");
                control.Tag = TextCollection1.EffectsLabRecolorTag;
                Ribbon.OnAction(control);
            }));
        }

        public void BlackAndWhiteBackgroundEffect()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                var control = new RibbonControl("BlackAndWhiteRecolorBackgroundMenu");
                control.Tag = TextCollection1.EffectsLabRecolorTag;
                Ribbon.OnAction(control);
            }));
        }

        public void GothamBackgroundEffect()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                var control = new RibbonControl("GothamRecolorBackgroundMenu");
                control.Tag = TextCollection1.EffectsLabRecolorTag;
                Ribbon.OnAction(control);
            }));
        }

        public void SepiaBackgroundEffect()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                var control = new RibbonControl("SepiaRecolorBackgroundMenu");
                control.Tag = TextCollection1.EffectsLabRecolorTag;
                Ribbon.OnAction(control);
            }));
        }

        public void PasteToFillSlide()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection1.PasteToFillSlideTag));
            }));
        }

        public void PasteAtCursorPosition()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection1.PasteAtCursorPositionTag));
            }));
        }

        public void PasteAtOriginalPosition()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection1.PasteAtOriginalPositionTag));
            }));
        }

        public void PasteIntoGroup()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection1.PasteIntoGroupTag));
            }));
        }

        public void ReplaceWithClipboard()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection1.ReplaceWithClipboardTag));
            }));
        }
    }
}
