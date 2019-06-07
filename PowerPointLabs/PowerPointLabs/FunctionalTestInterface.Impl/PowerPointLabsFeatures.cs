using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.EffectsLab;
using PowerPointLabs.FunctionalTestInterface.Impl.Controller;
using PowerPointLabs.TextCollection;
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

        public IELearningLabController ELearningLab
        {
            get { return ELearningLabController.Instance; }
        }

        private Ribbon1 Ribbon
        {
            get { return FunctionalTestExtensions.GetRibbonUi(); }
        }

        public void OpenWindow()
        {
            Task task = 
                new Task(() => UIThreadExecutor.Execute(() => Ribbon.OnAction(new RibbonControl(AnimationLabText.SettingsTag))));
            task.Start();
        }

        public void AutoCrop()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(CropLabText.CropToShapeTag));
            }));
        }

        public void CropOutPadding()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(CropLabText.CropOutPaddingTag));
            }));
        }

        public void CropToAspectRatioW1H10()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                RibbonControl control = new RibbonControl("CropToAspectRatioOption1_10");
                control.Tag = CropLabText.CropToAspectRatioTag;
                Ribbon.OnAction(control);
            }));
        }

        public void CropToSlide()
        {
            Ribbon.OnAction(new RibbonControl(CropLabText.CropToSlideTag));
        }

        public void CropToSame()
        {
            Ribbon.OnAction(new RibbonControl(CropLabText.CropToSameDimensionsTag));
        }

        public void AutoAnimate()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(AnimationLabText.AddAnimationSlideTag));
            }));
        }

        public void AnimateInSlide()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(AnimationLabText.AnimateInSlideTag));
            }));
        }

        public void AutoCaptions()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(CaptionsLabText.AddCaptionsTag));
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
                Ribbon.OnAction(new RibbonControl(ShortcutsLabText.ConvertToPictureTag));
            }));
        }

        public void DrillDown()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(ZoomLabText.DrillDownTag));
            }));
        }

        public void StepBack()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(ZoomLabText.StepBackTag));
            }));
        }

        public void AddZoomToArea()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(ZoomLabText.ZoomToAreaTag));
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
                Ribbon.OnAction(new RibbonControl(HighlightLabText.HighlightPointsTag));
            }));
        }

        public void HighlightBackground()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(HighlightLabText.HighlightBackgroundTag));
            }));
        }

        public void HighlightFragments()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(HighlightLabText.HighlightTextTag));
            }));
        }

        public void RemoveHighlight()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(HighlightLabText.RemoveHighlightTag));
            }));
        }

        public void AutoNarrate()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(NarrationsLabText.AddNarrationsTag));
            }));
        }

        public void GenerateTextAgenda()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(AgendaLabText.TextAgendaTag));
            }));
        }

        public void GenerateVisualAgenda()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(AgendaLabText.VisualAgendaTag));
            }));
        }

        public void GenerateBeamAgenda()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(AgendaLabText.BeamAgendaTag));
            }));
        }

        public void RemoveAgenda()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(AgendaLabText.RemoveAgendaTag));
            }));
        }

        public void SynchronizeAgenda()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(AgendaLabText.UpdateAgendaTag));
            }));
        }

        public void TransparentEffect()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(EffectsLabText.MakeTransparentTag));
            }));
        }

        public void MagnifyingGlassEffect()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(EffectsLabText.MagnifyTag));
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
                RibbonControl control = new RibbonControl("BlurSelectedOption90");
                control.Tag = EffectsLabText.BlurrinessTag;
                Ribbon.OnAction(control);
            }));
        }

        public void BlurRemainderEffect()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                RibbonControl control = new RibbonControl("BlurRemainderOption90");
                control.Tag = EffectsLabText.BlurrinessTag;
                Ribbon.OnAction(control);
            }));
        }

        public void BlurBackgroundEffect()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                RibbonControl control = new RibbonControl("BlurBackgroundOption90");
                control.Tag = EffectsLabText.BlurrinessTag;
                Ribbon.OnAction(control);
            }));
        }

        public void GreyScaleRemainderEffect()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                RibbonControl control = new RibbonControl("GrayScaleRecolorRemainderMenu");
                control.Tag = EffectsLabText.RecolorTag;
                Ribbon.OnAction(control);
            }));
        }

        public void BlackAndWhiteRemainderEffect()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                RibbonControl control = new RibbonControl("BlackAndWhiteRecolorRemainderMenu");
                control.Tag = EffectsLabText.RecolorTag;
                Ribbon.OnAction(control);
            }));
        }

        public void GothamRemainderEffect()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                RibbonControl control = new RibbonControl("GothamRecolorRemainderMenu");
                control.Tag = EffectsLabText.RecolorTag;
                Ribbon.OnAction(control);
            }));
        }

        public void SepiaRemainderEffect()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                RibbonControl control = new RibbonControl("SepiaRecolorRemainderMenu");
                control.Tag = EffectsLabText.RecolorTag;
                Ribbon.OnAction(control);
            }));
        }

        public void GreyScaleBackgroundEffect()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                RibbonControl control = new RibbonControl("GrayScaleRecolorBackgroundMenu");
                control.Tag = EffectsLabText.RecolorTag;
                Ribbon.OnAction(control);
            }));
        }

        public void BlackAndWhiteBackgroundEffect()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                RibbonControl control = new RibbonControl("BlackAndWhiteRecolorBackgroundMenu");
                control.Tag = EffectsLabText.RecolorTag;
                Ribbon.OnAction(control);
            }));
        }

        public void GothamBackgroundEffect()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                RibbonControl control = new RibbonControl("GothamRecolorBackgroundMenu");
                control.Tag = EffectsLabText.RecolorTag;
                Ribbon.OnAction(control);
            }));
        }

        public void SepiaBackgroundEffect()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                RibbonControl control = new RibbonControl("SepiaRecolorBackgroundMenu");
                control.Tag = EffectsLabText.RecolorTag;
                Ribbon.OnAction(control);
            }));
        }

        public void PasteToFillSlide()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(PasteLabText.PasteToFillSlideTag));
            }));
        }

        public void PasteAtCursorPosition()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(PasteLabText.PasteAtCursorPositionTag));
            }));
        }

        public void PasteAtOriginalPosition()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(PasteLabText.PasteAtOriginalPositionTag));
            }));
        }

        public void PasteIntoGroup()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(PasteLabText.PasteIntoGroupTag));
            }));
        }

        public void ReplaceWithClipboard()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(PasteLabText.ReplaceWithClipboardTag));
            }));
        }

        public void PasteToFitSlide()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                Ribbon.OnAction(new RibbonControl(PasteLabText.PasteToFitSlideTag));
            }));
        }
    }
}
