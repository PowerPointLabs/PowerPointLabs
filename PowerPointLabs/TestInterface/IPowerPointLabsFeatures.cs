using System;
using System.Drawing;

namespace TestInterface
{
    public interface IPowerPointLabsFeatures
    {
        void AutoCrop();
        void CropOutPadding();
        void CropToAspectRatioW1H10();
        void AutoAnimate();
        void AnimateInSlide();
        void AutoCaptions();
        void Spotlight();
        void SetSpotlightProperties(float newTransparency, float newSoftEdge, Color newColor);
        void OpenSpotlightDialog();
        void ConvertToPic();
        void CropToSlide();
        void CropToSame();
        void DrillDown();
        void StepBack();
        void AddZoomToArea();
        void OpenZoomLabSettings();
        void SetZoomProperties(bool backgroundChecked, bool multiSlideChecked);
        void HighlightPoints();
        void HighlightBackground();
        void HighlightFragments();
        void AutoNarrate();

        // Agenda Lab
        void GenerateTextAgenda();
        void GenerateVisualAgenda();
        void GenerateBeamAgenda();
        void RemoveAgenda();
        void SynchronizeAgenda();

        // Effects Lab
        void TransparentEffect();
        void MagnifyingGlassEffect();
        void SetTintForBlurSelected(bool isTinted);
        void SetTintForBlurRemainder(bool isTinted);
        void SetTintForBlurBackground(bool isTinted);
        void BlurSelectedEffect();
        void BlurRemainderEffect();
        void BlurBackgroundEffect();
        void GrayScaleRemainderEffect();
        void BlackAndWhiteRemainderEffect();
        void GothamRemainderEffect();
        void SepiaRemainderEffect();
        void GrayScaleBackgroundEffect();
        void BlackAndWhiteBackgroundEffect();
        void GothamBackgroundEffect();
        void SepiaBackgroundEffect();

        // Paste lab
        void PasteToFillSlide();
        void PasteToFitSlide();
        void PasteAtOriginalPosition();
        void PasteAtCursorPosition();
        void PasteIntoGroup();
        void ReplaceWithClipboard();

        IColorsLabController ColorsLab { get; }
        IShapesLabController ShapesLab { get; }
        IPositionsLabController PositionsLab { get; }
        ISyncLabController SyncLab { get; }
        ITimerLabController TimerLab { get; }
        IELearningLabController ELearningLab { get; }
    }
}
