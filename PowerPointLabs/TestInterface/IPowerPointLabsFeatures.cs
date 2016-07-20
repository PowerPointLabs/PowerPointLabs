using System.Drawing;

namespace TestInterface
{
    public interface IPowerPointLabsFeatures
    {
        void AutoCrop();
        void AutoAnimate();
        void AnimateInSlide();
        void AutoCaptions();
        void Spotlight();
        void SetSpotlightProperties(float newTransparency, float newSoftEdge, Color newColor);
        void OpenSpotlightDialog();
        void FitToWidth();
        void FitToHeight();
        void ConvertToPic();
        void DrillDown();
        void StepBack();
        void AddZoomToArea();
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
        void BlurrinessOverlay(string feature, bool pressed);
        void BlurSelectedEffect();
        void BlurRemainderEffect();
        void GreyScaleRemainderEffect();
        void GothamRemainderEffect();
        void SepiaRemainderEffect();
        void BlurBackgroundEffect();
        void BlackAndWhiteBackgroundEffect();
        void SepiaBackgroundEffect();

        IColorsLabController ColorsLab { get; }
        IShapesLabController ShapesLab { get; }
    }
}
