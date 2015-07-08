using System.Drawing;

namespace FunctionalTestInterface
{
    public interface IPowerPointLabsFeatures
    {
        void AutoCrop();
        void AutoAnimate();
        void RecreateAutoAnimate();
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
        void BlurRemainderEffect();
        void GreyScaleEffect();
        void BlackAndWhiteEffect();
        void GothamEffect();
        void SepiaEffect();

        IColorsLabController ColorsLab { get; }
        IShapesLabController ShapesLab { get; }
    }
}
