using System.Windows;

namespace TestInterface
{
    public interface IColorsLabController
    {
        void OpenPane();
        Point GetApplyTextButtonLocation();
        Point GetApplyLineButtonLocation();
        Point GetApplyFillButtonLocation();
        Point GetMainColorRectangleLocation();
        Point GetEyeDropperButtonLocation();

        void SlideBrightnessSlider(int value);
        void SlideSaturationSlider(int value);

        void ClickMonochromeRect(int index);
    }
}
