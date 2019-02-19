using System.Collections.Generic;
using System.Drawing;

using Point = System.Windows.Point;
using Rectangle = System.Windows.Shapes.Rectangle;

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
        void ClickAnalogousRect(int index);
        void ClickComplementaryRect(int index);
        void ClickTriadicRect(int index);
        void ClickTetradicRect(int index);

        void LoadFavoriteColors(string filePath);
        void ResetFavoriteColors();
        void ClearFavoriteColors();

        List<Color> GetCurrentFavoritePanel();

    }
}
