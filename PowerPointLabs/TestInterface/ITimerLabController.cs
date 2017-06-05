using System.Drawing;
using System.Windows.Forms;

namespace TestInterface
{
    public interface ITimerLabController
    {
        void OpenPane();

        void ClickCreateButton();

        void SetDurationTextBoxValue(double value);
        void SetHeightTextBoxValue(int value);
        void SetWidthTextBoxValue(int value);
        void SetHeightSliderValue(int value);
        void SetWidthSliderValue(int value);
    }
}
