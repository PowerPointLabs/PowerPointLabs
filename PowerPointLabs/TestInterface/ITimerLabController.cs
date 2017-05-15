using System.Drawing;
using System.Windows.Forms;

namespace TestInterface
{
    public interface ITimerLabController
    {
        void OpenPane();

        void ClickCreateButton();

        void SetHeightTextBoxValue(int value);
        int GetHeightTextBoxValue();
        void SetWidthTextBoxValue(int value);
        int GetWidthTextBoxValue();
    }
}
