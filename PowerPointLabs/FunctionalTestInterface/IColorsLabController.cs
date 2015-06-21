using System.Drawing;
using System.Windows.Forms;

namespace FunctionalTestInterface
{
    public interface IColorsLabController
    {
        void OpenPane();

        Panel GetDropletPanel();

        Button GetFontColorButton();

        Button GetLineColorButton();

        Button GetFillCollorButton();

        Panel GetMonoPanel1();

        Panel GetMonoPanel7();

        Panel GetFavColorPanel1();

        Button GetResetFavColorsButton();

        Button GetEmptyFavColorsButton();

        IColorsLabMoreInfoDialog ShowMoreColorInfo(Color color);
    }
}
