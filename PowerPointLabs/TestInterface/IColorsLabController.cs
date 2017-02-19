using System.Drawing;
using System.Windows.Forms;

namespace TestInterface
{
    public interface IColorsLabController
    {
        void OpenPane();

        Panel GetDropletPanel();

        Panel GetFontColorButton();

        Panel GetLineColorButton();

        Panel GetFillColorButton();

        Panel GetMonoPanel1();

        Panel GetMonoPanel7();

        Panel GetFavColorPanel1();

        Button GetResetFavColorsButton();

        Button GetEmptyFavColorsButton();

        IColorsLabMoreInfoDialog ShowMoreColorInfo(Color color);
    }
}
