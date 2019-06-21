using System.Drawing;
using System.Windows.Forms;

using MediaColor = System.Windows.Media.Color;

namespace PowerPointLabs.Utils.Windows
{
    class ColorDialogUtil
    {
        public static Color? RequestForColor(MediaColor selectedColor, bool fullOpen = true)
        {
            Color color = GraphicsUtil.DrawingColorFromMediaColor(selectedColor);
            return RequestForColor(color, fullOpen);
        }
        public static Color? RequestForColor(Color selectedColor, bool fullOpen = true)
        {
            return RequestForColorWinForm(selectedColor, fullOpen);
        }

        private static Color? RequestForColorWinForm(Color selectedColor, bool fullOpen)
        {
            ColorDialog dialog = new ColorDialog()
            {
                Color = selectedColor,
                FullOpen = fullOpen
            };
            if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                return null;
            }
            return dialog.Color;
        }
    }
}
