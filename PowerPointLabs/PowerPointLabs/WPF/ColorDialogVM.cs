using System.Windows.Media;
using PowerPointLabs.Utils;

namespace PowerPointLabs.WPF
{
    class ColorDialogVM
    {
        public bool FullOpen { get; set; }
        public Color Color { get; set; }

        public DialogResult ShowDialog()
        {
            return DialogResult.None;
        }

    }
}
