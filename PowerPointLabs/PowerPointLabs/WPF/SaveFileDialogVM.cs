using System.Windows.Forms;

namespace PowerPointLabs.WPF
{
    public class SaveFileDialogVM
    {
        public string DefaultExt { get; set; }
        public string Filter { get; set; }
        public string Title { get; set; }
        public string FileName { get; set; }
        public DialogResult LeftButton { get; set; }
        public DialogResult MiddleButton { get; set; }
        public DialogResult RightButton { get; set; }

        public DialogResult ShowDialog()
        {
            return DialogResult.None;
        }

    }
}
