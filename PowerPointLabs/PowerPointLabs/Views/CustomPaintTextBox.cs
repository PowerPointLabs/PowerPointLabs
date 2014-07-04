using System.Drawing;
using System.Windows.Forms;
using PPExtraEventHelper;

namespace PowerPointLabs.Views
{
    public class CustomPaintTextBox : NativeWindow
    {
        private readonly TextBox _parentTextBox;
        private readonly Bitmap _bitmap;
        private readonly Graphics _bufferGraphics;
        private readonly Graphics _textBoxGraphics;

        public CustomPaintTextBox(TextBox textBox)
        {
            _parentTextBox = textBox;
            _bitmap = new Bitmap(textBox.Width, textBox.Height);
            _bufferGraphics = Graphics.FromImage(_bitmap);
            _textBoxGraphics = Graphics.FromHwnd(textBox.Handle);

            AssignHandle(textBox.Handle);
        }

        ~CustomPaintTextBox()
        {
            ReleaseHandle();
        }

        private void CustomPaint()
        {
            _bufferGraphics.Clear(Color.Transparent);
            var labeledThumbnail = _parentTextBox.Parent.Parent as LabeledThumbnail;

            if (labeledThumbnail == null)
            {
                return;
            }

            TextRenderer.DrawText(_bufferGraphics, labeledThumbnail.NameLable, _parentTextBox.Font,
            _parentTextBox.ClientRectangle, _parentTextBox.ForeColor, _parentTextBox.BackColor,
            TextFormatFlags.TextBoxControl |
            TextFormatFlags.VerticalCenter |
            TextFormatFlags.HorizontalCenter |
            TextFormatFlags.WordBreak |
            TextFormatFlags.EndEllipsis);
            _textBoxGraphics.DrawImageUnscaled(_bitmap, 0, 0);
        }

        protected override void WndProc(ref Message m)
        {
            switch (m.Msg)
            {
                case (int)Native.Message.WM_PAINT:
                    _parentTextBox.Invalidate();
                    base.WndProc(ref m);
                    if (_parentTextBox.Enabled == false)
                    {
                        CustomPaint();
                    }
                    break;
                default:
                    base.WndProc(ref m);
                    break;
            }
        }
    }
}
