using System.Windows;
using System.Windows.Controls;

using MsoTextOrientation = Microsoft.Office.Core.MsoTextOrientation;

namespace PowerPointLabs.TooltipsLab
{
    /// <summary>
    /// Interaction logic for TextBoxWithDefaultText.xaml
    /// </summary>
    public partial class TextBoxWithDefaultText : UserControl
    {
        private bool isTextEdited = false;
        private bool isChangedByEvent = false;
        private string defaultText = "Enter Text Here";
        public TextBoxWithDefaultText()
        {
            InitializeComponent();
        }

        public void SetTextBoxFormat(MsoTextOrientation orientation, float left, float top, float width, float height)
        {
            textBox.TextChanged += OnTextChanged;
        }

        protected override void OnIsKeyboardFocusedChanged(DependencyPropertyChangedEventArgs e)
        {
            if (IsFocused)
            {
                textBox.Focus();
            }
            if (!isTextEdited)
            {
                return;
            }

            isChangedByEvent = true;
            if (IsFocused)
            {
                textBox.Text = "";
            }
            else
            {
                textBox.Text = defaultText;
            }
        }

        private void OnTextChanged(object sender, TextChangedEventArgs e)
        {
            if (!(sender is TextBox))
            {
                return;
            }
            TextBox source = sender as TextBox;
            if (isChangedByEvent)
            {
                isChangedByEvent = false;
                isTextEdited = source.Text == "";
                return;
            }
        }
    }
}
