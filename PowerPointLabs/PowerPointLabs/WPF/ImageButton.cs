using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace PowerPointLabs.WPF
{
    public class ImageButton : Button
    {
        public readonly int ImageButtonUniqueId;
        
        private readonly Image _image = null;
        private readonly Panel _panel = null;
        private bool _hasImage = false;
        private string _text = null;
        private static int _nextId = 0;


        public ImageButton()
        {
            Margin = new Thickness(3);
            Focusable = false;

            StackPanel panel = new StackPanel();
            panel.Orientation = Orientation.Horizontal;

            panel.Margin = new System.Windows.Thickness(1);

            _image = new Image();
            _image.Margin = new System.Windows.Thickness(0, 0, 0, 0);
            _image.Width = panel.Width;
            _image.Height = panel.Height;
            panel.Children.Add(_image);

            _panel = panel;

            ImageButtonUniqueId = _nextId;
            _nextId++;
        }

        // Properties
        public ImageSource Image
        {
            get
            {
                if (_image != null)
                {
                    return _image.Source;
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if (_image != null)
                {
                    _image.Source = value;
                    this.Content = _panel;
                    _hasImage = true;
                }
            }
        }

        public string Text
        {
            get { return _text; }
            set
            {
                _text = value;
                if (!_hasImage)
                {
                    this.Content = new TextBlock
                    {
                        TextAlignment = TextAlignment.Center,
                        Text = _text
                    };
                }
            }
        }
    }
}
