using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace PowerPointLabs
{
    public partial class LabeledThumbnail : UserControl
    {
        private bool _isHighlighted;
        private Image _imageSource;
        private string _imageSourcePath;
        private string _firstNameAssigned = string.Empty;

        # region Constructors
        public LabeledThumbnail()
        {
            Initialize();
        }

        public LabeledThumbnail(string imageName, string nameLable)
        {
            Initialize();

            NameLable = nameLable;

            var image = new Bitmap(imageName);
            ImageToThumbnail = CreateThumbnailImage(image, 50, 50);
        }
        # endregion

        # region Properties
        private const string InvalidCharacterError = @"'\' and '.' are not allowed in the name";
        public string NameLable
        {
            get { return labelTextBox.Text; }
            set
            {
                if (Verify(value))
                {
                    labelTextBox.Text = value;

                    if (_firstNameAssigned == string.Empty)
                    {
                        _firstNameAssigned = value;
                    }
                }
                else
                {
                    MessageBox.Show(InvalidCharacterError);
                    labelTextBox.SelectAll();
                }
            }
        }

        public Image ImageToThumbnail
        {
            get { return _imageSource; }
            
            private set
            {
                _imageSource = value;
                thumbnailPanel.BackgroundImage = CreateThumbnailImage(value, 50, 50);
            }
        }

        public string ImagePath
        {
            get { return _imageSourcePath; }
            set
            {
                _imageSourcePath = value;
                ImageToThumbnail = new Bitmap(value);
            }
        }
        # endregion

        # region Helper Functions
        // for names, we do not allow names involve '\' or '.'
        // Regex = [\\\.]
        private const string InvalidCharsRegex = "[\\\\\\.]";

        private double CalculateScalingRatio(Size oldSize, Size newSize)
        {
            double scalingRatio;

            if (oldSize.Width >= oldSize.Height)
            {
                scalingRatio = (double)newSize.Width / oldSize.Width;
            }
            else
            {
                scalingRatio = (double)newSize.Height / oldSize.Height;
            }

            return scalingRatio;
        }

        private Bitmap CreateThumbnailImage(Image oriImage, int width, int height)
        {
            var scalingRatio = CalculateScalingRatio(oriImage.Size, new Size(width, height));

            // calculate width and height after scaling
            var scaledWidth = (int)Math.Round(oriImage.Size.Width * scalingRatio);
            var scaledHeight = (int)Math.Round(oriImage.Size.Height * scalingRatio);

            // calculate left top corner position of the image in the thumbnail
            var scaledLeft = (width - scaledWidth) / 2;
            var scaledTop = (height - scaledHeight) / 2;

            // define drawing area
            var drawingRect = new Rectangle(scaledLeft, scaledTop, scaledWidth, scaledHeight);
            var thumbnail = new Bitmap(width, height);

            // here we set the thumbnail as the highest quality
            using (var thumbnailGraphics = Graphics.FromImage(thumbnail))
            {
                thumbnailGraphics.CompositingQuality = CompositingQuality.HighQuality;
                thumbnailGraphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                thumbnailGraphics.SmoothingMode = SmoothingMode.HighQuality;
                thumbnailGraphics.DrawImage(oriImage, drawingRect);
            }

            return thumbnail;
        }

        private void DeHighlight()
        {
            labelTextBox.BackColor = Color.FromKnownColor(KnownColor.Window);
            thumbnailPanel.BackColor = Color.FromKnownColor(KnownColor.Transparent);

            // if the name provided to the shape is not valid, and user de-focus the
            // current labled thumbnail, we shoud give the old name to the shape.
            if (!Verify(NameLable))
            {
                NameLable = _firstNameAssigned;
            }
        }

        private void Highlight()
        {
            labelTextBox.BackColor = Color.FromKnownColor(KnownColor.Highlight);
            thumbnailPanel.BackColor = Color.FromKnownColor(KnownColor.Highlight);
        }

        private void Initialize()
        {
            InitializeComponent();

            labelTextBox.TextChanged += OnNameLabelChanged;

            // let user specify the shape name
            labelTextBox.Enabled = true;
            labelTextBox.SelectAll();
        }

        private void RenameSource()
        {
            var oldName = Path.GetFileNameWithoutExtension(ImagePath);

            if (oldName != null)
            {
                ImagePath = ImagePath.Replace(oldName, NameLable);
            }
        }

        private bool Verify(string name)
        {
            var invalidChars = new Regex(InvalidCharsRegex);
            
            return !(string.IsNullOrWhiteSpace(name) || invalidChars.IsMatch(name));
        }
        # endregion

        # region API
        public void ToggleHighlight()
        {
            if (_isHighlighted)
            {
                DeHighlight();
            }
            else
            {
                Highlight();
            }

            _isHighlighted = !_isHighlighted;
        }

        public void EnableNameEdit()
        {
            // TODO: enable lable text box, start left mouse button hook
        }
        # endregion

        # region Event Handlers
        // delegates
        public delegate void NameLabelChangedHandler();

        // handlers
        public static event NameLabelChangedHandler NameLabelChanged;

        private void OnNameLabelChanged(object sender, EventArgs args)
        {
            // execute custom event handler if defined
            if (NameLabelChanged != null)
            {
                NameLabelChanged();
            }

            RenameSource();
        }
        # endregion
    }
}
