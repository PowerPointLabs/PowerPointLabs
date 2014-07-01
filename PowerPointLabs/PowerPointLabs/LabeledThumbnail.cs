using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using PPExtraEventHelper;

namespace PowerPointLabs
{
    public partial class LabeledThumbnail : UserControl
    {
        private bool _isHighlighted;
        private bool _isGoodName;
        private Image _imageSource;
        private string _imageSourcePath;
        private string _firstNameAssigned = string.Empty;

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
                    _isGoodName = true;

                    if (_firstNameAssigned == string.Empty)
                    {
                        _firstNameAssigned = value;
                    }
                }
                else
                {
                    MessageBox.Show(InvalidCharacterError);
                    labelTextBox.SelectAll();
                    _isGoodName = false;
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

        # region Constructors
        public LabeledThumbnail()
        {
            Initialize();
        }

        public LabeledThumbnail(string imagePath, string nameLable)
        {
            Initialize();

            NameLable = nameLable;

            var image = new Bitmap(imagePath);
            ImageToThumbnail = CreateThumbnailImage(image, 50, 50);
        }
        # endregion

        # region API
        public void StartNameEdit()
        {
            labelTextBox.Enabled = true;
            PPMouse.StartRegionClickHook();
        }

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

            PPMouse.Click += OnClickWhileEditing;

            // let user specify the shape name
            labelTextBox.Enabled = true;
            labelTextBox.SelectAll();
        }

        private bool PointInRectangle(Point point, Rectangle rect)
        {
            return point.X > rect.Left &&
                   point.X < rect.Right &&
                   point.Y > rect.Top &&
                   point.Y < rect.Bottom;
        }

        private void RenameSource()
        {
            var oldName = Path.GetFileNameWithoutExtension(ImagePath);

            if (oldName != null)
            {
                var newPath = ImagePath.Replace(oldName, NameLable);

                // rename the file on disk
                File.Move(ImagePath, newPath);

                // edit the image path on memory
                ImagePath = newPath;
            }
        }

        private bool Verify(string name)
        {
            var invalidChars = new Regex(InvalidCharsRegex);
            
            return !(string.IsNullOrWhiteSpace(name) || invalidChars.IsMatch(name));
        }
        # endregion

        # region Event Handlers
        private void OnClickWhileEditing(Point mousePosition)
        {
            var editingArea = labelTextBox.RectangleToScreen(labelTextBox.DisplayRectangle);

            // click outside the editing text box -> finish editing
            if (!PointInRectangle(mousePosition, editingArea))
            {
                NameLable = labelTextBox.Text;
                
                // if the name is accepted, end editing session and rename the file
                if (_isGoodName)
                {
                    PPMouse.StopRegionHook();
                    labelTextBox.Enabled = false;

                    RenameSource();
                }
            }
        }
        # endregion
    }
}
