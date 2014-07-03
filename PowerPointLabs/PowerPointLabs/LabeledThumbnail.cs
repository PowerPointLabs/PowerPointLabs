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
        private bool _isGoodName;
        private Image _imageSource;
        private string _imageSourcePath;

        public enum Status
        {
            Idle,
            Editing
        }

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

        public Status State { get; private set; }
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
            ImagePath = imagePath;
        }
        # endregion

        # region API
        public void DeHighlight()
        {
            motherPanel.BackColor = Color.FromKnownColor(KnownColor.Window);
            thumbnailPanel.BackColor = Color.FromKnownColor(KnownColor.Transparent);

            // dehighlight will hard-disable the text box editing
            labelTextBox.Enabled = false;
            State = Status.Idle;
        }

        public void Highlight()
        {
            motherPanel.BackColor = Color.FromKnownColor(KnownColor.Highlight);
            thumbnailPanel.BackColor = Color.FromKnownColor(KnownColor.Highlight);
        }

        public void StartNameEdit()
        {
            State = Status.Editing;
            
            Highlight();
            labelTextBox.Enabled = true;
            labelTextBox.Focus();
            labelTextBox.SelectAll();
        }

        public void FinishNameEdit()
        {
            NameLable = labelTextBox.Text;

            bool sameName;

            if (_isGoodName &&
                RenameSource(out sameName))
            {
                State = Status.Idle;

                labelTextBox.Enabled = false;
                NameChangedNotify(this, !sameName);
            }
            else
            {
                StartNameEdit();
            }
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
        private const string FileNameExistError = "File name is already used";

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

        private void Initialize()
        {
            InitializeComponent();

            motherPanel.Click += (sender, e) => Click(this, e);
            motherPanel.DoubleClick += (sender, e) => DoubleClick(this, e);
            
            foreach (Control control in motherPanel.Controls)
            {
                control.Click += (sender, e) => Click(this, e);
                control.DoubleClick += (sender, e) => DoubleClick(this, e);
            }

            labelTextBox.EnabledChanged += EnableChangedHandler;
            labelTextBox.KeyPress += EnterKeyWhileEditing;
            labelTextBox.LostFocus += NameLableLostFocus;

            // let user specify the shape name
            State = Status.Editing;
            labelTextBox.Enabled = true;
            labelTextBox.SelectAll();
        }

        private bool RenameSource(out bool sameName)
        {
            var oldName = Path.GetFileNameWithoutExtension(ImagePath);
            sameName = false;

            if (oldName != null)
            {
                // name doesn't change
                if (oldName == NameLable)
                {
                    sameName = true;
                    return true;
                }
                
                var newPath = ImagePath.Replace(oldName, NameLable);

                if (File.Exists(newPath))
                {
                    MessageBox.Show(FileNameExistError);
                    return false;
                }

                // rename the file on disk
                File.Move(ImagePath, newPath);
                ImagePath = newPath;
            }

            return true;
        }

        private bool Verify(string name)
        {
            var invalidChars = new Regex(InvalidCharsRegex);
            
            return !(string.IsNullOrWhiteSpace(name) || invalidChars.IsMatch(name));
        }
        # endregion

        # region Event Handlers
        public delegate void ClickEventDelegate(object sender, EventArgs e);

        public delegate void DoubleClickEventDelegate(object sender, EventArgs e);

        public delegate void NameChangedNotifyEventDelegate(object sender, bool nameChanged);

        public new event ClickEventDelegate Click;
        public new event DoubleClickEventDelegate DoubleClick;
        public event NameChangedNotifyEventDelegate NameChangedNotify;

        private void EnableChangedHandler(object sender, EventArgs e)
        {
            if (labelTextBox.Enabled == false)
            {
                labelTextBox.BackColor = Color.FromKnownColor(KnownColor.Window);
            }
        }

        private void EnterKeyWhileEditing(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Return)
            {
                FinishNameEdit();
                e.Handled = true;
            }
        }

        private void NameLableLostFocus(object sender, EventArgs args)
        {
            FinishNameEdit();
        }
        # endregion
    }
}
