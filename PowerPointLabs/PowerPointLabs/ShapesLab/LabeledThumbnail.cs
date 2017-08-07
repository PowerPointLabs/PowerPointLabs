using System;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

using PowerPointLabs.TextCollection;
using PowerPointLabs.Views;
using TestInterface;

namespace PowerPointLabs.ShapesLab
{
    public partial class LabeledThumbnail : UserControl, IShapesLabLabeledThumbnail
    {
        private bool _nameFinishHandled;
        private bool _isGoodName;
        
        private string _nameLabel;

        public enum Status
        {
            Idle,
            Editing
        }

        # region Properties
        public bool Highlighed { get; set; }

        public string NameLabel
        {
            get
            {
                return _nameLabel;
            }
            set
            {
                if (Verify(value))
                {
                    _nameLabel = value;
                    labelTextBox.Text = value;
                    _isGoodName = true;
                }
                else
                {
                    MessageBox.Show(Utils.ShapeUtil.IsShapeNameOverMaximumLength(value)
                                        ? CommonText.ErrorNameTooLong
                                        : CommonText.ErrorInvalidCharacter);

                    labelTextBox.SelectAll();
                    _isGoodName = false;
                }
            }
        }

        public string ImagePath { get; set; }

        public Status State { get; private set; }
        # endregion

        # region Constructors
        public LabeledThumbnail()
        {
            Initialize();
        }

        public LabeledThumbnail(string imagePath, string nameLable)
        {
            Initialize(imagePath, nameLable);
        }
        # endregion

        # region API
        public void DeHighlight()
        {
            motherPanel.BackColor = Color.FromKnownColor(KnownColor.Window);
            thumbnailPanel.BackColor = Color.FromKnownColor(KnownColor.Window);
            labelTextBox.BackColor = Color.FromKnownColor(KnownColor.Window);
            labelTextBox.ForeColor = Color.Black;

            Highlighed = false;
        }

        public void Highlight()
        {
            motherPanel.BackColor = Color.FromKnownColor(KnownColor.LightBlue);
            thumbnailPanel.BackColor = Color.FromKnownColor(KnownColor.LightBlue);

            if (!labelTextBox.Enabled)
            {
                labelTextBox.BackColor = Color.FromKnownColor(KnownColor.LightBlue);
            }

            Highlighed = true;
        }

        public void ToggleHighlight()
        {
            if (Highlighed)
            {
                DeHighlight();
            }
            else
            {
                Highlight();
            }
        }

        public void StartNameEdit()
        {
            // add the text box
            if (!motherPanel.Controls.Contains(labelTextBox))
            {
                motherPanel.Controls.Add(labelTextBox);
            }

            _nameFinishHandled = false;
            State = Status.Editing;
            
            Highlight();

            labelTextBox.Enabled = true;
            labelTextBox.Focus();
            labelTextBox.SelectAll();

            SetToolTip("Editing...");
        }

        public void FinishNameEdit()
        {
            // since messagebox will trigger LostFocus event, this method
            // has chance to be triggered mulitple times. To avoid this,
            // a flag will be set on the first time the function is called,
            // and skip the function by checking if the flag has been set.
            if (_nameFinishHandled)
            {
                return;
            }

            _nameFinishHandled = true;
            NameLabel = labelTextBox.Text;

            var oldName = Path.GetFileNameWithoutExtension(ImagePath);

            if (_isGoodName &&
                !IsDuplicateName(oldName))
            {
                State = Status.Idle;

                labelTextBox.Enabled = false;
                NameEditFinish(this, oldName);

                SetToolTip(NameLabel);
            }
            else
            {
                StartNameEdit();
            }
        }

        public void RenameWithoutEdit(string newName)
        {
            labelTextBox.Enabled = true;
            NameLabel = newName;
            labelTextBox.Enabled = false;

            SetToolTip(NameLabel);
        }
        # endregion

        # region Helper Functions
        // for names, we do not allow name involves
        // < (less than)
        // > (greater than)
        // : (colon)
        // " (double quote)
        // / (forward slash)
        // \ (backslash)
        // | (vertical bar or pipe)
        // ? (question mark)
        // * (asterisk)

        // Regex = [<>:"/\\|?*]
        private const string InvalidCharsRegex = "[<>:\"/\\\\|?*]";

        private void Initialize()
        {
            InitializeComponent();

            motherPanel.MouseDown += (sender, e) => Click(this, e);
            motherPanel.DoubleClick += (sender, e) => DoubleClick(this, e);

            thumbnailPanel.MouseDown += (sender, e) => Click(this, e);
            thumbnailPanel.DoubleClick += ThumbnailPanelDoubleClick;

            labelTextBox.DoubleClick += (sender, e) => labelTextBox.SelectAll();
            labelTextBox.EnabledChanged += EnableChangedHandler;
            labelTextBox.KeyPress += EnterKeyWhileEditing;
            labelTextBox.LostFocus += NameLabelLostFocus;
            
            // This is required for the thumbnail labels to show
            CustomPaintTextBox customPaintTextBox = new CustomPaintTextBox(labelTextBox);
        }

        private void Initialize(string imagePath, string nameLable)
        {
            Initialize();

            NameLabel = nameLable;
            SetToolTip(NameLabel);

            ImagePath = imagePath;

            // critical line, we need to free the reference to the image immediately after we've
            // finished thumbnail generation, else we could not modify (rename/ delete) the
            // image. Therefore, we use using keyword to ensure a collection.
            using (var bitmap = new Bitmap(ImagePath))
            {
                thumbnailPanel.BackgroundImage = Utils.GraphicsUtil.CreateThumbnailImage(bitmap, 50, 50);
            }

            State = Status.Idle;
            labelTextBox.Enabled = false;
        }

        private bool IsDuplicateName(string oldName)
        {
            // if the name hasn't changed, we don't need to check for duplicate name
            // since the default name/ old name is confirmed unique.
            if (oldName == NameLabel)
            {
                return false;
            }

            var newPath = ImagePath.Replace(oldName, NameLabel);

            // if the new name has been used, the new name is not allowed
            if (File.Exists(newPath))
            {
                MessageBox.Show(CommonText.ErrorFileNameExist);
                return true;
            }

            return false;
        }

        private void SetToolTip(string toolTip)
        {
            nameLabelToolTip.SetToolTip(motherPanel, toolTip);
            nameLabelToolTip.SetToolTip(thumbnailPanel, toolTip);
        }

        private bool Verify(string name)
        {
            var invalidChars = new Regex(InvalidCharsRegex);
            
            return !(string.IsNullOrWhiteSpace(name) ||
                     invalidChars.IsMatch(name) ||
                     Utils.ShapeUtil.IsShapeNameOverMaximumLength(name));
        }
        # endregion

        # region Event Handlers
        public delegate void ClickEventDelegate(object sender, MouseEventArgs e);
        public delegate void DoubleClickEventDelegate(object sender, EventArgs e);
        public delegate void NameEditFinishEventDelegate(object sender, string oldName);

        public new event ClickEventDelegate Click;
        public new event DoubleClickEventDelegate DoubleClick;
        public event NameEditFinishEventDelegate NameEditFinish;

        private void EnableChangedHandler(object sender, EventArgs e)
        {
            labelTextBox.BackColor = Color.FromKnownColor(KnownColor.Window);

            if (labelTextBox.Enabled == false)
            {
                labelTextBox.Text = string.Empty;
            }
            else
            {
                labelTextBox.ForeColor = Color.Black;
                labelTextBox.Text = NameLabel;
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

        private void NameLabelLostFocus(object sender, EventArgs args)
        {
            FinishNameEdit();
        }

        private void ThumbnailPanelDoubleClick(object sender, EventArgs e)
        {
            if (State == Status.Editing)
            {
                labelTextBox.SelectAll();
                return;
            }

            DoubleClick(this, e);
        }
        # endregion
    }
}
