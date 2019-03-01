using System;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media.Imaging;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ShapesLab.Views
{
    /// <summary>
    /// Interaction logic for CustomShapePaneItem.xaml
    /// </summary>
    public partial class CustomShapePaneItem : UserControl
    {

        private Bitmap image;

        public enum Status
        {
            Idle,
            Editing
        }

        #region Constructors

        public CustomShapePaneItem(string shapeName, string shapePath)
        {
            Initialize(shapeName, shapePath);
            /*
            editImage.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                Properties.Resources.SyncLabEditButton.GetHbitmap(),
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());
            pasteImage.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                Properties.Resources.SyncLabPasteButton.GetHbitmap(),
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions()); 
            deleteImage.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                Properties.Resources.SyncLabDeleteButton.GetHbitmap(),
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());*/

            //MouseDoubleClick += 
        }

        #endregion

        #region Properties

        public Bitmap Image
        {
            get
            {
                return image;
            }
            set
            {
                image = value;
                UpdateImage();
            }
        }

        public String Text
        {
            get
            {
                return textBox.Text;
            }
            set
            {
                if (!Verify(value))
                {
                    MessageBox.Show(Utils.ShapeUtil.IsShapeNameOverMaximumLength(value)
                                        ? CommonText.ErrorNameTooLong
                                        : CommonText.ErrorInvalidCharacter);

                    textBox.SelectAll();
                }
                textBox.Text = value;
                textBox.ToolTip = value;
            }
        }

        public bool Highlighted { get; set; }

        public string ImagePath { get; set; }

        public Status State { get; private set; }

        #endregion

        #region API

        public void RenameShape(string newShapeName)
        {
            string newPath = ImagePath.Replace(@"\" + Text, @"\" + newShapeName);
            Text = newShapeName;
            File.Move(ImagePath, newPath);
            ImagePath = newPath;
        }

        public void StartNameEdit()
        {
            //TODO
            /*
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
            */
        }

        public void FinishNameEdit()
        {
            //TODO
            /*
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

            string oldName = Path.GetFileNameWithoutExtension(ImagePath);

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
            */
        }

        # endregion

        #region Helper Functions
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

            textBox.MouseDoubleClick += (sender, e) => textBox.SelectAll();
            textBox.IsEnabledChanged += EnableChangedHandler;
            textBox.KeyDown += EnterKeyWhileEditing;
            textBox.LostFocus += NameLabelLostFocus;

            //TODO
            // This is required for the thumbnail labels to show
            //CustomPaintTextBox customPaintTextBox = new CustomPaintTextBox(labelTextBox);
        }

        private void Initialize(string shapeName, string shapePath)
        {
            Initialize();

            ImagePath = shapePath;

            // critical line, we need to free the reference to the image immediately after we've
            // finished thumbnail generation, else we could not modify (rename/ delete) the
            // image. Therefore, we use using keyword to ensure a collection.
            //TODO
            using (Bitmap bitmap = new Bitmap(ImagePath))
            {
                image = Utils.GraphicsUtil.CreateThumbnailImage(bitmap, 50, 50);
            }

            State = Status.Idle;
            //textBox.IsEnabled = false;
        }

        private bool IsDuplicateName(string newShapeName)
        {
            // if the name hasn't changed, we don't need to check for duplicate name
            // since the default name/ old name is confirmed unique.
            if (newShapeName == Text)
            {
                return false;
            }

            string newPath = ImagePath.Replace(newShapeName, Text);

            // if the new name has been used, the new name is not allowed
            if (File.Exists(newPath))
            {
                MessageBox.Show(CommonText.ErrorFileNameExist);
                return true;
            }

            return false;
        }

        private bool Verify(string name)
        {
            Regex invalidChars = new Regex(InvalidCharsRegex);

            return !(string.IsNullOrWhiteSpace(name) ||
                     invalidChars.IsMatch(name) ||
                     Utils.ShapeUtil.IsShapeNameOverMaximumLength(name));
        }

        #endregion

        #region Event Handlers

        private void UpdateImage()
        {
            // if image isn't set, fill the control with the label
            if (image == null)
            {
                imageBox.Visibility = Visibility.Hidden;
                col1.Width = new GridLength(0);
                return;
            }
            else
            {
                BitmapSource source = Imaging.CreateBitmapSourceFromHBitmap(
                                        image.GetHbitmap(),
                                        IntPtr.Zero,
                                        Int32Rect.Empty,
                                        BitmapSizeOptions.FromEmptyOptions());
                imageBox.Source = source;
                imageBox.Visibility = Visibility.Visible;
                col1.Width = new GridLength(60);
            }
        }

        private void EnableChangedHandler(object sender, DependencyPropertyChangedEventArgs e)
        {
            //TODO
            /*
            textBox.BackColor = Color.FromKnownColor(KnownColor.Window);

            if (labelTextBox.Enabled == false)
            {
                labelTextBox.Text = string.Empty;
            }
            else
            {
                labelTextBox.ForeColor = Color.Black;
                labelTextBox.Text = NameLabel;
            }
            */
        }

        private void EnterKeyWhileEditing(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Return)
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
                textBox.SelectAll();
                return;
            }
        }

        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            //TODO
            //parent.RemoveFormatItem(this);
        }

        //TODO
        private void OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ApplyFormatToSelected();
        }

        private void ApplyFormatToSelected()
        {
            MessageBox.Show(SyncLabText.ErrorShapeDeleted, SyncLabText.ErrorDialogTitle);
            this.StartNewUndoEntry();
        }

        #endregion
    }
}
