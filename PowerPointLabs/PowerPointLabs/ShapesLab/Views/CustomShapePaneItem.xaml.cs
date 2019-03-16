using System;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
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
        private CustomShapePaneWPF parent;
        private Status editStatus;
        private bool hasJustExitedFromTextBox = false;
        private string shapeName;

        public enum Status
        {
            Idle,
            Editing
        }

        #region Constructors

        public CustomShapePaneItem(CustomShapePaneWPF parent, string shapeName, string shapePath, bool isReadyForEditing)
        {
            Initialize(isReadyForEditing);
            this.parent = parent;
            ImagePath = shapePath;
            this.shapeName = shapeName;
            textBox.Text = shapeName;
            ToolTip = shapeName;

            // critical line, we need to free the reference to the image immediately after we've
            // finished thumbnail generation, else we could not modify (rename/ delete) the
            // image. Therefore, we use using keyword to ensure a collection.
            using (Bitmap bitmap = new Bitmap(shapePath))
            {
                Image = Utils.GraphicsUtil.CreateThumbnailImage(bitmap, 50, 50);
            }
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

        public Status EditStatus
        {
            get
            {
                return editStatus;
            }
            set
            {
                if (value == editStatus)
                {
                    return;
                }
                editStatus = value;
                switch (value)
                {
                    case Status.Editing:
                        StartNameEdit();
                        break;
                    case Status.Idle:
                        FinishNameEdit();
                        break;
                    default:
                        break;
                }
            }
        }

        public string ImagePath { get; set; }

        public string Text
        {
            get
            {
                return shapeName;
            }
        }

        #endregion

        #region API

        /// <summary>
        /// Updates UI
        /// </summary>
        /// <param name="newShapeName"></param>
        public void SyncRenameShape(string newShapeName)
        {
            textBox.Text = newShapeName;
            ToolTip = newShapeName;
            shapeName = newShapeName;
        }

        public void UnfocusTextBox()
        {
            //unfocus from the textbox
            DependencyObject scope = FocusManager.GetFocusScope(textBox);
            FocusManager.SetFocusedElement(scope, this as IInputElement);
        }

        public void StartNameEdit()
        {
            SetEditableTextBox();
        }

        public void FinishNameEdit()
        {
            SetReadOnlyTextBox();
            RenameShape(textBox.Text);
        }

        #endregion

        #region Context Menu

        private void AddShapeClick(object sender, RoutedEventArgs e)
        {
            parent.AddShapesToSlide();
        }

        private void EditShapeClick(object sender, RoutedEventArgs e)
        {
            EditStatus = Status.Editing;
        }

        private void MoveShapeClick(object sender, RoutedEventArgs e)
        {
            ShapesLabCategoryInfoDialogBox categoryInfoDialog = new ShapesLabCategoryInfoDialogBox(string.Empty, true);
            categoryInfoDialog.DialogConfirmedHandler += parent.MoveShapes;
            categoryInfoDialog.ShowDialog();
        }

        private void DeleteShapeClick(object sender, RoutedEventArgs e)
        {
            parent.RemoveAllSelectedShapes();
        }

        #endregion

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

        private void Initialize(bool isReadyForEdit)
        {
            InitializeComponent();

            canvas.MouseLeftButtonDown += TextBoxContainerClick;
            canvas.Focusable = true;
            textBox.KeyDown += EnterKeyWhileEditing;
            textBox.LostFocus += TextBoxLostFocus;

            if (isReadyForEdit)
            {
                SetEditableTextBox();
            }
            else
            {
                SetReadOnlyTextBox();
            }
        }

        private void SetReadOnlyTextBox()
        {
            textBox.IsEnabled = false;
            textBox.IsReadOnly = true;
            textBox.IsHitTestVisible = false;
            textBox.CaretIndex = 0;
            textBox.BorderThickness = new Thickness(0, 0, 0, 0);
            textBox.Background = Background;
        }

        private void SetEditableTextBox()
        {
            textBox.IsEnabled = true;
            textBox.IsReadOnly = false;
            textBox.IsHitTestVisible = true;
            textBox.CaretIndex = 0;
            textBox.BorderThickness = new Thickness(2, 2, 2, 2);
            textBox.Background = System.Windows.Media.Brushes.White;
            textBox.Focus();
            textBox.SelectAll();
        }

        private bool HasNameChanged(string newShapeName)
        {
            return newShapeName == shapeName;
        }

        private bool IsDuplicateName(string newShapeName)
        {
            string newPath = ImagePath.Replace(shapeName, newShapeName);

            // if the new name has been used, the new name is not allowed
            if (File.Exists(newPath))
            {
                MessageBox.Show(CommonText.ErrorFileNameExist);
                return true;
            }
            return false;
        }

        private bool IsValidName(string name)
        {
            Regex invalidChars = new Regex(InvalidCharsRegex);

            return !(string.IsNullOrWhiteSpace(name) ||
                     invalidChars.IsMatch(name) ||
                     Utils.ShapeUtil.IsShapeNameOverMaximumLength(name));
        }

        private void RenameShape(string newShapeName)
        {
            if (HasNameChanged(newShapeName))
            {
                return;
            }
            if (!IsValidName(newShapeName))
            {
                MessageBox.Show(Utils.ShapeUtil.IsShapeNameOverMaximumLength(newShapeName)
                                    ? CommonText.ErrorNameTooLong
                                    : CommonText.ErrorInvalidCharacter);
                textBox.Text = shapeName;
                EditStatus = Status.Editing;
                return;
            }

            if (IsDuplicateName(newShapeName))
            {
                EditStatus = Status.Editing;
                return;
            }
            //Update image
            string newPath = ImagePath.Replace(@"\" + shapeName, @"\" + newShapeName);
            if (File.Exists(ImagePath))
            {
                File.Move(ImagePath, newPath);
                this.GetAddIn().ShapePresentation.RenameShape(shapeName, newShapeName);
            }
            ImagePath = newPath;

            ShapesLabUtils.SyncShapeRename(this.GetAddIn(), shapeName, newShapeName, parent.CurrentCategory);
            parent.RenameCustomShape(shapeName, newShapeName);
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

        private void EnterKeyWhileEditing(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Return)
            {
                UnfocusTextBox();
                e.Handled = true;
            }
        }

        private void TextBoxLostFocus(object sender, EventArgs args)
        {
            hasJustExitedFromTextBox = true;
            EditStatus = Status.Idle;
        }
        
        private void OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            e.Handled = true;
            if (sender == canvas && e.ClickCount < 2)
            {
                return;
            }
            EditStatus = Status.Idle;
            parent.AddShapesToSlide();
        }

        private void TextBoxContainerClick(object sender, MouseButtonEventArgs e)
        {
            if (!parent.IsShapeSelected(this))
            {
                return;
            }
            canvas.Focus();
            new Thread(delegate ()
            {
                Thread.Sleep(1000);
                Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate ()
                {
                    FocusTextBox();
                }));
            }).Start();

            OnMouseDoubleClick(sender, e);
        }

        private void FocusTextBox()
        {
            if (hasJustExitedFromTextBox)
            {
                hasJustExitedFromTextBox = false;
                return;
            }
            if (!canvas.IsFocused)
            {
                return;
            }
            EditStatus = Status.Editing;
            SetEditableTextBox();
        }

        private void ClickOutsideTextBox(object sender, MouseButtonEventArgs e)
        {
            UnfocusTextBox();
        }

        #endregion
    }
}
