using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ShapesLab.Views
{
    /// <summary>
    /// Interaction logic for ShapesLabCategoryInfoDialogBox.xaml
    /// </summary>
    public partial class ShapesLabCategoryInfoDialogBox
    {
        public delegate void DialogConfirmedDelegate(string categoryName);
        public DialogConfirmedDelegate DialogConfirmedHandler { get; set; }

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
        // Windows reserved file names.

        // Regex = ^.*(?i)\bCON|PRN|AUX|NUL|COM1|COM2|COM3|COM4|COM5|COM6|COM7|COM8|COM9|LPT1|LPT2|LPT3|LPT4|LPT5|LPT6|LPT7|LPT8|LPT9\b.*$|[<>:"/\\|?*]
        private const string InvalidCharsRegex = "^.*(?i)(CON|PRN|AUX|NUL|COM1|COM2|COM3|COM4|COM5|COM6|COM7|COM8|COM9|LPT1|LPT2|LPT3|LPT4|LPT5|LPT6|LPT7|LPT8|LPT9)(?-i).*$|[<>:\"/\\\\|?*]";

        public ShapesLabCategoryInfoDialogBox()
        {
            InitializeComponent();
        }

        public ShapesLabCategoryInfoDialogBox(string categoryName)
            : this()
        {
            if (!string.IsNullOrEmpty(categoryName))
            {
                nameInput.Text = categoryName;
                nameInput.SelectAll();
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            string name = nameInput.Text;

            if (VerifyName(name) && VerifyCategory(name))
            {
                DialogConfirmedHandler(name);
                Close();
            }
        }

        #region Helper Functions
        private bool VerifyName(string name)
        {
            if (Utils.ShapeUtil.IsShapeNameOverMaximumLength(name))
            {
                MessageBox.Show(CommonText.ErrorNameTooLong);
                return false;
            }

            Regex invalidChars = new Regex(InvalidCharsRegex);

            if (string.IsNullOrWhiteSpace(name) || invalidChars.IsMatch(name))
            {
                MessageBox.Show(CommonText.ErrorInvalidCharacter);
                return false;
            }

            return true;
        }

        private bool VerifyCategory(string name)
        {
            if (this.GetAddIn().ShapePresentation.HasCategory(name))
            {
                MessageBox.Show(ShapesLabText.ErrorDuplicateCategoryName);
                return false;
            }

            return true;
        }
        # endregion
    }
}
