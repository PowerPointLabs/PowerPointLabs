using System.Text.RegularExpressions;
using System.Windows;

using PowerPointLabs.ActionFramework.Common.Extension;

namespace PowerPointLabs.Views
{
    /// <summary>
    /// Interaction logic for ShapesLabCategoryInfoDialogBox.xaml
    /// </summary>
    public partial class ShapesLabCategoryInfoDialogBox
    {
        private const string InvalidCharsRegex = "[<>:\"/\\\\|?*]";

        public enum Option
        {
            Ok,
            Cancel
        }

        public Option UserOption { get; private set; }
        public string CategoryName { get; private set; }
        

        public ShapesLabCategoryInfoDialogBox(string name)
        {
            InitializeComponent();
            
            if (!string.IsNullOrEmpty(name))
            {
                nameInput.Text = name;
                nameInput.SelectAll();
            }

            UserOption = Option.Cancel;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            string categoryName = nameInput.Text;

            if (VerifyName(categoryName) &&
                VerifyCategory(categoryName))
            {
                CategoryName = nameInput.Text;
                UserOption = Option.Ok;
                Close();
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            UserOption = Option.Cancel;
            Close();
        }

        #region Helper Functions
        private bool VerifyName(string name)
        {
            if (name.Length > 255)
            {
                MessageBox.Show(TextCollection.ErrorNameTooLong);
                return false;
            }

            var invalidChars = new Regex(InvalidCharsRegex);

            if (string.IsNullOrWhiteSpace(name) ||
                invalidChars.IsMatch(name))
            {
                MessageBox.Show(TextCollection.ErrorInvalidCharacter);
                return false;
            }

            return true;
        }

        private bool VerifyCategory(string name)
        {
            if (this.GetAddIn().ShapePresentation.HasCategory(name))
            {
                MessageBox.Show(TextCollection.CustomShapeDuplicateCategoryNameError);
                return false;
            }

            return true;
        }
        # endregion
    }
}
