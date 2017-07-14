using System.Text.RegularExpressions;
using System.Windows;

using PowerPointLabs.ActionFramework.Common.Extension;

using Forms = System.Windows.Forms;

namespace PowerPointLabs.ShapesLab.Views
{
    /// <summary>
    /// Interaction logic for ShapesLabCategoryInfoDialogBox.xaml
    /// </summary>
    public partial class ShapesLabCategoryInfoDialogBox
    {
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
        
        public Forms.DialogResult Result { get; private set; }
        public string CategoryName { get; private set; }
        

        public ShapesLabCategoryInfoDialogBox(string name)
        {
            InitializeComponent();
            
            if (!string.IsNullOrEmpty(name))
            {
                nameInput.Text = name;
                nameInput.SelectAll();
            }

            Result = Forms.DialogResult.OK;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            string categoryName = nameInput.Text;

            if (VerifyName(categoryName) && VerifyCategory(categoryName))
            {
                CategoryName = nameInput.Text;
                Result = Forms.DialogResult.OK;
                Close();
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Result = Forms.DialogResult.Cancel;
            Close();
        }

        #region Helper Functions
        private bool VerifyName(string name)
        {
            if (Utils.Graphics.IsShapeNameOverMaximumLength(name))
            {
                MessageBox.Show(TextCollection.ErrorNameTooLong);
                return false;
            }

            var invalidChars = new Regex(InvalidCharsRegex);

            if (string.IsNullOrWhiteSpace(name) || invalidChars.IsMatch(name))
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
