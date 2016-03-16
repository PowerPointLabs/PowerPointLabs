using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace PowerPointLabs.Views
{
    public partial class ShapesLabCategoryInfoForm : Form
    {
#pragma warning disable 0618
        private const string InvalidCharsRegex = "[<>:\"/\\\\|?*]";

        # region Properties
        public Option UserOption { get; private set; }

        public string CategoryName { get; private set; }
        # endregion

        # region Constructor
        public ShapesLabCategoryInfoForm(string intialName)
        {
            InitializeComponent();

            if (!string.IsNullOrEmpty(intialName))
            {
                categoryNameBox.Text = intialName;
                categoryNameBox.SelectAll();
            }
        }
        # endregion

        # region API
        public enum Option
        {
            Ok,
            Cancel
        }
        # endregion

        # region Event Handlers
        private void OkButtonClick(object sender, EventArgs e)
        {
            var categoryName = categoryNameBox.Text;

            if (VerifyName(categoryName) &&
                VerifyCategory(categoryName))
            {
                CategoryName = categoryNameBox.Text;
                UserOption = Option.Ok;
                Dispose();
            }
        }

        private void CancelButtonClick(object sender, EventArgs e)
        {
            UserOption = Option.Cancel;

            Dispose();
        }

        private void CategoryNameBoxKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == (char)Keys.Return)
            {
                e.Handled = true;
                OkButtonClick(sender, null);
            }
        }
        # endregion

        # region Helper Functions
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
            if (Globals.ThisAddIn.ShapePresentation.HasCategory(name))
            {
                MessageBox.Show(TextCollection.CustomShapeDuplicateCategoryNameError);
                return false;
            }

            return true;
        }
        # endregion
    }
}
