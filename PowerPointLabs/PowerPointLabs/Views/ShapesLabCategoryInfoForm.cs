using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace PowerPointLabs.Views
{
    public partial class ShapesLabCategoryInfoForm : Form
    {
        private const string InvalidCharsRegex = "[<>:\"/\\\\|?*]";

        # region Properties
        public Option UserOption { get; private set; }

        public string CategoryName { get; private set; }
        # endregion

        # region Constructor
        public ShapesLabCategoryInfoForm()
        {
            InitializeComponent();
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

            if (Verify(categoryName))
            {
                CategoryName = categoryNameBox.Text;
                UserOption = Option.Ok;
                Dispose();
            }
            else
            {
                MessageBox.Show(categoryName.Length > 255
                                    ? TextCollection.ErrorNameTooLong
                                    : TextCollection.ErrorInvalidCharacter);
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
        private bool Verify(string name)
        {
            var invalidChars = new Regex(InvalidCharsRegex);

            return !(string.IsNullOrWhiteSpace(name) ||
                     invalidChars.IsMatch(name) ||
                     name.Length > 255);
        }
        # endregion
    }
}
