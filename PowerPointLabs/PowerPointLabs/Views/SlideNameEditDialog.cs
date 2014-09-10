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
    public partial class SlideNameEditDialog : Form
    {
        private const string InvalidCharsRegex = "[<>:\"/\\\\|?*]";

        # region Properties
        public Option UserOption { get; private set; }

        public string SlideName { get; private set; }
        # endregion

        # region Constructor
        public SlideNameEditDialog()
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
            var slideName = slideNameBox.Text;

            if (VerifyName(slideName) &&
                VerifySlides(slideName))
            {
                SlideName = slideName;
                UserOption = Option.Ok;
                Dispose();
            }
        }

        private void CancelButtonClick(object sender, EventArgs e)
        {
            UserOption = Option.Cancel;

            Dispose();
        }

        private void SlieNameBoxKeyDown(object sender, KeyEventArgs e)
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

        private bool VerifySlides(string name)
        {
            if (Models.PowerPointCurrentPresentationInfo.Slides.Any(slide => slide.Name == name))
            {
                MessageBox.Show(TextCollection.SlideNameEditDuplicateSlideNameError);
                return false;
            }

            return true;
        }
        # endregion
    }
}
