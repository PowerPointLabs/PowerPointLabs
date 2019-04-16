using System.Reflection;

using MahApps.Metro.Controls.Dialogs;

using PowerPointLabs.ActionFramework.Common.Log;

namespace PowerPointLabs.PictureSlidesLab.Views
{
    partial class PictureSlidesLabWindow
    {
        private readonly TextBlockDialog _errorTextDialog = new TextBlockDialog();

        private void InitErrorTextDialog()
        {
            _errorTextDialog.GetType()
                    .GetProperty("OwningWindow", BindingFlags.Instance | BindingFlags.NonPublic)
                    .SetValue(_errorTextDialog, this, null);

            _errorTextDialog.OnOkButtonClick += () =>
            {
                _errorTextDialog.CloseDialog();
                this.HideMetroDialogAsync(_errorTextDialog, MetroDialogOptions);
            };
            Logger.Log("PSL init ErrorTextDialog done");
        }

        private void ShowErrorTextDialog(string text)
        {
            if (_errorTextDialog.IsOpen)
            {
                return;
            }

            _errorTextDialog.DialogTitleProperty.Text = "Error";
            _errorTextDialog.DialogTextBlockProperty.Text = text;
            _errorTextDialog.OpenDialog();
            this.ShowMetroDialogAsync(_errorTextDialog, MetroDialogOptions);
        }
    }
}
