using System;
using System.Drawing;
using System.Windows.Forms;

namespace PowerPointLabs.Views
{
    public partial class BulletAgendaSettingsDialog : Form
    {
        # region Event
        public delegate void UpdateSettingsDelegate(Color highlightColor, Color dimColor, Color defaultColor);
        public event UpdateSettingsDelegate SettingsHandler;
        # endregion

        # region Constructor
        public BulletAgendaSettingsDialog(Color highlightColor, Color dimColor, Color defaultColor)
        {
            InitializeComponent();

            higlightColorBox.BackColor = highlightColor;
            dimColorBox.BackColor = dimColor;
            defaultColorBox.BackColor = defaultColor;
        }
        # endregion

        # region Event Handlers
        private void DefaultColorBoxClick(object sender, EventArgs e)
        {
            var colorPicker = new ColorDialog
            {
                Color = defaultColorBox.BackColor,
                FullOpen = true
            };

            if (colorPicker.ShowDialog() != DialogResult.Cancel)
            {
                defaultColorBox.BackColor = colorPicker.Color;
            }
        }

        private void DimColorBoxClick(object sender, EventArgs e)
        {
            var colorPicker = new ColorDialog
            {
                Color = dimColorBox.BackColor,
                FullOpen = true
            };

            if (colorPicker.ShowDialog() != DialogResult.Cancel)
            {
                dimColorBox.BackColor = colorPicker.Color;
            }
        }

        private void HiglightColorBoxClick(object sender, EventArgs e)
        {
            var colorPicker = new ColorDialog
            {
                Color = higlightColorBox.BackColor,
                FullOpen = true
            };

            if (colorPicker.ShowDialog() != DialogResult.Cancel)
            {
                higlightColorBox.BackColor = colorPicker.Color;
            }
        }

        private void OkButtonClick(object sender, EventArgs e)
        {
            SettingsHandler(higlightColorBox.BackColor, dimColorBox.BackColor, defaultColorBox.BackColor);
            
            Close();
        }
        # endregion
    }
}
