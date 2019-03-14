using System;
using System.Drawing;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.TooltipsLab.Views
{
    /// <summary>
    /// Interaction logic for TooltipsLabSettingsAnimationEntry.xaml
    /// </summary>
    public partial class TooltipsLabSettingsAnimationEntry : UserControl
    {
        private MsoAnimEffect type;

        #region Constructors

        public TooltipsLabSettingsAnimationEntry(MsoAnimEffect type, Bitmap image)
        {
            InitializeComponent();
            Type = type;
            RemoveImage();
            //TODO - uncomment when proper images are added
            /*
            imageBox.Source = Imaging.CreateBitmapSourceFromHBitmap(
                image.GetHbitmap(),
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());*/
        }

        #endregion

        #region Properties

        public MsoAnimEffect Type
        {
            get
            {
                return type;
            }
            set
            {
                type = value;
                string nameForDisplay = RawNameToDisplayName(value.ToString());
                textBlock.Text = nameForDisplay;
                ToolTip = nameForDisplay;
            }
        }

        #endregion

        #region Helper functions

        private void RemoveImage()
        {
            imageBox.IsEnabled = false;
            imageBox.Height = 0;

            Height = 20;
            textBlock.Margin = new Thickness(0, 5, 0, 5);
        }

        private string RawNameToDisplayName(string name)
        {
            string trimmed = name.Replace(TooltipsLabConstants.AnimationNameHeader, "");

            if (string.IsNullOrWhiteSpace(trimmed))
            {
                return string.Empty;
            }

            StringBuilder newText = new StringBuilder(trimmed.Length * 2);
            newText.Append(trimmed[0]);
            for (int i = 1; i < trimmed.Length; i++)
            {
                if (char.IsUpper(trimmed[i]))
                {
                    newText.Append(' ');
                }
                newText.Append(trimmed[i]);
            }
            return newText.ToString();
        }

        #endregion
    }
}