using System;
using System.Drawing;
using System.IO;
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
                string nameForDisplay = value.ToString().Replace(
                    TooltipsLabConstants.AnimationNameHeader, "");
                textBlock.Text = nameForDisplay;
                ToolTip = nameForDisplay;
            }
        }

        #endregion
    }
}