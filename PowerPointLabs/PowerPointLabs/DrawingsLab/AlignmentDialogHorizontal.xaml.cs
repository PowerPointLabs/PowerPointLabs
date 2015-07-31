using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using PowerPointLabs.DataSources;

namespace PowerPointLabs.DrawingsLab
{
    /// <summary>
    /// Interaction logic for AlignmentDialogHorizontal.xaml
    /// </summary>
    public partial class AlignmentDialogHorizontal : Window
    {
        private DrawingsLabAlignmentDataSource dataSource;
        
        public AlignmentDialogHorizontal()
        {
            InitializeComponent();

            InitialiseDataSource();
        }

        private void InitialiseDataSource()
        {
            dataSource = FindResource("DataSource") as DrawingsLabAlignmentDataSource;
        }

        private void ButtomDialogOk_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
        }

        public float SourceAnchor
        {
            get { return dataSource.SourceAnchor; }
        }

        public float TargetAnchor
        {
            get { return dataSource.TargetAnchor; }
        }
    }

}
