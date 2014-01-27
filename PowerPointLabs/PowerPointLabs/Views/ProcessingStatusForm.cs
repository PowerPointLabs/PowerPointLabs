using System.Windows.Forms;

namespace AudioGen.Views
{
    public partial class ProcessingStatusForm : Form
    {
        private delegate void ProgressDelegate(int percentage);

        private delegate void SlideNumberDelegate(int current, int total);
        public ProcessingStatusForm()
        {
            InitializeComponent();
        }

        public void UpdateProgress(int percentage)
        {
            if (progressBar.InvokeRequired)
            {
                Invoke(new ProgressDelegate(UpdateProgress), new object[] {percentage});
            }
            else
            {
                progressBar.Value = percentage;
            }
        }

        public void UpdateSlideNumber(int current, int total)
        {
            if (slideNumber.InvokeRequired)
            {
                Invoke(new SlideNumberDelegate(UpdateSlideNumber), new object[] {current, total});
            }
            else
            {
                slideNumber.Text = current + "/" + total;
            }
        }
    }
}
