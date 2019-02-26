using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.ELearningWorkspace.Model;
using PowerPointLabs.ELearningLab.Service;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.Views
{
    /// <summary>
    /// Interaction logic for ProcessingStatusForm.xaml
    /// </summary>
    public partial class ProcessingStatusForm
    {
        private int totalValue;
        private BackgroundWorker worker;
        private BackgroundWorkerType workerType;
        public ProcessingStatusForm(int totalValue, BackgroundWorkerType workerType)
        {
            InitializeComponent();
            Dispatcher.Invoke(() => { progressBar.Value = 0; });
            label.Content = string.Format(ELearningLabText.ProgressStatusLabelFormat, 0);
            this.totalValue = totalValue;
            this.workerType = workerType;
            worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.WorkerSupportsCancellation = true;
            switch (workerType)
            {
                case BackgroundWorkerType.AudioGenerationService:
                    worker.DoWork += Worker_DoWorkForAudioGenerationService;
                    break;
                case BackgroundWorkerType.ELearningLabService:
                    worker.DoWork += Worker_DoWorkForELearningLabService;
                    break;
                default:
                    break;
            }
            worker.ProgressChanged += Worker_ProgressChanged;
            worker.RunWorkerCompleted += Worker_RunWorkerCompleted;
            worker.RunWorkerAsync();
        }

        private void Worker_DoWorkForELearningLabService(object sender, DoWorkEventArgs e)
        {
            for (int i = 0; i < totalValue; i++)
            {
                if (worker.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }
                int percentage = (int)Math.Round(((double)i + 1) / totalValue * 100);
                ELearningService.SyncAppearEffectAnimationsForSelfExplanationItem(i);
                (sender as BackgroundWorker).ReportProgress(percentage);
            }
        }

        private void Worker_DoWorkForAudioGenerationService(object sender, DoWorkEventArgs e)
        {
            for (int i = 0; i < totalValue; i++)
            {
                if (worker.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }
                int percentage = (int)Math.Round(((double)i + 1) / totalValue * 100);
                ComputerVoiceRuntimeService.EmbedSlideNotes(i);
                (sender as BackgroundWorker).ReportProgress(percentage);
            }
        }

        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            Dispatcher.Invoke( () =>
            {
                progressBar.Value = e.ProgressPercentage;
                label.Content = string.Format(ELearningLabText.ProgressStatusLabelFormat, e.ProgressPercentage);
            });
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            switch (workerType)
            {
                case BackgroundWorkerType.ELearningLabService:
                    ELearningService.SyncExitEffectAnimations();
                    break;
                case BackgroundWorkerType.AudioGenerationService:
                    break;
                default:
                    break;
            }
            Dispatcher.Invoke(() => { Close(); });
            if (AudioSettingService.IsPreviewEnabled)
            {
                ComputerVoiceRuntimeService.PreviewAnimations();
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            worker.CancelAsync();
            AzureRuntimeService.Cancel();
            Dispatcher.Invoke(() => { Close(); });
        }
    }
}
