﻿using System;
using System.Windows.Forms;
using Microsoft.Office.Tools;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.ELearningWorkspace.Views;
using PowerPointLabs.ELearningLab.Service;
using PowerPointLabs.ELearningLab.Service.StorageService;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Views;

namespace PowerPointLabs.ActionFramework.ELearningLab.ELearningTaskPane
{
    [ExportActionRibbonId(ELearningLabText.ELearningTaskPaneTag)]
    class ELearningActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            LoadingDialogBox splashView = new LoadingDialogBox();
            splashView.Show();
            AzureAccountStorageService.LoadUserAccount();
            WatsonAccountStorageService.LoadUserAccount();
            AudioSettingStorageService.LoadAudioSettingPreference();
            splashView.Close();
            this.RegisterTaskPane(typeof(ELearningLabTaskpane), ELearningLabText.ELearningTaskPaneLabel,
                ELearningTaskPaneVisibleValueChangedEventHandler);
            CustomTaskPane eLearningTaskpane = this.GetTaskPane(typeof(ELearningLabTaskpane));
            eLearningTaskpane.Visible = !eLearningTaskpane.Visible;
        }

        private void ELearningTaskPaneVisibleValueChangedEventHandler(object sender, EventArgs e)
        {
            CustomTaskPane eLearningTaskpane = this.GetTaskPane(typeof(ELearningLabTaskpane));
            if (eLearningTaskpane == null)
            {
                return;
            }
            ELearningLabTaskpane taskpane = eLearningTaskpane.Control as ELearningLabTaskpane;
            if (eLearningTaskpane.Visible)
            {
                taskpane.ELearningLabMainPanel.ReloadELearningLabOnSlideSelectionChanged();               
            }
            else
            {
                taskpane.ELearningLabMainPanel.SyncElearningLabOnSlideSelectionChanged();
            }
        }
    }
}
