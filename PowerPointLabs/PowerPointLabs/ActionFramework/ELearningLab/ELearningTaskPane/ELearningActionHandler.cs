using System;
using System.Windows.Forms;
using Microsoft.Office.Tools;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.ELearningWorkspace.Views;
using PowerPointLabs.ELearningLab.Views;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ELearningLab.ELearningTaskPane
{
    [ExportActionRibbonId(ELearningLabText.ELearningTaskPaneTag)]
    class ELearningActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.RegisterTaskPane(typeof(ELearningLabTaskpane), ELearningLabText.ELearningTaskPaneLabel,
                ELearningTaskPaneVisibleValueChangedEventHandler);
            CustomTaskPane eLearningTaskpane = this.GetTaskPane(typeof(ELearningLabTaskpane));
            eLearningTaskpane.Visible = !eLearningTaskpane.Visible;
        }

        private void ELearningTaskPaneVisibleValueChangedEventHandler(object sender, EventArgs e)
        {
            CustomTaskPane eLearningTaskpane = this.GetTaskPane(typeof(ELearningLabTaskpane));
            if (eLearningTaskpane.Visible)
            {
                ELearningLabMainPanel.GetInstance().HandleELearningPaneVisibilityChanged();
            }
            else if (!ELearningLabMainPanel.GetInstance().IsInSync())
            {
                DialogResult result = MessageBox.Show(
                    "ELearningLab detected that you have unsynced items in your workspace.\n" +
                    "Do you want to sync them now?", "ELearningLab", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    ELearningLabMainPanel.GetInstance().SyncClickItems();
                }
            }
        }
    }
}
