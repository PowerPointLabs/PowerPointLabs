using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ELearningLab.ELearningTaskPane
{
    [ExportEnabledRibbonId(ELearningLabText.ELearningTaskPaneTag)]
    class ELearningEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return PowerPointLabs.ELearningLab.Service.ELearningService.IsELearningWorkspaceEnabled;
        }
    }
}
