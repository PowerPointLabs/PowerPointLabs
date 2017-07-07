﻿using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.AnimationLab;

namespace PowerPointLabs.ActionFramework.AnimationLab
{
    [ExportActionRibbonId(TextCollection.AddAnimationSlideTag)]
    class AddAnimationSlideActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            AutoAnimate.AddAutoAnimation();
        }
    }
}
