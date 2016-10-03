using PowerPointLabs.ActionFramework.CheckBoxAction;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Common.Factory
{
    /// <summary>
    /// Factory for CheckBoxActionHandler
    /// </summary>
    public class CheckBoxActionHandlerFactory : BaseHandlerFactory<CheckBoxActionHandler>
    {
        protected override CheckBoxActionHandler GetEmptyHandler()
        {
            return new EmptyCheckBoxActionHandler();
        }
    }
}
