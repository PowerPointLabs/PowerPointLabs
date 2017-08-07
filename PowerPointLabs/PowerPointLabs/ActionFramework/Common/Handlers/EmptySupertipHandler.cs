using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Common.Handlers
{
    class EmptySupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return "";
        }
    }
}
