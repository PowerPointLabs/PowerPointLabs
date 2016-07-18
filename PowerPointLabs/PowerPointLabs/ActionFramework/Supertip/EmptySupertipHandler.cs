using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Supertip
{
    class EmptySupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId, string ribbonTag)
        {
            return "";
        }
    }
}
