using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Common.Handlers
{
    class EmptyContentHandler : ContentHandler
    {
        protected override string GetContent(string ribbonId)
        {
            return "";
        }
    }
}
