using PowerPointLabs.ActionFramework.Common.Handlers;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Common.Factory
{
    /// <summary>
    /// Factory for ContentHandler
    /// </summary>
    public class ContentHandlerFactory : BaseHandlerFactory<ContentHandler>
    {
        protected override ContentHandler GetEmptyHandler()
        {
            return new EmptyContentHandler();
        }
    }
}
