using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using System.Linq;
using System.Reflection;

using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Common.Factory
{
    /// <summary>
    /// Factory for THandler
    /// </summary>
    /// <typeparam name="THandler">Target handler type</typeparam>
    public abstract class BaseHandlerFactory<THandler>
    {
        [ImportMany]
        private IEnumerable<Lazy<THandler, IRibbonIdMetadata>> ImportedHandlers { get; set; }

        protected BaseHandlerFactory()
        {
            AggregateCatalog catalog = new AggregateCatalog(
                new AssemblyCatalog(Assembly.GetExecutingAssembly()));
            CompositionContainer container = new CompositionContainer(catalog);
            container.ComposeParts(this);
        }

        public THandler CreateInstance(string ribbonId, string ribbonTag)
        {
            foreach (Lazy<THandler, IRibbonIdMetadata> handler in ImportedHandlers)
            {
                if (handler.Metadata.RibbonIds.Contains(ribbonId)
                    || handler.Metadata.RibbonIds.Contains(ribbonTag))
                {
                    return handler.Value;
                }
            }
            return GetEmptyHandler();
        }

        protected abstract THandler GetEmptyHandler();
    }
}
