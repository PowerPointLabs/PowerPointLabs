using System.Collections.Generic;

namespace PowerPointLabs.ImageSearch.Domain
{
    public class StyleVariants
    {
        private readonly Dictionary<string, object> _variants;

        public StyleVariants(Dictionary<string, object> var)
        {
            _variants = var;
        }

        public void Apply(StyleOptions opt)
        {
            foreach (var pair in _variants)
            {
                var type = opt.GetType();
                var prop = type.GetProperty(pair.Key);
                prop.SetValue(opt, pair.Value, null);
            }
        }
    }
}
