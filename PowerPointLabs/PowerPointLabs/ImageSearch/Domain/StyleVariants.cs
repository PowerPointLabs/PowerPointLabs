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

        public void Set(string key, object newValue)
        {
            _variants[key] = newValue;
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

        /// <summary>
        /// return true, when applying variant to this style options has no effect (still same)
        /// </summary>
        /// <param name="opt"></param>
        public bool IsNoEffect(StyleOptions opt)
        {
            foreach (var pair in _variants)
            {
                if (pair.Key.Equals("OptionName") || pair.Value is bool)
                {
                    continue;
                }

                var type = opt.GetType();
                var prop = type.GetProperty(pair.Key);
                var optValue = prop.GetValue(opt, null);
                if (!pair.Value.Equals(optValue))
                {
                    return false;
                }
            }
            return true;
        }
    }
}
