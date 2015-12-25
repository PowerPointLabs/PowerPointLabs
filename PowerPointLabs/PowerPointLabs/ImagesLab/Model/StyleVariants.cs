using System.Collections.Generic;

namespace PowerPointLabs.ImagesLab.Model
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
        /// Copy corresponding variant info from the given style
        /// </summary>
        public StyleVariants Copy(StyleOptions opt, string givenOptionName = null)
        {
            var newVariants = new Dictionary<string, object>();
            foreach (var pair in _variants)
            {
                if (pair.Key.Equals("OptionName"))
                {
                    newVariants["OptionName"] = givenOptionName ?? "Reloaded";
                }
                else
                {
                    var type = opt.GetType();
                    var prop = type.GetProperty(pair.Key);
                    var optValue = prop.GetValue(opt, null);
                    newVariants[pair.Key] = optValue;
                }
            }
            return new StyleVariants(newVariants);
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
