using System.Collections.Generic;

namespace PowerPointLabs.PictureSlidesLab.Model
{
    public class StyleVariant
    {
        private readonly Dictionary<string, object> _variants;

        public StyleVariant(Dictionary<string, object> var)
        {
            _variants = var;
        }

        public Dictionary<string, object> GetVariants()
        {
            return new Dictionary<string, object>(_variants);
        }

        public void Set(string key, object newValue)
        {
            _variants[key] = newValue;
        }

        public object Get(string key)
        {
            return _variants[key];
        }

        public void Apply(StyleOption opt)
        {
            foreach (KeyValuePair<string, object> pair in _variants)
            {
                System.Type type = opt.GetType();
                System.Reflection.PropertyInfo prop = type.GetProperty(pair.Key);
                prop.SetValue(opt, pair.Value, null);
            }
        }

        /// <summary>
        /// Copy corresponding variant info from the given style
        /// </summary>
        public StyleVariant Copy(StyleOption opt, string givenOptionName = null)
        {
            Dictionary<string, object> newVariants = new Dictionary<string, object>();
            foreach (KeyValuePair<string, object> pair in _variants)
            {
                if (pair.Key.Equals("OptionName"))
                {
                    newVariants["OptionName"] = givenOptionName ?? "Reloaded";
                }
                else
                {
                    System.Type type = opt.GetType();
                    System.Reflection.PropertyInfo prop = type.GetProperty(pair.Key);
                    object optValue = prop.GetValue(opt, null);
                    newVariants[pair.Key] = optValue;
                }
            }
            return new StyleVariant(newVariants);
        }

        /// <summary>
        /// return true, when applying variant to this style options has no effect (still same)
        /// </summary>
        /// <param name="opt"></param>
        public bool IsNoEffect(StyleOption opt)
        {
            foreach (KeyValuePair<string, object> pair in _variants)
            {
                if (pair.Key.Equals("OptionName") || pair.Value is bool)
                {
                    continue;
                }

                System.Type type = opt.GetType();
                System.Reflection.PropertyInfo prop = type.GetProperty(pair.Key);
                object optValue = prop.GetValue(opt, null);
                if (!pair.Value.Equals(optValue))
                {
                    return false;
                }
            }
            return true;
        }
    }
}
