using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Excel_Generator.Utils
{
    public class SettingsList
    {
        public class SettingsObject
        {
            public enum valueType
            {
                STRING,
                INT,
                DOUBLE,
                BOOL,
                LIST,
                DICT,
                NONE
            }

            public valueType type = valueType.NONE;
            public string value_string = "";
            public int value_int = 0;
            public double value_double = 0;
            public bool value_bool = false;
            public List<SettingsObject> value_list = null;
            public Dictionary<string, SettingsObject> value_dict = null;

            public SettingsObject()
            {
                type = valueType.NONE;
            }

            public SettingsObject(int val)
            {
                type = valueType.INT;
                this.value_int = val;
            }

            public SettingsObject(double val)
            {
                type = valueType.DOUBLE;
                this.value_double = val;
            }

            public SettingsObject(string val)
            {
                type = valueType.STRING;
                this.value_string = val;
            }

            public SettingsObject(bool val)
            {
                type = valueType.BOOL;
                this.value_bool = val;
            }

            public SettingsObject(List<SettingsObject> val)
            {
                type = valueType.LIST;
                this.value_list = val;
            }

            public SettingsObject(Dictionary<string, SettingsObject> val)
            {
                type = valueType.DICT;
                this.value_dict = val;
            }

            public override string ToString()
            {
                switch (type)
                {
                    case valueType.STRING:
                        return "\"" + Utils.EscapeString(value_string) + "\"";
                    case valueType.INT:
                        return value_int.ToString(CultureInfo.InvariantCulture);
                    case valueType.DOUBLE:
                        return value_double.ToString(CultureInfo.InvariantCulture);
                    case valueType.BOOL:
                        return value_bool ? "true" : "false";
                    case valueType.LIST:
                        {
                            StringBuilder builder = new StringBuilder();
                            throw new NotImplementedException("Listen wurden noch nicht implementiert!");
                            return builder.ToString();
                        }
                    case valueType.DICT:
                        {
                            StringBuilder builder = new StringBuilder();
                            throw new NotImplementedException("Dictionaries wurden noch nicht implementiert!");
                            return builder.ToString();
                        }
                    default:
                        return "";
                }
            }

            public static SettingsObject FromString(string data)
            {
                if (String.IsNullOrWhiteSpace(data))
                    return new SettingsObject();
                if (data.StartsWith('"') && data.EndsWith('"'))
                    return new SettingsObject(Utils.UnEscapeString(data.Substring(1, data.Length - 2)));
                if (data.StartsWith('[') && data.EndsWith(']'))
                    return FromListString(data.Substring(1, data.Length - 2));
                if (data.StartsWith('{') && data.EndsWith('}'))
                    return FromDictString(data.Substring(1, data.Length - 2));
                if (data.Equals("false"))
                    return new SettingsObject(false);
                if (data.Equals("true"))
                    return new SettingsObject(true);
                if (data.Contains("."))
                    if (double.TryParse(data, NumberStyles.Float, CultureInfo.InvariantCulture, out double val_double))
                        return new SettingsObject(val_double);
                if (int.TryParse(data, NumberStyles.Integer, CultureInfo.InvariantCulture, out int val_int))
                    return new SettingsObject(val_int);

                return new SettingsObject();
            }

            private static SettingsObject FromListString(string data)
            {
                List<SettingsObject> list = new List<SettingsObject>();
                throw new NotImplementedException("Listen wurden noch nicht implementiert!");
                return new SettingsObject(list);
            }

            private static SettingsObject FromDictString(string data)
            {
                Dictionary<string, SettingsObject> dict = new Dictionary<string, SettingsObject>();
                throw new NotImplementedException("Dictionaries wurden noch nicht implementiert!");
                return new SettingsObject(dict);
            }
        }
    }
}
