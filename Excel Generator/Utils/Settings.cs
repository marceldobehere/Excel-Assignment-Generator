using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Excel_Generator.Utils.SettingsList;

namespace Excel_Generator.Utils
{
    public partial class Settings
    {
        public static readonly string SETTINGS_PATH = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "/mztools/Excel_Generator/";
        private static SettingsList settingsList = null;
        public static string Language
        {
            get
            {
                return settingsList.GetSetting(SettingsList.Setting.Language).value_string;
            }
            set
            {
                settingsList.GetSetting(SettingsList.Setting.Language).value_string = value;
                settingsList.SaveSettings();
            }
        }
        public static List<string> LanguageList = new List<string>();

        public static void Init(string folderPath)
        {
            Console.Clear();
            Console.WriteLine("Marcels Debug Konsole");
            Console.WriteLine("---------------------");
            Console.WriteLine("> Lade Einstellungen...");
            settingsList = new SettingsList(folderPath);
            Console.WriteLine("> Fertig!");
            Console.WriteLine();

            LanguageList.Clear();
            Console.WriteLine();
            Console.WriteLine("> Lade Sprachen...");
            foreach (var lang in Utils.GetAllResourcesThatStartWith("LANG_"))
            {
                string langName = lang.Substring(5);
                Console.WriteLine($" - Sprache: \"{langName}\"");
                LanguageList.Add(langName);
            }
            
            Console.WriteLine("> Fertig!");

        }

        private partial class SettingsList
        {
            public enum Setting
            {
                Language
            };

            private Dictionary<Setting, SettingsObject> settingsObjects;

            public SettingsList(string folderPath)
            {
                settingsObjects = new Dictionary<Setting, SettingsObject>();
                LoadSettings(folderPath);
            }

            public SettingsObject GetSetting(Setting setting)
            {
                return settingsObjects[setting];
            }

            public void SetSetting(Setting setting, SettingsObject settingsObject)
            {
                settingsObjects[setting] = settingsObject;
            }

            public void LoadSettings(string folderPath)
            {
                settingsObjects.Clear();
                if (!Directory.Exists(folderPath))
                {
                    Console.WriteLine("> Creating Settings File");
                    Directory.CreateDirectory(folderPath);
                    Directory.CreateDirectory(folderPath + "Klassen");
                    File.Create(folderPath + "settings.cfg");
                }
                else
                {
                    using (StreamReader reader = new StreamReader(folderPath + "settings.cfg"))
                    {
                        while (!reader.EndOfStream)
                        {
                            string line = reader.ReadLine().Trim();
                            int splitIndex = line.IndexOf(':');
                            if (splitIndex == -1 || splitIndex == line.Length - 1)
                                continue;
                            AddSetting(line.Substring(0, splitIndex).Trim(), line.Substring(splitIndex + 1).Trim());
                        }
                    }
                }

                // Set All Settings to default values if they dont exist
                if (!settingsObjects.ContainsKey(Setting.Language))
                    settingsObjects.Add(Setting.Language, new SettingsObject("DE"));

                SaveSettings(folderPath);
            }

            private void AddSetting(string name, string value)
            {
                //Console.WriteLine($" - Adding Setting - {name}: \"{Utils.EscapeString(value)}\"");
                Setting setting = Setting.Language;
                switch (name)
                {
                    case "Language":
                        {
                            setting = Setting.Language;
                            break;
                        }

                    default:
                        return;
                }
                SettingsObject obj = SettingsObject.FromString(value);
                settingsObjects.Add(setting, obj);
                //Console.WriteLine($" + - Setting value: {obj}");
                //Console.WriteLine($" + - Setting value: {{{obj.value_string}}}");
            }

            public void SaveSettings(string folderPath)
            {
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                    Directory.CreateDirectory(folderPath + "Klassen");
                    File.Create(folderPath + "settings.cfg");
                }
                using (StreamWriter writer = new StreamWriter(folderPath + "settings.cfg"))
                {
                    foreach (Setting setting in settingsObjects.Keys)
                        writer.WriteLine($"{setting}: {settingsObjects[setting]}");
                }
            }

            public void LoadSettings()
            {
                LoadSettings(SETTINGS_PATH);
            }

            public void SaveSettings()
            {
                SaveSettings(SETTINGS_PATH);
            }
        }
    }
}
