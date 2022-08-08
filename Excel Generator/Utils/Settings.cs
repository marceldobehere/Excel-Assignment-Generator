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

        public static string selectedYear = "";
        public static string selectedClass = "";
        public static string selectedAssignment = "";
        public static string selectedStudent = "";

        public static List<string> YearList
        {
            get
            {
                Console.WriteLine("> Lade Jahres-Liste...");
                List<string> yearList = new List<string>();
                foreach (string yearFolder in Directory.GetDirectories(SETTINGS_PATH + "Jahre"))
                    yearList.Add(Path.GetFileName(yearFolder));
                yearList.Add(LocalizationManager.GetPhrase(LocalizationManager.LanguagePhraseList.Phrase.Main_SelectYearTextNew));

                Console.WriteLine("> Fertig.");
                return yearList;
            }
        }


        public static List<string> ClassList
        {
            get
            {
                Console.WriteLine("> Lade Klassen-Liste...");
                List<string> classList = new List<string>();
                if (selectedYear.Equals(""))
                    goto end;
                if (!Directory.Exists(SETTINGS_PATH + "Jahre/" + selectedYear + "/Klassen"))
                    goto end;

                foreach (string classFolder in Directory.GetDirectories(SETTINGS_PATH + "Jahre/" + selectedYear + "/Klassen"))
                    classList.Add(Path.GetFileName(classFolder));
                classList.Add(LocalizationManager.GetPhrase(LocalizationManager.LanguagePhraseList.Phrase.Main_SelectClassTextNew));

            end:
                Console.WriteLine("> Fertig.");
                return classList;
            }
        }

        public static List<string> StudentList
        {
            get
            {
                Console.WriteLine("> Lade Schüler-Liste...");
                List<string> studentList = new List<string>();
                if (selectedYear.Equals(""))
                    goto end;
                if (selectedClass.Equals(""))
                    goto end;

                string filePath = SETTINGS_PATH + "Jahre/" + selectedYear + "/Klassen/" + selectedClass + "/Klassenliste.txt";

                if (!File.Exists(filePath))
                    goto end;

                using (StreamReader reader = new StreamReader(filePath))
                {
                    while (!reader.EndOfStream)
                    {
                        studentList.Add(reader.ReadLine());
                    }
                }
                studentList.Add(LocalizationManager.GetPhrase(LocalizationManager.LanguagePhraseList.Phrase.Class_SelectStudentTextNew));

            end:
                Console.WriteLine("> Fertig.");
                return studentList;
            }
        }

        public static List<string> AssignmentList
        {
            get
            {
                Console.WriteLine("> Lade Aufgaben-Liste...");
                List<string> assignmentList = new List<string>();
                if (selectedYear.Equals(""))
                    goto end;
                if (selectedClass.Equals(""))
                    goto end;
                if (!Directory.Exists(SETTINGS_PATH + "Jahre/" + selectedYear + "/Klassen/" + selectedClass + "/Aufgaben"))
                    goto end;

                foreach (string classFolder in Directory.GetDirectories(SETTINGS_PATH + "Jahre/" + selectedYear + "/Klassen/" + selectedClass + "/Aufgaben"))
                    assignmentList.Add(Path.GetFileName(classFolder));
                assignmentList.Add(LocalizationManager.GetPhrase(LocalizationManager.LanguagePhraseList.Phrase.Main_SelectAssignmentTextNew));

            end:
                Console.WriteLine("> Fertig.");
                return assignmentList;
            }
        }

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
            selectedClass = "";
            selectedYear = "";

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
