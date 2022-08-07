using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Excel_Generator.Utils.SettingsList;

namespace Excel_Generator.Utils
{
    public class LocalizationManager
    {
        public static readonly string SETTINGS_PATH = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "/mztools/Excel_Generator/";
        private static LanguagePhraseList langPhraseList = null;

        public static string GetPhrase(LanguagePhraseList.Phrase phrase)
        {
            return langPhraseList.GetPhrase(phrase).value_string;
        }

        public static void UpdateLang(string langname)
        {
            if (!Utils.ResourceExists("LANG_" + langname))
                return;

            Console.WriteLine("> Update Sprachdaten...");

            Settings.Language = langname;

            langPhraseList = new LanguagePhraseList("LANG_" + Settings.Language);

            MainWindowHost.GlobalUpdateText();
            Console.WriteLine("> Fertig!");
            Console.WriteLine();
        }

        public static void Init()
        {
            Console.WriteLine("> Lade Sprachdaten...");

            if (!Utils.ResourceExists("LANG_" + Settings.Language))
                Settings.Language = "DE";

            langPhraseList = new LanguagePhraseList("LANG_" + Settings.Language);
            Console.WriteLine("> Fertig!");
            Console.WriteLine();

        }

        public class LanguagePhraseList
        {
            /*
            Main_TestText: "Das ist ein Test!"
            Main_SettingsButton: "Einstellungen"
            Settings_TitleText: "Einstellungen"
            Settings_LanguageSelectionText: "Sprache"
            */
            public enum Phrase
            {
                Main_TestText,
                Main_SettingsButton,
                Main_TitleText,
                Settings_TitleText,
                Settings_LanguageSelectionText
            };

            private Dictionary<Phrase, SettingsObject> phrases;

            public LanguagePhraseList(string filePath)
            {
                phrases = new Dictionary<Phrase, SettingsObject>();
                LoadLanguageData(filePath);
            }

            public SettingsObject GetPhrase(Phrase phrase)
            {
                return phrases[phrase];
            }

            public void SetPhrase(Phrase setting, SettingsObject settingsObject)
            {
                phrases[setting] = settingsObject;
            }

            public void LoadLanguageData(string filePath)
            {
                phrases.Clear();
                using (StreamReader reader = Utils.GetResourceFileStreamReader(filePath))
                {
                    while (!reader.EndOfStream)
                    {
                        string line = reader.ReadLine().Trim();
                        int splitIndex = line.IndexOf(':');
                        if (splitIndex == -1 || splitIndex == line.Length - 1)
                            continue;
                        AddPhrase(line.Substring(0, splitIndex).Trim(), line.Substring(splitIndex + 1).Trim());
                    }
                }
            }

            private void AddPhrase(string name, string value)
            {
                //Console.WriteLine($"Phrase: {name}: {Utils.EscapeString(value)}");

                if (!Enum.TryParse(name, out Phrase phrase))
                    return;

                SettingsObject obj = SettingsObject.FromString(value);
                phrases.Add(phrase, obj);
            }
        }
    }
}
