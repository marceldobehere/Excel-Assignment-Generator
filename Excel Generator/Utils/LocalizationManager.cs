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

            LanguagePhraseList.DisplayAllLanguageTextAsEnum("LANG_DE");

            if (!Utils.ResourceExists("LANG_" + Settings.Language))
                Settings.Language = "DE";

            langPhraseList = new LanguagePhraseList("LANG_" + Settings.Language);
            Console.WriteLine("> Fertig!");
            Console.WriteLine();

        }

        public class LanguagePhraseList
        {
            public enum Phrase
            {
                Main_TestText,
                Main_SettingsButton,
                Main_ClassMenuButton,
                Main_AssignmentMenuButton,
                Main_TitleText,
                Main_SelectYearText,
                Main_SelectYearTextNew,
                Main_SelectYearTextNewText,
                Main_SelectYearTextDeleteText,
                Main_SelectClassText,
                Main_SelectClassTextNew,
                Main_SelectClassTextNewText,
                Main_SelectClassTextDeleteText,
                Main_SelectAssignmentText,
                Main_SelectAssignmentTextNew,
                Main_SelectAssignmentTextNewText,
                Main_SelectAssignmentTextDeleteText,
                Input_TitleText,
                Input_ConfirmButton,
                Input_CancelButton,
                Input_YesButton,
                Input_NoButton,
                Input_InvalidInputTitleText,
                Input_InvalidInputText,
                Input_WarningTitleText,
                Warning_TitleText,
                Warning_StudentStartText,
                Warning_StudentDuplicateText,
                Warning_StudentDoneAssignmentText,
                Warning_StudentReviewedAssignmentText,
                Warning_YesButton,
                Warning_NoButton,
                Settings_TitleText,
                Settings_LanguageSelectionText,
                Class_TitleText,
                Class_StudentText,
                Class_StudentNameText,
                Class_StudentNumberText,
                Class_SelectStudentText,
                Class_SelectStudentTextNew,
                Class_SelectStudentTextNewText,
                Class_SelectStudentTextDeleteText,
                Assignment_TitleText,
                Assignment_ActiveText,
                Assignment_StudentListText,
                Assignment_FlipSelectionText,
                Assignment_ClearSelectionText,
                Assignment_CreateAssignmentText,
                Assignment_DeleteAssignmentText,
                Assignment_GradeAssignmentText,
                Assignment_ViewAssignmentText,
                Assignment_UploadAssignmentText
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

            public static void DisplayAllLanguageTextAsEnum(string filePath)
            {
                using (StreamReader reader = Utils.GetResourceFileStreamReader(filePath))
                {
                    bool first = false;
                    Console.WriteLine("----------------------");
                    Console.WriteLine("public enum Phrase \n{");
                    while (!reader.EndOfStream)
                    {
                        string line = reader.ReadLine().Trim();
                        int splitIndex = line.IndexOf(':');
                        if (splitIndex == -1 || splitIndex == line.Length - 1)
                            continue;
                        if (!first)
                            first = true;
                        else
                            Console.Write(",\n");
                        Console.Write($"\t{line.Substring(0, splitIndex).Trim()}");
                    }
                    Console.WriteLine("\n}");
                    Console.WriteLine("----------------------\n");
                }
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
