using Excel_Generator.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using static Excel_Generator.Utils.LocalizationManager.LanguagePhraseList;

namespace Excel_Generator.Pages
{
    /// <summary>
    /// Interaktionslogik für Test2.xaml
    /// </summary>
    public partial class SettingsPage : UserControl
    {
        public SettingsPage()
        {
            InitializeComponent();
            langBox.ItemsSource = Utils.Settings.LanguageList;
            langBox.SelectedItem = Utils.Settings.Language;

            UpdateText();
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count == 0)
                return;
            
            if (sender == langBox)
            {
                Console.WriteLine($"Lang changed to: {e.AddedItems[0]}");
                LocalizationManager.UpdateLang(e.AddedItems[0] as string);
            }
        }

        private static List<string> consoleOptions = new List<string>();
        private static bool consoleState = false;
        public static bool ConsoleState
        {
            get
            {
                return consoleState;
            }
            set
            {
                consoleState = value;
                if (consoleState)
                {
                    Utils.Utils.ShowConsole();
                }
                else
                {
                    Utils.Utils.HideConsole();
                }
            }
        }

        public void UpdateText()
        {
            settingsTitleLabel.Text = LocalizationManager.GetPhrase(Phrase.Settings_TitleText);
            settingsLangSelectionLabel.Text = LocalizationManager.GetPhrase(Phrase.Settings_LanguageSelectionText);
            settingsConsoleSelectionLabel.Text = LocalizationManager.GetPhrase(Phrase.Settings_ConsoleSelectionText);

            consoleOptions = new List<string>();
            consoleOptions.Add(LocalizationManager.GetPhrase(Phrase.Settings_ConsoleSelectionHiddenText));
            consoleOptions.Add(LocalizationManager.GetPhrase(Phrase.Settings_ConsoleSelectionShownText));

            consoleBox.ItemsSource = consoleOptions;
            
            if (consoleState)
                consoleBox.SelectedItem = LocalizationManager.GetPhrase(Phrase.Settings_ConsoleSelectionShownText);
            else
                consoleBox.SelectedItem = LocalizationManager.GetPhrase(Phrase.Settings_ConsoleSelectionHiddenText);
        }

        private void closeSettingsMenuButton_Click(object sender, RoutedEventArgs e)
        {
            this.Visibility = Visibility.Hidden;
        }

        private void consoleBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count == 0)
                return;

            ConsoleState = (e.AddedItems[0] == LocalizationManager.GetPhrase(Phrase.Settings_ConsoleSelectionShownText));
        }
    }
}
