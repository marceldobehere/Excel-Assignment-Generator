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
using System.Windows.Shapes;
using static Excel_Generator.Utils.LocalizationManager.LanguagePhraseList;

namespace Excel_Generator
{
    /// <summary>
    /// Interaktionslogik für MainWindowHost.xaml
    /// </summary>
    public partial class MainWindowHost : Window
    {
        public static MainWindowHost globalHost;
        private bool doneInit;

        public MainWindowHost()
        {
            doneInit = false;
            globalHost = this;
            Utils.Utils.AllocConsole();
            //Utils.Utils.HideConsole();

            Pages.SettingsPage.ConsoleState = false;

            Settings.Init(Utils.Settings.SETTINGS_PATH);
            LocalizationManager.Init();
            InitializeComponent();
            mainPage.Visibility = Visibility.Visible;
            settingsPage.Visibility = Visibility.Hidden;
            classMenuPage.Visibility = Visibility.Hidden;
            assignmentMenuPage.Visibility = Visibility.Hidden;

            doneInit = true;
            GlobalUpdateText();
        }

        public static void GlobalUpdateText()
        {
            if (!globalHost.doneInit)
                return;

            globalHost.UpdateText();
            globalHost.settingsPage.UpdateText();
            globalHost.classMenuPage.UpdateText();
            globalHost.assignmentMenuPage.UpdateText();
            globalHost.mainPage.UpdateText();
        }

        public void UpdateText()
        {
            settingsButton.Content = LocalizationManager.GetPhrase(Phrase.Main_SettingsButton);
        }

        private void settingsButton_Click(object sender, RoutedEventArgs e)
        {
            switch (settingsPage.Visibility)
            {
                case Visibility.Visible:
                    {
                        settingsPage.Visibility = Visibility.Hidden;
                        break;
                    }
                case Visibility.Hidden:
                    {
                        settingsPage.Visibility = Visibility.Visible;
                        break;
                    }
                default:
                    {
                        settingsPage.Visibility = Visibility.Hidden;
                        break;
                    }
            }
        }
    }
}
