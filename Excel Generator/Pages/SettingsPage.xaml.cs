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

namespace Excel_Generator
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
    
    
        public void UpdateText()
        {
            settingsTitleLabel.Text = LocalizationManager.GetPhrase(Phrase.Settings_TitleText);
            settingsLangSelectionLabel.Text = LocalizationManager.GetPhrase(Phrase.Settings_LanguageSelectionText);
        }
    }
}
