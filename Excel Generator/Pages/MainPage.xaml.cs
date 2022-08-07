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
    /// Interaktionslogik für Test1.xaml
    /// </summary>
    public partial class MainPage : UserControl
    {
        public MainPage()
        {
            InitializeComponent();
        }

        public void UpdateText()
        {
            testTextLabel.Text = LocalizationManager.GetPhrase(Phrase.Main_TestText);
            titleLabel.Text = LocalizationManager.GetPhrase(Phrase.Main_TitleText);
        }
    }
}
