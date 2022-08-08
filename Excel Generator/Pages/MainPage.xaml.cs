using Excel_Generator.Utils;
using Excel_Generator.Windows;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
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
using static Excel_Generator.Windows.Input;

namespace Excel_Generator.Pages
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
            titleLabel.Text = LocalizationManager.GetPhrase(Phrase.Main_TitleText);
            classMenuButton.Content = LocalizationManager.GetPhrase(Phrase.Main_ClassMenuButton);

            selectYearLabel.Text = LocalizationManager.GetPhrase(Phrase.Main_SelectYearText);
            selectYearBox.ItemsSource = Settings.YearList;
            deleteYearButton.IsEnabled = !Settings.selectedYear.Equals("");

            selectClassLabel.Text = LocalizationManager.GetPhrase(Phrase.Main_SelectClassText);
            selectClassBox.ItemsSource = Settings.ClassList;
            deleteClassButton.IsEnabled = !Settings.selectedClass.Equals("");
            classMenuButton.IsEnabled = deleteClassButton.IsEnabled;

            selectAssignmentLabel.Text = LocalizationManager.GetPhrase(Phrase.Main_SelectAssignmentText);
            selectAssignmentBox.ItemsSource = Settings.AssignmentList;
            deleteAssignmentButton.IsEnabled = !Settings.selectedAssignment.Equals("");
            openInFolderButton.IsEnabled = deleteAssignmentButton.IsEnabled;
        }

        private void selectYearBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count == 0)
                return;

            string selected = e.AddedItems[0] as string;
            if (selected.Equals(LocalizationManager.GetPhrase(Phrase.Main_SelectYearTextNew)))
            {
                Utils.Settings.selectedYear = "";
                InputInfo info = Input.ShowInputBox(LocalizationManager.GetPhrase(Phrase.Input_TitleText), LocalizationManager.GetPhrase(Phrase.Main_SelectYearTextNewText), LocalizationManager.GetPhrase(Phrase.Input_ConfirmButton), LocalizationManager.GetPhrase(Phrase.Input_CancelButton));
                Console.WriteLine($"Result: {info.state} - \"{info.inputText}\"");
                if (info.state == InputState.CONFIRMED)
                {
                    if (Utils.Utils.CheckFolderName(info.inputText))
                    {
                        Directory.CreateDirectory(Utils.Settings.SETTINGS_PATH + "Jahre/" + info.inputText + "/Klassen");
                        selectYearBox.ItemsSource = Settings.YearList;
                        Utils.Settings.selectedYear = info.inputText;
                        selectYearBox.SelectedValue = info.inputText;

                    }
                    else
                    {
                        MessageBox.Show(LocalizationManager.GetPhrase(Phrase.Input_InvalidInputText), LocalizationManager.GetPhrase(Phrase.Input_InvalidInputTitleText));
                        selectYearBox.SelectedValue = "";
                        Utils.Settings.selectedYear = "";
                    }
                }
                else
                {
                    selectYearBox.SelectedValue = "";
                    Utils.Settings.selectedYear = "";
                }
            }
            else
            {
                Utils.Settings.selectedYear = selected;
            }
            selectClassBox.SelectedValue = "";
            Utils.Settings.selectedClass = "";
            Settings.selectedStudent = "";
            selectAssignmentBox.SelectedValue = "";
            Utils.Settings.selectedAssignment = "";
            UpdateText();
            MainWindowHost.globalHost.classMenuPage.UpdateText();
        }

        private void selectClassBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count == 0)
                return;

            string selected = e.AddedItems[0] as string;
            if (selected.Equals(LocalizationManager.GetPhrase(Phrase.Main_SelectClassTextNew)))
            {
                Utils.Settings.selectedClass = "";
                InputInfo info = Input.ShowInputBox(LocalizationManager.GetPhrase(Phrase.Input_TitleText), LocalizationManager.GetPhrase(Phrase.Main_SelectClassTextNewText), LocalizationManager.GetPhrase(Phrase.Input_ConfirmButton), LocalizationManager.GetPhrase(Phrase.Input_CancelButton));
                Console.WriteLine($"Result: {info.state} - \"{info.inputText}\"");
                if (info.state == InputState.CONFIRMED)
                {
                    if (Utils.Utils.CheckFolderName(info.inputText))
                    {
                        Directory.CreateDirectory(Utils.Settings.SETTINGS_PATH + "Jahre/" + Utils.Settings.selectedYear + "/Klassen/" + info.inputText + "/Aufgaben");
                        File.Create(Utils.Settings.SETTINGS_PATH + "Jahre/" + Utils.Settings.selectedYear + "/Klassen/" + info.inputText + "/Klassenliste.txt");
                        selectClassBox.ItemsSource = Settings.ClassList;
                        Utils.Settings.selectedClass = info.inputText;
                        selectClassBox.SelectedValue = info.inputText;
                    }
                    else
                    {
                        MessageBox.Show(LocalizationManager.GetPhrase(Phrase.Input_InvalidInputText), LocalizationManager.GetPhrase(Phrase.Input_InvalidInputTitleText));
                        selectClassBox.SelectedValue = "";
                        Utils.Settings.selectedClass = "";
                    }
                }
                else
                {
                    selectClassBox.SelectedValue = "";
                    Utils.Settings.selectedClass = "";
                }
            }
            else
            {
                Utils.Settings.selectedClass = selected;
            }
            selectAssignmentBox.SelectedValue = "";
            Utils.Settings.selectedAssignment = "";
            Settings.selectedStudent = "";
            UpdateText();
            MainWindowHost.globalHost.classMenuPage.UpdateText();
        }

        private void selectAssignmentBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count == 0)
                return;

            string selected = e.AddedItems[0] as string;
            if (selected.Equals(LocalizationManager.GetPhrase(Phrase.Main_SelectAssignmentTextNew)))
            {
                Utils.Settings.selectedAssignment = "";
                InputInfo info = Input.ShowInputBox(LocalizationManager.GetPhrase(Phrase.Input_TitleText), LocalizationManager.GetPhrase(Phrase.Main_SelectAssignmentTextNewText), LocalizationManager.GetPhrase(Phrase.Input_ConfirmButton), LocalizationManager.GetPhrase(Phrase.Input_CancelButton));
                Console.WriteLine($"Result: {info.state} - \"{info.inputText}\"");
                if (info.state == InputState.CONFIRMED)
                {
                    if (Utils.Utils.CheckFolderName(info.inputText))
                    {
                        string pathToAssignment = Settings.SETTINGS_PATH + "Jahre/" + Settings.selectedYear + "/Klassen/" + Settings.selectedClass + "/Aufgaben/" + info.inputText;
                        Directory.CreateDirectory(pathToAssignment);
                        Directory.CreateDirectory(pathToAssignment + "/Aufgaben");
                        Directory.CreateDirectory(pathToAssignment + "/Abgegebene Aufgaben");
                        Directory.CreateDirectory(pathToAssignment + "/Loesungen");
                        selectAssignmentBox.ItemsSource = Settings.AssignmentList;
                        Utils.Settings.selectedAssignment = info.inputText;
                        selectAssignmentBox.SelectedValue = info.inputText;



                        // Vorlage
                        byte[] data = Utils.Utils.GetResourceFileByteArray("EXAMPLE_Vorlage");
                        using (BinaryWriter writer = new BinaryWriter(new FileStream(pathToAssignment + "/Vorlage.xlsx", FileMode.Create)))
                        {
                            writer.Write(data);
                        }
                    }
                    else
                    {
                        MessageBox.Show(LocalizationManager.GetPhrase(Phrase.Input_InvalidInputText), LocalizationManager.GetPhrase(Phrase.Input_InvalidInputTitleText));
                        selectAssignmentBox.SelectedValue = "";
                        Utils.Settings.selectedAssignment = "";
                    }
                }
                else
                {
                    selectAssignmentBox.SelectedValue = "";
                    Utils.Settings.selectedAssignment = "";
                }
            }
            else
            {
                Utils.Settings.selectedAssignment = selected;
            }
            UpdateText();
            MainWindowHost.globalHost.classMenuPage.UpdateText();
        }

        private void deleteYearButton_Click(object sender, RoutedEventArgs e)
        {
            Console.WriteLine($"> Delete Year {Settings.selectedYear}?");
            MessageBoxResult res = MessageBox.Show(LocalizationManager.GetPhrase(Phrase.Main_SelectYearTextDeleteText), LocalizationManager.GetPhrase(Phrase.Input_WarningTitleText), MessageBoxButton.YesNo);
            if (res == MessageBoxResult.Yes)
            {
                Directory.Delete(Utils.Settings.SETTINGS_PATH + "Jahre/" + Settings.selectedYear, true);
                Settings.selectedYear = "";
                selectYearBox.SelectedValue = "";
                UpdateText();
                MainWindowHost.globalHost.classMenuPage.UpdateText();
            }
        }

        private void deleteClassButton_Click(object sender, RoutedEventArgs e)
        {
            Console.WriteLine($"> Delete Class {Settings.selectedClass}?");
            MessageBoxResult res = MessageBox.Show(LocalizationManager.GetPhrase(Phrase.Main_SelectClassTextDeleteText), LocalizationManager.GetPhrase(Phrase.Input_WarningTitleText), MessageBoxButton.YesNo);
            if (res == MessageBoxResult.Yes)
            {
                Directory.Delete(Utils.Settings.SETTINGS_PATH + "Jahre/" + Settings.selectedYear + "/Klassen/" + Settings.selectedClass, true);
                Settings.selectedClass = "";
                selectClassBox.SelectedValue = "";
                UpdateText();
                MainWindowHost.globalHost.classMenuPage.UpdateText();
            }
        }

        private void deleteAssignmentButton_Click(object sender, RoutedEventArgs e)
        {
            Console.WriteLine($"> Delete Assignment {Settings.selectedAssignment}?");
            MessageBoxResult res = MessageBox.Show(LocalizationManager.GetPhrase(Phrase.Main_SelectAssignmentTextDeleteText), LocalizationManager.GetPhrase(Phrase.Input_WarningTitleText), MessageBoxButton.YesNo);
            if (res == MessageBoxResult.Yes)
            {
                Directory.Delete(Utils.Settings.SETTINGS_PATH + "Jahre/" + Settings.selectedYear + "/Klassen/" + Settings.selectedClass + "/Aufgaben/" + Settings.selectedAssignment, true);
                Settings.selectedAssignment = "";
                selectAssignmentBox.SelectedValue = "";
                UpdateText();
                MainWindowHost.globalHost.classMenuPage.UpdateText();
            }
        }

        private void openInFolderButton_Click(object sender, RoutedEventArgs e)
        {
            Utils.Utils.WindowsExplorerOpen(Utils.Settings.SETTINGS_PATH + "Jahre/" + Settings.selectedYear + "/Klassen/" + Settings.selectedClass + "/Aufgaben/" + Settings.selectedAssignment);
        }

        private void classMenuButton_Click(object sender, RoutedEventArgs e)
        {
            MainWindowHost.globalHost.classMenuPage.Visibility = Visibility.Visible;
        }
    }
}
