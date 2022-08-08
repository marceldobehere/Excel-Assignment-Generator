using Excel_Generator.Utils;
using Excel_Generator.Windows;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
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
    /// Interaktionslogik für ClassMenuPage.xaml
    /// </summary>
    public partial class ClassMenuPage : UserControl
    {
        public ClassMenuPage()
        {
            InitializeComponent();
        }

        public void UpdateText()
        {
            classMenuTitleLabel.Text = LocalizationManager.GetPhrase(Phrase.Class_TitleText);
            selectStudentLabel.Text = LocalizationManager.GetPhrase(Phrase.Class_SelectStudentText);
            selectStudentBox.ItemsSource = Settings.StudentList;
            deleteStudentButton.IsEnabled = !Settings.selectedStudent.Equals("");
        }

        private void closeClassMenuButton_Click(object sender, RoutedEventArgs e)
        {
            MainWindowHost.globalHost.classMenuPage.Visibility = Visibility.Hidden;
        }

        private void selectStudentBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count == 0)
                return;

            string selected = e.AddedItems[0] as string;
            if (selected.Equals(LocalizationManager.GetPhrase(Phrase.Class_SelectStudentTextNew)))
            {
                Utils.Settings.selectedAssignment = "";
                InputInfo info = Input.ShowInputBox(LocalizationManager.GetPhrase(Phrase.Input_TitleText), LocalizationManager.GetPhrase(Phrase.Class_SelectStudentTextNewText), LocalizationManager.GetPhrase(Phrase.Input_ConfirmButton), LocalizationManager.GetPhrase(Phrase.Input_CancelButton));
                Console.WriteLine($"Result: {info.state} - \"{info.inputText}\"");
                if (info.state == InputState.CONFIRMED)
                {
                    if (!info.inputText.Contains("\n") && info.inputText.Split(",").Length == 2)
                    {
                        List<string> studentList = Settings.StudentList;
                        studentList.Add(info.inputText);
                        studentList.Remove(LocalizationManager.GetPhrase(LocalizationManager.LanguagePhraseList.Phrase.Class_SelectStudentTextNew));

                        using (StreamWriter writer = new StreamWriter(new FileStream(Settings.SETTINGS_PATH + "Jahre/" + Settings.selectedYear + "/Klassen/" + Settings.selectedClass + "/Klassenliste.txt", FileMode.Create)))
                        {
                            foreach (string student in studentList)
                                writer.WriteLine(student);
                        }

                        selectStudentBox.ItemsSource = Settings.StudentList;
                        Utils.Settings.selectedStudent = info.inputText;
                        selectStudentBox.SelectedValue = info.inputText;
                    }
                    else
                    {
                        MessageBox.Show(LocalizationManager.GetPhrase(Phrase.Input_InvalidInputText), LocalizationManager.GetPhrase(Phrase.Input_InvalidInputTitleText));
                        selectStudentBox.SelectedValue = "";
                        Utils.Settings.selectedStudent = "";
                    }
                }
                else
                {
                    selectStudentBox.SelectedValue = "";
                    Utils.Settings.selectedStudent = "";
                }
            }
            else
            {
                Settings.selectedStudent = selected;
            }
            UpdateText();
        }

        private void deleteStudentButton_Click(object sender, RoutedEventArgs e)
        {
            Console.WriteLine($"> Delete Student {Settings.selectedStudent}?");
            MessageBoxResult res = MessageBox.Show(LocalizationManager.GetPhrase(Phrase.Class_SelectStudentTextDeleteText), LocalizationManager.GetPhrase(Phrase.Input_WarningTitleText), MessageBoxButton.YesNo);
            if (res == MessageBoxResult.Yes)
            {
                List<string> studentList = Settings.StudentList;
                studentList.Remove(Settings.selectedStudent);
                studentList.Remove(LocalizationManager.GetPhrase(LocalizationManager.LanguagePhraseList.Phrase.Class_SelectStudentTextNew));

                using (StreamWriter writer = new StreamWriter(new FileStream(Settings.SETTINGS_PATH + "Jahre/" + Settings.selectedYear + "/Klassen/" + Settings.selectedClass + "/Klassenliste.txt", FileMode.Create)))
                {
                    foreach (string student in studentList)
                        writer.WriteLine(student);
                }

                Settings.selectedStudent = "";
                selectStudentBox.SelectedValue = "";
                UpdateText();
            }
        }

        private void uploadClassListButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                // openFileDialog.FileName
                List<string> studentList = new List<string>();

                using (StreamReader reader = new StreamReader(openFileDialog.FileName))
                {
                    while (!reader.EndOfStream)
                    {
                        string line = reader.ReadLine();
                        if (!line.Contains("\n") && line.Split(",").Length == 2)
                            studentList.Add(line);
                    }
                }

                using (StreamWriter writer = new StreamWriter(new FileStream(Settings.SETTINGS_PATH + "Jahre/" + Settings.selectedYear + "/Klassen/" + Settings.selectedClass + "/Klassenliste.txt", FileMode.Create)))
                {
                    foreach (string student in studentList)
                        writer.WriteLine(student);
                }

                Settings.selectedStudent = "";
                selectStudentBox.SelectedValue = "";
                UpdateText();
            }
        }
    }
}
