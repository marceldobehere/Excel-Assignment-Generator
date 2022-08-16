using Excel_Generator.Utils;
using Excel_Generator.Windows;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
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
using static Excel_Generator.Pages.AssignmentMenuPage.CheckBoxUnit;
using static Excel_Generator.Utils.LocalizationManager.LanguagePhraseList;
using static Excel_Generator.Utils.Utils;

namespace Excel_Generator.Pages
{
    /// <summary>
    /// Interaktionslogik für AssignmentMenuPage.xaml
    /// </summary>
    public partial class AssignmentMenuPage : UserControl
    {
        public ObservableCollection<CheckBoxUnit> StudentCheckBoxList = new ObservableCollection<CheckBoxUnit>();

        public class CheckBoxUnit : INotifyPropertyChanged
        {
            public event PropertyChangedEventHandler PropertyChanged;

            private void RaisePropertyChanged([CallerMemberName] string caller = "")
                    => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(caller));

            private string text;
            public string Text
            {
                get => text;
                set
                {
                    text = value;
                    RaisePropertyChanged();
                }
            }

            private bool _checked;
            public bool Checked
            {
                get => _checked;
                set
                {
                    _checked = value;
                    RaisePropertyChanged();
                }
            }

            private int id;
            public int Id
            {
                get => id;
                set
                {
                    id = value;
                    RaisePropertyChanged();
                }
            }

            public string GetBackgroundColor
            {
                get
                {
                    return assignmentColors[CurrentAssignmentState];
                }
            }

            private AssignmentState _assignmentState;

            public AssignmentState CurrentAssignmentState
            {
                get
                {
                    return _assignmentState;
                }
                set
                {
                    _assignmentState = value;
                    RaisePropertyChanged();
                }
            }

            public enum AssignmentState
            {
                NO_ASSIGNMENT_CREATED,
                ASSIGNMENT_NOT_UPLOADED,
                ASSIGNMENT_IN_REVIEW,
                ASSIGNMENT_DONE
            }

            public Dictionary<AssignmentState, string> assignmentColors = new Dictionary<AssignmentState, string>()
            {
                { AssignmentState.NO_ASSIGNMENT_CREATED, "#FFDFDFDF" },
                { AssignmentState.ASSIGNMENT_NOT_UPLOADED, "#FFF4B9B9" },
                { AssignmentState.ASSIGNMENT_IN_REVIEW, "#FFE8E6CF" },
                { AssignmentState.ASSIGNMENT_DONE, "#FFCDE8BD" }
            };

            public CheckBoxUnit(string text, int id)
            {
                Text = text;
                Id = id;
                Checked = false;
                CurrentAssignmentState = AssignmentState.NO_ASSIGNMENT_CREATED;
            }
        }

        public void CreateCheckBoxList()
        {
            StudentCheckBoxList.Clear();
            foreach (var studentName in Settings.StudentList)
            {
                if (studentName.Equals(LocalizationManager.GetPhrase(Phrase.Class_SelectStudentTextNew)))
                    continue;
                StudentObject studentObj = Utils.Utils.ConvertStringToStudentStruct(studentName);
                if (studentObj != null)
                    StudentCheckBoxList.Add(new CheckBoxUnit(studentObj.name, studentObj.id));
            }

            UpdateCheckBoxColors();

            studentList.ItemsSource = StudentCheckBoxList;
        }


        public AssignmentMenuPage()
        {
            InitializeComponent();
            UpdateText();
        }

        public void UpdateText()
        {
            assignmentMenuTitleLabel.Text = LocalizationManager.GetPhrase(Phrase.Assignment_TitleText);
            activeAssignmentLabel.Text = LocalizationManager.GetPhrase(Phrase.Assignment_ActiveText) + $"\"{Settings.selectedAssignment}\"";
            CreateCheckBoxList();

            studentListLabel.Text = LocalizationManager.GetPhrase(Phrase.Assignment_StudentListText);
            flipSelectionButton.Content = LocalizationManager.GetPhrase(Phrase.Assignment_FlipSelectionText);
            clearSelectionButton.Content = LocalizationManager.GetPhrase(Phrase.Assignment_ClearSelectionText);

            addAssignmentButton.Content = LocalizationManager.GetPhrase(Phrase.Assignment_CreateAssignmentText);
            removeAssignmentButton.Content = LocalizationManager.GetPhrase(Phrase.Assignment_DeleteAssignmentText);
            checkAssignmentButton.Content = LocalizationManager.GetPhrase(Phrase.Assignment_GradeAssignmentText);
            viewAssignmentButton.Content = LocalizationManager.GetPhrase(Phrase.Assignment_ViewAssignmentText);
            uploadAssignmentButton.Content = LocalizationManager.GetPhrase(Phrase.Assignment_UploadAssignmentText);
        }

        private void flipSelectionButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in studentList.ItemsSource)
            {
                CheckBoxUnit student = item as CheckBoxUnit;
                //Console.WriteLine($" - Student: {student.Text}: {student.Checked}");
                student.Checked = !student.Checked;
            }
        }

        private void clearSelectionButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in studentList.ItemsSource)
            {
                CheckBoxUnit student = item as CheckBoxUnit;
                student.Checked = false;
            }
        }

        private List<CheckBoxUnit> GetSelectedBoxes()
        {
            List<CheckBoxUnit> units = new List<CheckBoxUnit>();
            foreach (var item in studentList.ItemsSource)
            {
                CheckBoxUnit student = item as CheckBoxUnit;
                if (student.Checked)
                    units.Add(student);
            }
            return units;
        }

        private List<StudentObject> GetSelectedStudents()
        {
            List<StudentObject> studentList = new List<StudentObject>();
            foreach (var studentName in Settings.StudentList)
            {
                if (studentName.Equals(LocalizationManager.GetPhrase(Phrase.Class_SelectStudentTextNew)))
                    continue;
                StudentObject studentObj = ConvertStringToStudentStruct(studentName);
                if (studentObj != null)
                    studentList.Add(studentObj);
            }

            List<CheckBoxUnit> units = GetSelectedBoxes();
            List<StudentObject> students = new List<StudentObject>();

            foreach (CheckBoxUnit unit in units)
            {
                StudentObject foundStudent = null;

                foreach (StudentObject tempStudent in studentList)
                    if (tempStudent.id == unit.Id)
                    {
                        foundStudent = tempStudent;
                        break;
                    }

                if (foundStudent != null)
                    students.Add(foundStudent);
            }

            return students;
        }

        private void studentList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Console.WriteLine("Selection Changed!");
            CheckBox a;
        }

        private void closeAssignmentMenuButton_Click(object sender, RoutedEventArgs e)
        {
            MainWindowHost.globalHost.assignmentMenuPage.Visibility = Visibility.Hidden;
        }

        // TODO:
        // Add Assignment Statistics
        // Add Student Statistics
        // Add Class Statistics

        // Add Manual Grading
        // Improve Vorlage


        private void UpdateCheckBoxColors()
        {
            if (Settings.selectedYear == "")
                return;
            if (Settings.selectedClass == "")
                return;
            if (Settings.selectedAssignment == "")
                return;

            string assignmentFolderPath = Settings.SETTINGS_PATH + "Jahre/" +
                                          Settings.selectedYear + "/Klassen/" +
                                          Settings.selectedClass + "/Aufgaben/" +
                                          Settings.selectedAssignment + "/";


            Console.WriteLine("\nLösungen");
            List<int> solutionIds = Excel_API.MainExcelAPI.GetStudentIDsFromFolder(assignmentFolderPath + "Loesungen/EXCEL");
            Console.WriteLine("\nAbgegeben");
            List<int> uploadedIds = Excel_API.MainExcelAPI.GetStudentIDsFromFolder(assignmentFolderPath + "Abgegebene Aufgaben");
            Console.WriteLine("\nFertig");
            List<int> reviewedIds = Excel_API.MainExcelAPI.GetStudentIDsFromFolder(assignmentFolderPath + "Fertige Aufgaben/EXCEL");

            foreach (var student in StudentCheckBoxList)
            {
                if (reviewedIds.Contains(student.Id))
                    student.CurrentAssignmentState = AssignmentState.ASSIGNMENT_DONE;
                else if (uploadedIds.Contains(student.Id))
                    student.CurrentAssignmentState = AssignmentState.ASSIGNMENT_IN_REVIEW;
                else if (solutionIds.Contains(student.Id))
                    student.CurrentAssignmentState = AssignmentState.ASSIGNMENT_NOT_UPLOADED;
                else
                    student.CurrentAssignmentState = AssignmentState.NO_ASSIGNMENT_CREATED;
            }
        }

        private void addAssignmentButton_Click(object sender, RoutedEventArgs e)
        {
            string assignmentFolderPath = Settings.SETTINGS_PATH + "Jahre/" +
                                          Settings.selectedYear + "/Klassen/" +
                                          Settings.selectedClass + "/Aufgaben/" +
                                          Settings.selectedAssignment + "/";

            string solFile = assignmentFolderPath + "Vorlage.xlsx";
            string solFolder = assignmentFolderPath + "Loesungen";
            string queFolder = assignmentFolderPath + "Aufgaben";


            List<StudentObject> selectedStudents = GetSelectedStudents();


            {
                Console.WriteLine("\nLösungen");
                List<int> solutionIds = Excel_API.MainExcelAPI.GetStudentIDsFromFolder(assignmentFolderPath + "Loesungen/EXCEL");

                for (int i = 0; i < selectedStudents.Count; i++)
                {
                    var student = selectedStudents[i];

                    if (solutionIds.Contains(student.id))
                    {
                        // Student already has assignment

                        var res = StudentWarning.ShowWarningBox(
                            LocalizationManager.GetPhrase(Phrase.Warning_TitleText),
                            LocalizationManager.GetPhrase(Phrase.Warning_StudentStartText),
                            student.name,
                            LocalizationManager.GetPhrase(Phrase.Warning_StudentDuplicateText),
                            LocalizationManager.GetPhrase(Phrase.Warning_YesButton),
                            LocalizationManager.GetPhrase(Phrase.Warning_NoButton)
                            );

                        if (res == StudentWarning.InputState.NO)
                        {
                            selectedStudents.RemoveAt(i);
                            i--;
                        }
                    }
                }
            }



            Excel_API.MainExcelAPI.ErrorRes error = Excel_API.MainExcelAPI.GenerateAssignmentsForStudents(selectedStudents.ToArray(), solFile, solFolder, queFolder);
            if (error != null)
                MessageBox.Show($"Error: {error.exception}", "Error");

            CreateCheckBoxList();
        }

        private void removeAssignmentButton_Click(object sender, RoutedEventArgs e)
        {
            string assignmentFolderPath = Settings.SETTINGS_PATH + "Jahre/" +
                              Settings.selectedYear + "/Klassen/" +
                              Settings.selectedClass + "/Aufgaben/" +
                              Settings.selectedAssignment + "/";

            List<StudentObject> selectedStudents = GetSelectedStudents();

            {
                Console.WriteLine("\nAbgegeben");
                Dictionary<int, string> uploadedIds = Excel_API.MainExcelAPI.GetStudentIDsAndFilenamesFromFolder(assignmentFolderPath + "Abgegebene Aufgaben");



                for (int i = 0; i < selectedStudents.Count; i++)
                {
                    var student = selectedStudents[i];

                    if (uploadedIds.ContainsKey(student.id))
                    {
                        // Student already has assignment

                        var res = StudentWarning.ShowWarningBox(
                            LocalizationManager.GetPhrase(Phrase.Warning_TitleText),
                            LocalizationManager.GetPhrase(Phrase.Warning_StudentStartText),
                            student.name,
                            LocalizationManager.GetPhrase(Phrase.Warning_StudentDoneAssignmentText),
                            LocalizationManager.GetPhrase(Phrase.Warning_YesButton),
                            LocalizationManager.GetPhrase(Phrase.Warning_NoButton)
                            );

                        if (res == StudentWarning.InputState.NO)
                        {
                            selectedStudents.RemoveAt(i);
                            i--;
                        }
                    }
                }


                if (selectedStudents.Count == 0)
                    return;


                Console.WriteLine("\nLösungen");
                Dictionary<int, string> solutionIds = Excel_API.MainExcelAPI.GetStudentIDsAndFilenamesFromFolder(assignmentFolderPath + "Loesungen/EXCEL");
                Console.WriteLine("\nAbgegeben");
                //Dictionary<int, string>uploadedIds = Excel_API.MainExcelAPI.GetStudentIDsAndFilenamesFromFolder(assignmentFolderPath + "Abgegebene Aufgaben");
                Console.WriteLine("\nFertig");
                Dictionary<int, string> reviewedIds = Excel_API.MainExcelAPI.GetStudentIDsAndFilenamesFromFolder(assignmentFolderPath + "Fertige Aufgaben/EXCEL");



                foreach (var student in selectedStudents)
                {
                    Console.WriteLine($" - Deleting data from: {student.name}");
                    if (uploadedIds.ContainsKey(student.id))
                        File.Delete(uploadedIds[student.id]);
                    
                    if (reviewedIds.ContainsKey(student.id))
                    {
                        string name = System.IO.Path.GetFileNameWithoutExtension(solutionIds[student.id]);
                        string tempPath = Directory.GetParent(System.IO.Path.GetDirectoryName(solutionIds[student.id])).FullName;

                        Console.WriteLine($" - Name: {name}");
                        Console.WriteLine($" - Path: {tempPath}");
                        Console.WriteLine();

                        //File.Delete(tempPath + "/EXCEL/" + name + ".xlsx");
                        File.Delete(reviewedIds[student.id]);
                        File.Delete(tempPath + "/TXT/" + name + ".txt");
                    }

                    if (solutionIds.ContainsKey(student.id))
                    {
                        string name = System.IO.Path.GetFileNameWithoutExtension(solutionIds[student.id]);
                        string tempPath = Directory.GetParent(System.IO.Path.GetDirectoryName(solutionIds[student.id])).FullName;

                        Console.WriteLine($" - Name: {name}");
                        Console.WriteLine($" - Path: {tempPath}");
                        Console.WriteLine();

                        File.Delete(tempPath + "/EXCEL/" + name + ".xlsx");
                        File.Delete(tempPath + "/TXT/" + name + ".txt");

                    }
                }

            }

            CreateCheckBoxList();
        }

        private void checkAssignmentButton_Click(object sender, RoutedEventArgs e)
        {
            string assignmentFolderPath = Settings.SETTINGS_PATH + "Jahre/" +
                                          Settings.selectedYear + "/Klassen/" +
                                          Settings.selectedClass + "/Aufgaben/" +
                                          Settings.selectedAssignment + "/";

            List<StudentObject> selectedStudents = GetSelectedStudents();

            {
                Console.WriteLine("\nFertig");
                Dictionary<int, string> reviewedIds = Excel_API.MainExcelAPI.GetStudentIDsAndFilenamesFromFolder(assignmentFolderPath + "Fertige Aufgaben/EXCEL");



                for (int i = 0; i < selectedStudents.Count; i++)
                {
                    var student = selectedStudents[i];

                    if (reviewedIds.ContainsKey(student.id))
                    {
                        // Student already has assignment

                        var res = StudentWarning.ShowWarningBox(
                            LocalizationManager.GetPhrase(Phrase.Warning_TitleText),
                            LocalizationManager.GetPhrase(Phrase.Warning_StudentStartText),
                            student.name,
                            LocalizationManager.GetPhrase(Phrase.Warning_StudentReviewedAssignmentText),
                            LocalizationManager.GetPhrase(Phrase.Warning_YesButton),
                            LocalizationManager.GetPhrase(Phrase.Warning_NoButton)
                            );

                        if (res == StudentWarning.InputState.NO)
                        {
                            selectedStudents.RemoveAt(i);
                            i--;
                        }
                    }
                }


                if (selectedStudents.Count == 0)
                    return;


                Console.WriteLine("\nAbgegeben");
                Dictionary<int, string> uploadedIds = Excel_API.MainExcelAPI.GetStudentIDsAndFilenamesFromFolder(assignmentFolderPath + "Abgegebene Aufgaben");

               
                List<(StudentObject student, string path)> toGrade = new List<(StudentObject student, string path)>();

                foreach (var student in selectedStudents)
                    if (uploadedIds.ContainsKey(student.id))
                    {
                        Console.WriteLine($" - Grading: {student.name}");
                        toGrade.Add((student, uploadedIds[student.id]));
                    }

                if (selectedStudents.Count == 0)
                    return;
                Excel_API.MainExcelAPI.GradeWorksheets(toGrade.ToArray(), assignmentFolderPath + "Loesungen", assignmentFolderPath + "Fertige Aufgaben");
            }

            CreateCheckBoxList();
        }

        private void viewAssignmentButton_Click(object sender, RoutedEventArgs e)
        {
            string assignmentFolderPath = Settings.SETTINGS_PATH + "Jahre/" +
                                                      Settings.selectedYear + "/Klassen/" +
                                                      Settings.selectedClass + "/Aufgaben/" +
                                                      Settings.selectedAssignment + "/";

            List<StudentObject> selectedStudents = GetSelectedStudents();

            {
                Console.WriteLine("\nFertig");
                Dictionary<int, string> reviewedIds = Excel_API.MainExcelAPI.GetStudentIDsAndFilenamesFromFolder(assignmentFolderPath + "Fertige Aufgaben/EXCEL");

                List<string> toGrade = new List<string>();

                foreach (var student in selectedStudents)
                    if (reviewedIds.ContainsKey(student.id))
                    {
                        Console.WriteLine($" - Showing: {student.name}");
                        Console.WriteLine($" - Path: {reviewedIds[student.id]}");
                        OpenWithDefaultProgram(reviewedIds[student.id]);
                    }

            }

        }

        private void uploadAssignmentButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog test = new OpenFileDialog();
            test.Multiselect = true;
            test.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            var res = test.ShowDialog();
            if (!res.HasValue)
                return;
            if (!res.Value)
                return;

            Console.WriteLine("Files:");
            foreach (string filename in test.FileNames)
            {
                Console.WriteLine($" - {filename}");
                string newFilename = Settings.SETTINGS_PATH + "Jahre/" +
                                     Settings.selectedYear + "/Klassen/" +
                                     Settings.selectedClass + "/Aufgaben/" +
                                     Settings.selectedAssignment + "/Abgegebene Aufgaben/EXCEL/" +
                                     System.IO.Path.GetFileName(filename);
                File.Copy(filename, newFilename, true);
            }

            CreateCheckBoxList();
        }
    }
}
