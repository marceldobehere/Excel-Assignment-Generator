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

namespace Excel_Generator.Windows
{
    /// <summary>
    /// Interaktionslogik für StudentWarning.xaml
    /// </summary>
    public partial class StudentWarning : Window
    {
        public enum InputState
        {
            YES,
            NO
        }

        public InputState state = InputState.NO;

        public StudentWarning(string title, string studentFirstText, string studentName, string studentWarhingMainText, string yes, string no)
        {
            InitializeComponent();
            Title = title;
            StudentNameText.Text = studentName;
            StudentWarningText.Text = studentWarhingMainText;
            studentStartText.Text = studentFirstText;
            yesButton.Content = yes;
            noButton.Content = no;
        }

        private void yesButton_Click(object sender, RoutedEventArgs e)
        {
            state = InputState.YES;
            Close();
        }

        private void noButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        public static InputState ShowWarningBox(string title, string studentFirstText, string studentName, string studentWarhingMainText, string yes, string no)
        {
            StudentWarning warning = new StudentWarning(title, studentFirstText, studentName, studentWarhingMainText, yes, no);
            warning.ShowDialog();
            return warning.state;
        }
    }
}
