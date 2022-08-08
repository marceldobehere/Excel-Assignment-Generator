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
    /// Interaktionslogik für Input.xaml
    /// </summary>
    public partial class Input : Window
    {
        public struct InputInfo
        {
            public string inputText;
            public InputState state;
        }

        public enum InputState
        {
            CONFIRMED,
            CANCELED
        }

        public InputState state = InputState.CANCELED;

        public Input(string title, string question, string confirmText, string cancelText)
        {
            InitializeComponent();
            Title = title;
            questionText.Text = question;
            confirmButton.Content = confirmText;
            cancelButton.Content = cancelText;
            inputText.Focus();
        }

        public string InputText
        {
            get { return inputText.Text; }
        }

        private void confirmButton_Click(object sender, RoutedEventArgs e)
        {
            state = InputState.CONFIRMED;
            Close();
        }

        private void cancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        public static InputInfo ShowInputBox(string title, string question, string confirmText, string cancelText)
        {
            Input input = new Input(title, question, confirmText, cancelText);
            input.ShowDialog();
            InputInfo info = new InputInfo();
            info.inputText = input.InputText;
            info.state = input.state;
            return info;
        }

        private void inputText_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                state = InputState.CONFIRMED;
                Close();
                return;
            }
            if (e.Key == Key.Escape)
            {
                Close();
                return;
            }
        }
    }
}
