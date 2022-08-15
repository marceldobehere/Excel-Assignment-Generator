using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace Excel_Generator.Utils
{
    public class Utils
    {
        [DllImport("Kernel32")]
        public static extern void AllocConsole();

        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;

        public static void HideConsole()
        {
            ShowWindow(GetConsoleWindow(), SW_HIDE);
        }

        public static void ShowConsole()
        {
            ShowWindow(GetConsoleWindow(), SW_SHOW);
        }

        public static void WindowsExplorerOpen(string path)
        {
            CommandLine(path, $"start \"\" \"{path}\"");
        }

        private static void CommandLine(string workingDirectory, string Command)
        {
            ProcessStartInfo processInfo;
            
            processInfo = new ProcessStartInfo("cmd.exe", "/K \"" + Command + "\" && exit");
            processInfo.WorkingDirectory = workingDirectory;
            processInfo.CreateNoWindow = true;
            processInfo.UseShellExecute = true;
            processInfo.WindowStyle = ProcessWindowStyle.Hidden;

            Process process = Process.Start(processInfo);
            process.WaitForExit();
        }

        public static List<string> GetAllResourcesThatStartWith(string start)
        {
            List<string> resources = new List<string>();

            if (Excel_Generator.Properties.Resources.ResourceManager.GetResourceSet(CultureInfo.CurrentUICulture, true, true) != null)
                foreach (var x in Excel_Generator.Properties.Resources.ResourceManager.GetResourceSet(CultureInfo.CurrentUICulture, true, true))
                {
                    string? name = (x as DictionaryEntry?).Value.Key as string;

                    if (name != null)
                    {
                        //Console.WriteLine($"Resource: {name}");
                        if (name.StartsWith(start))
                            resources.Add(name);
                    }
                }

            return resources;
        }

        public static bool ResourceExists(string filename)
        {
            return Excel_Generator.Properties.Resources.ResourceManager.GetObject(filename) != null;
        }

        public static StreamReader GetResourceFileStreamReader(string filename)
        {
            //var assembly = Assembly.GetExecutingAssembly();
            //var stream = assembly.GetManifestResourceStream($"Excel_Generator.{filename}");
            byte[] byteArr = Properties.Resources.ResourceManager.GetObject(filename) as byte[];
            return new StreamReader(new MemoryStream(byteArr));
        }

        public static byte[] GetResourceFileByteArray(string filename)
        {
            return Properties.Resources.ResourceManager.GetObject(filename) as byte[];
        }

        public class StudentObject
        {
            public string name = "";
            public int id = -1;
            public StudentObject(string name, int id)
            {
                this.name = name;
                this.id = id;
            }
        }

        public static StudentObject ConvertStringToStudentStruct(string name)
        {
            string[] split = name.Split(",");
            int id = -1;
            if (int.TryParse(split[0], out id))
            {
                StudentObject student = new StudentObject(split[1], id);
                return student;
            }
            else if (int.TryParse(split[1], out id))
            {
                StudentObject student = new StudentObject(split[0], id);
                return student;
            }
            else
            {
                Console.WriteLine($"ERROR: Student \"{name}\" cannot be parsed!");
                return null;
            }
        }

        private static string allowedChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_ ";
        public static bool CheckFolderName(string name)
        {
            if (name.Length == 0)
                return false;
            
            foreach (char c in name)
                if (!allowedChars.Contains(c))
                    return false;

            return true;
        }

        public static string EscapeString(string original)
        {
            StringBuilder builder = new StringBuilder();

            for (int i = 0; i < original.Length; i++)
            {
                switch (original[i])
                {
                    case '\\':
                        {
                            builder.Append("\\\\");
                            break;
                        }
                    case '\n':
                        {
                            builder.Append("\\n");
                            break;
                        }
                    case '"':
                        {
                            builder.Append("\\\"");
                            break;
                        }
                    default:
                        {
                            builder.Append(original[i]);
                            break;
                        }
                }
            }

            return builder.ToString();
        }

        public static void OpenWithDefaultProgram(string path)
        {
            //using Process appLauncher = new Process();
            //appLauncher.StartInfo.FileName = path;
            //appLauncher.Start();
            //System.Diagnostics.Process.Start(path);
            CommandLine(path, $"start \"\" \"{path}\"");
        }

        public static string UnEscapeString(string original)
        {
            StringBuilder builder = new StringBuilder();

            for (int i = 0; i < original.Length; i++)
            {
                if (original[i] == '\\' && i + 1 < original.Length)
                {
                    switch (original[i + 1])
                    {
                        case '\\':
                            {
                                builder.Append("\\");
                                break;
                            }
                        case 'n':
                            {
                                builder.Append("\n");
                                break;
                            }
                        case '"':
                            {
                                builder.Append('"');
                                break;
                            }
                        default:
                            {
                                builder.Append(original[i + 1]);
                                break;
                            }
                    }
                    i++;
                }
                else
                    builder.Append(original[i]);
            }

            return builder.ToString();
        }
    }
}
