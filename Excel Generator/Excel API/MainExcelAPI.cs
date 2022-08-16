using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using Excel_Generator.Excel_API.Utils;
using static Excel_Generator.Excel_API.Utils.Utils;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.Globalization;
using static Excel_Generator.Excel_API.Utils.Utils.SolutionClass;
using static Excel_Generator.Utils.Utils;

namespace Excel_Generator.Excel_API
{
    public class MainExcelAPI
    {


        static void Clear()
        {
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.DarkGreen;
            Console.WriteLine("Marcel Zietek 2BHIF Excel Aufgaben Werkzeug");
            Console.WriteLine("-------------------------------------------");
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.White;
        }

        public class ErrorRes
        {
            public string msg;
            public Exception exception;

            public ErrorRes()
            {
                msg = "";
                exception = null;
            }
            public ErrorRes(string msg)
            {

                this.msg = msg;
                exception = null;
            }
            public ErrorRes(string msg, Exception error)
            {
                this.msg = msg;
                exception = error;
            }
        }

        public static ErrorRes GenerateAssignmentsForStudents(Excel_Generator.Utils.Utils.StudentObject[] students, string solutionFilePath, string solutionFolderPath, string questionFolderPath, string password = "12345")
        {
            try
            {
                GenerateQuestionsAndSolutions(students, solutionFilePath, solutionFolderPath, questionFolderPath, password);
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Beim Generieren der Fragen ist ein Fehler aufgetreten: {e.Message}");
                Console.ForegroundColor = ConsoleColor.White;
                return new ErrorRes("Generating Failed", e);
            }

            return null;
        }

        public static Dictionary<int, string> GetStudentIDsAndFilenamesFromFolder(string foldername)
        {
            Dictionary<int, string> studentData = new Dictionary<int, string>();

            Console.WriteLine("IDs:");


            if (!Directory.Exists(foldername))
            {
                Console.WriteLine(" - Folder doesn't exist!");
                return studentData;
            }

            foreach (string file in Directory.GetFiles(foldername))
            {
                if (!file.EndsWith(".xlsx"))
                    continue;

                try
                {
                    IWorkbook book = WorkbookFactory.Create(file);

                    ISheet main = book.GetSheetAt(0);

                    int id = (int)(GetCellFromXY(main, 0, 999).NumericCellValue - 100);

                    studentData.Add(id, file);
                    Console.WriteLine($" - {id}");

                    book.Close();
                }
                catch (Exception e)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Bei der Datei \"{file}\" ist ein Fehler aufgetreten! ({e})");
                    Console.ForegroundColor = ConsoleColor.White;
                }
            }



            return studentData;
        }

        public static List<int> GetStudentIDsFromFolder(string foldername)
        {
            List<int> studentIds = new List<int>();

            Console.WriteLine("IDs:");

            if (!Directory.Exists(foldername))
            {
                Console.WriteLine(" - Folder doesn't exist!");
                return studentIds;
            }

            foreach (string file in Directory.GetFiles(foldername))
            {
                if (!file.EndsWith(".xlsx"))
                    continue;

                try
                {
                    IWorkbook book = WorkbookFactory.Create(file);

                    ISheet main = book.GetSheetAt(0);

                    int id = (int)(GetCellFromXY(main, 0, 999).NumericCellValue - 100);

                    studentIds.Add(id);
                    Console.WriteLine($" - {id}");

                    book.Close();
                }
                catch (Exception e)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Bei der Datei \"{file}\" ist ein Fehler aufgetreten! ({e})");
                    Console.ForegroundColor = ConsoleColor.White;
                }
            }



            return studentIds;
        }


        public static void FakeMain(string[] args)
        {
            Clear();
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("Was wollen Sie tun?");
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine("1: Angaben erstellen");
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("2: Abgegebene Angaben automatisch beurteilen");
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine();

            string input = "";
            while (!input.Equals("1") && !input.Equals("2"))
            {
                Console.Write("> ");
                input = Console.ReadLine();
            }

            //if (input.Equals("1"))
            //{
            //    string solutionFilePath = "Vorlage.xlsx", solutionFolderPath = "Loesungen", questionFolderPath = "Aufgaben";

            //    Console.WriteLine("Bitte geben Sie den Pfad zur Angabedatei an.");
            //    while (!File.Exists(solutionFilePath))
            //    {
            //        Console.Write("> ");
            //        solutionFilePath = Console.ReadLine();
            //    }
            //    Console.WriteLine($"Angabedatei: \"{solutionFilePath}\"");
            //    Console.WriteLine();


            //    Console.WriteLine("Bitte geben Sie den Pfad zum Loesungsordner an.");
            //    while (!Directory.Exists(solutionFolderPath))
            //    {
            //        Console.Write("> ");
            //        solutionFolderPath = Console.ReadLine();
            //    }
            //    Console.WriteLine($"Loesungsordner: \"{solutionFolderPath}\"");
            //    Console.WriteLine();


            //    Console.WriteLine("Bitte geben Sie den Pfad zum Angabenordner an.");
            //    while (!Directory.Exists(questionFolderPath))
            //    {
            //        Console.Write("> ");
            //        questionFolderPath = Console.ReadLine();
            //    }
            //    Console.WriteLine($"Angabenordner: \"{questionFolderPath}\"");
            //    Console.WriteLine();

            //    Console.WriteLine("Geben Sie die Anzahl an Angaben an.");
            //    string amount = "";
            //    while (!int.TryParse(amount, out _))
            //    {
            //        Console.Write("> ");
            //        amount = Console.ReadLine();
            //    }
            //    Console.WriteLine();

            //    Console.WriteLine("Geben Sie das Passwort ein: (Leer lassen um es ohne Passwort zu schützen)");
            //    Console.Write("> ");
            //    string password = Console.ReadLine();
            //    Console.WriteLine();

            //    try
            //    {
            //        GenerateQuestionsAndSolutions(int.Parse(amount), solutionFilePath, solutionFolderPath, questionFolderPath, password);
            //    }
            //    catch (Exception e)
            //    {
            //        Console.ForegroundColor = ConsoleColor.Red;
            //        Console.WriteLine($"Beim Generieren der Fragen ist ein Fehler aufgetreten: {e.Message}");
            //        Console.ForegroundColor = ConsoleColor.White;
            //    }
            //}
            //else if (input.Equals("2"))
            //{
            //    string solutionFolderPath = "Loesungen", questionFolderPath = "Abgegebene Aufgaben", gradedFolderPath = "Verbesserte Aufgaben";

            //    Console.WriteLine("Bitte geben Sie den Pfad zu den Angaben an.");
            //    while (!Directory.Exists(questionFolderPath))
            //    {
            //        Console.Write("> ");
            //        questionFolderPath = Console.ReadLine();
            //    }
            //    Console.WriteLine($"Angabeordner: \"{questionFolderPath}\"");
            //    Console.WriteLine();

            //    Console.WriteLine("Bitte geben Sie den Pfad zum Loesungsordner an.");
            //    while (!Directory.Exists(solutionFolderPath))
            //    {
            //        Console.Write("> ");
            //        solutionFolderPath = Console.ReadLine();
            //    }
            //    Console.WriteLine($"Loesungsordner: \"{solutionFolderPath}\"");
            //    Console.WriteLine();


            //    Console.WriteLine("Bitte geben Sie den Pfad vom Ordner an, indem die Aufgaben verbessert werden sollen.");
            //    while (!Directory.Exists(gradedFolderPath))
            //    {
            //        Console.Write("> ");
            //        gradedFolderPath = Console.ReadLine();
            //    }
            //    Console.WriteLine($"Verbesserte Angaben-Order: \"{gradedFolderPath}\"");
            //    Console.WriteLine();
            //    try
            //    {
            //        GradeWorksheets(questionFolderPath, solutionFolderPath, gradedFolderPath);
            //    }
            //    catch (Exception e)
            //    {
            //        Console.ForegroundColor = ConsoleColor.Red;
            //        Console.WriteLine($"Beim Bewerten der Abgegebenen Angaben ist ein Fehler aufgetreten: {e.Message}");
            //        Console.ForegroundColor = ConsoleColor.White;
            //    }
            //}




            Console.WriteLine("\nEnde.");
            Console.ReadLine();
        }



        static void GenerateQuestionsAndSolutions(Excel_Generator.Utils.Utils.StudentObject[] students, string solutionFilePath, string solutionFolderPath, string questionFolderPath, string password = "")
        {
            Clear();
            IWorkbook OGbook = WorkbookFactory.Create(solutionFilePath);
            //IWorkbook book = WorkbookFactory.Create("../files/Anlagenverkauf 2 Prüfung.xlsx");

            Console.WriteLine("> Lade Blätter...");
            ISheet OGquestions = OGbook.GetSheet("Aufgabe");
            ISheet OGsolutions = OGbook.GetSheet("Lösung");
            ISheet OGcfgSheet = OGbook.GetSheet("Konfiguration");
            Console.WriteLine("> Blätter geladen.");
            Console.WriteLine();

            {
                if (!Directory.Exists(solutionFolderPath))
                    Directory.CreateDirectory(solutionFolderPath);

                if (!Directory.Exists(solutionFolderPath + "/TXT"))
                    Directory.CreateDirectory(solutionFolderPath + "/TXT");

                if (!Directory.Exists(solutionFolderPath + "/EXCEL"))
                    Directory.CreateDirectory(solutionFolderPath + "/EXCEL");

                if (!Directory.Exists(questionFolderPath))
                    Directory.CreateDirectory(questionFolderPath);
            }


            {
                //string[] files = Directory.GetFiles(questionFolderPath);
                //foreach (string file in files)
                //    File.Delete(file);

                //files = Directory.GetFiles(solutionFolderPath);
                //foreach (string file in files)
                //    File.Delete(file);

                //files = Directory.GetFiles(solutionFolderPath + "/TXT");
                //foreach (string file in files)
                //    File.Delete(file);

                //files = Directory.GetFiles(solutionFolderPath + "/EXCEL");
                //foreach (string file in files)
                //    File.Delete(file);
            }

            if (students.Length == 0)
                return;

            ConfigThing config = ParseConfig(OGcfgSheet);

            int amountOfStudents = students.Length;

            XSSFWorkbook[] workbooks = new XSSFWorkbook[amountOfStudents];
            SolutionClass[] sols = new SolutionClass[amountOfStudents];


            XSSFFormulaEvaluator[] eval = new XSSFFormulaEvaluator[amountOfStudents];
            ISheet[] sol = new ISheet[amountOfStudents];
            XSSFWorkbook[] solBooks = new XSSFWorkbook[amountOfStudents];

            XSSFFormulaEvaluator[] eval2 = new XSSFFormulaEvaluator[amountOfStudents];
            ISheet[] sol2 = new ISheet[amountOfStudents];
            XSSFWorkbook[] solBooks2 = new XSSFWorkbook[amountOfStudents];

            GradingConfig gradingConfig = new GradingConfig();

            //XSSFFormulaEvaluator.EvaluateAllFormulaCells();

            Console.WriteLine("> Kopiere Angaben...");
            for (int i = 0; i < amountOfStudents; i++)
            {
                Console.WriteLine($"> Angabe {i + 1}/{amountOfStudents}.");

                {
                    XSSFWorkbook solBook = new XSSFWorkbook();
                    OGsolutions.CopyTo(solBook, "Loesung", true, true);
                    //XSSFFormulaEvaluator.EvaluateAllFormulaCells(solBook);
                    sol[i] = solBook.GetSheetAt(0);
                    eval[i] = new XSSFFormulaEvaluator(solBook);
                    solBooks[i] = solBook;
                    //solBook.Write(new FileStream("Test123.xlsx", FileMode.Create));
                }

                {
                    XSSFWorkbook solBook = new XSSFWorkbook();
                    OGquestions.CopyTo(solBook, "Loesung", true, true);
                    //XSSFFormulaEvaluator.EvaluateAllFormulaCells(solBook);
                    sol2[i] = solBook.GetSheetAt(0);
                    eval2[i] = new XSSFFormulaEvaluator(solBook);
                    solBooks2[i] = solBook;
                    //solBook.Write(new FileStream("Test123.xlsx", FileMode.Create));
                }


                XSSFWorkbook tempBook = new XSSFWorkbook();
                workbooks[i] = tempBook;
                sols[i] = new SolutionClass();


                //style.CloneStyleFrom
                OGquestions.CopyTo(tempBook, "Aufgabe", true, true);

                XSSFSheet sheet = tempBook.GetSheet("Aufgabe") as XSSFSheet;

                SetCellFromXY(sheet, 100 + students[i].id, 0, 999); //101+i

                System.Collections.IEnumerator enumerator = sheet.GetRowEnumerator();
                enumerator.MoveNext();


                while (true)
                {
                    IRow row = (IRow)enumerator.Current;

                    foreach (ICell cell in row.Cells)
                    {
                        //Console.WriteLine($" - Zelle: {Utils.GetCellValueAsString(cell)}");
                        //cell.CellStyle = new XSSFCellStyle()

                        //XSSFCellStyle Cstyle = new XSSFCellStyle(tempBook.GetStylesSource());
                        //XSSFCellStyle tempS = (XSSFCellStyle)cell.CellStyle;
                        //if (tempS != null)
                        //{
                        //    try
                        //    {
                        //        Cstyle.CloneStyleFrom(tempS);
                        //    }
                        //    catch (Exception e) { }
                        //}
                        //Cstyle.IsLocked = true;

                        //cell.CellStyle = Cstyle;

                        try
                        {
                            cell.CellStyle.IsLocked = true;
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine($">   Die Zelle {cell.Address} hatte ungültige Formatierungen.");
                            //Console.WriteLine($" -> Error at {cell.Address}: {e}");

                            Console.WriteLine($"- Exception: {e.Message}");

                            cell.CellStyle = new XSSFCellStyle(tempBook.GetStylesSource());
                            cell.CellStyle.CloneStyleFrom(tempBook.GetStylesSource().GetStyleAt(1));
                            cell.CellStyle.IsLocked = true;
                        }

                        // CellUtil.SetCellStyleProperty(cell, "locked", true);
                    }
                    //Console.WriteLine();
                    if (!enumerator.MoveNext())
                        break;
                }

            }
            Console.WriteLine("> Angaben kopiert.");

            Random rnd = new Random();

            Console.WriteLine("> Füge die Zufallswerte ein...");
            //Console.WriteLine("Cells:");
            ICell[,] cells = GetCellsFromXY(OGsolutions, 0, 0, 100, 100);

            for (int y = 0; y < cells.GetLength(0); y++)
            {
                for (int x = 0; x < cells.GetLength(1); x++)
                {
                    ICell cell = cells[x, y];
                    if (cell == null)
                        continue;

                    XSSFColor col = (XSSFColor)cell.CellStyle.FillForegroundColorColor;

                    if (config.randomValDict.Exists(col))
                    {
                        //Console.WriteLine("> Setze Farbe");
                        List<double> list = config.randomValDict.Getval(col);

                        for (int i = 0; i < amountOfStudents; i++)
                        {
                            XSSFWorkbook tempBook = workbooks[i];

                            XSSFSheet sheet = (XSSFSheet)tempBook.GetSheetAt(0);

                            double vall = list[rnd.Next(list.Count)];
                            //Console.WriteLine($"Setze {x}, {y} auf: {vall}");
                            SetCellFromXY(sheet, vall, x, y);
                            SetCellFromXY(sol[i], vall, x, y);
                            SetCellFromXY(sol2[i], vall, x, y);

                            //Console.WriteLine($"WERT: {Utils.GetCellFromXY(sheet, x, y)}");
                        }
                    }
                }
                //Console.WriteLine();
            }
            Console.WriteLine("> Werte Eingefügt.");
            Console.WriteLine();



            Console.WriteLine("> Fülle die Loesungen ein...");
            //Console.WriteLine("Cells:");

            for (int y = 0; y < cells.GetLength(0); y++)
            {
                for (int x = 0; x < cells.GetLength(1); x++)
                {
                    ICell cell = cells[x, y];
                    if (cell == null)
                        continue;

                    XSSFColor col = (XSSFColor)cell.CellStyle.FillForegroundColorColor;
                    if (config.scoreDict.Exists(col))
                    {
                        for (int i = 0; i < amountOfStudents; i++)
                        {
                            XSSFWorkbook tempBook = workbooks[i];

                            ISheet sheet = tempBook.GetSheetAt(0);

                            ICell Tcell = GetCellFromXY(sheet, x, y);

                            CellValue Cval = eval[i].Evaluate(GetCellFromXY(sol[i], x, y));
                            //Console.WriteLine($"Wert bei {x} {y}: {GetCellValueAsString(Cval)}");

                            sols[i].AddVal(x, y, GetCellValueAsString(Cval), config.scoreDict.Getval(col));

                            SetCellFromXY(sol2[i], GetCellValueAsObject(Cval), x, y);

                            SetCellFromXY(sol2[i], 100 + students[i].id, 0, 999); //101+i

                            try
                            {
                                Tcell.CellStyle.IsLocked = false;
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine($" -> Fehler bei ({x},{y} = {Tcell.Address}): {e}");
                                Tcell.CellStyle = new XSSFCellStyle(tempBook.GetStylesSource());
                                Tcell.CellStyle.CloneStyleFrom(tempBook.GetStylesSource().GetStyleAt(1));
                                Tcell.CellStyle.IsLocked = false;
                            }
                        }
                    }
                }
                //Console.WriteLine();
            }
            Console.WriteLine("> Loesungen ausgefüllt.");
            Console.WriteLine();


            Console.WriteLine("> Schreibe Konfiguration...");
            //Console.WriteLine("Cells:");

            for (int y = 0; y < cells.GetLength(0); y++)
            {
                for (int x = 0; x < cells.GetLength(1); x++)
                {
                    ICell cell = cells[x, y];
                    if (cell == null)
                        continue;

                    XSSFColor col = (XSSFColor)cell.CellStyle.FillForegroundColorColor;
                    if (config.gradingDict.Exists(col))
                    {
                        GradingConfig.GradingThings type = config.gradingDict.Getval(col);
                        if (type == GradingConfig.GradingThings.totalScore)
                        {
                            gradingConfig.things.Add((x, y), GradingConfig.GradingThings.totalScore);
                        }
                        else if (type == GradingConfig.GradingThings.score)
                        {
                            gradingConfig.things.Add((x, y), GradingConfig.GradingThings.score);
                        }
                    }
                }
                //Console.WriteLine();
            }
            gradingConfig.SaveToFile(solutionFolderPath + "/config.cfg");
            Console.WriteLine("> Konfiguration geschrieben.");
            Console.WriteLine();







            Console.WriteLine("> Speichere Angaben und Loesungen...");
            for (int i = 0; i < amountOfStudents; i++)
            {
                Console.WriteLine($"> Angabe {i + 1}/{amountOfStudents}.");
                string Qname = $"{questionFolderPath}/Angabe {students[i].name}.xlsx";
                string Sname = $"{solutionFolderPath}/TXT/Loesung {students[i].id}.txt";
                string Sname2 = $"{solutionFolderPath}/EXCEL/Loesung {students[i].id}.xlsx";

                XSSFWorkbook tempBook = workbooks[i];

                XSSFSheet sheet = (XSSFSheet)tempBook.GetSheetAt(0);

                sheet.EnableLocking();
                sheet.ProtectSheet(password);
                sheet.LockSelectLockedCells(false);

                tempBook.Write(new FileStream(Qname, FileMode.Create));
                tempBook.Close();

                solBooks2[i].Write(new FileStream(Sname2, FileMode.Create));
                solBooks2[i].Close();

                solBooks[i].Close();


                sols[i].WriteToFile(Sname);
            }
            Console.WriteLine("> Dateien gespeichert.");


            OGbook.RemoveSheetAt(OGbook.GetSheetIndex("Lösung"));
            OGbook.RemoveSheetAt(OGbook.GetSheetIndex("Konfiguration"));
            OGquestions.ProtectSheet(password);

            //OGbook.Write(new FileStream("Vorlage test.xlsx", FileMode.Create));
        }

        static string getColValasString(XSSFColor col)
        {
            return $"{col.ARGBHex} ({col.Theme}, {Math.Round(col.Tint, 3).ToString(CultureInfo.InvariantCulture)})";
        }

        static ConfigThing ParseConfig(ISheet configSheet)
        {
            ConfigThing cfg = new ConfigThing();


            Console.WriteLine("> Lade Daten der Konfiguration...");
            //Console.WriteLine("Cells:");
            ICell[,] cells = GetCellsFromXY(configSheet, 0, 0, 50, 50);




            for (int y = 0; y < cells.GetLength(0); y++)
            {
                ICell currCell = cells[0, y];
                if (currCell == null)
                    continue;

                if (String.IsNullOrWhiteSpace(GetCellValueAsString(currCell)))
                    continue;

                string cfgval = GetCellValueAsString(currCell).ToLower();

                //Console.WriteLine($"> Konfiguration für \"{cfgval}\" gefunden.");

                if (cfgval.StartsWith("punkte"))
                {
                    int y1 = y;
                    while (true)
                    {
                        y1++;
                        ICell cell = cells[0, y1];
                        if (cell == null)
                            break;
                        if (String.IsNullOrWhiteSpace(GetCellValueAsString(cell)))
                            break;

                        XSSFColor col = (XSSFColor)cell.CellStyle.FillForegroundColorColor;
                        string val = GetCellValueAsString(cell);
                        string colval = getColValasString(col);

                        //Console.WriteLine($"> - Farbe: {colval}, Wert: {val}");

                        cfg.scoreDict.Add(col, double.Parse(val, NumberStyles.Float, CultureInfo.InvariantCulture));

                    }
                    //Console.WriteLine();
                    y = y1 - 1;
                    continue;
                }

                if (cfgval.StartsWith("bewertung"))
                {
                    int y1 = y;
                    while (true)
                    {
                        y1++;
                        ICell cell = cells[0, y1];
                        if (cell == null)
                            break;
                        if (String.IsNullOrWhiteSpace(GetCellValueAsString(cell)))
                            break;

                        XSSFColor col = (XSSFColor)cell.CellStyle.FillForegroundColorColor;
                        string val = GetCellValueAsString(cell).ToLower();

                        if (val.StartsWith("gesamt"))
                        {
                            cfg.gradingDict.Add(col, GradingConfig.GradingThings.totalScore);
                        }
                        else if (val.StartsWith("erreicht"))
                        {
                            cfg.gradingDict.Add(col, GradingConfig.GradingThings.score);
                        }


                        //Console.WriteLine($"> - Farbe: {col.ARGBHex}, Wert: {val}");

                        //cfg.scoreDict.Add(col, double.Parse(val, NumberStyles.Float, CultureInfo.InvariantCulture));

                    }
                    //Console.WriteLine();
                    y = y1 - 1;
                    continue;
                }

                if (cfgval.StartsWith("zufall"))
                {
                    int y1 = y;
                    while (true)
                    {
                        y1++;
                        ICell cell = cells[0, y1];
                        if (cell == null)
                            break;

                        XSSFColor col = (XSSFColor)cell.CellStyle.FillForegroundColorColor;

                        if (col.ARGBHex.Equals("FFFFFFFF"))
                            break;

                        if (!cfg.randomValDict.Exists(col))
                            cfg.randomValDict.Add(col, new List<double>());

                        List<double> list = cfg.randomValDict.Getval(col);
                        //Console.Write($"> - Farbe: {col.ARGBHex}, Werte:");

                        int x1 = 0;
                        while (true)
                        {
                            x1++;
                            ICell cellA = cells[x1, y1];
                            if (cellA == null)
                                break;
                            if (String.IsNullOrWhiteSpace(GetCellValueAsString(cellA)))
                                break;

                            string val = GetCellValueAsString(cellA);
                            //Console.Write($" {val}");
                            list.Add(double.Parse(val, NumberStyles.Float, CultureInfo.InvariantCulture));
                        }
                        //Console.WriteLine();

                    }
                    //Console.WriteLine();
                    y = y1 - 1;
                    continue;
                }


            }
            Console.WriteLine("> Daten geladen.");
            Console.WriteLine();




            return cfg;
        }


        public static void GradeWorksheets((StudentObject student, string path)[] studentFileArray, string solutionFolderPath, string gradedFolderPath)
        {
            {
                if (!Directory.Exists(gradedFolderPath))
                    Directory.CreateDirectory(gradedFolderPath);

                if (!Directory.Exists(gradedFolderPath + "/EXCEL"))
                    Directory.CreateDirectory(gradedFolderPath + "/EXCEL");

                if (!Directory.Exists(gradedFolderPath + "/TXT"))
                    Directory.CreateDirectory(gradedFolderPath + "/TXT");

                //string[] files_ = Directory.GetFiles(gradedFolderPath);
                //foreach (string file in files_)
                //    File.Delete(file);
            }

            GradingConfig cfg = GradingConfig.LoadFromFile(solutionFolderPath + "/config.cfg");


            Console.WriteLine();
            Console.WriteLine("Lange Ergebnisliste:");
            Console.WriteLine("----------------------------------------------------------");
            Console.WriteLine();

            foreach (var file in studentFileArray)
            {
                Console.WriteLine($"> Schaue Datei \"{file}\" an:");

                try
                {
                    IWorkbook book = WorkbookFactory.Create(file.path);

                    ISheet main = book.GetSheetAt(0);

                    int id = (int)(GetCellFromXY(main, 0, 999).NumericCellValue - 100);
                    //Console.WriteLine($"ID: {id}");

                    GradeWorkSheet(book, $"{solutionFolderPath}/TXT/Loesung {id}.txt", file.path, true, id, $"{gradedFolderPath}/EXCEL/{Path.GetFileName(file.path)}", $"{gradedFolderPath}/TXT/Statistik {file.student.id}.txt", cfg);

                    book.Close();
                }
                catch (Exception e)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Bei der Datei \"{file.path}\" ist ein Fehler aufgetreten! ({e})");
                    Console.ForegroundColor = ConsoleColor.White;
                }
            }

            Console.WriteLine();
            Console.WriteLine("Kurze Ergebnisliste:");
            Console.WriteLine("----------------------------------------------------------");
            Console.WriteLine();

            foreach (var file in studentFileArray)
            {
                try
                {
                    IWorkbook book = WorkbookFactory.Create(file.path);

                    ISheet main = book.GetSheetAt(0);

                    int id = (int)(GetCellFromXY(main, 0, 999).NumericCellValue - 100);
                    //Console.WriteLine($"ID: {id}");

                    GradeWorkSheet(book, $"{solutionFolderPath}/TXT/Loesung {id}.txt", file.path, false, id, "", "", cfg);

                    book.Close();
                }
                catch (Exception e)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Bei der Datei \"{file.path}\" ist ein Fehler aufgetreten! ({e})");
                    Console.ForegroundColor = ConsoleColor.White;
                }
            }
        }

        static void GradeWorkSheet(IWorkbook book, string solutionPath, string filename, bool show_deb, int id, string gradedExcelFilePath, string gradedTextFilePath, GradingConfig cfg)
        {
            ISheet main = book.GetSheetAt(0);
            XSSFFormulaEvaluator eval = new XSSFFormulaEvaluator(book);
            SolutionClass sol = new SolutionClass(solutionPath);

            List<double> pointsGained = new List<double>();

            double score = 0, maxScore = 0;

            foreach (SolutionThing thing in sol.vals)
            {
                if (thing.points == 0)
                    continue;

                maxScore += thing.points;

                string val = "";
                ICell cell = GetCellFromXY(main, thing.x, thing.y);

                if (cell != null)
                {
                    try
                    {
                        val = GetCellValueAsString(eval.Evaluate(cell));
                    }
                    catch
                    {
                        val = "<ERROR>";
                    }
                }

                if (show_deb) Console.Write($"        > Überprüfe bei {GetCellFromXY(main, thing.x, thing.y).Address} ({thing.x}, {thing.y}): \"{val}\" == \"{thing.value}\" ? ({thing.points} Punkte) -> ");


                if (val.ToLower().Equals(thing.value.ToLower()))
                {
                    Console.ForegroundColor = ConsoleColor.DarkGreen;
                    score += thing.points;
                    if (show_deb)
                    {
                        Console.Write($"Passt. {thing.points}/{thing.points}");
                        pointsGained.Add(thing.points);

                        if (cell != null)
                        {
                            XSSFCellStyle style = (XSSFCellStyle)book.CreateCellStyle();
                            style.CloneStyleFrom(cell.CellStyle);
                            style.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightGreen.Index;
                            cell.CellStyle = style;
                        }
                    }
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.DarkRed;
                    if (show_deb)
                    {
                        Console.Write($"Passt nicht. 0/{thing.points}");
                        pointsGained.Add(0);

                        if (cell != null)
                        {
                            XSSFCellStyle style = (XSSFCellStyle)book.CreateCellStyle();
                            style.CloneStyleFrom(cell.CellStyle);
                            style.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Red.Index;
                            cell.CellStyle = style;
                        }
                    }
                }
                Console.ForegroundColor = ConsoleColor.Yellow;
                if (show_deb) Console.WriteLine($" - ({score}/{maxScore} Punkte)");
                Console.ForegroundColor = ConsoleColor.White;
            }

            //if (show_deb) Console.Write("\n\t");
            //Console.WriteLine($"> Schüler: \"{GetCellValueAsString(GetCellFromXY(main, 1, 0))}\" (\"{filename}\")");
            //if (show_deb) Console.Write("\t");
            //Console.WriteLine($"> Punkteanzahl: {score}/{maxScore}. ({Math.Round(100 * (score / maxScore), 2)}%)");
            //if (show_deb) Console.WriteLine("-------------------------------------------------");
            //Console.WriteLine();

            if (show_deb)
            {

                foreach (var thing in cfg.things)
                {
                    switch (thing.Value)
                    {
                        case GradingConfig.GradingThings.totalScore:
                            SetCellFromXY(main, maxScore, thing.Key.x, thing.Key.y);
                            break;

                        case GradingConfig.GradingThings.score:
                            SetCellFromXY(main, score, thing.Key.x, thing.Key.y);
                            break;
                    }
                }

                using (StreamWriter writer = new StreamWriter(gradedTextFilePath))
                {
                    writer.WriteLine("Punkte:");
                    writer.WriteLine(score.ToString(CultureInfo.InvariantCulture));
                    writer.WriteLine("Von:");
                    writer.WriteLine(maxScore.ToString(CultureInfo.InvariantCulture));
                    writer.WriteLine("Aufgaben:");
                    writer.WriteLine(pointsGained.Count.ToString(CultureInfo.InvariantCulture));
                    writer.WriteLine("Liste:");
                    foreach (var x in pointsGained)
                        writer.WriteLine(x.ToString(CultureInfo.InvariantCulture));
                }

                book.Write(new FileStream(gradedExcelFilePath, FileMode.Create));
            }


            PrintStudent(
                GetCellValueAsString(GetCellFromXY(main, 1, 0)),
                filename,
                score,
                maxScore,
                Math.Round(100 * (score / maxScore), 2),
                show_deb,
                id
            );
        }

        static ConsoleColor[] scoreCols = new ConsoleColor[] { ConsoleColor.DarkRed, ConsoleColor.Red, ConsoleColor.Yellow, ConsoleColor.Green, ConsoleColor.DarkGreen };

        static void PrintStudent(string studentName, string filename, double score, double maxScore, double percent, bool show_deb, int id)
        {
            {
                if (show_deb) Console.Write("\n\t");

                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.Write($"> Schüler: ");

                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.Write($"\"{studentName}\" (Gruppe: {id})");

                Console.ForegroundColor = ConsoleColor.White;
                Console.Write("    -  ");


                Console.ForegroundColor = ConsoleColor.DarkYellow;
                Console.Write($"(\"{filename}\")");

                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine();
            }

            int grade = 5;
            ConsoleColor scoreColor = scoreCols[0];
            {
                /* 
                1: 100-91
                2: 90-81
                3: 80-66
                4: 65-50
                5: < 50
                */

                if (percent >= 91)
                    grade = 4;
                else if (percent >= 81)
                    grade = 3;
                else if (percent >= 66)
                    grade = 2;
                else if (percent >= 50)
                    grade = 1;
                else
                    grade = 0;

                scoreColor = scoreCols[grade];
                grade = 5 - grade;
            }

            {

                if (show_deb) Console.Write("\t");
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.Write($"> Punkteanzahl: ");

                Console.ForegroundColor = scoreColor;
                Console.Write(score);

                Console.ForegroundColor = ConsoleColor.White;
                Console.Write("/");

                Console.ForegroundColor = ConsoleColor.Blue;
                Console.Write(maxScore);


                Console.ForegroundColor = ConsoleColor.White;
                Console.Write($". (");

                Console.ForegroundColor = scoreColor;
                Console.Write($"{grade} - {percent}%");

                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine(")");
            }


            Console.ForegroundColor = ConsoleColor.White;

            if (show_deb) Console.WriteLine("-------------------------------------------------");
            Console.WriteLine();
        }





        /*
         //IRow row = sheet.GetRow(2); // y

            //ICell cell = row.GetCell(1); // x

            if (false)
            {
                ICell cell = Utils.GetCellFromXY(sheet, 1, 2);

                Console.WriteLine($"Value old: {Utils.GetCellValueAsString(cell)}");

                //cell.SetCellValue("LMAO");

                Console.WriteLine($"Value new: {Utils.GetCellValueAsString(cell)}");

                Utils.SetCellFromXY(sheet, 3.1415, 1, 2);

                Console.WriteLine($"Value new 2: {Utils.GetCellValueAsString(cell)}");

                Utils.SetCellFromXY(sheet, new Utils.Formula("B3*10"), 2, 2);
            }
         */
    }
}
