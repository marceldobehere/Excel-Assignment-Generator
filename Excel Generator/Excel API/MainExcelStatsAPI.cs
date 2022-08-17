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
using System.Windows;
using Excel_Generator.Utils;
using static Excel_Generator.Utils.LocalizationManager.LanguagePhraseList;

namespace Excel_Generator.Excel_API
{
    public class MainExcelStatsAPI
    {
        /*
            Punkte:
            3
            Von:
            3
            Aufgaben:
            3
            Liste:
            1
            1
            1
         */


        public static void CreateAssignmentStatistics(StudentObject[] studentsToShow, StudentObject[] studentsNotUploaded, string dataFilepath, string statFilepath, string statTitle, string year, string className, string assignment)
        {
            try
            {
                Dictionary<int, string> files = new Dictionary<int, string>();
                double totalPoints = 0;
                Dictionary<int, double?> points = new Dictionary<int, double?>();

                foreach (var student in studentsToShow)
                    points[student.id] = null;

                foreach (var student in studentsNotUploaded)
                    points[student.id] = 0;

                foreach (var file in Directory.GetFiles(dataFilepath))
                {
                    var arr = Path.GetFileNameWithoutExtension(file).Split(' ');

                    if (arr.Length == 0)
                        continue;
                    if (!int.TryParse(arr[arr.Length - 1], out int id))
                        continue;

                    using (StreamReader reader = new StreamReader(file))
                    {
                        List<string> data = new List<string>();

                        while (!reader.EndOfStream)
                            data.Add(reader.ReadLine());

                        if (data.Count < 4)
                            continue;

                        if (!double.TryParse(data[1], NumberStyles.Float, CultureInfo.InvariantCulture, out double _points))
                            continue;
                        if (!double.TryParse(data[3], NumberStyles.Float, CultureInfo.InvariantCulture, out double _total))
                            continue;

                        totalPoints = _total;
                        points[id] = _points;
                    }

                    files[id] = file;
                }


                var workbook = new XSSFWorkbook();

                var sheet = workbook.CreateSheet(statTitle);

                SetCellFromXY(sheet, statTitle, 0, 0);
                SetCellFromXY(sheet, year, 1, 0);
                SetCellFromXY(sheet, className, 2, 0);
                SetCellFromXY(sheet, assignment, 3, 0);


                SetCellFromXY(sheet, LocalizationManager.GetPhrase(Phrase.Statistics_StudentText), 0, 2);
                SetCellFromXY(sheet, LocalizationManager.GetPhrase(Phrase.Statistics_PointsText), 1, 2);
                SetCellFromXY(sheet, LocalizationManager.GetPhrase(Phrase.Statistics_TotalPointsText), 2, 2);
                SetCellFromXY(sheet, LocalizationManager.GetPhrase(Phrase.Statistics_PercentText), 3, 2);
                SetCellFromXY(sheet, LocalizationManager.GetPhrase(Phrase.Statistics_GradeText), 4, 2);

                IFont font = workbook.CreateFont();
                font.FontHeightInPoints = 11;
                font.FontName = "Arial";
                font.Boldweight = (short)FontBoldWeight.Bold;

                for (int x = 0; x < 4; x++)
                {
                    CellUtil.SetAlignment(sheet.GetRow(0).GetCell(x), NPOI.SS.UserModel.HorizontalAlignment.Left);
                    sheet.GetRow(0).GetCell(x).CellStyle.SetFont(font);
                }

                {
                    for (int x = 0; x < 5; x++)
                    {
                        CellUtil.SetAlignment(sheet.GetRow(2).GetCell(x), NPOI.SS.UserModel.HorizontalAlignment.Center);

                        sheet.GetRow(2).GetCell(x).CellStyle.SetFont(font);
                    }


                    int row = 3;
                    foreach (var student in studentsToShow)
                    {
                        SetCellFromXY(sheet, student.name, 0, row);
                        if (points[student.id] != null)
                        {
                            SetCellFromXY(sheet, points[student.id], 1, row);
                            SetCellFromXY(sheet, totalPoints, 2, row);
                            SetCellFromXY(sheet, $"{Math.Round(points[student.id].Value / totalPoints * 100, 2)}%", 3, row);
                            SetCellFromXY(sheet, GetGrade((int)(points[student.id] / totalPoints * 100)), 4, row);
                        }
                        else
                        {
                            SetCellFromXY(sheet, "-", 1, row);
                            SetCellFromXY(sheet, "-", 2, row);
                            SetCellFromXY(sheet, "-", 3, row);
                            SetCellFromXY(sheet, "-", 4, row);
                        }


                        try
                        {
                            CellUtil.SetAlignment(sheet.GetRow(row).GetCell(0), NPOI.SS.UserModel.HorizontalAlignment.Left);
                            for (int x = 1; x < 5; x++)
                            {
                                CellUtil.SetAlignment(sheet.GetRow(row).GetCell(x), NPOI.SS.UserModel.HorizontalAlignment.Center);
                            }
                        }
                        catch (Exception e)
                        {
                            
                        }

                        row++;
                    }
                }

                for (int row = 2; row < sheet.PhysicalNumberOfRows; row++)
                {
                    if (sheet.GetRow(row) == null)
                        continue;
                    
                    for (int col = 0; col < sheet.GetRow(row).PhysicalNumberOfCells; col++)
                    {
                        if (sheet.GetRow(row).GetCell(col) == null)
                            continue;
                        
                        sheet.GetRow(row).GetCell(col).CellStyle.BorderBottom = BorderStyle.Thin;
                        sheet.GetRow(row).GetCell(col).CellStyle.BorderLeft = BorderStyle.Thin;
                        sheet.GetRow(row).GetCell(col).CellStyle.BorderRight = BorderStyle.Thin;
                        sheet.GetRow(row).GetCell(col).CellStyle.BorderTop = BorderStyle.Thin;
                    }
                }

                for (int i = 0; i < 10; i++)
                    sheet.AutoSizeColumn(i);


                workbook.Write(new FileStream(statFilepath, FileMode.Create));

                workbook.Close();

                Excel_Generator.Utils.Utils.OpenWithDefaultProgram(statFilepath);
            }
            catch (Exception e)
            {
                Console.WriteLine($"Fehler: {e.Message}");
                MessageBox.Show($"Fehler: {e.Message}");
            }
        }


    }
}
