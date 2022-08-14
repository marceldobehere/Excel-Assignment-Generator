using System;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Excel_Generator.Excel_API.Utils
{
    static class Utils
    {
        public static string GetCellValueAsString(ICell cell)
        {
            if (cell == null)
                return "";

            CellType type = cell.CellType;

            if (type == CellType.Unknown || type == CellType.Blank)
                return "";
            if (type == CellType.Boolean)
                return cell.BooleanCellValue.ToString(CultureInfo.InvariantCulture);
            if (type == CellType.String)
                return cell.StringCellValue;
            if (type == CellType.Numeric)
                return Math.Round(cell.NumericCellValue, 3).ToString(CultureInfo.InvariantCulture);
            if (type == CellType.Error)
                return cell.ErrorCellValue.ToString(CultureInfo.InvariantCulture);
            if (type == CellType.Formula)
                return cell.CellFormula;

            return "";
        }

        public static string GetCellValueAsString(CellValue cell)
        {
            if (cell == null)
                return "";

            CellType type = cell.CellType;

            if (type == CellType.Unknown || type == CellType.Blank)
                return "";
            if (type == CellType.Boolean)
                return cell.BooleanValue.ToString(CultureInfo.InvariantCulture);
            if (type == CellType.String)
                return cell.StringValue;
            if (type == CellType.Numeric)
                return Math.Round(cell.NumberValue, 3).ToString(CultureInfo.InvariantCulture);
            if (type == CellType.Error)
                return cell.ErrorValue.ToString(CultureInfo.InvariantCulture);


            return "";
        }

        public static object GetCellValueAsObject(CellValue cell)
        {
            if (cell == null)
                return null;

            CellType type = cell.CellType;

            if (type == CellType.Unknown || type == CellType.Blank)
                return "";
            if (type == CellType.Boolean)
                return cell.BooleanValue;
            if (type == CellType.String)
                return cell.StringValue;
            if (type == CellType.Numeric)
                return cell.NumberValue;
            if (type == CellType.Error)
                return cell.ErrorValue;


            return null;
        }

        public static ICell GetCellFromXY(ISheet sheet, int x, int y)
        {
            if (sheet == null)
                return null;

            IRow row = sheet.GetRow(y);
            if (row == null)
                row = sheet.CreateRow(y);

            ICell cell = row.GetCell(x);
            if (cell == null)
                cell = row.CreateCell(y);

            return cell;
        }

        //static ICell SetCellFromXY(ISheet sheet, ICell cell, int x, int y)
        //{
        //    if (sheet == null)
        //        throw new Exception("Worksheet is null!");

        //    IRow row = sheet.GetRow(y);
        //    if (row == null)
        //        return null;

        //    ICell cell2 = row.GetCell(x);
        //    if (cell2 == null)
        //    {
        //        row.CreateCell(x);
        //        cell2 = row.GetCell(x);
        //    }

        //    cell2.SetCellValue(GetCellValueAsString(cell));

        //    return cell2;
        //}

        public class Formula
        {
            public string formula = "";
            public Formula(string formula)
                => this.formula = formula;
        }

        public static ICell SetCellFromXY(ISheet sheet, object val, int x, int y)
        {
            if (sheet == null)
                throw new Exception("Worksheet is null!");

            IRow row = sheet.GetRow(y);
            if (row == null)
                row = sheet.CreateRow(y);


            ICell cell = row.GetCell(x);
            if (cell == null)
            {
                row.CreateCell(x);
                cell = row.GetCell(x);
            }

            if (val is bool)
            {
                cell.SetCellValue((bool)(object)val);
                cell.SetCellType(CellType.Boolean);
            }
            else if (val is string)
            {
                cell.SetCellValue((string)(object)val);
                cell.SetCellType(CellType.String);
            }
            else if (val is DateTime)
            {
                cell.SetCellValue((DateTime)(object)val);
                cell.SetCellType(CellType.String);
            }
            else if (val is double || val is int)
            {
                if (val is int)
                    cell.SetCellValue((int)val * 1.0d);
                else
                    cell.SetCellValue((double)val);

                cell.SetCellType(CellType.Numeric);
            }
            else if (val is IRichTextString)
            {
                cell.SetCellValue((IRichTextString)(object)val);
                cell.SetCellType(CellType.String);
            }
            else if (val is Formula)
            {
                cell.SetCellFormula(((Formula)val).formula);
                cell.SetCellType(CellType.Formula);
            }
            else if (val == null)
            {
                cell.SetBlank();
                cell.SetCellType(CellType.Blank);
            }


            return cell;
        }


        public static ICell[,] GetCellsFromXY(ISheet sheet, int x1, int y1, int x2, int y2)
        {
            if (sheet == null)
                return null;

            if (x1 > x2)
                (x1, x2) = (x2, x1);
            if (y1 > y2)
                (y1, y2) = (y2, y1);

            int w = x2 - x1 + 1;
            int h = y2 - y1 + 1;

            ICell[,] cells = new ICell[w, h];


            for (int y = y1; y <= y2; y++)
            {
                IRow row = sheet.GetRow(y);
                if (row != null)
                    for (int x = x1; x <= x2; x++)
                        cells[x - x1, y - y1] = row.GetCell(x);
            }

            return cells;
        }


        public class GradingConfig
        {
            public enum GradingThings
            {
                totalScore,
                score
            }

            public Dictionary<(int x, int y), GradingThings> things;

            public GradingConfig()
            {
                things = new Dictionary<(int x, int y), GradingThings>();
            }

            public void SaveToFile(string filename)
            {
                using (StreamWriter writer = new StreamWriter(filename))
                {
                    foreach (var thing in things)
                    {
                        writer.WriteLine($"{thing.Key.x}|{thing.Key.y}|{thing.Value}");
                    }
                }
            }


            public static GradingConfig LoadFromFile(string filename)
            {
                GradingConfig cfg = new GradingConfig();

                if (File.Exists(filename))
                {
                    using (StreamReader reader = new StreamReader(filename))
                    {
                        while (!reader.EndOfStream)
                        {
                            try
                            {
                                string line = reader.ReadLine();
                                string[] parts = line.Split('|');
                                if (parts.Length == 3)
                                {
                                    int x = int.Parse(parts[0]);
                                    int y = int.Parse(parts[1]);
                                    GradingThings thing = (GradingThings)Enum.Parse(typeof(GradingThings), parts[2]);

                                    cfg.things.Add((x, y), thing);
                                }
                            }
                            catch
                            {
                                // ignore
                            }
                        }
                    }
                }
                else
                {

                    cfg.SaveToFile(filename);
                }


                return cfg;
            }
        }




        public class ConfigThing
        {
            public ColDict<double, XSSFColor> scoreDict = new ColDict<double, XSSFColor>();
            public ColDict<List<double>, XSSFColor> randomValDict = new ColDict<List<double>, XSSFColor>();
            public ColDict<GradingConfig.GradingThings, XSSFColor> gradingDict = new ColDict<GradingConfig.GradingThings, XSSFColor>();


        }

        public static bool ColEqual(XSSFColor a, XSSFColor b)
        {
            if (a == null || b == null)
            {
                if (a == null && b == null)
                    return true;
                else
                    return false;
            }

            if (!a.ARGBHex.Equals(b.ARGBHex))
                return false;

            if (a.Tint != b.Tint)
                return false;


            if (a.Theme != b.Theme)
                return false;


            return true;
        }

        public class SolutionClass
        {
            public struct SolutionThing
            {
                public int x, y;
                public string value;
                public double points;
            }

            public List<SolutionThing> vals;

            public SolutionClass()
            {
                vals = new List<SolutionThing>();
            }

            public SolutionClass(string path)
            {
                vals = new List<SolutionThing>();

                StreamReader reader = new StreamReader(path);

                while (!reader.EndOfStream)
                {
                    string[] parts = reader.ReadLine().Split('|');
                    vals.Add(new SolutionThing()
                    {
                        x = int.Parse(parts[0]),
                        y = int.Parse(parts[1]),
                        value = parts[2],
                        points = double.Parse(parts[3], NumberStyles.Float, CultureInfo.InvariantCulture)
                    });
                }
            }

            public void AddVal(int x, int y, string value, double points)
            {
                vals.Add(new SolutionThing() { x = x, y = y, value = value, points = points });
            }

            public void WriteToFile(string path)
            {
                StreamWriter writer = new StreamWriter(path);

                foreach (SolutionThing thing in vals)
                    writer.WriteLine($"{thing.x}|{thing.y}|{thing.value}|{thing.points.ToString(CultureInfo.InvariantCulture)}");

                writer.Close();
            }
        }

        public class ColDict<T1, T2>
        {
            public List<T1> values;
            public List<T2> cols;

            public ColDict()
            {
                cols = new List<T2>();
                values = new List<T1>();
            }

            public bool Exists(T2 col)
            {
                for (int i = 0; i < cols.Count; i++)
                    if (ColEqual((XSSFColor)(object)col, (XSSFColor)(object)cols[i]))//(col.Equals(cols[i]))
                        return true;

                return false;
            }


            public T1 Getval(T2 col)
            {
                for (int i = 0; i < cols.Count; i++)
                    if (ColEqual((XSSFColor)(object)col, (XSSFColor)(object)cols[i]))//(col.Equals(cols[i]))
                        return values[i];

                return default(T1);
            }

            public void Add(T2 col, T1 val)
            {
                if (Exists(col))
                    throw new Exception("Colour already exists!");

                cols.Add(col);
                values.Add(val);
            }
        }
    }
}
