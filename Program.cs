using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DichotomieDeuxDimensions
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var width = 16;
            var height = 16;

            var aera = new Aera(0, 0, width, height, 0);

            aera = CutUpAera(aera);

            var aeraToDisplay = new List<Aera>();
            SetId(aera, null, aeraToDisplay);

            ExcelGenerator.CreateExcelFile(@$"C:\tmp\decoupage-{width}x{height}_{DateTime.Now.Ticks}.xlsx", width, height, aeraToDisplay);

            var collections = new List<Aera>();
            GetAllAeras(aera, collections);

            var maxDeepth = collections.Select(a => a.Deepth).Max();
            var dic = new Dictionary<int, IEnumerable<Aera>>();

            foreach(var deepth in Enumerable.Range(0, maxDeepth))
            {
                dic.Add(deepth, collections.Where(t => t.Deepth == deepth).ToList());
            }

            ExcelGenerator.CreateExcelFile(@$"C:\tmp\decoupages-{width}x{height}_{DateTime.Now.Ticks}.xlsx", height, dic);

            Console.WriteLine("Hello world");
        }

        private static Aera CutUpAera(Aera aera)
        {
            if (aera.Length / 2d < 1 && aera.Height / 2d < 1)
                return aera;

            var x = 0;
            var y = 0;
            var xp = 0; 
            var yp = 0;

            if(aera.Length / 2d < 1 && aera.Height / 2d >= 1)
            {
                x = aera.Length;
                y = (int)Math.Ceiling(aera.Height / 2d);
                yp = aera.Height % 2 == 0 ? y : y - 1;

                aera.SubAera = new List<Aera>
                {
                    CutUpAera(new Aera(aera.X, aera.Y, x, y, aera.Deepth + 1)),
                    CutUpAera(new Aera(aera.X, aera.Y + y, x, yp, aera.Deepth + 1))
                };
            }
            else if(aera.Length / 2d >= 1 && aera.Height / 2d < 1)
            {
                x = (int)Math.Ceiling(aera.Length / 2d);
                y = aera.Height;
                xp = aera.Length % 2 == 0 ? x : x - 1;

                aera.SubAera = new List<Aera>
                {
                    CutUpAera(new Aera(aera.X, aera.Y, x, y, aera.Deepth + 1)),
                    CutUpAera(new Aera(aera.X + x, aera.Y, xp, y, aera.Deepth + 1))
                };
            }
            else
            {
                x = (int)Math.Ceiling(aera.Length / 2d);
                y = (int)Math.Ceiling(aera.Height / 2d);

                xp = aera.Length %2 == 0 ? x : x - 1;
                yp = aera.Height  % 2 == 0 ? y : y - 1;

                aera.SubAera = new List<Aera>
                {
                    CutUpAera(new Aera(aera.X, aera.Y, x, y, aera.Deepth + 1)),
                    CutUpAera(new Aera(aera.X, aera.Y + y, x, yp, aera.Deepth + 1)),
                    CutUpAera(new Aera(aera.X + x, aera.Y, xp, y, aera.Deepth + 1)),
                    CutUpAera(new Aera(aera.X + x, aera.Y + y, xp, yp, aera.Deepth + 1)),
                };
            }

            return aera;
        }

        private static void SetId(Aera aera, Aera parent, List<Aera> collection)
        {
            if (aera.SubAera == null)
            {
                var parentCopy = new Aera(aera, parent.Id);
                collection.Add(parentCopy);
            }
            else
                foreach (var a in aera.SubAera)
                    SetId(a, aera, collection);
        }

        private static void GetAllAeras(Aera aera, List<Aera> collection)
        {
            collection.Add(aera);

            if(aera.SubAera != null)
                foreach(var a in aera.SubAera)
                    GetAllAeras(a, collection);
        }
    }

    public static class ExcelHelper
    {
        public static string GetExcelColumnName(int columnNumber)
        {
            string columnName = string.Empty;

            while (columnNumber > 0) 
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }

            return columnName;
        }

        public static Cell GetColumn(this SheetData sheetData, int idRow, int idColumn)
            => sheetData.Elements<Row>().First(t => t.RowIndex == idRow).Elements<Cell>().First(t => t.CellReference == $"{GetExcelColumnName(idColumn)}{idRow}");

        public static Row GetRow(this SheetData sheetData, int idRow)
            => sheetData.Elements<Row>().First(t => t.RowIndex == idRow);
    }

    public class ExcelGenerator
    {
        public static void CreateExcelFile(string filePath, int width,  int height, IEnumerable<Aera> aeras) 
        {
            var qtyColoredAears = aeras.GroupBy(t => t.Id, (k, v) => k).Distinct().Count();
            var dic = aeras.GroupBy(t => t.Id, (k, v) => k).Distinct().Select((k, i) => new { Id = k.ToString(), Index = i + 1 }).ToDictionary(t => t.Id, t => t.Index);
            var dicStr = aeras.GroupBy(t => t.Id, (k, v) => k).Distinct().Select((k, i) => new { Id = k.ToString(), Str = ExcelHelper.GetExcelColumnName(i + 1) }).ToDictionary(t => t.Id, t => t.Str);

            using (var spreadsheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = spreadsheetDocument.AddWorkbookPart();
                spreadsheetDocument.WorkbookPart.Workbook = new Workbook();
                spreadsheetDocument.WorkbookPart.Workbook.Sheets = new Sheets();
                uint sheetId = 1;
                var sheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                sheetPart.Worksheet = new Worksheet(sheetData);

                var stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylePart.Stylesheet = CreateStyleSheet(qtyColoredAears);
                stylePart.Stylesheet.Save();

                var sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                var relationshipId = spreadsheetDocument.WorkbookPart.GetIdOfPart(sheetPart);
                var sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = "Test" };
                sheets.Append(sheet);

                foreach(var idRow in Enumerable.Range(1, height))
                {
                    var row = new Row { RowIndex = (UInt32)idRow };
                    foreach(var idCol in Enumerable.Range(1, width))
                    {
                        var str = aeras.First(t => t.X == idCol - 1 && t.Y == height - idRow).Id.ToString();
                        var cell = new Cell { CellReference = $"{ExcelHelper.GetExcelColumnName(idCol)}{idRow}" };
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(dicStr[str]);
                        cell.StyleIndex = UInt32Value.FromUInt32((UInt32)dic[str]);
                        row.AppendChild(cell);
                    }
                    sheetData.AppendChild(row); 
                }

                workbookPart.Workbook.Save();
            }
        
        }

        public static void CreateExcelFile(string filepPath, int height, Dictionary<int, IEnumerable<Aera>> aeraDic)
        {
            var maxKey = aeraDic.Select(t => t.Key).Max();
            var qtyColoredAeras = aeraDic[maxKey].Count();

            using (var spreadsheetDocument = SpreadsheetDocument.Create(filepPath, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = spreadsheetDocument.AddWorkbookPart();
                spreadsheetDocument.WorkbookPart.Workbook = new Workbook();
                spreadsheetDocument.WorkbookPart.Workbook.Sheets = new Sheets();
                uint sheetId = 0;

                var stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylePart.Stylesheet = CreateStyleSheet(qtyColoredAeras);
                stylePart.Stylesheet.Save();

                foreach(var keyValuePair in aeraDic)
                {
                    ++sheetId;

                    var sheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    sheetPart.Worksheet = new Worksheet(sheetData);
                    var sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                    var relationshipId = spreadsheetDocument.WorkbookPart.GetIdOfPart(sheetPart);

                    var sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = $"Deepth {keyValuePair.Key}" };
                    sheets.Append(sheet);

                    var aeras = keyValuePair.Value;
                    var dic = aeras.GroupBy(t => t.Id, (k, v) => k).Distinct().Select((k, i) => new { Id = k.ToString(), Index = i + 1 }).ToDictionary(t => t.Id, t => t.Index);
                    var dicStr = aeras.GroupBy(t => t.Id, (k, v) => k).Distinct().Select((k, i) => new { Id = k.ToString(), Str = ExcelHelper.GetExcelColumnName(i + 1) }).ToDictionary(t => t.Id, t => t.Str);

                    foreach(var idRow in Enumerable.Range(1, height))
                    {
                        var row = new Row { RowIndex = (UInt32)idRow };
                        sheetData.AppendChild(row); 
                    }

                    var idx = 0;

                    foreach(var a in aeras)
                    {
                        idx++;
                        foreach(var i in Enumerable.Range(1, a.Length))
                        {
                            foreach(var j in Enumerable.Range(1,a.Height))
                            {
                                var iRef = a.X + i;
                                var jRef = a.Y + j;
                                var cellReference = $"{ExcelHelper.GetExcelColumnName(iRef)}{jRef}";
                                var cell = new Cell { CellReference = cellReference };
                                cell.StyleIndex = UInt32Value.FromUInt32((UInt32)idx);
                                var row = sheetData.GetRow(jRef);
                                row.AppendChild(cell);
                            }
                        }
                    }
                }

                workbookPart.Workbook.Save(workbookPart);
            }
        }

        private static Stylesheet CreateStyleSheet(int colorQuantity)
        {
            var styleSheet = new Stylesheet();

            //----------------------------------------------------------
            // Fonts
            //----------------------------------------------------------
            styleSheet.Fonts = new DocumentFormat.OpenXml.Spreadsheet.Fonts();

            var font = new DocumentFormat.OpenXml.Spreadsheet.Font();
            styleSheet.Fonts.Append(font);

            font = new DocumentFormat.OpenXml.Spreadsheet.Font();
            font.Bold = new DocumentFormat.OpenXml.Spreadsheet.Bold();
            font.Bold.Val = BooleanValue.FromBoolean(true);
            styleSheet.Fonts.Append(font);

            styleSheet.Fonts.Count = UInt32Value.FromUInt32((UInt32)styleSheet.Fonts.ChildElements.Count);

            //----------------------------------------------------------
            // Fills
            //----------------------------------------------------------
            styleSheet.Fills = new Fills();

            var fill = new Fill();
            var PatternFillPreset = new PatternFill();
            PatternFillPreset.PatternType = PatternValues.None;
            fill.PatternFill = PatternFillPreset;
            styleSheet.Fills.Append(fill);

            //Fill Index 1. Defaults by Microsoft
            fill = new Fill();
            PatternFillPreset = new PatternFill();
            PatternFillPreset.PatternType = PatternValues.Gray125;
            fill.PatternFill = PatternFillPreset;
            styleSheet.Fills.Append(fill);

            //Fill Index 2 - Custom & Gold
            fill = new Fill();
            var patternFill = new PatternFill();
            patternFill.PatternType = PatternValues.Solid;
            patternFill.ForegroundColor = new ForegroundColor();
            patternFill.ForegroundColor.Rgb = HexBinaryValue.FromString("F9DF02");
            fill.PatternFill = patternFill;
            styleSheet.Fills.Append(fill);

            var temp = new List<string>();

            foreach(var idColor in Enumerable.Range(1, colorQuantity))
            {
                var colorHex = ColorGenerator.GetRGB(idColor + 200).ToString("X");
                temp.Add(colorHex);
                fill = new Fill();
                patternFill = new PatternFill();
                patternFill.PatternType = PatternValues.Solid;
                patternFill.ForegroundColor = new ForegroundColor();
                patternFill.ForegroundColor.Rgb = HexBinaryValue.FromString(colorHex);
                fill.PatternFill = patternFill;
                styleSheet.Fills.Append(fill);
            }

            styleSheet.Fills.Count = UInt32Value.FromUInt32((UInt32)styleSheet.Fills.ChildElements.Count);

            //----------------------------------------------------------
            // Borders
            //----------------------------------------------------------
            styleSheet.Borders = new Borders();

            var border = new Border();
            styleSheet.Borders.Append(border);

            styleSheet.Borders.Count = UInt32Value.FromUInt32((UInt32)styleSheet.Borders.ChildElements.Count);

            //----------------------------------------------------------
            // Cell Formats
            //----------------------------------------------------------
            styleSheet.CellFormats = new CellFormats(); 

            // index 0 : Default call format
            var cellFormat = new CellFormat();
            styleSheet.CellFormats.Append(cellFormat);

            foreach(var idColor in Enumerable.Range(1, colorQuantity))
            {
                cellFormat = new CellFormat();
                cellFormat.FillId = UInt32Value.FromUInt32((UInt32)(idColor + 2));
                styleSheet.CellFormats.Append(cellFormat);
            }

            styleSheet.CellFormats.Count = UInt32Value.FromUInt32((UInt32)styleSheet.CellFormats.ChildElements.Count);

            return styleSheet;
        }
    }

    public class Aera
    {
        public int X { get; }
        public int Y { get; }
        public int Length { get; }
        public int Height { get; }
        public Guid Id { get; }
        public int Deepth { get; }

        public List<Aera> SubAera { get; set; }

        public Aera(int x, int y, int length, int height, int deepth)
        {
            X = x;
            Y = y;
            Length = length;
            Height = height;
            Deepth = deepth;
            Id = Guid.NewGuid();
        }

        public Aera(Aera aera)
        {
            X = aera.X;
            Y = aera.Y;
            Length = aera.Length;
            Height = aera.Height;
            Id = aera.Id;
            Deepth = aera.Deepth;
        }

        public Aera(Aera aera, Guid id)
        {
            X = aera.X;
            Y = aera.Y;
            Length = aera.Length;
            Height = aera.Height;
            Id = id;
            Deepth = aera.Deepth;
        }
    }

    //https://stackoverflow.com/questions/309149
    public class ColorGenerator
    {
        public static int GetRGB(int index)
        {
            int[] p = GetPattern(index);
            return GetElement(p[0]) << 16 | GetElement(p[1]) << 8 | GetElement(p[2]);
        }

        private static int GetElement(int index)
        {
            int value = index - 1;
            int v = 0;
            for(int i = 0; i < 8; i++)
            {
                v = v | (value & 1);
                v <<= 1;
                value >>= 1;
            }

            v >>= 1;
            return v & 0xFF;
        }

        private static int[] GetPattern(int index)
        {
            int n = (int)Math.Cbrt(index);
            index -= (n * n * n);
            var p = Enumerable.Range(0, 3).Select(t => n).ToArray();

            if (index == 0)
                return p;

            index--;
            int v = index % 3;
            index = index / 3;
            if(index < n)
            {
                p[v] = index % n;
                return p;
            }

            index -= n;
            p[v] = index / n;
            p[++v % 3] = index % n;
            return p;
        }
    }
}
