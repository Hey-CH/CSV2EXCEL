using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.VisualBasic.FileIO;
using System.Xml.Linq;
using System.IO;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
namespace CSV2EXCEL {
    public class Cs {
        public const string _settingsxml = "settings.xml";

        public const string _settings = "settings";
        public const string _csvsetting = "csvsetting";
        public const string _headerrowcount = "headerrowcount";
        public const string _encoding = "encoding";
        public const string _delimiter = "delimiter";
        public const string _exchangesettings = "exchangesettings";
        public const string _exsetting = "exsetting";
        public const string _csvcol = "csvcol";
        public const string _excelcol = "excelcol";
        public const string _header = "header";
        public const string _type = "type";

        public static string LogPath {
            get {
                return Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "error.log");
            }
        }
        public static void Log(string msg) {
            try {
                File.AppendAllText(LogPath, "[" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "]" + msg+ "\r\n");
            } catch { }
        }
    }
    public static class Ex {
        public static List<string[]> GetCSVData(this string path,string delimiter = ",") {
            try {
                var ret = new List<string[]>();
                using(var fs = new TextFieldParser(path, Encoding.UTF8)) {
                    fs.SetDelimiters(delimiter);
                    string[] fields;
                    while((fields = fs.ReadFields()) != null) {
                        ret.Add(fields);
                    }
                }
                return ret;
            } catch {
                throw new ApplicationException("settings.xmlロード中にエラーが発生しました。");
            }
        }
        public static void CreateExcel(this string path, Dictionary<int, List<string>> data, Settings settings) {
            using(SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook)) {

                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                //SheetData（シートの内容）構築
                var sheetData = new SheetData();
                var map = settings.ExSettings
                                .OrderBy(es=>ExcelColCnt(es.ExcelColumn)).ToList();//CSVの○列をExcel×列に配置するという情報
                var maxrow = data.Max(kv => kv.Value.Count);

                uint rowIndex = 0;
                //設定ファイルでヘッダーが設定されてる場合
                if(settings.ExSettings.Any(es => !string.IsNullOrEmpty(es.Header))) {
                    rowIndex++;
                    Row row = new Row() { RowIndex = rowIndex };
                    map.ForEach(m => {
                        string cellReference = m.ExcelColumn + rowIndex;
                        Cell newCell = new Cell() { CellReference = cellReference ,DataType=CellValues.String , CellValue = new CellValue(m.Header) };
                        row.Append(newCell);
                    });
                    sheetData.Append(row);
                }
                for(int i = 0; i < maxrow; i++) {
                    rowIndex++;
                    Row row = new Row() { RowIndex = rowIndex };
                    map.ForEach(m => {
                        var text = data[m.CSVColumn][i];
                        string cellReference = m.ExcelColumn + rowIndex;
                        Cell newCell = new Cell() { CellReference = cellReference, DataType = GetType(m.FormatType), CellValue = new CellValue(text) };
                        row.Append(newCell);
                    });
                    sheetData.Append(row);
                }

                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                    AppendChild<Sheets>(new Sheets());

                Sheet sheet = new Sheet() {
                    Id = spreadsheetDocument.WorkbookPart.
                    GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1"
                };
                sheets.Append(sheet);

                workbookpart.Workbook.Save();

                spreadsheetDocument.Close();
            }
        }
        static int ExcelColCnt(string col) {
            var alpha = new[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            int ret = 0;
            for(int i = 0; i < col.Length; i++) {
                var a = alpha.IndexOf(col[i].ToString());
                ret += a * (i + 1) * alpha.Length;
            }
            return ret;
        }
        public static int IndexOf<T>(this IEnumerable<T> source, T value) {
            int index = 0;
            var comparer = StringComparer.OrdinalIgnoreCase;
            foreach(T item in source) {
                if(comparer.Equals(item, value)) return index;
                index++;
            }
            return -1;
        }
        static CellValues GetType(string type) {
            if(type == null || type.ToLower() == "string")
                return CellValues.String;
            else if(type.ToLower() == "boolean")
                return CellValues.Boolean;
            else if(type.ToLower() == "date")
                return CellValues.Date;
            else if(type.ToLower() == "error")
                return CellValues.Error;
            else if(type.ToLower() == "number")
                return CellValues.Number;
            else
                return CellValues.String;
        }
    }
    public class Settings {
        public int CSVHeader { get; set; }
        public string Encoding { get; set; }
        public string Delimiter { get; set; }
        public List<ExchangeSetting> ExSettings { get; set; }

        public static Settings Load(string path) {
            var doc = XDocument.Load(path, LoadOptions.PreserveWhitespace | LoadOptions.SetBaseUri);
            return new Settings(doc);
        }
        public Settings(XDocument settings) {
            var csvsetting = settings.Root.Element(Cs._csvsetting);
            CSVHeader = csvsetting.Attribute(Cs._headerrowcount) == null ? 0 : int.Parse(csvsetting.Attribute(Cs._headerrowcount).Value);
            Encoding = csvsetting.Attribute(Cs._encoding) == null ? "utf-8" : csvsetting.Attribute(Cs._encoding).Value;
            ExSettings = settings.Descendants(Cs._exsetting).Select(e => new ExchangeSetting(e)).ToList();
            Delimiter = csvsetting.Attribute(Cs._encoding) == null ? "," : csvsetting.Attribute(Cs._encoding).Value;
        }
    }
    public class ExchangeSetting {
        public int CSVColumn { get; set; }
        public string ExcelColumn { get; set; }
        public string FormatType { get; set; }
        public string Header { get; set; }
        public ExchangeSetting(XElement exsetting) {
            this.CSVColumn = int.Parse(exsetting.Attribute(Cs._csvcol).Value);
            this.ExcelColumn = exsetting.Attribute(Cs._excelcol).Value;
            this.Header = exsetting.Attribute(Cs._header)==null ? null : exsetting.Attribute(Cs._header).Value;
            this.FormatType = exsetting.Attribute(Cs._type) == null ? "string" : exsetting.Attribute(Cs._type).Value;
        }
    }
}
