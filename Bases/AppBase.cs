using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Security.RightsManagement;
using System.Text;
using System.Threading.Tasks;

namespace TCS_Polynom_data_actualiser
{
    public enum ActualisationStatus
    {
        Актуализирован,
        Не_актуализирован
    }
    public static class AppBase
    {
        public static void Initialize() { }
        public static XLWorkbook SettingsWorkBook { get; set; } = new XLWorkbook(Program.settingsWorkbookPath);
        public static Lazy<XLWorkbook> ElementsActualisationWorkBook { get; set; } = new Lazy<XLWorkbook>(() => 
        {
            if (!File.Exists(ElementsActualisationSettings.FilePath))
                throw new Exception("Файл \"Актуализация элементов\" не найден.");
            return new XLWorkbook(ElementsActualisationSettings.FilePath);
        });
        public static Lazy<XLWorkbook> GroupsActualisationWorkBook { get; set; } = new Lazy<XLWorkbook>(() =>
        {
            if (!File.Exists(GroupsActualisationSettings.FilePath))
                throw new Exception("Файл \"Актуализация групп\" не найден.");
            return new XLWorkbook(GroupsActualisationSettings.FilePath);
        });
        public static class CommonSettings
        {
            public static IXLWorksheet CommonSheet { get; set; } = SettingsWorkBook.Worksheets.Single(x => x.Name == "Общие");
            public static List<string> Types { get; } = new Func<List<string>>(() =>
            {
                List<string> types = new List<string>();
                foreach (var startSell in CommonSheet.Search("вкл"))
                {
                    types.Add(startSell.WorksheetRow().Cell("A").Value.ToString());
                }
                return types;
            }).Invoke();
            public static Dictionary<string, string> TypesTranslit { get; } = new Func<Dictionary<string, string>>(() =>
            {
                Dictionary<string, string> typesTranslit = new Dictionary<string, string>();
                foreach (var startSell in CommonSheet.Search("вкл"))
                {
                    var name = startSell.WorksheetRow().Cell("A").Value.ToString();
                    var translit = startSell.WorksheetRow().Cell("B").Value.ToString();
                    typesTranslit.Add(name, translit);
                }
                return typesTranslit;
            }).Invoke();
            public static Dictionary<string, string> TypesDef { get; } = new Func<Dictionary<string, string>>(() =>
            {
                Dictionary<string, string> typesDef = new Dictionary<string, string>();
                foreach (var startSell in CommonSheet.Search("вкл"))
                {
                    var name = startSell.WorksheetRow().Cell("A").Value.ToString();
                    var def = startSell.WorksheetRow().Cell("C").Value.ToString();
                    typesDef.Add(name, def);
                }
                return typesDef;
            }).Invoke();
            public static int CursorPosition
            {
                get
                {
                    return Convert.ToInt32(CommonSheet.Search("<=").First().WorksheetRow().Cell("G").Value.ToString());
                }
                set
                {
                    var range = CommonSheet.Range("G1:I100");
                    CommonSheet.Search("<=").Value = "";
                    range.Search(value.ToString()).Single(x => Convert.ToInt32(x.Value) == value).WorksheetRow().Cell("I").Value = "<="; // костыль. Берет 10 и 1, если ищешь 1.
                    SettingsWorkBook.Save();
                }
            }
        }
        public static class TCSSettings
        {
            public static IXLWorksheet TCSSheet { get; set; } = SettingsWorkBook.Worksheets.Single(x => x.Name == "ТКС");
            public static string TCSConnectionString { get; } = new Func<string>(() =>
            {
                return TCSSheet.Search("Строка подключения").First().WorksheetRow().Cell("B").Value.ToString();
            }).Invoke();
            public static Dictionary<string, string> TcsQueryFilePaths { get; } = new Func<Dictionary<string, string>>(() =>
            {
                Dictionary<string, string> tcsQueryFilePaths = new Dictionary<string, string>();
                foreach (var pair in CommonSettings.TypesTranslit)
                {
                    var cell = TCSSheet.Search(pair.Value).First();
                    string path = cell.WorksheetRow().Cell("B").Value.ToString();
                    tcsQueryFilePaths.Add(pair.Key, path);
                }
                return tcsQueryFilePaths;
            }).Invoke();
        }
        public static class PolynomSettings
        {
            public static IXLWorksheet PolynomSheet { get; set; } = SettingsWorkBook.Worksheets.Single(x => x.Name == "Полином");
            //public static string Role = "Администраторы";
            //public static string Name = @"MZ\kozlov_vi";
            //public static string Password = "uw39sccvb";
        }
        public static class ElementsActualisationSettings
        {
            public static IXLWorksheet ElementActualsationSheet => SettingsWorkBook.Worksheets.Single(x => x.Name == "Актуализатор элементов");
            public static Dictionary<string, List<string>> TcsByPolynomTypes { get; } = new Func<Dictionary<string, List<string>>>(() =>
            {
                Dictionary<string, List<string>> tcsToPolynomTypes = new Dictionary<string, List<string>>();
                foreach (var pair in CommonSettings.Types)
                {
                    List<string> plynomGroups = new List<string>();
                    var cell = ElementActualsationSheet.Search(pair).First().CellBelow();
                    int parce;
                    while (int.TryParse(cell.Value.ToString(), out parce))
                    {
                        plynomGroups.Add(cell.WorksheetRow().Cell("B").Value.ToString());
                        cell = cell.CellBelow();
                    }
                    tcsToPolynomTypes.Add(pair, plynomGroups);
                }
                return tcsToPolynomTypes;
            }).Invoke();
            public static ActualisationStatus Status
            {
                get
                {
                    UpdateSettingsWorkBook();
                    string status = ElementActualsationSheet.Search("Статус документа").First().CellBelow().Value.ToString();
                    if (status == "Актуализирован")
                        return ActualisationStatus.Актуализирован;
                    else
                        return ActualisationStatus.Не_актуализирован;
                }
                set
                {
                    if(value.ToString() == "Актуализирован")
                        ElementActualsationSheet.Search("Статус документа").First().CellBelow().Value = "Актуализирован";
                    else
                        ElementActualsationSheet.Search("Статус документа").First().CellBelow().Value = "Не актуализирован";
                    SettingsWorkBook.Save();
                }
            }
            public static string FilePath { get; } = new Func<string>(() =>
            {
                string path = ElementActualsationSheet.Search("Путь до экселя").First().CellRight().Value.ToString();
                return path + "Актуализация элементов.xlsx";
            }).Invoke();
            public static string ArchivePath { get; } = new Func<string>(() =>
            {
                string path = ElementActualsationSheet.Search("Путь до архива").First().CellRight().Value.ToString();
                return path + $"Актуализация элементов-{(int)DateTime.Now.Ticks}.xlsx";
            }).Invoke();
        }
        public static class GroupsActualisationSettings
        {
            public static IXLWorksheet GroupActualisationSheet => SettingsWorkBook.Worksheets.Single(x => x.Name == "Актуализатор групп");
            public static ActualisationStatus Status
            {
                get
                {
                    UpdateSettingsWorkBook();
                    string status = GroupActualisationSheet.Search("Статус документа").First().CellBelow().Value.ToString();
                    if (status == "Актуализирован")
                        return ActualisationStatus.Актуализирован;
                    else
                        return ActualisationStatus.Не_актуализирован;
                }
                set
                {
                    if (value.ToString() == "Актуализирован")
                        GroupActualisationSheet.Search("Статус документа").First().CellBelow().Value = "Актуализирован";
                    else
                        GroupActualisationSheet.Search("Статус документа").First().CellBelow().Value = "Не актуализирован";
                    SettingsWorkBook.Save();
                }
            }
            public static string FilePath { get; } = new Func<string>(() =>
            {
                string path = GroupActualisationSheet.Search("Путь до экселя").First().CellRight().Value.ToString();
                return path + $"Актуализация групп.xlsx";
            }).Invoke();
            public static string ArchivePath { get; } = new Func<string>(() =>
            {
                string path = GroupActualisationSheet.Search("Путь до архива").First().CellRight().Value.ToString();
                return path + $"Актуализация групп{(int)DateTime.Now.Ticks}.xlsx";
            }).Invoke();
        }
        public static class ImportFileSettings
        {
            public static IXLWorksheet ImportFileSheet { get; set; } = SettingsWorkBook.Worksheets.Single(x => x.Name == "Файл импорта");
            public static string FilePath { get; } = new Func<string>(() =>
            {
                var path = ImportFileSheet.Search("Путь до файла импорта").First().CellRight().Value.ToString();
                return path + $"Файл импорта.xlsx";
            }).Invoke();
            public static string ArchivePath { get; } = new Func<string>(() =>
            {
                string path = ImportFileSheet.Search("Путь до архива").First().CellRight().Value.ToString();
                return path + $"Файл импорта{(int)DateTime.Now.Ticks}.xlsx";
            }).Invoke();
        }
        public static void UpdateSettingsWorkBook()
        {
            SettingsWorkBook = new XLWorkbook(Program.settingsWorkbookPath);
        }
    }
}