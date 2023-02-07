using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
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
    public class AppSettings
    {
        public AppSettings()
        {
            CommonSettings common = new CommonSettings();
            TCSSettings tcsSettings = new TCSSettings();
            PolynomSettings polynomSettings = new PolynomSettings();
            ElementsActualisationSettings elementsSettings = new ElementsActualisationSettings();
            GroupsActualisationSettings groupsAcrualisationSettings = new GroupsActualisationSettings();
        }
        public static XLWorkbook AppSettingsWorkBook = new XLWorkbook("D:\\ascon_obmen\\kozlov_vi\\Полином\\Приложения\\TCS_Polynom_data_actualiser\\Настройки.xlsx");
        public static IXLWorksheet CommonSheet = AppSettingsWorkBook.Worksheets.Single(x => x.Name == "Общие");
        public static IXLWorksheet TCSSheet = AppSettingsWorkBook.Worksheets.Single(x => x.Name == "ТКС");
        public static IXLWorksheet PolynomSheet = AppSettingsWorkBook.Worksheets.Single(x => x.Name == "Полином");
        public static IXLWorksheet ElementActualsationSheet = AppSettingsWorkBook.Worksheets.Single(x => x.Name == "Актуализатор элементов");
        public static IXLWorksheet GroupActualisationSheet = AppSettingsWorkBook.Worksheets.Single(x => x.Name == "Актуализатор групп");
        public static IXLWorksheet ImportFileSheet = AppSettingsWorkBook.Worksheets.Single(x => x.Name == "Файл импорта");
        public class CommonSettings
        {
            public static List<string> Types { get; } = new Func<List<string>>(() =>
            {
                List<string> types = new List<string>();
                foreach (var startSell in CommonSheet.Search("вкл"))
                {
                    types.Add(startSell.WorksheetRow().Cell("A").Value.ToString());
                }
                return types;
            }).Invoke();
            public static Dictionary<string, string> TypesTranslit { get; }
                = new Func<Dictionary<string, string>>(() =>
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
                    CommonSheet.Search("<=").Value = "";
                    var range = CommonSheet.Range("G4:G20");
                    range.Search(value.ToString()).First().WorksheetRow().Cell("I").Value = "<=";
                    AppSettingsWorkBook.Save();
                }
            }
        }
        public class TCSSettings
        {
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
        public class PolynomSettings
        {
            //public static string Role = "Администраторы";
            //public static string Name = @"MZ\kozlov_vi";
            //public static string Password = "uw39sccvb";
        }
        public class ElementsActualisationSettings
        {
            public static Dictionary<string, List<string>> TcsToPolynomTypes { get; } = new Func<Dictionary<string, List<string>>>(() =>
            {
                Dictionary<string, List<string>> tcsToPolynomTypes = new Dictionary<string, List<string>>();
                foreach (var pair in CommonSettings.TypesTranslit)
                {
                    List<string> plynomGroups = new List<string>();
                    var cell = ElementActualsationSheet.Search(pair.Value).First().CellBelow();
                    int parce;
                    while (int.TryParse(cell.Value.ToString(), out parce))
                    {
                        plynomGroups.Add(cell.WorksheetRow().Cell("B").Value.ToString());
                        cell = cell.CellBelow();
                    }
                    tcsToPolynomTypes.Add(pair.Key, plynomGroups);
                }
                return tcsToPolynomTypes;
            }).Invoke();

            public static ActualisationStatus Status
            {
                get
                {
                    string status = ElementActualsationSheet.Search("Статус документа").First().CellRight().Value.ToString();
                    if (status == "Актуализирован")
                        return ActualisationStatus.Актуализирован;
                    else
                        return ActualisationStatus.Не_актуализирован;
                }
                set
                {
                    ElementActualsationSheet.Search("Статус документа").First().CellRight().Value = value.ToString();
                    AppSettingsWorkBook.Save();
                }
            }
            public static string FilePath { get; } = new Func<string>(() =>
            {
                string path = ElementActualsationSheet.Search("Путь до файла эксель \"Актуализация элементов\"").First().CellRight().Value.ToString();
                return path + "Актуализация элементов.xlsx";
            }).Invoke();

            public static string ArchivePath { get; } = new Func<string>(() =>
            {
                string path = ElementActualsationSheet.Search("Путь до архива файла эксель \"Актуализация элементов\"").First().CellRight().Value.ToString();
                return path + $"Актуализация элементов{(int)DateTime.Now.Ticks}.xlsx";
            }).Invoke();
        }
        public class GroupsActualisationSettings
        {
            public static ActualisationStatus Status
            {
                get
                {
                    string status = GroupActualisationSheet.Search("Статус документа").First().CellRight().Value.ToString();
                    if (status == "Актуализирован")
                        return ActualisationStatus.Актуализирован;
                    else
                        return ActualisationStatus.Не_актуализирован;
                }
                set
                {
                    GroupActualisationSheet.Search("Статус документа").First().CellRight().Value = value.ToString();
                    AppSettingsWorkBook.Save();
                }
            }
            public static string FilePath { get; } = new Func<string>(() =>
            {
                string path = GroupActualisationSheet.Search("Путь до файла эксель \"Актуализация групп\"").First().CellRight().Value.ToString();
                return path + $"Актуализация групп.xlsx";
            }).Invoke();
            public static string ArchivePath { get; } = new Func<string>(() =>
            {
                string path = GroupActualisationSheet.Search("Путь до архива файла эксель \"Актуализация групп\"").First().CellRight().Value.ToString();
                return path + $"Актуализация групп{(int)DateTime.Now.Ticks}.xlsx";
            }).Invoke();
        }
        public class ImportFileSettings
        {
            public static string FilePath { get; } = new Func<string>(() =>
            {
                var path = ImportFileSheet.Search("Путь до файла импорта").First().CellRight().Value.ToString();
                return path + $"Файл импорта.xlsx";
            }).Invoke();
            public static string ArchivePath { get; } = new Func<string>(() =>
            {
                string path = ImportFileSheet.Search("Путь до архива файла импорта").First().CellRight().Value.ToString();
                return path + $"Файл импорта{(int)DateTime.Now.Ticks}.xlsx";
            }).Invoke();
        }
    }
}