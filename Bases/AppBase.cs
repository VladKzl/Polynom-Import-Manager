using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
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
        public static void Initialize() 
        {
            CommonSettings.Initialize();
            ElementsFileSettings.Initialize();
            GroupsFileSettings.Initialize();
            PropertiesSettings.Initialize();
            ImportFileSettings.Initialize();
        }
        public static string AppDir { get; set; } = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location).Replace("\\bin\\Debug", "");
        public static XLWorkbook SettingsWorkBook { get; set; } = new XLWorkbook(Program.settingsWorkbookPath);
        public static Lazy<XLWorkbook> ElementsActualisationWorkBook { get; set; } = new Lazy<XLWorkbook>(() => 
        {
            if (!File.Exists(ElementsFileSettings.FilePath))
                throw new Exception("Файл \"Актуализация элементов\" не найден.");
            return new XLWorkbook(ElementsFileSettings.FilePath);
        });
        public static Lazy<XLWorkbook> GroupsActualisationWorkBook { get; set; } = new Lazy<XLWorkbook>(() =>
        {
            if (!File.Exists(GroupsFileSettings.FilePath))
                throw new Exception("Файл \"Актуализация групп\" не найден.");
            return new XLWorkbook(GroupsFileSettings.FilePath);
        });
        
        public static class CommonSettings
        {
            public static void Initialize() { }
            public static IXLWorksheet CommonSheet { get; set; } = SettingsWorkBook.Worksheets.Single(x => x.Name == "Общие");
            public static int CursorPosition
            {
                get
                {
                    if(Algorithm.targetModule == Algorithm.RunModule.Elements)
                        return Convert.ToInt32(CommonSheet.Range("B3:D10").Search("<=").First().WorksheetRow().Cell("B").Value.ToString());
                    if (Algorithm.targetModule == Algorithm.RunModule.Propertyes)
                        return Convert.ToInt32(CommonSheet.Range("B12:D20").Search("<=").First().WorksheetRow().Cell("B").Value.ToString());
                    /*if (Algorithm.targetModule == Algorithm.RunModule.Concepts)
                        return Convert.ToInt32(CommonSheet.Range("B12:D20").Search("<=").First().WorksheetRow().Cell("B").Value.ToString());*/
                    return 0;
                }
                set
                {
                    IXLRange range = null;
                    if (Algorithm.targetModule == Algorithm.RunModule.Elements) range = CommonSheet.Range("B3:D10");
                    if (Algorithm.targetModule == Algorithm.RunModule.Propertyes) range = CommonSheet.Range("B12:D20");
                    /*if (Algorithm.targetModule == Algorithm.RunModule.Concepts) range = CommonSheet.Range("B3:D10");*/

                    range.Search("<=").Value = "";
                    range.Search(value.ToString()).Single(x => Convert.ToInt32(x.Value) == value).WorksheetRow().Cell("D").Value = "<="; // костыль. Берет 10 и 1, если ищешь 1.
                    SettingsWorkBook.Save();
                }
            }
            public static string TCSConnectionString { get; } = new Func<string>(() =>
            {
                return CommonSheet.Search("Строка подключения").First().CellRight().Value.ToString();
            }).Invoke();
        }
        public static class ElementsFileSettings
        {
            public static void Initialize()
            {
                Types = new Func<List<string>>(() =>
                {
                    List<string> types = new List<string>();
                    foreach (var startSell in ElementActualsationSheet.Search("вкл"))
                    {
                        types.Add(startSell.CellLeft(3).Value.ToString());
                    }
                    return types;
                }).Invoke();
                TypesTranslit = new Func<Dictionary<string, string>>(() =>
                {
                    Dictionary<string, string> typesTranslit = new Dictionary<string, string>();
                    foreach (var startSell in ElementActualsationSheet.Search("вкл"))
                    {
                        var name = startSell.CellLeft(3).Value.ToString();
                        var translit = startSell.CellLeft(2).Value.ToString();
                        typesTranslit.Add(name, translit);
                    }
                    return typesTranslit;
                }).Invoke();
                TypesDef = new Func<Dictionary<string, string>>(() =>
                {
                    Dictionary<string, string> typesDef = new Dictionary<string, string>();
                    foreach (var startSell in ElementActualsationSheet.Search("вкл"))
                    {
                        var name = startSell.CellLeft(3).Value.ToString();
                        var def = startSell.CellLeft(1).Value.ToString();
                        typesDef.Add(name, def);
                    }
                    return typesDef;
                }).Invoke();
                FilePath = Path.Combine(AppDir, "Актуализация элементов.xlsx");
                ArchivePath = Path.Combine(AppDir, $"Архив актуализация элементов\\Актуализация элементов-{(int)DateTime.Now.Ticks}.xlsx");
                TcsByPolynomTypes = new Func<Dictionary<string, List<string>>>(() =>
                {
                    Dictionary<string, List<string>> tcsToPolynomTypes = new Dictionary<string, List<string>>();
                    foreach (var pair in Types)
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
                TcsTypesSQLQueries = new Func<Dictionary<string, string>>(() =>
                {
                    Dictionary<string, string> tcsTypesSQLQueryes = new Dictionary<string, string>();
                    
                    foreach (string tcsType in Types)
                    {
                        var subdir = $"Запросы SQL\\{tcsType}.sql";
                        var fullPath = Path.Combine(AppDir, subdir);
                        tcsTypesSQLQueryes.Add(tcsType, File.ReadAllText(fullPath));
                    }
                    return tcsTypesSQLQueryes;
                }).Invoke();
            }
            public static IXLWorksheet ElementActualsationSheet => SettingsWorkBook.Worksheets.Single(x => x.Name == "Актуализатор элементов");
            public static Dictionary<string, List<string>> TcsByPolynomTypes { get; set; }
            public static Dictionary<string, string> TcsTypesSQLQueries { get; set; }
            public static List<string> Types { get; set; }
            public static Dictionary<string, string> TypesTranslit { get; set; }
            public static Dictionary<string, string> TypesDef { get; set; }
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
            public static string FilePath { get; set; }
            public static string ArchivePath { get; set; }
        }
        public static class GroupsFileSettings
        {
            public static void Initialize() { }
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
            public static string FilePath { get; set; } = Path.Combine(AppDir, "Aктуализация групп.xlsx");
            public static string ArchivePath { get; set; } = Path.Combine(AppDir, $"Архив актуализация групп\\Актуализация групп{(int)DateTime.Now.Ticks}.xlsx");
        }
        public static class PropertiesSettings
        {
            public static void Initialize() { }
            public static Lazy<XLWorkbook> WorkBook { get; set; } = new Lazy<XLWorkbook>(() =>
            {
                if (!File.Exists(FilePath))
                    throw new Exception("Файл \"Актуализация свойств\" не найден.");
                return new XLWorkbook(FilePath);
            });
            public static IXLWorksheet PropertyesSheet => WorkBook.Value.Worksheets.Single(x => x.Name == "Свойства");
            public static IXLWorksheet StatusSheet => WorkBook.Value.Worksheets.Single(x => x.Name == "Статус");
            public static string TcsPropertyesSQLQuery { get; set; } = File.ReadAllText(Path.Combine(AppDir, $"Запросы SQL\\Свойства.sql"));
            public static ActualisationStatus StatusProperties
            {
                get
                {
                    UpdateWorkBook();
                    string status = StatusSheet.Search("Статус свойств").First().CellBelow().Value.ToString();
                    if (status == "Актуализирован")
                        return ActualisationStatus.Актуализирован;
                    else
                        return ActualisationStatus.Не_актуализирован;
                }
                set
                {
                    if (value.ToString() == "Актуализирован")
                        StatusSheet.Search("Статус свойств").First().CellBelow().Value = "Актуализирован";
                    else
                        StatusSheet.Search("Статус свойств").First().CellBelow().Value = "Не актуализирован";
                    WorkBook.Value.Save();
                }
            }
            public static ActualisationStatus StatusGroups
            {
                get
                {
                    UpdateWorkBook();
                    string status = StatusSheet.Search("Статус групп").First().CellBelow().Value.ToString();
                    if (status == "Актуализирован")
                        return ActualisationStatus.Актуализирован;
                    else
                        return ActualisationStatus.Не_актуализирован;
                }
                set
                {
                    if (value.ToString() == "Актуализирован")
                        StatusSheet.Search("Статус групп").First().CellBelow().Value = "Актуализирован";
                    else
                        StatusSheet.Search("Статус групп").First().CellBelow().Value = "Не актуализирован";
                    WorkBook.Value.Save();
                }
            }
            public static string FilePath { get; set; } = Path.Combine(AppDir, "Актуализация свойств.xlsx");
            public static string ArchivePath { get; set; } = Path.Combine(AppDir, $"Архив актуализация свойств\\Актуализация свойств-{(int)DateTime.Now.Ticks}.xlsx");
            public static void UpdateWorkBook()
            {
                if (!File.Exists(FilePath))
                    throw new Exception("Файл \"Актуализация свойств\" не найден.");

                WorkBook = new Lazy<XLWorkbook>(() =>
                {
                    return new XLWorkbook(FilePath);
                });
            }
        }
        public static class ImportFileSettings
        {
            public static void Initialize() { }
            public static string FilePath { get; set; } = Path.Combine(AppDir, "Файл импорта.xlsx");
            public static string ArchivePath { get; set; } = Path.Combine(AppDir, $"Архив файл импорта\\Файл импорта-{(int)DateTime.Now.Ticks}.xlsx");
        }
        public static void UpdateSettingsWorkBook()
        {
            SettingsWorkBook = new XLWorkbook(Program.settingsWorkbookPath);
        }
    }
}