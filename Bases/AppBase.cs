using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.SqlServer.Server;
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
using Path = System.IO.Path;

namespace Polynom_Import_Manager
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
            ElementsSettings.Initialize();

            ElementsFile.Initialize();
            PropertiesFile.Initialize();

            _ImportFile.Initialize();
        }
        public static XLWorkbook SettingsWorkBook { get; set; } = new XLWorkbook(Program.settingsWorkbookPath);
        public static void ReconnectSettingsWorkBook()
        {
            if (SettingsWorkBook != null)
                SettingsWorkBook.Dispose();
            if (!File.Exists(Program.settingsWorkbookPath))
            {
                throw new Exception("Систесмый файл \"Настройки\" не найден! Его нельзя удалять. Восстановите.");
            }
            SettingsWorkBook = new XLWorkbook(Program.settingsWorkbookPath);
        }
        public static class CommonSettings
        {
            public static void Initialize() { }
            public static IXLWorksheet AlgorythmSheet => SettingsWorkBook.Worksheet("Алгоритм");
            public static IXLWorksheet SystemsSheet => SettingsWorkBook.Worksheet("Системное");
            public static int CursorPosition
            {
                get
                {
                    if(Algorithm.targetModule == Algorithm.RunModule.Elements)
                        return Convert.ToInt32(AlgorythmSheet.Range("A3:C7").Search("<=").First().WorksheetRow().Cell("A").Value);
                    if (Algorithm.targetModule == Algorithm.RunModule.Propertyes)
                        return Convert.ToInt32(AlgorythmSheet.Range("A9:C14").Search("<=").First().WorksheetRow().Cell("A").Value);
                    /*if (Algorithm.targetModule == Algorithm.RunModule.Concepts)
                        return Convert.ToInt32(AlgorythmSheet.Range("B12:D20").Search("<=").First().WorksheetRow().Cell("B").Value.ToString());*/
                    return 0;
                }
                set
                {
                    IXLRange range = null;
                    if (Algorithm.targetModule == Algorithm.RunModule.Elements) 
                        range = AlgorythmSheet.Range("A3:C7");
                    if (Algorithm.targetModule == Algorithm.RunModule.Propertyes) 
                        range = AlgorythmSheet.Range("A9:C14");
                    /*if (Algorithm.targetModule == Algorithm.RunModule.Concepts) range = AlgorythmSheet.Range("B3:D10");*/

                    range.Search("<=").Value = "";
                    range.Search(value.ToString()).Single(x => Convert.ToInt32(x.Value) == value).WorksheetRow().Cell("C").Value = "<="; // костыль. Берет 10 и 1, если ищешь 1.
                    SettingsWorkBook.Save();
                }
            }
            public static string ConnectionString { get; } = new Func<string>(() =>
            {
                return SystemsSheet.Search("Строка подключения").First().CellRight().Value.ToString();
            }).Invoke();
        }
        public static class ElementsSettings
        {
            public static void Initialize()
            {
                List<IXLCell> vklCells = Sheet.Search("Вкл").ToList();
                vklCells.RemoveAt(0);
                if(vklCells.Count == 0)
                    throw new Exception("Зайдайте хотя бы одной переносимной группе статус \"Вкл\".");

                Types = vklCells.Select(x => x.WorksheetRow().Cell("A").Value.ToString()).ToList();
                TypesTranslit = new Func<List<(string, string)>>(() =>
                {
                    List<(string type, string translit)> typesTranslit = new List<(string, string)>();
                    foreach (var vklCell in vklCells)
                    {
                        var row = vklCell.WorksheetRow();
                        typesTranslit.Add((row.Cell("A").Value.ToString(), row.Cell("B").Value.ToString()));
                    }
                    return typesTranslit;
                }).Invoke();
                TypesCode = new Func<List<(string, string)>>(() =>
                {
                    List<(string type, string typeCode)> typesCode = new List<(string, string)>();
                    foreach (var vklCell in vklCells)
                    {
                        var row = vklCell.WorksheetRow();
                        typesCode.Add((row.Cell("A").Value.ToString(), row.Cell("C").Value.ToString()));
                    }
                    return typesCode;
                }).Invoke();
                TypePolynomPaths = new Func<Dictionary<string, List<string>>>(() =>
                {
                    Dictionary<string, List<string>> typePolynomPaths = new Dictionary<string, List<string>>();
                    foreach (var vklCell in vklCells)
                    {
                        var row = vklCell.WorksheetRow();

                        string typeName = row.Cell("A").Value.ToString();
                        List<string> paths = new List<string>();

                        StringReader reader = new StringReader(row.Cell("D").Value.ToString());
                        string path;
                        while ((path = reader.ReadLine()) != null)
                        {
                            paths.Add(path);
                        }

                        typePolynomPaths.Add(typeName, paths);
                    }
                    return typePolynomPaths;
                }).Invoke(); // Проверить как работает с пустой строкой. Должен быть пустой лист.
                TypesSQLQueries = new Func<Dictionary<string, string>>(() =>
                {
                    Dictionary<string, string> tcsTypesSQLQueryes = new Dictionary<string, string>();

                    foreach (string type in Types)
                    {
                        var subdir = $"Запросы SQL\\{type}.sql";
                        var fullPath = Path.Combine(Program.appDir, subdir);
                        tcsTypesSQLQueryes.Add(type, File.ReadAllText(fullPath));
                    }
                    return tcsTypesSQLQueryes;
                }).Invoke();
            }
            public static IXLWorksheet Sheet => SettingsWorkBook.Worksheet("Актуализация элементов");
            public static List<string> Types { get; set; }
            public static List<(string type, string translit)> TypesTranslit { get; set; }
            public static List<(string type, string typeCode)> TypesCode { get; set; }
            public static Dictionary<string, List<string>> TypePolynomPaths { get; set; }
            public static Dictionary<string, string> TypesSQLQueries { get; set; }
        }
        public static class ElementsFile
        {
            public static void Initialize() { }
            public static XLWorkbook WorkBook { get; set; }
            public static IXLWorksheet StatusSheet => WorkBook.Worksheet("Статус");
            public static ActualisationStatus StatusElements
            {
                get
                {
                    ReconnectWorkBook(false);
                    string status = StatusSheet.Search("Статус элементов").First().CellBelow().Value.ToString();
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
                    WorkBook.Save();
                }
            }
            public static ActualisationStatus StatusGroups
            {
                get
                {
                    ReconnectWorkBook(false);
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
                    WorkBook.Save();
                }
            }
            public static string FilePath { get; set; } = Path.Combine(Program.appDir, "Актуализация элементов.xlsx");
            public static string ArchivePath { get; set; } = Path.Combine(Program.appDir, $"Архив актуализация элементов\\Актуализация элементов-{(int)DateTime.Now.Ticks}.xlsx");
            public static void ReconnectWorkBook(bool createIfNotExists)
            {
                if(WorkBook != null)
                    WorkBook.Dispose();
                bool exists = File.Exists(FilePath);
                if (!exists)
                {
                    if (createIfNotExists)
                    {
                        WorkBook = new XLWorkbook();
                        return;
                    }
                    else
                    {
                        throw new Exception("Файл \"Актуализация элементов\" не найден.");
                    }
                }
                WorkBook = new XLWorkbook(FilePath);
            }
        }
        public static class PropertiesFile
        {
            public static void Initialize() { }
            public static Lazy<XLWorkbook> WorkBook { get; set; } = new Lazy<XLWorkbook>(() =>
            {
                if (!File.Exists(FilePath))
                    throw new Exception("Файл \"Актуализация свойств\" не найден.");
                return new XLWorkbook(FilePath);
            });
            public static IXLWorksheet PropertiesSheet => WorkBook.Value.Worksheets.Worksheet("Свойства");
            public static IXLWorksheet StatusSheet => WorkBook.Value.Worksheet("Статус");
            public static ActualisationStatus StatusProperties
            {
                get
                {
                    ReconnectWorkBook();
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
                    ReconnectWorkBook();
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
            public static string PropertyesSQLQuery { get; set; } = File.ReadAllText(Path.Combine(Program.appDir, $"Запросы SQL\\Свойства.sql"));
            public static string FilePath { get; set; } = Path.Combine(Program.appDir, "Актуализация свойств.xlsx");
            public static string ArchivePath { get; set; } = Path.Combine(Program.appDir, $"Архив актуализация свойств\\Актуализация свойств-{(int)DateTime.Now.Ticks}.xlsx");
            public static void ReconnectWorkBook()
            {
                if (!File.Exists(FilePath))
                    throw new Exception("Файл \"Актуализация свойств\" не найден.");

                WorkBook = new Lazy<XLWorkbook>(() =>
                {
                    return new XLWorkbook(FilePath);
                });
            }
        }
        public static class _ImportFile
        {
            public static void Initialize() { }
            public static XLWorkbook WorkBook { get; set; }
            public static string FilePath { get; set; } = Path.Combine(Program.appDir, "Файл импорта.xlsx");
            public static string ArchivePath { get; set; } = Path.Combine(Program.appDir, $"Архив файл импорта\\Файл импорта-{(int)DateTime.Now.Ticks}.xlsx");
        }
    }
}