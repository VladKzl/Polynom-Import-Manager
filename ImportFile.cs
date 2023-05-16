using Ascon.Polynom.Api;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace TCS_Polynom_data_actualiser
{
    public class ImportFile
    {
        public enum Sheets
        {
            Propertyes,
            Сoncepts
        }
        public static void AddGroupsAndElements()
        {
            XLWorkbook ImportFileWorkBook = new XLWorkbook();
            CreateDoc();

            foreach (string type in AppBase.ElementsFileSettings.Types)
            {
                var workSheet = ImportFileWorkBook.Worksheet(type);

                var elementAndGroupPairs = PolynomElementsCreation.CreatedElementAndGroupByTcsType[type];
                for (int i = 0; i < elementAndGroupPairs.Count(); i++)
                {
                    string referenceName = elementAndGroupPairs[i].group.ParentCatalog.Reference.Name;
                    string catalogName = elementAndGroupPairs[i].group.ParentCatalog.Name;
                    List<string> groupsNames = GetGroups();
                    string elementName = elementAndGroupPairs[i].element.Name;
                    FillRowData();

                    List<string> GetGroups()
                    {
                        List<string> groups = new List<string>();
                        IGroup groupToAdd = elementAndGroupPairs[i].group;
                        groups.Add(groupToAdd.Name);
                        while (groupToAdd.ParentGroup != null)
                        {
                            groups.Add(groupToAdd.ParentGroup.Name);
                            groupToAdd = groupToAdd.ParentGroup;
                        }
                        groups.Reverse();
                        return groups;
                    }
                    void FillRowData()
                    {
                        int elementNameGorisontalSellPosition = 0;
                        var nextRow = workSheet.Row(i+2);

                        nextRow.Cell("C").Value = referenceName;
                        nextRow.Cell("D").Value = catalogName;
                        FillGroupsData();
                        nextRow.Cell(elementNameGorisontalSellPosition).Value = elementName;

                        void FillGroupsData()
                        {
                            for (int n = 0; n < groupsNames.Count(); n++)
                            {
                                int gorisontalCellPosition = n + 5;
                                nextRow.Cell(gorisontalCellPosition).Value = groupsNames[n];
                                elementNameGorisontalSellPosition = gorisontalCellPosition + 1;
                            }
                        }
                    }
                }
            }
            FormattingStructure();
            ImportFileWorkBook.Save();


            void CreateDoc()
            {
                if (File.Exists(AppBase.ImportFileSettings.FilePath))
                {
                    // Если файл слуществует и уже были добалвены страницы ранее, значит создаем новый а старый в архив.
                    ImportFileWorkBook = new XLWorkbook(AppBase.ImportFileSettings.FilePath);
                    if (AreWorksheetsExists())
                    {
                        File.Move(AppBase.ImportFileSettings.FilePath, AppBase.ImportFileSettings.ArchivePath);
                        ImportFileWorkBook = new XLWorkbook();
                        CreateSheets();
                        ImportFileWorkBook.SaveAs(AppBase.ImportFileSettings.FilePath);
                        return;
                    }
                    // Если файл существует и станицы не были добавлены, то работаем в том же файле.
                    CreateSheets();
                    ImportFileWorkBook.Save();
                }
                ImportFileWorkBook = new XLWorkbook();
                CreateSheets();
                ImportFileWorkBook.SaveAs(AppBase.ImportFileSettings.FilePath);

                void CreateSheets()
                {
                    foreach (string type in AppBase.ElementsFileSettings.Types)
                    {
                        ImportFileWorkBook.AddWorksheet(type);
                    }
                }
                bool AreWorksheetsExists()
                {
                    IXLWorksheet sheet;
                    if (ImportFileWorkBook.TryGetWorksheet(AppBase.ElementsFileSettings.Types.First(), out sheet))
                        return true;
                    return false;
                }
            }
            void FormattingStructure()
            {
                foreach (string type in AppBase.ElementsFileSettings.Types)
                {
                    var workSheet = ImportFileWorkBook.Worksheet(type);
                    var rows = workSheet.RangeUsed().Rows();
                    foreach (var row in rows)
                    {
                        string lastCellValue = row.LastCellUsed().Value.ToString();
                        row.LastCellUsed().Value = "";
                        row.LastCell().Value = lastCellValue;
                    }

                    var columnsUsed = workSheet.ColumnsUsed().ToList();
                    columnsUsed[0].Cell(1).Value = "REFERENCE";
                    columnsUsed[1].Cell(1).Value = "CATALOGS";

                    for (int i = 2; i < columnsUsed.Count; i++)
                    {
                        columnsUsed[i].Cell(1).Value = "GROUP";
                    }
                    columnsUsed.Last().Cell(1).Value = "NAME";
                }
            }
        }
        public static void AddProperties()
        {
            XLWorkbook ImportFileWorkBook = new XLWorkbook();
            CreateDoc();
            var importSheet = ImportFileWorkBook.Worksheet("PROPERTIES");

            var propRows = AppBase.PropertiesSettings.PropertyesSheet.RowsUsed().ToList();
            propRows.RemoveRange(0, 2);
            var rowsCount = propRows.First().Cell("B").WorksheetColumn().CellsUsed().Count();

            for (int i = 0; i < rowsCount; i++)
            {
                string propName = propRows[i].Cell("B").Value.ToString(); //имя
                string propPath = propRows[i].Cell("C").Value.ToString(); //путь
                string realPath = GetRealPath();

                try
                {
                    string code = TCSBase.Propertyes.GetColumnValueFromPropRows<string>(propName, propPath, TCSBase.Propertyes.RowColumnsForSearch.CODE); //тут ошибка
                    string typeName = TCSBase.Propertyes.GetColumnValueFromPropRows<string>(propName, propPath, TCSBase.Propertyes.RowColumnsForSearch.TYPE);
                    string masureetity = TCSBase.Propertyes.GetColumnValueFromPropRows<string>(propName, propPath, TCSBase.Propertyes.RowColumnsForSearch.MEASUREENTITY);
                    string lov = TCSBase.Propertyes.GetColumnValueFromPropRows<string>(propName, propPath, TCSBase.Propertyes.RowColumnsForSearch.LOV);
                    string description = TCSBase.Propertyes.GetColumnValueFromPropRows<string>(propName, propPath, TCSBase.Propertyes.RowColumnsForSearch.DESCRIPTION);
                    importSheet.Cell(i + 1, "A").Value = propName;
                    importSheet.Cell(i + 1, "B").Value = code;
                    importSheet.Cell(i + 1, "C").Value = typeName;
                    importSheet.Cell(i + 1, "D").Value = masureetity;
                    importSheet.Cell(i + 1, "E").Value = lov != null ? FormatLov(lov) : "";
                    importSheet.Cell(i + 1, "F").Value = description;
                    importSheet.Cell(i + 1, "G").Value = realPath;
                }
                catch
                {
                    Console.WriteLine($"Свойство в экселе {propName} с группой {propPath} не найдено в sql запросе на свойства. Скорее всего эксель изменил значение из за форматирования.");
                    continue;
                }

                string GetRealPath()
                {
                    string polynomPaths = propRows[i].Cell("D").Value.ToString();
                    int? polynomIndex = null;
                    if (propRows[i].Cell("E").Value.ToString() != string.Empty)
                        polynomIndex = Convert.ToInt32(propRows[i].Cell("E").Value);

                    if (polynomPaths != string.Empty && polynomIndex != null)
                    {
                        var separatedPaths1 = Regex.Replace(polynomPaths, "[0-100]", ";");
                        var separatedPaths2 = separatedPaths1.Remove(1,3);
                        List<string> splitedPaths = separatedPaths2.Split(';').ToList();
                        splitedPaths.RemoveAt(0);
                        return splitedPaths[polynomIndex.Value];
                    }
                    return propPath;
                }       
            }
            ImportFileWorkBook.Save();

            string FormatLov(string lov)
            {
                StringBuilder stringBuilder = new StringBuilder();

                var splitString = lov.Split(';').ToList();
                for (int i = 0; i < splitString.Count; i++)
                {
                    if (i < splitString.Count - 1)
                        stringBuilder.AppendLine(splitString[i]);
                    if (i == splitString.Count - 1)
                        stringBuilder.Append(splitString[i]);
                }
                return stringBuilder.ToString();
            }
            void CreateDoc()
            {
                if (File.Exists(AppBase.ImportFileSettings.FilePath))
                {
                    // Если файл слуществует и уже были добалвены страницы ранее, значит создаем новый а старый в архив.
                    ImportFileWorkBook = new XLWorkbook(AppBase.ImportFileSettings.FilePath);
                    if (AreWorksheetsExists())
                    {
                        File.Move(AppBase.ImportFileSettings.FilePath, AppBase.ImportFileSettings.ArchivePath);
                        ImportFileWorkBook = new XLWorkbook();
                        CreateSheet();
                        ImportFileWorkBook.SaveAs(AppBase.ImportFileSettings.FilePath);
                        return;
                    }
                    // Если файл существует и станицы не были добавлены, то работаем в том же файле.
                    CreateSheet();
                    ImportFileWorkBook.Save();
                    return;
                }
                ImportFileWorkBook = new XLWorkbook();
                CreateSheet();
                ImportFileWorkBook.SaveAs(AppBase.ImportFileSettings.FilePath);

                void CreateSheet()
                {
                    var sheet = ImportFileWorkBook.AddWorksheet("PROPERTIES");
                    sheet.Cell(1, "A").Value = "NAME";
                    sheet.Cell(1, "B").Value = "CODE";
                    sheet.Cell(1, "C").Value = "TYPE";
                    sheet.Cell(1, "D").Value = "MEASUREENTITY";
                    sheet.Cell(1, "E").Value = "LOV";
                    sheet.Cell(1, "F").Value = "DESCRIPTION";
                    sheet.Cell(1, "G").Value = "FOLDER";
                }
                bool AreWorksheetsExists()
                {
                    IXLWorksheet sheet;
                    if (ImportFileWorkBook.TryGetWorksheet("PROPERTIES", out sheet))
                        return true;
                    return false;
                }
            }
        }
    }
}
