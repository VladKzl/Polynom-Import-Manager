using Ascon.Polynom.Api;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using static Polynom_Import_Manager.AppBase;

namespace Polynom_Import_Manager
{
    public class ImportFile
    {
        public enum Sheets
        {
            ELEMENTS,
            PROPERTIES,
            CONCEPTS
        }
        public static void AddProperties()
        {
            CreateDoc(Sheets.PROPERTIES);
            PropertiesFile.ReconnectWorkBook();

            var sheet = _ImportFile.WorkBook.Worksheet(Sheets.PROPERTIES.ToString());

            var propRows = PropertiesFile.PropertiesSheet.RowsUsed().ToList();
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
                    sheet.Cell(i + 2, "A").Value = propName;
                    sheet.Cell(i + 2, "B").Value = code;
                    sheet.Cell(i + 2, "C").Value = typeName;
                    sheet.Cell(i + 2, "D").Value = masureetity;
                    sheet.Cell(i + 2, "E").Value = lov != null ? FormatLov(lov) : "";
                    sheet.Cell(i + 2, "F").Value = description;
                    sheet.Cell(i + 2, "G").Value = realPath;
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
            _ImportFile.WorkBook.Save();


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
        }
        public static void AddElements()
        {
            CreateDoc(Sheets.ELEMENTS);
            ElementsFile.ReconnectWorkBook(false);

            foreach (string type in ElementsSettings.Types)
            {
                var elementsSheet = ElementsFile.WorkBook.Worksheet(type);
                var importSheet = _ImportFile.WorkBook.Worksheet(type);

                List<IXLRow> rows = elementsSheet.Column("C").CellsUsed().Select(x => x.WorksheetRow()).ToList();
                rows.RemoveRange(0, 1);
                for (int i = 0; i < rows.Count; i++)
                {
                    string elementName = rows[i].Cell("B").Value.ToString();
                    string polynomPath = new Func<string>(() =>
                    {
                        string path = rows[i].Cell("C").Value.ToString();
                        string polynomPaths = rows[i].Cell("D").Value.ToString();
                        int? index = null;
                        if (rows[i].Cell("E").Value.ToString() != string.Empty)
                        {
                            index = Convert.ToInt32(rows[i].Cell("E").Value);
                        }

                        if (polynomPaths == string.Empty || index == null)
                        {
                            return path; 
                        }
                        return PolynomBase.GetPolynomPath(polynomPaths, index);
                    }).Invoke();
                    List<string> splitPath = CommonCode.GetSplitPath(polynomPath);
                    string reference = splitPath[0];
                    string catalog = splitPath[1];
                    List<string> groups = splitPath.GetRange(2, splitPath.Count - 2);

                    FillRowData();

                    void FillRowData()
                    {
                        int elementNameGorisontalSellPosition = 0;
                        var nextRow = importSheet.Row(i + 2);

                        nextRow.Cell("C").Value = reference;
                        nextRow.Cell("D").Value = catalog;
                        FillGroupsData();
                        nextRow.Cell(elementNameGorisontalSellPosition).Value = elementName;

                        void FillGroupsData()
                        {
                            for (int n = 0; n < groups.Count(); n++)
                            {
                                int gorisontalCellPosition = n + 5;
                                nextRow.Cell(gorisontalCellPosition).Value = groups[n];
                                elementNameGorisontalSellPosition = gorisontalCellPosition + 1;
                            }
                        }
                    }
                }
                FormattingStructure();

                void FormattingStructure()
                {
                    var _rows = importSheet.RowsUsed().ToList();
                    _rows.RemoveRange(0, 1);
                    var lastCellIndex = importSheet.ColumnsUsed().ToList().Count + 2;
                    foreach (var _row in _rows)
                    {
                        string lastCellValue = _row.LastCellUsed().Value.ToString();
                        _row.LastCellUsed().Value = "";
                        _row.Cell(lastCellIndex).Value = lastCellValue;
                    }

                    var columnsUsed = importSheet.ColumnsUsed().ToList();
                    columnsUsed[0].Cell(1).Value = "REFERENCE";
                    columnsUsed[1].Cell(1).Value = "CATALOGS";

                    for (int i = 2; i < columnsUsed.Count; i++)
                    {
                        columnsUsed[i].Cell(1).Value = "GROUP";
                    }
                    columnsUsed.Last().Cell(1).Value = "NAME";
                }
            }
            _ImportFile.WorkBook.Save();
        }
        public static void CreateDoc(Sheets targetSheet)
        {
            if (File.Exists(_ImportFile.FilePath))
            {
                _ImportFile.WorkBook = new XLWorkbook(_ImportFile.FilePath);
                // Если файл слуществует и уже были добавлены страницы ранее, значит создаем новый а старый в архив.
                if (AreWorksheetsExists())
                {
                    File.Move(_ImportFile.FilePath, _ImportFile.ArchivePath);
                    _ImportFile.WorkBook = new XLWorkbook();
                    CreateSheet();
                    _ImportFile.WorkBook.SaveAs(_ImportFile.FilePath);
                    return;
                }
                // Если файл существует и станицы не были добавлены, то работаем в том же файле.
                CreateSheet();
                _ImportFile.WorkBook.Save();
                return;
            }
            _ImportFile.WorkBook = new XLWorkbook();
            CreateSheet();
            _ImportFile.WorkBook.SaveAs(_ImportFile.FilePath);

            bool AreWorksheetsExists()
            {
                IXLWorksheet sheet;
                switch (targetSheet)
                {
                    case Sheets.ELEMENTS:
                        foreach (var sheetName in ElementsSettings.Types)
                        {
                            if (_ImportFile.WorkBook.TryGetWorksheet(sheetName, out sheet))
                                return true;
                        }
                        return false;
                    case Sheets.PROPERTIES:
                        if (_ImportFile.WorkBook.TryGetWorksheet(Sheets.PROPERTIES.ToString(), out sheet))
                            return true;
                        return false;
                    default: throw new Exception();
                        /*                if (targetSheet == Sheets.CONCEPTS)
                {
                    IXLWorksheet sheet;
                    if (_ImportFile.WorkBook.TryGetWorksheet("PROPERTIES", out sheet))
                        return true;
                    return false;
                }*/
                }

            }
            void CreateSheet()
            {
                IXLWorksheet sheet;
                switch (targetSheet)
                {
                    case Sheets.ELEMENTS:
                        foreach (var sheetName in ElementsSettings.Types)
                        {
                            _ImportFile.WorkBook.AddWorksheet(sheetName);
                        }
                        break;
                    case Sheets.PROPERTIES:
                        sheet = _ImportFile.WorkBook.AddWorksheet(Sheets.PROPERTIES.ToString());
                        sheet.Cell(1, "A").Value = "NAME";
                        sheet.Cell(1, "B").Value = "CODE";
                        sheet.Cell(1, "C").Value = "TYPE";
                        sheet.Cell(1, "D").Value = "MEASUREENTITY";
                        sheet.Cell(1, "E").Value = "LOV";
                        sheet.Cell(1, "F").Value = "DESCRIPTION";
                        sheet.Cell(1, "G").Value = "FOLDER";
                        break;
                    default: throw new Exception();
                }
            }
        }
    }
}
