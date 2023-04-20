using Ascon.Polynom.Api;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static TCS_Polynom_data_actualiser.AppBase;

namespace TCS_Polynom_data_actualiser
{
    public class ImportFile
    {
        public enum Sheets
        {
            Propertyes,
            Сoncepts
        }
        public static XLWorkbook ImportFileWorkBook { get; set; } = null;
        public static void AddGroupsAndElements()
        {
            CreateDoc();

            foreach (string type in CommonSettings.Types)
            {
                var workSheet = ImportFileWorkBook.Worksheet(type);

                var elementAndGroupPairs = PolynomObjectsCreation.CreatedElementAndGroupByTcsType[type];
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
                if (ImportFileWorkBook != null)
                {
                    CreateSheets();
                    return;
                }
                if (!File.Exists(ImportFileSettings.FilePath))
                {
                    ImportFileWorkBook = new XLWorkbook();
                    CreateSheets();
                    ImportFileWorkBook.SaveAs(ImportFileSettings.FilePath);
                    return;
                }
                File.Move(ImportFileSettings.FilePath, ImportFileSettings.ArchivePath);
                ImportFileWorkBook = new XLWorkbook();
                CreateSheets();
                ImportFileWorkBook.SaveAs(ImportFileSettings.FilePath);

                void CreateSheets()
                {
                    foreach (string type in CommonSettings.Types)
                    {
                        IXLWorksheet sheet;
                        if (ImportFileWorkBook.TryGetWorksheet(type, out sheet))
                            continue;

                        ImportFileWorkBook.AddWorksheet(type);
                    }
                }
            }
            void FormattingStructure()
            {
                foreach (string type in CommonSettings.Types)
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
        private static void CreateDocForOthers(Sheets sheet)
        {
            string sheetName = sheet.ToString();

            if (ImportFileWorkBook != null)
            {
                CreateSheet();
                return;
            }
            if (!File.Exists(ImportFileSettings.FilePath))
            {
                ImportFileWorkBook = new XLWorkbook();
                CreateSheet();
                ImportFileWorkBook.Worksheet(sheetName).Clear();
                ImportFileWorkBook.SaveAs(ImportFileSettings.FilePath);
                return;
            }
            File.Copy(ImportFileSettings.FilePath, ImportFileSettings.ArchivePath);
            ImportFileWorkBook = new XLWorkbook(ImportFileSettings.FilePath);
            CreateSheet();
            ImportFileWorkBook.Worksheet(sheetName).Clear();

            void CreateSheet()
            {
                IXLWorksheet worksheet;
                if (ImportFileWorkBook.TryGetWorksheet(sheetName, out worksheet))
                    return;
                ImportFileWorkBook.AddWorksheet(sheetName);
            }
        }
    }
}
