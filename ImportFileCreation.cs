using Ascon.Polynom.Api;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static TCS_Polynom_data_actualiser.AppSettings;

namespace TCS_Polynom_data_actualiser
{
    public class ImportFileCreation
    {
        public static XLWorkbook NewImportFileWorkBook = new XLWorkbook();
        public static XLWorkbook CurrentImportFileWorkBook = new XLWorkbook(ImportFileSettings.FilePath);
        public static void CreateFile()
        {
            if (File.Exists(ImportFileSettings.FilePath))
                File.Move(ImportFileSettings.FilePath, ImportFileSettings.ArchivePath);

            foreach (string type in CommonSettings.Types)
            {
                var workSheet = NewImportFileWorkBook.AddWorksheet(type);

                var elementAndGroupPairs = PolynomObjectsCreation.CreatedGroupAndElementPairByType[type];
                for(int i = 0; i < elementAndGroupPairs.Count(); i++)
                {
                    var elementAndGroupPair = elementAndGroupPairs[i];
                    int elementNameGorisontalSellPosition = 0;
                    string elementName = elementAndGroupPair.element.Name;
                    string referenceName = elementAndGroupPair.group.ParentCatalog.Reference.Name;
                    string catalogName = elementAndGroupPair.group.ParentCatalog.Name;
                    List<string> groupsNames = new Func<List<string>>(() =>
                    {
                        List<string> groups = new List<string>();
                        IGroup groupToAdd = elementAndGroupPair.group;
                        groups.Add(groupToAdd.Name);
                        while (groupToAdd.ParentGroup != null)
                        {
                            groups.Add(groupToAdd.ParentGroup.Name);
                            groupToAdd = groupToAdd.ParentGroup;
                        }
                        groups.Reverse();
                        return groups;
                    }).Invoke();

                    int verticalCellPosition = i + 2;
                    workSheet.Cell(verticalCellPosition, "C").Value = referenceName;
                    workSheet.Cell(verticalCellPosition, "D").Value = catalogName;

                    for (int n = 0; n < groupsNames.Count(); n++)
                    {
                        int gorisontalCellPosition = n + 5;
                        workSheet.Cell(verticalCellPosition, gorisontalCellPosition).Value = groupsNames[n];
                        elementNameGorisontalSellPosition = gorisontalCellPosition + 1;
                    }
                    workSheet.Cell(verticalCellPosition, elementNameGorisontalSellPosition).Value = elementName;
                }
            }
            /*NewImportFileWorkBook.SaveAs(ImportFileSettings.FilePath);*/
            FormattingStructure();
        }
        public static void FormattingStructure()
        {
            foreach(string type in CommonSettings.Types)
            {
                var workSheet = NewImportFileWorkBook.Worksheet(type);
                var rows = workSheet.RangeUsed().Rows();
                foreach(var row in rows)
                {
                    string lastCellValue = row.LastCellUsed().Value.ToString();
                    row.LastCellUsed().Value = "";
                    row.LastCell().Value = lastCellValue;
                }

                var columns = workSheet.ColumnsUsed().ToList();
                columns[0].Cell(1).Value = "REFERENCE";
                columns[1].Cell(1).Value = "CATALOGS";

                var columnsUsed = workSheet.ColumnsUsed().ToList();
                for (int i = 2; i < columnsUsed.Count; i++)
                {
                    columnsUsed[i].Cell(1).Value = "GROUP";
                }
                columnsUsed.Last().Cell(1).Value = "NAME";
            }
            NewImportFileWorkBook.SaveAs(ImportFileSettings.FilePath);
        }
    }
}
