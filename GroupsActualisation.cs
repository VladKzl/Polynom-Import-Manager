using Ascon.Polynom.Api;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static TCS_Polynom_data_actualiser.AppSettings;

namespace TCS_Polynom_data_actualiser
{
    public class GroupsActualisation
    {
        public static XLWorkbook NewGroupsActualisationWorkBook = new XLWorkbook();
        public static XLWorkbook CurrentGroupsActualisationWorkBook = new XLWorkbook(GroupsActualisationSettings.FilePath);
        public static Dictionary<string, List<string>> ActualisedGroupsByType { get; } = new Func<Dictionary<string, List<string>>>(() =>
        {
            Dictionary<string, List<string>> actualisedGroupsByType = new Dictionary<string, List<string>>();
            XLWorkbook actualisedElementsWorkBook = new XLWorkbook(ElementsActualisationSettings.FilePath);
            foreach(var type in CommonSettings.Types)
            {
                List<string> groupsNamesDistincted = new List<string>();
                Dictionary<string, string> distinctChecker = new Dictionary<string, string>();
                var workSheet = actualisedElementsWorkBook.Worksheet(type);
                var groupsNames = workSheet.Range(workSheet.Cell(3, "C"), workSheet.Column("C")
                                                    .LastCellUsed()).CellsUsed(true)
                                                    .Select(cell => (string)cell.Value).ToList();
                foreach (string name in groupsNames)
                {
                    try
                    {
                        distinctChecker.Add(name, "");
                    }
                    catch
                    {
                        continue;
                    }
                    groupsNamesDistincted.Add(name);
                }
                actualisedGroupsByType.Add(type, groupsNamesDistincted);
            }
            return actualisedGroupsByType;
        }).Invoke();
        public static void FillDocumentData()
        {
            if (File.Exists(GroupsActualisationSettings.FilePath))
                File.Move(GroupsActualisationSettings.FilePath, GroupsActualisationSettings.ArchivePath);

            foreach (string type in CommonSettings.Types)
            {
                var sheet = NewGroupsActualisationWorkBook.AddWorksheet(type);
                sheet.Cell(1, "A").Value = "TCS";
                sheet.Cell(1, "B").Value = "Polynom";
            }
            foreach(var groupsPair in ActualisedGroupsByType)
            {
                var workSheet = NewGroupsActualisationWorkBook.Worksheet(groupsPair.Key);
                for (int i = 0; i < groupsPair.Value.Count(); i++) // Может не быть последнего элемента
                {
                    string group = groupsPair.Value[i];
                    int cellRowNum = i + 2;
                    workSheet.Cell(cellRowNum, "A").Value = group;
                    IGroup findedGroup = null;
                    foreach(string polynomType in ElementsActualisationSettings.TcsToPolynomTypes[groupsPair.Key])
                    {
                        if (PolynomBase.TrySearchGroupInGroupByName(polynomType, group, out findedGroup))
                        {
                            workSheet.Cell(cellRowNum, "B").Value = findedGroup.Name;
                            break;
                        }
                        workSheet.Cell(cellRowNum, "B").Value = "";
                    }
                    /*NewGroupsActualisationWorkBook.Worksheet(groupsPair.Key).Cell(cellRowNum, "B").Value = PolynomBase.TrySearchGroupInAllReferencesByTcsName(group) ? group : "";*/
                }
            }
            NewGroupsActualisationWorkBook.SaveAs(GroupsActualisationSettings.FilePath);
        }
    }
}
