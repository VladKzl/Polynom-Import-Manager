using Ascon.Polynom.Api;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static TCS_Polynom_data_actualiser.AppBase;

namespace TCS_Polynom_data_actualiser
{
    public class GroupsActualisation
    {
        static GroupsActualisation()
        {
            DistinctedTCSGroupsNamesByType = new Func<Dictionary<string, List<string>>>(() =>
            {
                Dictionary<string, List<string>> distinctedGroupsByType = new Dictionary<string, List<string>>();
                foreach (var type in CommonSettings.Types)
                {
                    var workSheet = ElementsActualisationWorkBook.Value.Worksheet(type);
                    List<string> groupsNames = workSheet.Range(workSheet.Cell(3, "C"), workSheet.Column("C").LastCellUsed()).CellsUsed(true).Select(cell => (string)cell.Value).ToList();
                    List<string> distinctedGroupsNames = new List<string>();
                    foreach (string groupName in groupsNames)
                    {
                        if (!distinctedGroupsNames.Any(x => x == groupName))
                            distinctedGroupsNames.Add(groupName);
                    }
                    distinctedGroupsByType.Add(type, distinctedGroupsNames);
                }
                return distinctedGroupsByType;
            }).Invoke();
        }
        private static XLWorkbook NewGroupsActualisationWorkBook { get; set; } = new XLWorkbook();
        public static Dictionary<string, List<string>> DistinctedTCSGroupsNamesByType { get; set; }
        public static void CreateAndFillGroupsDocument()
        {
            if (File.Exists(GroupsActualisationSettings.FilePath))
                File.Move(GroupsActualisationSettings.FilePath, GroupsActualisationSettings.ArchivePath);
            CreateDocAndSheets();
            ActualiseGroups();
            using (NewGroupsActualisationWorkBook)
            {
                NewGroupsActualisationWorkBook.SaveAs(GroupsActualisationSettings.FilePath);
            }

            void CreateDocAndSheets()
            {
                foreach (string type in CommonSettings.Types)
                {
                    var sheet = NewGroupsActualisationWorkBook.AddWorksheet(type);
                    sheet.Cell(1, "A").Value = "TCS";
                    sheet.Cell(1, "B").Value = "Polynom";
                }
            }
            void ActualiseGroups()
            {
                foreach (var groupsPair in DistinctedTCSGroupsNamesByType)
                {
                    var workSheet = NewGroupsActualisationWorkBook.Worksheet(groupsPair.Key);
                    for (int i = 0; i < groupsPair.Value.Count(); i++) // Может не быть последнего элемента
                    {
                        string group = groupsPair.Value[i];
                        int cellRowNum = i + 2;
                        workSheet.Cell(cellRowNum, "A").Value = group;
                        IGroup findedGroup = null;
                        foreach (string polynomType in ElementsActualisationSettings.TcsByPolynomTypes[groupsPair.Key])
                        {
                            if (PolynomBase.TrySearchGroupInGroup(polynomType, group, out findedGroup))
                            {
                                workSheet.Cell(cellRowNum, "B").Value = findedGroup.Name;
                                break;
                            }
                            workSheet.Cell(cellRowNum, "B").Value = "";
                        }
                    }
                }
            }
        }
    }
}
