using Ascon.Polynom.Api;
using Ascon.Vertical.Application.Configuration;
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
                foreach (var type in ElementsFileSettings.Types)
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
            if (File.Exists(GroupsFileSettings.FilePath))
                File.Move(GroupsFileSettings.FilePath, GroupsFileSettings.ArchivePath);
            CreateDocAndSheets();
            ActualiseGroups();
            using (NewGroupsActualisationWorkBook)
            {
                NewGroupsActualisationWorkBook.SaveAs(GroupsFileSettings.FilePath);
            }

            void CreateDocAndSheets()
            {
                foreach (string type in ElementsFileSettings.Types)
                {
                    var sheet = NewGroupsActualisationWorkBook.AddWorksheet(type);
                    sheet.Cell(1, "A").Value = "TCS";
                    sheet.Cell(1, "B").Value = "Polynom";
                    sheet.Cell(1, "C").Value = "Выбор группы Polynom";
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
                        List<IGroup> findedAllGroups = new List<IGroup>();
                        foreach (string polynomType in ElementsFileSettings.TcsByPolynomTypes[groupsPair.Key])
                        {
                            List<IGroup> findedGroups = null;
                            if (PolynomBase.TrySearchGroupsInGroup(polynomType, group, out findedGroups))
                                findedAllGroups.AddRange(findedGroups);
                        }
                        if(findedAllGroups != null)
                            workSheet.Cell(cellRowNum, "B").Value = CreatePaths(findedAllGroups);
                        if (findedAllGroups == null)
                            workSheet.Cell(cellRowNum, "B").Value = "";
                    }
                }
                string CreatePaths(List<IGroup> groups)
                {
                    StringBuilder pathsBouilder = new StringBuilder();
                    StringBuilder pathBuilder = new StringBuilder();

                    for(int g = 0; g < groups.Count; g++)
                    {
                        var paths = groups[g].GetPath();
                        for (int p = 0; p < paths.Count; p++)
                        {
                            if(paths[p] is IReference)
                            {
                                IReference reference = (IReference)paths[p];
                                pathBuilder.Append(reference.Name + "/");
                            }
                            if (paths[p] is ICatalog)
                            {
                                ICatalog catalog = (ICatalog)paths[p];
                                pathBuilder.Append(catalog.Name = "/");
                            }
                            if (paths[p] is IGroup)
                            {
                                IGroup _group = (IGroup)paths[p];
                                pathBuilder.Append(_group.Name + "/");
                            }
                        }
                        if(g < groups.Count - 1)
                            pathsBouilder.AppendLine($"{g} - {pathBuilder}");
                        if (g == groups.Count - 1)
                            pathsBouilder.Append($"{g} - {pathBuilder}");
                        pathBuilder.Clear();
                    }
                    return pathsBouilder.ToString();
                }
            }
        }
    }
}
