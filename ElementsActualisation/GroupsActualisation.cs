using Ascon.Polynom.Api;
using Ascon.Vertical.Application.Configuration;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Polynom_Import_Manager.AppBase;

namespace Polynom_Import_Manager
{
    public class GroupsActualisation
    {
        static GroupsActualisation()
        {
            /*DistinctedTCSGroupsNamesByType = new Func<Dictionary<string, List<string>>>(() =>
            {
                Dictionary<string, List<string>> distinctedGroupsByType = new Dictionary<string, List<string>>();
                foreach (var type in ElementsFile.Types)
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
            }).Invoke();*/
        }
        /*public static Dictionary<string, List<string>> DistinctedTCSGroupsNamesByType { get; set; }*/
        public static void CreateAndFillGroupsDocument()
        {
            ElementsFile.ReconnectWorkBook(false);
            FillSheets();
            ElementsFile.WorkBook.Save();


            void FillSheets()
            {
                foreach(var type in ElementsSettings.Types)
                {
                    var sheet = ElementsFile.WorkBook.Worksheet(type);
                    ArrayList typePolynomPathObjects = PolynomBase.Elements.TypePolynomPathObjects[type];

                    List<IXLRow> rows = sheet.Column("C").CellsUsed().Select(x => x.WorksheetRow()).ToList();
                    rows.RemoveRange(0,1);
                    for(int i = 0; i < rows.Count; i++)
                    {
                        string targetGroupName = CommonCode.GetLastGroupFromPath(rows[i].Cell("C").Value.ToString());

                        List<IGroup> allFindedGroups = new List<IGroup>();
                        FindeSimilarGroups();
                        if(allFindedGroups.Count > 0)
                            rows[i].Cell("D").Value = CreateAllPaths();

                        Console.Write($"\rОбработали групп {i}/{rows.Count}");


                        string CreateAllPaths()
                        {
                            StringBuilder allPathsBuilder = new StringBuilder();
                            /*                            //Добавляем базовые группы
                            List<string> polynomPaths = ElementsSettings.TypePolynomPaths[type];
                            for (int _i = 0; _i < polynomPaths.Count; _i++)
                            {
                                allPathsBuilder.AppendLine($"{_i} - {polynomPaths[_i]}/Нераспределенные");
                            }*/
                            // Добавляем похожие группы
                            for (int g = 0; g < allFindedGroups.Count; g++)
                            {
                                string path = CreatePath(allFindedGroups[g], g);

                                if (g < allFindedGroups.Count - 1)
                                    allPathsBuilder.AppendLine($"{g} - {path}");
                                if (g == allFindedGroups.Count - 1)
                                    allPathsBuilder.Append($"{g} - {path}");
                            }
                            return allPathsBuilder.ToString();

                            string CreatePath(IGroup group, int g)
                            {
                                StringBuilder pathBuilder = new StringBuilder();

                                foreach (var pathPart in group.GetPath().ToList())
                                {
                                    if (pathPart is IReference)
                                    {
                                        var referencePart = (IReference)pathPart;
                                        pathBuilder.Append(referencePart.Name + "/");
                                    }
                                    if (pathPart is ICatalog)
                                    {
                                        var catalogPart = (ICatalog)pathPart;
                                        pathBuilder.Append(catalogPart.Name + "/");
                                    }
                                    if(pathPart is IGroup)
                                    {
                                        var groupPart = (IGroup)pathPart;
                                        pathBuilder.Append(groupPart.Name + "/");
                                    }
                                }
                                return pathBuilder.ToString().TrimEnd('/');
                            }
                        }
                        void FindeSimilarGroups()
                        {
                            foreach (var pathObject in typePolynomPathObjects)
                            {
                                List<IGroup> findedGroups = new List<IGroup>();
                                if (pathObject is IGroup)
                                {
                                    IGroup group = (IGroup)pathObject;
                                    PolynomBase.TrySearchGroupsInGroup(group, targetGroupName, out findedGroups);
                                    allFindedGroups.AddRange(findedGroups);
                                }
                                if (pathObject is ICatalog)
                                {
                                    ICatalog catalog = (ICatalog)pathObject;
                                    PolynomBase.TrySearchGroupsInCatalog(catalog, targetGroupName, out findedGroups);
                                    allFindedGroups.AddRange(findedGroups);
                                }
                            }
                        }
                    }
                    Console.Write("\n");
                }
                
            }
            /*void FillSheets1()
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
                        foreach (string polynomType in ElementsFile.TypePolynomPaths[groupsPair.Key])
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
            }*/
        }
    }
}
