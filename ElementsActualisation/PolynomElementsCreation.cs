using Ascon.Polynom.Api;
using Ascon.Vertical.Application.Configuration;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.RightsManagement;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace TCS_Polynom_data_actualiser
{
    public class PolynomElementsCreation
    {
        public static Dictionary<string, IGroup> RootGroupByTcsType = new Dictionary<string, IGroup>();
        public static Dictionary<string, List<IGroup>> CreatedGroupsByTcsType = new Dictionary<string, List<IGroup>>();
        public static Dictionary<string, List<(IElement element, IGroup group)>> CreatedElementAndGroupByTcsType = new Dictionary<string, List<(IElement element, IGroup group)>>();
        private static void RootGroupCreation()
        {
            foreach (string type in AppBase.ElementsFileSettings.Types)
            {
                IGroup rootGroup = null;
                string rootGroupName = $"Нераспределенные {type}";
                var parentCatalog = PolynomBase.ElementsActualisation.GroupsByTcsType[type].First().ParentCatalog;
                if (!PolynomBase.TrySearchGroupInAllReferences(rootGroupName, out rootGroup))
                    rootGroup = parentCatalog.CreateGroup(rootGroupName);
                RootGroupByTcsType.Add(type, rootGroup);
            }
        }
        public static void GroupsCreation()
        {
            RootGroupCreation();
            foreach (string type in AppBase.ElementsFileSettings.Types)
            {
                List<IGroup> groupElements = new List<IGroup>();
                var rows = AppBase.GroupsActualisationWorkBook.Value.Worksheet(type).RangeUsed().Rows().ToList();
                rows.RemoveAt(0);
                foreach (var row in rows)
                {
                    string groupNameTcs = (string)row.Cell("A").Value;
                    string groupsPolynom = (string)row.Cell("B").Value;
                    int groupIndex = 0;
                    if (groupsPolynom != "")
                        groupIndex = Convert.ToInt32(row.Cell("C").Value);
                    
                    if (groupsPolynom == "")
                    {
                        CreateNewGroupAtRoot();
                        continue;
                    }

                    List<IGroup> findedAllGroups = new List<IGroup>();
                    foreach (string groupName in AppBase.ElementsFileSettings.TcsByPolynomTypes[type])
                    {
                        List<IGroup> findedGroups = null;
                        if (PolynomBase.TrySearchGroupsInGroup(groupName, groupNameTcs, out findedGroups))
                            findedAllGroups.AddRange(findedGroups);
                    }
                    if(findedAllGroups != null)
                        groupElements.Add(findedAllGroups.ElementAt(groupIndex));
                    if(findedAllGroups == null)
                        Console.WriteLine($"Файл устарел. Группа \"{groupNameTcs}\" не найдена. Актуализируйте элементы и группы заново");

                    void CreateNewGroupAtRoot()
                    {
                        var newGroup = RootGroupByTcsType[type].CreateGroup(groupNameTcs);
                        groupElements.Add(newGroup);
                        CommonCode.GetPercent(rows.Count(), row.RowNumber() - 1, type + "-");
                    }
                }
                CreatedGroupsByTcsType.Add(type, groupElements);
            }
        }
        public static void CreationAPI()
        {
            foreach (string type in AppBase.ElementsFileSettings.Types)
            {
                List<IGroup> typeGroups = CreatedGroupsByTcsType[type];
                var workSheet = AppBase.ElementsActualisationWorkBook.Value.Worksheet(type);
                // Так как исапольузется Range то буквы колонок меняются на A,B,C и тд.
                var elementsRows = workSheet.Range(workSheet.Cell(3, "B"), workSheet.Column("C").LastCellUsed().CellRight()).Rows();

                foreach (var elementRow in elementsRows)
                {
                    string elementNameTcs = (string)elementRow.Cell("A").Value;
                    string groupNameTcs = (string)elementRow.Cell("B").Value;
                    string elementNamePolynomDouble = elementRow.Cell("C").IsEmpty() ? "" : (string)elementRow.Cell("C").Value;
                    Validation();

                    IGroup group = typeGroups.Single(x => x.Name == groupNameTcs);
                    group.CreateElement(elementNameTcs);

                    CommonCode.GetPercent(elementsRows.Count(), elementRow.RowNumber() - 2, type + "-");

                    void Validation()
                    {
                        if (elementNamePolynomDouble == "")
                            return;

                        List<IElement> findedElements = null;
                        foreach (var createdGroup in CreatedGroupsByTcsType[type])
                        {
                            if (PolynomBase.TrySearchElementsInGroup(elementNamePolynomDouble, createdGroup, out findedElements))
                            {
                                findedElements.ForEach(x => x.Delete());
                                break;
                            }
                        }
                        if (findedElements == null)
                            Console.WriteLine($"Элемент-дубль {elementNamePolynomDouble} столбца \"D\" типа {type} не был найден.");
                    }
                }
            }
            Console.WriteLine("Применяем изменения..");
            PolynomBase.Transaction.Commit();
            Console.WriteLine("Изменения применены.");
        }
        public static void CreationImportFile()
        {
            foreach (string type in AppBase.ElementsFileSettings.Types)
            {
                List<IGroup> typeGroups = CreatedGroupsByTcsType[type];
                List<(IElement element, IGroup group)> elementAndGroup = new List<(IElement element, IGroup group)>();
                var workSheet = AppBase.ElementsActualisationWorkBook.Value.Worksheet(type);
                // Так как исапольузется Range то буквы колонок меняются на A,B,C и тд.
                var elementsRows = workSheet.Range(workSheet.Cell(3, "B"), workSheet.Column("C").LastCellUsed().CellRight()).Rows();

                foreach (var elementRow in elementsRows)
                {
                    string elementNameTcs = (string)elementRow.Cell("A").Value;
                    string groupNameTcs = (string)elementRow.Cell("B").Value;
                    string elementNamePolynomDouble = elementRow.Cell("C").IsEmpty() ? "" : (string)elementRow.Cell("C").Value;
                    Validation();

                    IGroup group = typeGroups.Single(x => x.Name == groupNameTcs);
                    var element = group.CreateElement(elementNameTcs);
                    elementAndGroup.Add((element, group));
                    CommonCode.GetPercent(elementsRows.Count(), elementRow.RowNumber() - 2, type + "-");

                    void Validation()
                    {
                        if (elementNamePolynomDouble == "")
                            return;

                        List<IElement> findedElements = null;
                        foreach (var createdGroup in CreatedGroupsByTcsType[type])
                        {
                            if (PolynomBase.TrySearchElementsInGroup(elementNamePolynomDouble, createdGroup, out findedElements))
                            {
                                findedElements.ForEach(x => x.Delete());
                                break;
                            }
                        }
                        if (findedElements == null)
                            Console.WriteLine($"Элемент-дубль {elementNamePolynomDouble} столбца \"D\" типа {type} не был найден.");
                    }
                }
                CreatedElementAndGroupByTcsType.Add(type, elementAndGroup);
            }
            ImportFile.AddGroupsAndElements();
            PolynomBase.Transaction.Rollback();
        }   
    }
}