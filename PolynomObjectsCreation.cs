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
using static TCS_Polynom_data_actualiser.AppBase;

namespace TCS_Polynom_data_actualiser
{
    public class PolynomObjectsCreation
    {
        public static Dictionary<string, IGroup> RootGroupByTcsType = new Dictionary<string, IGroup>();
        public static Dictionary<string, List<IGroup>> CreatedGroupsByTcsType = new Dictionary<string, List<IGroup>>();
        public static Dictionary<string, List<(IElement element, IGroup group)>> CreatedElementAndGroupByTcsType = new Dictionary<string, List<(IElement element, IGroup group)>>();
        private static ITransaction StartedTransaction = PolynomBase.Session.Objects.StartTransaction();
        private static void RootGroupCreation()
        {
            foreach (string type in CommonSettings.Types)
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
            foreach (string type in CommonSettings.Types)
            {
                List<IGroup> groupElements = new List<IGroup>();
                var rows = GroupsActualisationWorkBook.Value.Worksheet(type).RangeUsed().Rows().ToList();
                rows.RemoveAt(0);
                foreach (var row in rows)
                {
                    string groupNameTcs = (string)row.Cell("A").Value;
                    string groupNamePolynom = (string)row.Cell("B").Value;
                    
                    if (groupNamePolynom == "")
                    {
                        var newGroup = RootGroupByTcsType[type].CreateGroup(groupNameTcs);
                        groupElements.Add(newGroup);
                        CommonCode.GetPercent(rows.Count(), row.RowNumber() - 1, type + "-");
                        continue;
                    }
                    IGroup findedGroup = null;
                    foreach (string groupName in ElementsActualisationSettings.TcsByPolynomTypes[type])
                    {
                        if (PolynomBase.TrySearchGroupInGroup(groupName, groupNamePolynom, out findedGroup))
                        {
                            groupElements.Add(findedGroup);
                            break;
                        }
                    }
                    if(findedGroup == null)
                    {
                        var newGroup = RootGroupByTcsType[type].CreateGroup(groupNameTcs);
                        groupElements.Add(newGroup);
                        CommonCode.GetPercent(rows.Count(), row.RowNumber(), type + "-");

                        Console.WriteLine($"Группа \"{groupNamePolynom}\" не найдена. Вы не верно указали имя группы.\n" +
                            $"Но группа создана в родительском каталоге и элемент будет помещен в нее.");
                    }
                }
                CreatedGroupsByTcsType.Add(type, groupElements);
            }
        }
        public static void ElementsCreation()
        {
            foreach (string type in CommonSettings.Types)
            {
                List<IGroup> typeGroups = CreatedGroupsByTcsType[type];
                var workSheet = ElementsActualisationWorkBook.Value.Worksheet(type);
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
            StartedTransaction.Commit();
            Console.WriteLine("Изменения применены.");
        }
        public static void ElementsCreationImportFile()
        {
            foreach (string type in CommonSettings.Types)
            {
                List<IGroup> typeGroups = CreatedGroupsByTcsType[type];
                List<(IElement element, IGroup group)> elementAndGroup = new List<(IElement element, IGroup group)>();
                var workSheet = ElementsActualisationWorkBook.Value.Worksheet(type);
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
            StartedTransaction.Rollback();
        }
    }
}
