using Ascon.Polynom.Api;
using Ascon.Vertical.Application.Configuration;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using static TCS_Polynom_data_actualiser.AppSettings;

namespace TCS_Polynom_data_actualiser
{
    public class PolynomObjectsCreation
    {
        public static XLWorkbook CurrentGroupsActualisationWorkBook = new XLWorkbook(GroupsActualisationSettings.FilePath);
        public static Lazy<XLWorkbook> CurrentElementsActualisationWorkBook = new Lazy<XLWorkbook>(new Func<XLWorkbook>(() => new XLWorkbook(ElementsActualisationSettings.FilePath)));
        public static Dictionary<string, List<(string groupName, IGroup group)>> CreatedGroupsAndNamesByType = new Dictionary<string, List<(string name, IGroup group)>>();
        public static Dictionary<string, List<(IElement element, IGroup group)>> CreatedGroupAndElementPairByType = new Dictionary<string, List<(IElement element, IGroup group)>>();
        public static Dictionary<string, IGroup> RootGroupByType = new Dictionary<string, IGroup>();
        private static ITransaction StartedTransaction = PolynomBase.Session.Objects.StartTransaction();
        public static void RootGroupCreation()
        {
            foreach (string type in CommonSettings.Types)
            {
                IGroup rootGroup = null;
                string rootGroupName = $"Нераспределенные {type}";
                var parentCatalog = PolynomBase.GroupsByType[type].First().ParentCatalog;
                if (PolynomBase.TrySearchGroupInAllReferencesByName(rootGroupName, out rootGroup))
                {}
                else
                {
                    rootGroup = parentCatalog.CreateGroup(rootGroupName);
                }
                RootGroupByType.Add(type, rootGroup);
            }
        }
        public static void GroupsCreation()
        {
            RootGroupCreation();
            foreach (string type in CommonSettings.Types)
            {
                List<(string groupName, IGroup group)> groupAndNamePairs = new List<(string groupName, IGroup group)>();
                var rows = CurrentGroupsActualisationWorkBook.Worksheet(type).RangeUsed().Rows().ToList();
                rows.RemoveAt(0);
                foreach (var row in rows)
                {
                    string groupNameTcs = (string)row.Cell("A").Value;
                    string groupNamePolynom = (string)row.Cell("B").Value;
                    
                    IGroup rootGroup;
                    if (groupNamePolynom == "")
                    {
                        RootGroupByType.TryGetValue(type, out rootGroup);
                        var newGroup = rootGroup.CreateGroup(groupNameTcs);
                        groupAndNamePairs.Add((groupName: groupNameTcs, group: newGroup));
                        CommonCode.GetPercent(rows.Count(), row.RowNumber() - 1, type + "-");
                    }
                    else
                    {
                        IGroup findedGroup = null;
                        foreach (string baseGroup in ElementsActualisationSettings.TcsToPolynomTypes[type])
                        {
                            if (PolynomBase.TrySearchGroupInGroupByName(baseGroup, groupNamePolynom, out findedGroup))
                            {
                                groupAndNamePairs.Add((groupName: groupNameTcs, group: findedGroup));
                                break;
                            }
                        }
                        if(findedGroup == null)
                        {
                            RootGroupByType.TryGetValue(type, out rootGroup);
                            var newGroup = rootGroup.CreateGroup(groupNameTcs);
                            groupAndNamePairs.Add((groupName: groupNameTcs, group: findedGroup));
                            CommonCode.GetPercent(rows.Count(), row.RowNumber(), type + "-");

                            Console.WriteLine($"Группа \"{groupNamePolynom}\" не найдена. Вы не верно указали имя группы.\n" +
                                $"Но группа создана в родительском каталоге и элемент будет помещен в нее.");
                        }
                    }
                }
                CreatedGroupsAndNamesByType.Add(type, groupAndNamePairs);
            }
        }
        public static void ElementsCreation(bool _useImportFile)
        {
            bool useImportFile = _useImportFile;
            foreach (string type in CommonSettings.Types)
            {
                List<(string groupName, IGroup group)> groupAndNamePairs = CreatedGroupsAndNamesByType[type];
                List<(IElement element, IGroup group)> elementAndGroupPair = new List<(IElement element, IGroup group)>();
                var workSheet = CurrentElementsActualisationWorkBook.Value.Worksheet(type);
                // Так как исапольузется Range то буквы колонок меняются на A,B,C и тд.
                var elementsRows = workSheet.Range(workSheet.Cell(3, "B"), workSheet.Column("C").LastCellUsed().CellRight().CellRight()).Rows();

                if (ElementsActualisationSettings.Status == ActualisationStatus.Не_актуализирован)
                {
                    foreach(var elementRow in elementsRows)
                    {
                        string elementNameTcs = (string)elementRow.Cell("A").Value;
                        string groupNameTcs = (string)elementRow.Cell("B").Value;

                        IGroup polynomGroupObject;
                        if (groupAndNamePairs.Where(x => x.groupName == groupNameTcs).Any())
                        {
                            polynomGroupObject = groupAndNamePairs.Where(x => x.groupName == groupNameTcs).First().group;
                            if (useImportFile)
                            {
                                var element = polynomGroupObject.CreateElement(elementNameTcs);
                                elementAndGroupPair.Add((element, polynomGroupObject));
                                CommonCode.GetPercent(elementsRows.Count(), elementRow.RowNumber() - 2, type + "-");
                            }
                            else
                            {
                                polynomGroupObject.CreateElement(elementNameTcs);
                                CommonCode.GetPercent(elementsRows.Count(), elementRow.RowNumber() - 2, type + "-");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"Элемент не добавлен, так как группа не найдена.");
                        }
                    }
                }
                else
                {
                    foreach (var elementRow in elementsRows)
                    {
                        string elementNameTcs = (string)elementRow.Cell("A").Value;
                        string groupNameTcs = (string)elementRow.Cell("B").Value;
                        string whatToDo = (string)elementRow.Cell("C").Value;
                        string currentPolynomName = (string)elementRow.Cell("D").Value;

                        IElement polynomElementObject;
                        IGroup polynomGroupObject;
                        if (currentPolynomName != "")
                        {
                            if(PolynomBase.TrySearchElementInAllReferencesByName(currentPolynomName, out polynomElementObject))
                            {
                                polynomElementObject.Name = elementNameTcs;
                            }
                            else
                            {
                                Console.WriteLine($"Элемент не переименован, так как не найден в Полиноме.");
                            }
                        }
                        else
                        {
                            if (groupAndNamePairs.Where(x => x.groupName == groupNameTcs).Any())
                            {
                                polynomGroupObject = groupAndNamePairs.Where(x => x.groupName == groupNameTcs).First().group;
                                if (useImportFile)
                                {
                                    var element = polynomGroupObject.CreateElement(elementNameTcs);
                                    elementAndGroupPair.Add((element, polynomGroupObject));
                                    CommonCode.GetPercent(elementsRows.Count(), elementRow.RowNumber(), type + "-");
                                }
                                else
                                    polynomGroupObject.CreateElement(elementNameTcs);
                            }
                            else
                            {
                                Console.WriteLine($"Элемент не добавлен, так как группа не найдена.");
                            }
                        }
                    }
                }
                if (useImportFile)
                    CreatedGroupAndElementPairByType.Add(type, elementAndGroupPair);
            }
            if (useImportFile)
            {
                ImportFileCreation.CreateFile();
                StartedTransaction.Rollback();
            }
            else
            {
                Console.WriteLine($"Изменения применяются, ожидайте.");
                StartedTransaction.Commit();
            }
                /*StartedTransaction.Rollback();*/
        }
    }
}
