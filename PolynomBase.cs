using Ascon.Polynom.Api;
using Ascon.Polynom.Login;
using Ascon.Vertical.Technology;
using DocumentFormat.OpenXml.Office.PowerPoint.Y2021.M06.Main;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Runtime.Remoting.Messaging;
using System.Security.AccessControl;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static TCS_Polynom_data_actualiser.AppSettings;

namespace TCS_Polynom_data_actualiser
{
    public class PolynomBase
    {
        public PolynomBase()
        {
            if (Session == null)
                Session = GetSession();
        }
        public static ISession Session;
        public static Dictionary<string, List<IGroup>> GroupsByType { get; } = new Func<Dictionary<string, List<IGroup>>>(() =>
        {
            Session = GetSession();
            Dictionary<string, List<IGroup>> groupsByType = new Dictionary<string, List<IGroup>>();
            foreach (var TcsToPolynomTypesPair in ElementsActualisationSettings.TcsToPolynomTypes)
            {
                List<IGroup> groups = new List<IGroup>();
                foreach (string polynomTypeName in TcsToPolynomTypesPair.Value)
                {
                    IGroup group;
                    if(TrySearchGroupInAllReferencesByName(polynomTypeName, out group))
                    {
                        groups.Add(group);
                    }
                    else
                        throw new Exception($"Группа(Папка/каталог) \"{polynomTypeName}\" не найдена в Полниноме. Проверьте правильность имени типа.");

                }
                groupsByType.Add(TcsToPolynomTypesPair.Key, groups);
            }
            return groupsByType;
        }).Invoke();
        public static Dictionary<string, List<IElement>> ElementsByType { get; set; } = new Dictionary<string, List<IElement>>();
        public static Dictionary<string, List<string>> ElementsNamesByType { get; set; } = new Dictionary<string, List<string>>();
        private static ISession GetSession()
        {
            ISession session;
            ConnectionInfo connectionInfo;
            bool isSessionOpen =  LoginManager.TryOpenSession(Guid.Empty, SessionOptions.None, ClientType.Editor,
                out session, 
                out connectionInfo);
            if (!isSessionOpen)
            {
                throw new Exception("Не удалось получить сессию");
            }
            Console.WriteLine($"Сессия \"{session.Id}\" успешно получена.");
            return session;
        }
        public static void FillElementsPropertyes()
        {
            foreach (var groupByTypePair in GroupsByType)
            {
                List<IElement> elements = new List<IElement>();
                List<string> elementsNames = new List<string>();
                foreach (IGroup _group in groupByTypePair.Value)
                {
                    var concept = Session.Objects.GetKnownConcept(KnownConceptKind.Element);
                    var propDef = Session.Objects.GetKnownPropertyDefinition(KnownPropertyDefinitionKind.Description);

                    var condition = Session.Objects.CreateSimpleCondition(
                        concept,
                        propDef,
                        (int)StringCompareOperation.NullValue,
                        null,
                        null,
                        (int)StringCompareOptions.None);
                    var resultScope = _group.Intersect(condition);

                    Console.WriteLine($"{groupByTypePair.Key} | Получаем элементы из группы полинома \"{_group.Name}\". Ожидайте...");

                    List<IElement> _elements = resultScope.GetEnumerable<IElement>().ToList();
                    if (_elements.Count == 0 || _elements == null)
                    {
                        Console.WriteLine($"В {_group.Name} нет эллементов.");
                        continue;
                    }
                    List<string> _elementsNames = _elements.Select(element => element.Name).ToList();
                    elements.AddRange(_elements);
                    elementsNames.AddRange(_elementsNames);

                    Console.WriteLine($"{groupByTypePair.Key} | Получили {_elements.Count} шт.");
                }
                if (elementsNames.Count != 0)
                {
                    ElementsByType.Add(groupByTypePair.Key, elements);
                    ElementsNamesByType.Add(groupByTypePair.Key, elementsNames);
                }
                Console.WriteLine();
            }
        }
        public static bool TrySearchGroupInGroupByName(string baseGroupName, string targetGroupName, out IGroup findedGroup)
        {
            IGroup _findedGroup;
            if(TrySearchGroupInAllReferencesByName(baseGroupName, out _findedGroup))
            {
                var concept = Session.Objects.GetKnownConcept(KnownConceptKind.Group);
                var propDef = Session.Objects.GetKnownPropertyDefinition(KnownPropertyDefinitionKind.Name);

                var condition = Session.Objects.CreateSimpleCondition(
                    concept,
                    propDef,
                    (int)StringCompareOperation.Equal,
                    ((IStringPropertyDefinition)propDef).CreateStringPropertyValueData(targetGroupName),
                    null,
                    (int)StringCompareOptions.None);
                var resultScope = _findedGroup.Intersect(condition);
                IGroup group = resultScope.GetEnumerable<IGroup>().FirstOrDefault();
                if (group != null)
                {
                    findedGroup = group;
                    return true;
                }
                findedGroup = null;
                return false;
            }
            findedGroup = null;
            Console.WriteLine($"Базовая группа \"{baseGroupName}\" не найдена при попытке найти группу в группе");
            return false;
        }
        public static bool TrySearchGroupInAllReferencesByName(string groupName, out IGroup findedGroup)
        {
            foreach (var reference in Session.Objects.AllReferences)
            {
                var concept = Session.Objects.GetKnownConcept(KnownConceptKind.Group);
                var propDef = Session.Objects.GetKnownPropertyDefinition(KnownPropertyDefinitionKind.Name);

                var condition = Session.Objects.CreateSimpleCondition(
                    concept,
                    propDef,
                    (int)StringCompareOperation.Equal,
                    ((IStringPropertyDefinition)propDef).CreateStringPropertyValueData(groupName),
                    null,
                    (int)StringCompareOptions.None);
                var resultScope = reference.Intersect(condition);
                IGroup group = resultScope.GetEnumerable<IGroup>().FirstOrDefault();
                if (group != null)
                {
                    findedGroup = group;
                    return true;
                }
            };
            findedGroup = null;
            return false;
        }
        public static bool TrySearchGroupsInAllReferencesByName(string groupName, out List<IGroup> findedGroup)
        {
            List<IGroup> groups = new List<IGroup>();
            foreach (var reference in Session.Objects.AllReferences)
            {
                var concept = Session.Objects.GetKnownConcept(KnownConceptKind.Group);
                var propDef = Session.Objects.GetKnownPropertyDefinition(KnownPropertyDefinitionKind.Name);

                var condition = Session.Objects.CreateSimpleCondition(
                    concept,
                    propDef,
                    (int)StringCompareOperation.Equal,
                    ((IStringPropertyDefinition)propDef).CreateStringPropertyValueData(groupName),
                    null,
                    (int)StringCompareOptions.None);
                var resultScope = reference.Intersect(condition);
                List<IGroup> _groups = resultScope.GetEnumerable<IGroup>().ToList();
                if (_groups.Count() != 0)
                {
                    groups.AddRange(_groups);
                }
            };
            if(groups.Count() == 0)
            {
                findedGroup = groups;
                return false;
            }
            findedGroup = groups;
            return true;
        }
        public static bool TrySearchElementInAllReferencesByName(string elementName, out IElement findedElement)
        {
            foreach (var reference in Session.Objects.AllReferences)
            {
                var concept = Session.Objects.GetKnownConcept(KnownConceptKind.Element);
                var propDef = Session.Objects.GetKnownPropertyDefinition(KnownPropertyDefinitionKind.Name);

                var condition = Session.Objects.CreateSimpleCondition(
                    concept,
                    propDef,
                    (int)StringCompareOperation.Equal,
                    ((IStringPropertyDefinition)propDef).CreateStringPropertyValueData(elementName),
                    null,
                    (int)StringCompareOptions.None);
                var resultScope = reference.Intersect(condition);
                IElement element = resultScope.GetEnumerable<IElement>().FirstOrDefault();
                if (element != null)
                {
                    findedElement = element;
                    return true;
                }
            };
            findedElement = null;
            return false;
        }
        public static bool TrySearchElementsInAllReferencesByName(string elementName, out List<IElement> findedElement)
        {
            List<IElement> elements = new List<IElement>();
            foreach (var reference in Session.Objects.AllReferences)
            {
                var concept = Session.Objects.GetKnownConcept(KnownConceptKind.Element);
                var propDef = Session.Objects.GetKnownPropertyDefinition(KnownPropertyDefinitionKind.Name);

                var condition = Session.Objects.CreateSimpleCondition(
                    concept,
                    propDef,
                    (int)StringCompareOperation.Equal,
                    ((IStringPropertyDefinition)propDef).CreateStringPropertyValueData(elementName),
                    null,
                    (int)StringCompareOptions.None);
                var resultScope = reference.Intersect(condition);
                List<IElement> _elements = resultScope.GetEnumerable<IElement>().ToList();
                if (_elements.Count() != 0)
                {
                    elements.AddRange(_elements);
                }
            };
            if (elements.Count() == 0)
            {
                findedElement = elements;
                return false;
            }
            findedElement = elements;
            return true;
        }

        //private static IElement GetElement(ISession session)
        //{
        //    var reference = session.Objects.AllReferences.FirstOrDefault(r => r.DisplayName == ReferenceName);
        //    if (reference != null)
        //    {
        //        var catalog = reference.Catalogs.FirstOrDefault(c => c.DisplayName == CatalogName);
        //        if (catalog != null)
        //        {
        //            var group1 = catalog.Groups.FirstOrDefault(g => g.DisplayName == Group1Name);
        //            if (group1 != null)
        //            {
        //                var group2 = GetFirstGroupByName(group1, Group2Name);
        //                if (group2 != null)
        //                {
        //                    var group3 = GetFirstGroupByName(group2, Group3Name);
        //                    if (group3 != null)
        //                    {
        //                        var group4 = GetFirstGroupByName(group3, Group4Name);
        //                        if (group4 != null)
        //                        {
        //                            var elements = group4.Elements;
        //                            var _object = elements.FirstOrDefault(e => e.Name == ObjectName);

        //                            return _object;
        //                        }
        //                        Console.WriteLine($"Не найдена группа \"{Group4Name}\".");
        //                    }
        //                    else
        //                    {
        //                        Console.WriteLine($"Не найдена группа \"{Group3Name}\".");
        //                    }
        //                }
        //                else
        //                {
        //                    Console.WriteLine($"Не найдена группа \"{Group2Name}\".");
        //                }
        //            }
        //            else
        //            {
        //                Console.WriteLine($"Не найдена группа \"{Group1Name}\".");
        //            }
        //        }
        //        else
        //        {
        //            Console.WriteLine($"Не найден каталог \"{CatalogName}\".");
        //        }
        //    }
        //    else
        //    {
        //        Console.WriteLine($"Не найден справочник \"{ReferenceName}\".");
        //    }

        //    return null;
        //}
    }
}
