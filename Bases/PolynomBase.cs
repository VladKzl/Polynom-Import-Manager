using Ascon.Polynom.Api;
using Ascon.Polynom.Login;
using Ascon.Vertical.Technology;
using DocumentFormat.OpenXml.Office.PowerPoint.Y2021.M06.Main;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Linq.Expressions;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Runtime.Remoting.Messaging;
using System.Security.AccessControl;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static TCS_Polynom_data_actualiser.AppBase;

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
        public class ElementsActualisation
        {
            // GroupsByTcsType инициализируется отдельно, так как остальные очень долгие.
            public static void Initialize()
            {
                Console.WriteLine("Получаем данные из Полином. Подождите..");
                FillElementsByTypeAndElementsNamesByTypePropertyes();
                Console.WriteLine("Получили данные из Полином.\n");
            }
            public static Dictionary<string, List<IGroup>> GroupsByTcsType { get; set; } = new Func<Dictionary<string, List<IGroup>>>(() =>
            {
                Dictionary<string, List<IGroup>> groupsByType = new Dictionary<string, List<IGroup>>();
                foreach (var tcsByPolynomGroupsPair in ElementsActualisationSettings.TcsByPolynomTypes)
                {
                    List<IGroup> groups = new List<IGroup>();
                    foreach (string polynomTypeName in tcsByPolynomGroupsPair.Value)
                    {
                        IGroup group;
                        if (!TrySearchGroupInAllReferences(polynomTypeName, out group))
                            throw new Exception($"Группа(Папка/каталог) \"{polynomTypeName}\" не найдена в Полниноме. Проверьте правильность имени группы/подгруппы Полином.");
                        groups.Add(group);

                    }
                    groupsByType.Add(tcsByPolynomGroupsPair.Key, groups);
                }
                return groupsByType;
            }).Invoke();
            public static Dictionary<string, List<IElement>> ElementsByTcsType { get; set; } = new Dictionary<string, List<IElement>>();
            public static Dictionary<string, List<string>> ElementsNameByTcsType { get; set; } = new Dictionary<string, List<string>>();
            private static void FillElementsByTypeAndElementsNamesByTypePropertyes()
            {
                foreach (var groupsElementsByTypePair in GroupsByTcsType)
                {
                    List<IElement> elements = new List<IElement>();
                    List<string> elementsNames = new List<string>();
                    foreach (IGroup _group in groupsElementsByTypePair.Value)
                    {
                        Console.WriteLine($"- Получаем элементы из группы полинома \"{_group.Name}\". Ожидайте...");

                        List<IElement> findedElements = new List<IElement>();
                        if (!TrySearchAllElementsInGroup(_group, out findedElements))
                        {
                            Console.WriteLine($"В {_group.Name} нет эллементов.");
                            continue;
                        }
                        List<string> _elementsNames = findedElements.Select(element => element.Name).ToList();
                        elements.AddRange(findedElements);
                        elementsNames.AddRange(_elementsNames);

                        Console.WriteLine($"-- Получили {findedElements.Count} шт.");
                    }
                    if (elementsNames.Count != 0)
                    {
                        ElementsByTcsType.Add(groupsElementsByTypePair.Key, elements);
                        ElementsNameByTcsType.Add(groupsElementsByTypePair.Key, elementsNames);
                    }
                }
            }
        }
        public class PropertyesActualisation
        {

        }
        public static bool TrySearchGroupInGroup(string baseGroupName, string targetGroupName, out IGroup findedGroup)
        {
            IGroup _findedGroup;
            if(TrySearchGroupInAllReferences(baseGroupName, out _findedGroup))
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
        public static bool TrySearchGroupInAllReferences(string groupName, out IGroup findedGroup)
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
        public static bool TrySearchGroupsInAllReferences(string groupName, out List<IGroup> findedGroup)
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
        public static bool TrySearchElementInAllReferences(string elementName, out IElement findedElement)
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
        public static bool TrySearchElementsInGroup(string elementName, IGroup groupElement, out List<IElement> findedElements)
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
            var resultScope = groupElement.Intersect(condition);
            List<IElement> elements = resultScope.GetEnumerable<IElement>().ToList();
            if (elements.Count > 0 || elements == null)
            {
                findedElements = elements;
                return true;
            }
            findedElements = null;
            return false;
        }
        public static bool TrySearchAllElementsInGroup(IGroup groupElement, out List<IElement> findedElements)
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
            var resultScope = groupElement.Intersect(condition);

            var timer = new Stopwatch();
            timer.Restart();
            List<IElement> elements = resultScope.GetEnumerable<IElement>().ToList();
            Console.WriteLine($"*DEBUG*| Получили List<Elements> за {timer.Elapsed.TotalSeconds} сек.");
            timer.Stop();

            findedElements = elements;
            if (elements.Count() == 0 || elements == null)
                return false;
            return true;
        }
        public static bool TrySearchElementsInAllReferences(string elementName, out List<IElement> findedElement)
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
            findedElement = elements;
            if (elements.Count() == 0)
                return false;
            return true;
        }
    }
}
