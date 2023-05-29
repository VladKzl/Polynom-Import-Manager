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
using static Polynom_Import_Manager.AppBase;
using Ascon.Polynom.Api;
using System.Collections;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace Polynom_Import_Manager
{
    public class PolynomBase
    {
        public PolynomBase()
        {
            if (Session == null)
                Session = GetSession();
            Transaction = Session.Objects.StartTransaction();
        }
        public static ISession Session;
        public static ITransaction Transaction;
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
        public class Elements
        {
            static Elements()
            {
                TypePolynomPathObjects = new Func<Dictionary<string, ArrayList>>(() =>
                {
                    var typePolynomPathObjects = new Dictionary<string, ArrayList>();
                    foreach (var type in ElementsSettings.Types)
                    {
                        List<string> typePaths = ElementsSettings.TypePolynomPaths[type];
                        if (typePaths.Count == 0) // Если путь не указан в экселе, то у типа value пустой объект
                        {
                            typePolynomPathObjects.Add(type, new ArrayList());
                            continue;
                        }
                        List<List<string>> splitedPaths = new List<List<string>>();
                        foreach (var typePath in typePaths)
                        {
                            splitedPaths.Add(CommonCode.GetSplitPath(typePath));
                        }
                        // Получаем конечные объекты path для типа
                        ArrayList arrayList = new ArrayList();
                        foreach (List<string> splitedPath in splitedPaths)
                        {
                            arrayList.Add(GetLastGroupOrCatalogFromPath(splitedPath));
                        }
                        typePolynomPathObjects.Add(type, arrayList);


                        object GetLastGroupOrCatalogFromPath(List<string> splitedPath)
                        {
                            if (splitedPath.Count == 1)
                                throw new Exception($"Путь до элементов полинома от \"{type}\" должен заканчиваться не справочником, а каталогом или группой.");

                            string referenceName = splitedPath[0];
                            string catalogName = splitedPath[1];
                            List<string> groupNames = splitedPath.Count >= 3 ? splitedPath.GetRange(2, splitedPath.Count - 2) : null;

                            IReference referenceObject = null;
                            ICatalog catalogObject = null;

                            // Ищем справочник
                            if (!Session.Objects.AllReferences.Any(x => x.Name == referenceName))
                                throw new Exception($"Справочника \"{referenceName}\" не существует. Проверьте путь типа \"{type}\".");
                            referenceObject = Session.Objects.AllReferences.Single(x => x.Name == referenceName);
                            // Ищем каталог
                            if (!referenceObject.Catalogs.Any(x => x.Name == catalogName))
                                throw new Exception($"Каталога \"{catalogName}\" не существует. Проверьте путь типа \"{type}\".");
                            catalogObject = referenceObject.Catalogs.Single(x => x.Name == catalogName);

                            if (splitedPath.Count == 2)
                            {
                                return catalogObject;
                            }
                            if (splitedPath.Count >= 3)
                            {
                                return GetLastGroupFromPath();
                            }
                            return null;


                            IGroup GetLastGroupFromPath()
                            {
                                IGroup groupObj = null;
                                foreach (var groupName in groupNames)
                                {
                                    if (groupObj == null)
                                    {
                                        if (!catalogObject.Groups.Any(x => x.Name == groupName))
                                            throw new Exception($"Группы \"{groupName}\" не существует в каталоге \"{catalogName}\". Проверьте путь типа \"{type}\".");
                                        groupObj = catalogObject.Groups.Single(x => x.Name == groupName);
                                        continue;
                                    }
                                    if (!groupObj.Groups.Any(x => x.Name == groupName))
                                        throw new Exception($"Группы \"{groupName}\" не существует в группе \"{groupObj.Name}\". Проверьте путь типа \"{type}\".");
                                    groupObj = groupObj.Groups.Single(g => g.Name == groupName);
                                }
                                return groupObj;
                            }
                        }
                    }
                    return typePolynomPathObjects;
                }).Invoke();
            }
            public static Dictionary<string, ArrayList> TypePolynomPathObjects { get; set; }
            public static Lazy<Dictionary<string, List<IElement>>> ElementsByTcsType { get; set; } = new Lazy<Dictionary<string, List<IElement>>>(new Func<Dictionary<string, List<IElement>>>(() => 
            {
                Console.WriteLine("Получаем элементы из Полином. Подождите..");

                Dictionary<string, List<IElement>> elementsByTcsType = new Dictionary<string, List<IElement>>();
                foreach (var type in ElementsSettings.Types)
                {
                    ArrayList pathObjects = TypePolynomPathObjects[type];

                    List<IElement> typeElements = new List<IElement>();
                    foreach (var pathObject in pathObjects)
                    {
                        if(pathObject is IGroup)
                        {
                            List<IElement> findedElements = new List<IElement>();
                            IGroup group = (IGroup)pathObject;
                            TrySearchAllElementsInGroup(group, out findedElements);
                            typeElements.AddRange(findedElements);
                        }
                        if (pathObject is ICatalog)
                        {
                            List<IElement> findedElements = new List<IElement>();
                            ICatalog catalog = (ICatalog)pathObject;
                            TrySearchElementsInCatalog(catalog, out findedElements);
                            typeElements.AddRange(findedElements);
                        }
                    }
                    Console.WriteLine($"- получили {type}.\n");
                    elementsByTcsType.Add(type, typeElements);
                }
                Console.WriteLine("Получили все элементы из Полином.\n");
                return elementsByTcsType;
            })); // Могут быть дубли элментов. Фильтрануть

            /* public static Dictionary<string, List<IGroup>> GroupsByTcsType { get; set; } = new Func<Dictionary<string, List<IGroup>>>(() =>
             {
                 Dictionary<string, List<IGroup>> groupsByType = new Dictionary<string, List<IGroup>>();
                 foreach (var tcsByPolynomGroupsPair in ElementsFile.TypePolynomPaths)
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
             }).Invoke();*/
        }
        public class PropertyesActualisation
        {
            static PropertyesActualisation()
            {
                Console.WriteLine("Получаем свойства из Полином. Подождите..");
                Properties = Session.Objects.AllPropertyDefinitions.ToList();
                PropertiesNames = Session.Objects.AllPropertyDefinitions.Select(x => x.Name).ToList();
                Console.WriteLine("Получили свойства из Полином.\n");
            }
            public static List<IPropertyDefinition> Properties { get; set; }
            public static List<string> PropertiesNames { get; set; }
        }
        public static string GetPolynomPath(string polynomPaths, int? index)
        {
            var separatedPaths1 = Regex.Replace(polynomPaths, @"\d", ";");
            var separatedPaths2 = separatedPaths1.Replace(" - ", "");
            List<string> splitedPaths = separatedPaths2.Split(';').ToList();
            splitedPaths.RemoveAt(0);
            return splitedPaths[index.Value];
        }
        public static bool TrySearchGroupsInGroup(IGroup baseGroupName, string targetGroupName, out List<IGroup> findedGroups)
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
            var resultScope = baseGroupName.Intersect(condition);
            List<IGroup> groups = resultScope.GetEnumerable<IGroup>().ToList();

            findedGroups = groups;
            if (groups.Count > 0)
                return true;
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
        public static bool TrySearchElementsInGroup(IGroup groupElement, string elementName, out List<IElement> findedElements)
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

            findedElements = elements;
            if (elements.Count > 0)
                return true;
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
            if (elements.Count == 0)
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
        public static bool TrySearchElementsInCatalog(ICatalog catalogObj, out List<IElement> findedElements)
        {
            var concept = Session.Objects.GetKnownConcept(KnownConceptKind.Element);
            var propDef = Session.Objects.GetKnownPropertyDefinition(KnownPropertyDefinitionKind.Description);

            var condition = Session.Objects.CreateSimpleCondition(
                concept,
                propDef,
                (int)StringCompareOperation.NullValue,// Тут можно просто None
                null,
                null,
                (int)StringCompareOptions.None);
            var resultScope = catalogObj.Intersect(condition);
            List<IElement> elements = resultScope.GetEnumerable<IElement>().ToList();
            if (elements.Count > 0)
            {
                findedElements = elements;
                return true;
            }
            findedElements = new List<IElement>();
            return false;
        }
        public static bool TrySearchGroupsInCatalog(ICatalog catalogObj, string targetGroupName, out List<IGroup> findedGroups)
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
            var resultScope = catalogObj.Intersect(condition);
            List<IGroup> groups = resultScope.GetEnumerable<IGroup>().ToList();

            findedGroups = groups;
            if (groups.Count > 0)
                return true;
            return false;
        }
        /*        public static bool TrySearchPropertiesInAllReferences()
                {
            

                    var concept = Session.Objects.GetKnownConcept(KnownConceptKind.);

                    var propDefCatalog = (IPropertyOwnerScope)Session.Objects.PropDefCatalog.PropDefGroups.First().PropertyDefinitions.First().;



                    var condition = Session.Objects.CreateSimpleCondition(
                        concept,
                        null,
                        (int)StringCompareOperation.None,
                        null,
                        null,
                        (int)StringCompareOptions.None);

                    condition.In
                    var resultScope = propDefCatalog.Intersect(condition);

                    return true;
                }*/
    }
}
