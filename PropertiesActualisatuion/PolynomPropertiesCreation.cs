using Ascon.Polynom.Api;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static TCS_Polynom_data_actualiser.AppBase;

namespace TCS_Polynom_data_actualiser
{
    public class PolynomPropertiesCreation
    {
        public static void UseApi()
        {
            List<IPropDefGroup> oldPolynomGroups = new List<IPropDefGroup>();
            GetGroupsFromAllPropertyes();
            IPropDefGroup rootGroup = null;
            if (!PolynomBase.Session.Objects.PropDefCatalog.PropDefGroups.Any(x => x.Name == "Новые свойства"))
            {
                rootGroup = PolynomBase.Session.Objects.PropDefCatalog.CreatePropDefGroup("Новые свойства");
            }

            var propRows = PropertiesSettings.PropertyesSheet.Column("B").CellsUsed().Select(x => x.WorksheetRow()).ToList();
            propRows.RemoveRange(0, 2);

            for(int i = 0; i < propRows.Count; i++)
            {
                string propName = propRows[i].Cell("B").Value.ToString(); //имя
                string propPath = propRows[i].Cell("C").Value.ToString(); //путь
                string propGroupName = TCSBase.Propertyes.GetPropGroupFromPath(propPath);
                List<string> propSplitPath = TCSBase.Propertyes.GetSplitPath(propPath);
                string polynomGroups = propRows[i].Cell("D").Value.ToString(); //полином пути
                int? polynomGroupIndex = null; //индекс пути
                if (propRows[i].Cell("E").Value.ToString() != string.Empty)
                    polynomGroupIndex = Convert.ToInt32(propRows[i].Cell("E").Value);

                List<IPropDefGroup> polynomPropGroups = new List<IPropDefGroup>();
                GetGroupFromAllPropertyes();

                string typeName;
                string code;
                string measureentity;
                string lov;
                string description;
                try
                {
                    typeName = TCSBase.Propertyes.GetColumnValueFromPropRows<string>(propName, propPath, TCSBase.Propertyes.RowColumnsForSearch.TYPE);
                    code = TCSBase.Propertyes.GetColumnValueFromPropRows<string>(propName, propPath, TCSBase.Propertyes.RowColumnsForSearch.CODE);
                    measureentity = TCSBase.Propertyes.GetColumnValueFromPropRows<string>(propName, propPath, TCSBase.Propertyes.RowColumnsForSearch.MEASUREENTITY);
                    lov = TCSBase.Propertyes.GetColumnValueFromPropRows<string>(propName, propPath, TCSBase.Propertyes.RowColumnsForSearch.LOV);
                    description = TCSBase.Propertyes.GetColumnValueFromPropRows<string>(propName, propPath, TCSBase.Propertyes.RowColumnsForSearch.DESCRIPTION);
                }
                catch
                {
                    Console.WriteLine($"Свойство в экселе {propName} с группой {propPath} не найдено в sql запросе на свойства. Скорее всего эксель изменил значение из за форматирования.");
                    continue;
                }

                IPropDefGroup propGroup;
                if (polynomGroups != string.Empty && polynomGroupIndex != null)
                {
                    List<IPropDefGroup> propPossibleGroups = polynomPropGroups.Where(x => x.Name == propGroupName).ToList();
                    propGroup = propPossibleGroups.ElementAt((int)polynomGroupIndex);
                    continue;
                }
                propGroup = CreatePropGroup(propSplitPath);
                CreateProperty();

                Console.Write($"\rОбработали групп {i}/{propRows.Count}");


                void GetGroupFromAllPropertyes()
                {
                    var groups = PolynomBase.Session.Objects.AllPropertyDefinitions.Select(x => x.OwnerGroup).ToList();
                    foreach (var group in groups)
                    {
                        if (!polynomPropGroups.Any(g => g.Id == group.Id))
                        {
                            polynomPropGroups.Add(group);
                        }
                    }
                }
                void CreateProperty()
                {
                    try
                    {
                        if (typeName == "string")
                        {

                            var propDefenition = propGroup.CreateStringPropertyDefinition(propName);
                            propDefenition.Code = code;
                        }
                        if (typeName == "date")
                        {
                            var propDefenition = propGroup.CreateDateTimePropertyDefinition(propName);
                            propDefenition.Code = code;
                        }
                        if (typeName == "double")
                        {
                            var propDefenition = propGroup.CreateDoublePropertyDefinition(propName);
                            propDefenition.Code = code;
                        }
                        if (typeName == "bool")
                        {
                            var propDefenition = propGroup.CreateBooleanPropertyDefinition(propName);
                            propDefenition.Code = code;
                        }
                        if (typeName == "int")
                        {
                            var propDefenition = propGroup.CreateIntegerPropertyDefinition(propName);
                            propDefenition.Code = code;
                        }
                        if (typeName == "enum")
                        {
                            if (lov == null)
                            {
                                throw new Exception($"Нет значений LOV. {propName} загружен без значений.");//
                            }
                            var propDefenition = propGroup.CreateEnumPropertyDefinition(propName);
                            propDefenition.Code = code;
                            FormatLov(lov).ForEach(x =>
                            {
                                propDefenition.AddItem(x);
                            });
                        }
                    }
                    catch
                    {
                        Console.WriteLine("Неизвестная ошибка при создании свойства.\n" +
                            "не пофикшено - \"1\". изменяется на \"1\" в экселе");
                        return;
                    }
                }
            }
            PolynomBase.Transaction.Commit();


            IPropDefGroup CreatePropGroup(List<string> tcsSplitPath)
            {
                IPropDefGroup propGroup = rootGroup;
                foreach (var groupName in tcsSplitPath)
                {
                    if(propGroup.PropDefGroups.Any(x => x.Name == groupName))
                    {
                        propGroup = propGroup.PropDefGroups.Single(x => x.Name == groupName);
                        continue;
                    }
                    propGroup = propGroup.CreatePropDefGroup(groupName);
                }
                return propGroup;
            }
            void GetGroupsFromAllPropertyes()
            {
                var groups = PolynomBase.Session.Objects.AllPropertyDefinitions.Select(x => x.OwnerGroup).ToList();
                foreach (var group in groups)
                {
                    if (!oldPolynomGroups.Any(g => g.Id == group.Id))
                    {
                        oldPolynomGroups.Add(group);
                    }
                }
            }
            List<string> FormatLov(string lov)
            {
                return lov.Split(';').ToList();
            }
        }
        public static void UseImportFile()
        {
            ImportFile.AddProperties();
            PolynomBase.Transaction.Rollback();
        }
    }
}
