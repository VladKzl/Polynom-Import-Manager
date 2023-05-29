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
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using static Polynom_Import_Manager.AppBase;

namespace Polynom_Import_Manager
{
    public class PolynomElementsCreation
    {
        public static void UseApi()
        {
            ElementsFile.ReconnectWorkBook(true);

            foreach (string type in ElementsSettings.Types)
            {
                var sheet = ElementsFile.WorkBook.Worksheet(type);

                List<IXLRow> rows = sheet.Column("C").CellsUsed().Select(x => x.WorksheetRow()).ToList();
                rows.RemoveRange(0, 1);
                for (int i = 0; i < rows.Count; i++)
                {
                    string elementName = rows[i].Cell("B").Value.ToString();
                    string path = rows[i].Cell("C").Value.ToString();
                    string polynomPaths = rows[i].Cell("D").Value.ToString();
                    int? index = null;
                    if(rows[i].Cell("E").Value.ToString() != string.Empty)
                    {
                        index = Convert.ToInt32(rows[i].Cell("E").Value);
                    }

                    IGroup group = null;
                    if (polynomPaths == string.Empty || index == null)
                    {
                        group = GetGroup(path);
                        group.CreateElement(elementName);
                        Console.Write($"\rЗагружено {type} - {i}/{rows.Count}");
                        continue;
                    }
                    string polynomPath = PolynomBase.GetPolynomPath(polynomPaths, index);
                    group = GetGroup(polynomPath);
                    if(!group.Elements.Any(x => x.Name == elementName))
                        group.CreateElement(elementName);

                    Console.Write($"\rЗагружено {type} - {i}/{rows.Count}");


                    IGroup GetGroup(string _path)
                    {
                        List<string> splitPath = CommonCode.GetSplitPath(_path);
                        var referenceName = splitPath[0];
                        var catalogName = splitPath[1];
                        List<string> groupsNames = splitPath.GetRange(2, splitPath.Count - 2);

                        IReference referenceObj = null;
                        ICatalog catalogObj = null;
                        IGroup groupObj = null;
                        //Справочник
                        if (PolynomBase.Session.Objects.AllReferences.Any(x => x.Name == referenceName))
                            referenceObj = PolynomBase.Session.Objects.AllReferences.Single(x => x.Name == referenceName);
                        else
                            referenceObj = PolynomBase.Session.Objects.CreateReference(referenceName);
                        //Каталог
                        if (referenceObj.Catalogs.Any(x => x.Name == catalogName))
                            catalogObj = referenceObj.Catalogs.Single(x => x.Name == catalogName);
                        else
                            catalogObj = referenceObj.CreateCatalog(catalogName);
                        //Группы
                        foreach (var groupName in groupsNames)
                        {
                            if (groupObj == null)
                            {
                                if (catalogObj.Groups.Any(x => x.Name == groupName))
                                    groupObj = catalogObj.Groups.Single(x => x.Name == groupName);
                                else
                                    groupObj = catalogObj.CreateGroup(groupName);
                            }
                            else
                            {
                                if (groupObj.Groups.Any(x => x.Name == groupName))
                                    groupObj = groupObj.Groups.Single(x => x.Name == groupName);
                                else
                                    groupObj = groupObj.CreateGroup(groupName);
                            }
                        }
                        return groupObj;
                    }

                }
                Console.Write("\n");
            }

            PolynomBase.Transaction.Commit();
        }
        public static void UseImportFile()
        {
            ImportFile.AddElements();
            PolynomBase.Transaction.Rollback();
        }
    }
}