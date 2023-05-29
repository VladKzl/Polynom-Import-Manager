using Ascon.Polynom.Api;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using static Polynom_Import_Manager.AppBase;

namespace Polynom_Import_Manager
{
    public class PropertyesGroupsActualisation
    {
        public static void ActualiseGroups()
        {
            List<IPropDefGroup> polynomPropGroups = new List<IPropDefGroup>();
            GetGroupFromAllPropertyes();

            List<string> extractedGroups = new List<string>();
            ExtractGroupFromPaths();

            for (int i = 0; i < extractedGroups.Count(); i++)
            {    
                var propRows = PropertiesFile.PropertiesSheet.RowsUsed().ToList();
                string propGroupName = extractedGroups[i];
                List<IPropDefGroup> propPossibleGroups = polynomPropGroups.Where(x => x.Name == propGroupName).ToList();

                if (propPossibleGroups.Count == 0)
                    continue;

                propRows[i + 2].Cell("D").Value = CreateAllPaths();

                Console.Write($"\rОбработали групп {i}/{extractedGroups.Count}");

                string CreateAllPaths()
                {
                    StringBuilder allPaths = new StringBuilder();
                    for (int g = 0; g < propPossibleGroups.Count; g++)
                    {
                        CreatePath(propPossibleGroups[g], g);
                    }
                    return allPaths.ToString();

                    void CreatePath(IPropDefGroup group, int g)
                    {
                        StringBuilder path = new StringBuilder();
                        path.Append(group.Name + "/");

                        while (group != null)
                        {
                            group = group.ParentGroup;
                            if(group != null)
                                path.Insert(0, group.Name + "/");
                        }
                        if (g < propPossibleGroups.Count - 1)
                            allPaths.AppendLine($"{g} - {path}");
                        if (g == propPossibleGroups.Count - 1)
                            allPaths.Append($"{g} - {path}");
                    }
                }
            }

            PropertiesFile.WorkBook.Value.Save();

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
            void ExtractGroupFromPaths()
            {
                List<string> propPaths = new List<string>();
                var pathsCells = PropertiesFile.PropertiesSheet.Column("C").CellsUsed().ToList();
                pathsCells.RemoveRange(0, 1);
                propPaths = pathsCells.Select(x => x.Value.ToString()).ToList();

                foreach (var propPath in propPaths)
                {
                    string propGroup = GetLastGroupFromPath(propPath);
                    extractedGroups.Add(propGroup);

                    string GetLastGroupFromPath(string path)
                    {
                        return path.Split('/').ToList().Last();
                    }
                }
            }
        }
    }
}
