using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using static TCS_Polynom_data_actualiser.AppBase;
using System.Collections;
using System;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Office2010.ExcelAc;

namespace TCS_Polynom_data_actualiser
{
    public class TCSBase
    {
        public TCSBase()
        {
            TCSConnection = new SqlConnection(CommonSettings.TCSConnectionString);
            TCSConnection.Open();
        }
        public static SqlConnection TCSConnection { get; set; }
        public class Elements
        {
            static Elements()
            {
                Console.WriteLine("Получаем элементы из Tcs. Подождите..");
                ElementsNameByTcsType = new Func<Dictionary<string, List<string>>>(() =>
                {
                    var elementsNamesByType = new Dictionary<string, List<string>>();
                    foreach (var queryPair in ElementsFileSettings.TcsTypesSQLQueries)
                    {
                        SqlCommand tcsCommand = new SqlCommand(queryPair.Value, TCSConnection);
                        SqlDataReader reader = tcsCommand.ExecuteReader();

                        List<string> names = new List<string>();
                        while (reader.Read())
                        {
                            names.Add((string)reader.GetValue(0));
                        }
                        elementsNamesByType.Add(queryPair.Key, names);
                        reader.Close();
                        Console.WriteLine($"- Получили {queryPair.Key} из ТКС");
                    }
                    return elementsNamesByType;
                }).Invoke();
                AllTypesElementAndGroupNames = new Func<List<(string elementName, string groupName)>>(() =>
                {
                    var elementsNamesByType = new List<(string elementName, string groupName)>();
                    foreach (var queryPair in ElementsFileSettings.TcsTypesSQLQueries)
                    {
                        SqlCommand tcsCommand = new SqlCommand(queryPair.Value, TCSConnection);
                        SqlDataReader reader = tcsCommand.ExecuteReader();

                        while (reader.Read())
                        {
                            elementsNamesByType.Add(((string)reader.GetValue(0), (string)reader.GetValue(1)));
                        }
                        reader.Close();
                    }
                    return elementsNamesByType;
                }).Invoke();
                Console.WriteLine("Получили элементы из Tcs.\n");
            }
            public static Dictionary<string, List<string>> ElementsNameByTcsType { get; set; }
            public static List<(string elementName, string groupName)> AllTypesElementAndGroupNames { get; set; }
        }
        public class Propertyes
        {
            public enum RowColumnsForSearch
            {
                CODE,
                TYPE,
                MEASUREENTITY,
                LOV,
                DESCRIPTION,

            }
            static Propertyes()
            {
                Console.WriteLine("Получаем свойства из Tcs. Подождите..");

                SqlCommand tcsCommand = new SqlCommand(PropertiesSettings.TcsPropertyesSQLQuery, TCSConnection);
                SqlDataReader reader = tcsCommand.ExecuteReader();
                DataTable table = new DataTable();
                table.Load(reader);
                var propRowsDistincted = table.AsEnumerable().GroupBy(x => (x.Field<string>("NAME"), x.Field<string>("FOLDER"))).Select(x => x.First()).ToList();

                PropertiesRows = propRowsDistincted;
                PropertiesAndSplitPath = new Func<List<(string prop, List<string> splitPath)>>(() =>
                {
                    List<(string prop, List<string> splitPath)> propertyAndSplitPath = new List<(string prop, List<string> splitPath)>();
                    foreach(var propRow in propRowsDistincted)
                    {
                        string propName = propRow.Field<string>("NAME");
                        List<string> splitPath = GetSplitPath(propRow.Field<string>("FOLDER"));
                        propertyAndSplitPath.Add((propName, splitPath));
                    }
                    return propertyAndSplitPath;
                }).Invoke();
                PropertiesAndPath = new Func<List<(string prop, string path)>>(() =>
                {
                    List<(string prop, string path)> propertyAndPath = new List<(string prop, string path)>();
                    foreach (var propRow in propRowsDistincted)
                    {
                        string propName = propRow.Field<string>("NAME");
                        string path = propRow.Field<string>("FOLDER");
                        propertyAndPath.Add((propName, path));
                    }
                    return propertyAndPath;
                }).Invoke();
                PropertiesNames = new Func<List<string>>(() =>
                {
                    List<string> propertys = new List<string>();
                    foreach (var propRow in propRowsDistincted)
                    {
                        string propName = propRow.Field<string>("NAME");
                        propertys.Add(propName);
                    }
                    return propertys;
                }).Invoke();
                Console.WriteLine("Получили свойства из Tcs.\n");
            }
            public static List<DataRow> PropertiesRows { get; set; }
            public static List<(string prop, List<string> splitPath)> PropertiesAndSplitPath { get; set; }
            public static List<(string prop, string path)> PropertiesAndPath { get; set; }
            public static List<string> PropertiesNames { get; set; }
            public static List<string> GetSplitPath(string path)
            {
                return path.Split('/').ToList();
            }
            public static string GetPropGroupFromPath(string path)
            {
                return path.Split('/').ToList().Last();
            }
            public static T GetColumnValueFromPropRows<T>(string propName, string propPath, RowColumnsForSearch searchedColumn)
            {
                var row = PropertiesRows.Single(x => x.Field<string>("NAME") == propName && x.Field<string>("FOLDER") == propPath);
                return row.Field<T>(searchedColumn.ToString());
                // Может завершиться с ошибкой из за того что эксель форматирует изначальные значения и они отличаются от вводимых (нужно фиксить числа в дату)
            }
        }
    }
}