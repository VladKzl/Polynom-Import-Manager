using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using static Polynom_Import_Manager.AppBase;
using System.Collections;
using System;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Office2010.ExcelAc;

namespace Polynom_Import_Manager
{
    public class TCSBase
    {
        public TCSBase()
        {
            TCSConnection = new SqlConnection(CommonSettings.ConnectionString);
            TCSConnection.Open();
        }
        public static SqlConnection TCSConnection { get; set; }
        public class Elements
        {
            static Elements()
            {
                Console.WriteLine("Получаем элементы из Tcs. Подождите..");

                ElementAndPathByTcsType = new Func<Dictionary<string, List<(string element, string path)>>>(() =>
                {
                    var elementAndGroupByTcsType = new Dictionary<string, List<(string element, string group)>>();
                    foreach (var queryPair in ElementsSettings.TypesSQLQueries)
                    {
                        SqlCommand tcsCommand = new SqlCommand(queryPair.Value, TCSConnection);
                        SqlDataReader reader = tcsCommand.ExecuteReader();

                        var names = new List<(string, string)>();
                        while (reader.Read())
                        {
                            var element = reader.GetValue(0).ToString();
                            var group = reader.GetValue(1).ToString();
                            names.Add((element, group));
                        }
                        elementAndGroupByTcsType.Add(queryPair.Key, names);
                        reader.Close();
                        Console.WriteLine($"- Получили {queryPair.Key} из ТКС");
                    }
                    return elementAndGroupByTcsType;
                }).Invoke(); // вроде как нужно фильтрануть по element и path

                Console.WriteLine("Получили элементы из Tcs.\n");
            }
            public static Dictionary<string, List<(string element, string path)>> ElementAndPathByTcsType { get; set; }
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

                SqlCommand tcsCommand = new SqlCommand(PropertiesFile.PropertyesSQLQuery, TCSConnection);
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
                        List<string> splitPath = CommonCode.GetSplitPath(propRow.Field<string>("FOLDER"));
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

            public static T GetColumnValueFromPropRows<T>(string propName, string propPath, RowColumnsForSearch searchedColumn)
            {
                var row = PropertiesRows.Single(x => x.Field<string>("NAME") == propName && x.Field<string>("FOLDER") == propPath);
                return row.Field<T>(searchedColumn.ToString());
                // Может завершиться с ошибкой из за того что эксель форматирует изначальные значения и они отличаются от вводимых (нужно фиксить числа в дату)
            }
        }
    }
}