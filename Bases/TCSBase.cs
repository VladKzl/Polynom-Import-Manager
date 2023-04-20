using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using static TCS_Polynom_data_actualiser.AppBase;
using System.Collections;
using System;

namespace TCS_Polynom_data_actualiser
{
    public class TCSBase
    {
        public TCSBase()
        {
            TCSConnection = new SqlConnection(TCSSettings.TCSConnectionString);
            TCSConnection.Open();
        }
        public static SqlConnection TCSConnection { get; set; }
        public class ElementsActualisation
        {
            static ElementsActualisation()
            {
                Console.WriteLine("Получаем данные из Tcs. Подождите..");
                ElementsNameByTcsType = new Func<Dictionary<string, List<string>>>(() =>
                {
                    var elementsNamesByType = new Dictionary<string, List<string>>();
                    foreach (var queryFilePath in TCSSettings.TcsQueryFilePaths)
                    {
                        string selectQuery = File.ReadAllText(queryFilePath.Value);
                        SqlCommand tcsCommand = new SqlCommand(selectQuery, TCSConnection);
                        SqlDataReader reader = tcsCommand.ExecuteReader();

                        List<string> names = new List<string>();
                        while (reader.Read())
                        {
                            names.Add((string)reader.GetValue(0));
                        }
                        elementsNamesByType.Add(queryFilePath.Key, names);
                        reader.Close();
                        Console.WriteLine($"- Получили {queryFilePath.Key} из ТКС");
                    }
                    return elementsNamesByType;
                }).Invoke();
                ElementsNameAndGroupForAllTypes = new Func<List<(string elementName, string groupName)>>(() =>
                {
                    var elementsNamesByType = new List<(string elementName, string groupName)>();
                    foreach (var queryFilePath in TCSSettings.TcsQueryFilePaths)
                    {
                        string selectQuery = File.ReadAllText(queryFilePath.Value);
                        SqlCommand tcsCommand = new SqlCommand(selectQuery, TCSConnection);
                        SqlDataReader reader = tcsCommand.ExecuteReader();

                        while (reader.Read())
                        {
                            elementsNamesByType.Add(((string)reader.GetValue(0), (string)reader.GetValue(1)));
                        }
                        reader.Close();
                    }
                    return elementsNamesByType;
                }).Invoke();
                Console.WriteLine("Получили данные из Tcs.\n");
            }
            public static Dictionary<string, List<string>> ElementsNameByTcsType { get; set; }
            public static List<(string elementName, string groupName)> ElementsNameAndGroupForAllTypes { get; set; }
        }
        public class PropertyesActualisation
        {

        }
    }
}