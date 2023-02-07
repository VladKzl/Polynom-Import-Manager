using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using static TCS_Polynom_data_actualiser.AppSettings;
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
            FillTCSElementsNamesByType();
        }
        private static SqlConnection TCSConnection { get; set; }
        public static Dictionary<string, List<string>> NamesByType = new Dictionary<string, List<string>>();
        public static List<List<string>> ElementsNamesAndGroup = new List<List<string>>();
        public static void FillTCSElementsNamesByType()
        {
            foreach(var queryFilePath in TCSSettings.TcsQueryFilePaths)
            {
                if(queryFilePath.Value != "" && queryFilePath.Value != null)
                {
                    string selectQuery = File.ReadAllText(queryFilePath.Value);
                    SqlCommand tcsCommand = new SqlCommand(selectQuery, TCSConnection);
                    SqlDataReader reader = tcsCommand.ExecuteReader();

                    List<string> names = new List<string>();
                    while (reader.Read())
                    {
                        if (reader.GetValue(0) != DBNull.Value)
                        {
                            names.Add((string)reader.GetValue(0));
                            ElementsNamesAndGroup.Add(new List<string>() { (string)reader.GetValue(0), (string)reader.GetValue(1) });
                        }
                    }
                    NamesByType.Add(queryFilePath.Key, names);
                    reader.Close();
                    Console.WriteLine($"Получили {queryFilePath.Key} из ТКС");
                }
                else
                    throw new Exception("Заполните пути до sql файлов tcs правильно.");
            }
            
        }
    }
}