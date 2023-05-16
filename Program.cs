using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static TCS_Polynom_data_actualiser.PolynomBase;
using static TCS_Polynom_data_actualiser.ElementsActualisation;
using DocumentFormat.OpenXml.Spreadsheet;
using static TCS_Polynom_data_actualiser.Algorithm;
using Ascon.Polynom.Api;

namespace TCS_Polynom_data_actualiser
{
    public class Program
    {
        public static string settingsWorkbookPath = "D:\\ascon_obmen\\kozlov_vi\\Полином\\Приложения\\TCS_Polynom_data_actualiser\\Настройки.xlsx";
        public static TCSBase TcsBase { get; set;  }
        public static PolynomBase PolynomBase { get; set; }
        //public static string Role = "Администраторы";
        //public static string Name = @"MZ\kozlov_vi";
        //public static string Password = "uw39sccvb";

        [STAThread]
        static void Main(string[] args)
        {
            AppConfiguration();
            while (true)
            {
                Console.WriteLine
                (
                    "\nВыберите номер режима импорта в Полином.\n" +
                    "1 - Импорт элементов и групп через API\n" +
                    "2 - Импорт свойств и групп через API\n" +
                    "3 - Импорт элементов и групп через обменный файл\n" +
                    "4 - Импорт свойств и групп через обменный файл\n" +
                    "5 - Откат изменений"
                );
                
                switch (Convert.ToInt32(Console.ReadLine()))
                {
                    case 1: RunAlgorithm(RunModule.Elements, RunMode.API); break;
                    case 2: RunAlgorithm(RunModule.Propertyes, RunMode.API); break;
                    case 3: RunAlgorithm(RunModule.Elements, RunMode.ImportFile); break;
                    case 4: RunAlgorithm(RunModule.Propertyes, RunMode.ImportFile); break;
                    case 5: ChangesRollback.AllRolbacks(); break;
                    default: break;
                }
                Console.WriteLine("Работа завершена. Перейти к режимам загрузки?");
                if (!CommonCode.UserValidationPlusOrMinus("да", "выйти"))
                    break;
            }
            Console.WriteLine("Работа завершена окончательно.");
        }
        private static void AppConfiguration()
        {
            Console.WriteLine("Начали конфигурацию приложения");
            AppBase.Initialize();
            TcsBase = new TCSBase();
            PolynomBase = new PolynomBase();
            Console.WriteLine("Завершили конфигурацию приложения");
        }
    }
}

// Перед началом работы нужно изменить системный разделитель excel с , на .
