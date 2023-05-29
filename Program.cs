using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Polynom_Import_Manager.PolynomBase;
using static Polynom_Import_Manager.ElementsActualisation;
using DocumentFormat.OpenXml.Spreadsheet;
using static Polynom_Import_Manager.Algorithm;
using Ascon.Polynom.Api;
using System.Reflection;

namespace Polynom_Import_Manager
{
    public class Program
    {
        public static string settingsWorkbookPath = "D:\\ascon_obmen\\kozlov_vi\\Полином\\Приложения\\Polynom Import Manager\\Настройки.xlsx";
        public static string appDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location).Replace("\\bin\\Debug", "");
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
                    "Через Api\n" +
                    "   1 - Элементы и их группы\n" +
                    "   2 - Свойства и их группы\n" +
                    "Через файл импорта\n" +
                    "   3 - Элементы и их группы\n" +
                    "   4 - Свойства и их группы\n" +
                    "Прочее\n" +
                    "   5 - Откат изменений"
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
