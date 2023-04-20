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

namespace TCS_Polynom_data_actualiser
{
    public class Program
    {
        public static string settingsWorkbookPath = "D:\\ascon_obmen\\kozlov_vi\\Полином\\Приложения\\TCS_Polynom_data_actualiser\\Настройки.xlsx";
        public static TCSBase TcsBase { get; set;  }
        public static PolynomBase PolynomBase { get; set; }

        [STAThread]
        static void Main(string[] args)
        {
            AppConfiguration();
            Console.WriteLine
            (
                "\nВыберите режим импорта в Полином. Введите число от 1 до 3.\n" +
                "1 - Импорт элементов и групп через API\n" +
                "2 - Импорт элементов и групп через обменный файл\n" +
                "3 - Откат изменений"
            );
            switch (Convert.ToInt32(Console.ReadLine()))
            {
                case 1: Algorithm.RunAlgorithm(false); break;
                case 2: Algorithm.RunAlgorithm(true); break;
                case 3: ChangesRollback.AllRolbacks(); break;
                default: break;
            }
            Console.WriteLine("Работа завершена");
            Console.ReadLine();
        }
        private static void AppConfiguration()
        {
            Console.WriteLine(" Начали конфигурацию приложения");
            AppBase.Initialize();
            TcsBase = new TCSBase();
            PolynomBase = new PolynomBase();
            Console.WriteLine("Завершили конфигурацию приложения");
        }
    }
}
