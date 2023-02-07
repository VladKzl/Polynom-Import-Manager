using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static TCS_Polynom_data_actualiser.AppSettings;
using static TCS_Polynom_data_actualiser.PolynomBase;
using static TCS_Polynom_data_actualiser.ElementsActualisation;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TCS_Polynom_data_actualiser
{
    public class Program
    {
        public static AppSettings Settings { get; set; }
        public static TCSBase TcsBase { get; set;  }
        public static PolynomBase PolynomBase { get; set; }
        public static ElementsActualisation ElementsActualisation { get; set; }

        [STAThread]
        static void Main(string[] args)
        {
            Console.WriteLine("Начали конфигурацию приложения");
            Settings = new AppSettings();
            TcsBase = new TCSBase();
            PolynomBase = new PolynomBase();
            Console.WriteLine("Завершили конфигурацию приложения");

            Console.WriteLine
                (
                    "Выберите режим работы. Введите число от 1 до 3.\n" +
                    "1 - Импорт элементов ТКС в Полином через API\n" +
                    "2 - Импорт элементов ТКС в Полином через обменный файл\n" +
                    "3 - Откат изменений"
                );
            switch (Convert.ToInt32(Console.ReadLine()))
            {
                case 1: Algorithm.RunAlgorithm(1); break;
                case 2: Algorithm.RunAlgorithm(2); break;
                case 3: ChangesRollback.AllRolbacks(); break;
                default: break;
            }
            Console.WriteLine("Работа завершена");
            Console.ReadLine();
        }
    }
}
