using DocumentFormat.OpenXml.Office2013.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using static TCS_Polynom_data_actualiser.AppBase;

namespace TCS_Polynom_data_actualiser
{
    public class Algorithm
    {
        public static bool UseImportFile = false;
        private static bool Move1IsDone = false;
        private static bool Move2IsDone = false;
        private static bool Move3IsDone = false;
        private static bool Move4IsDone = false;
        private static bool Move5IsDone = false;
        private static bool Move6IsDone = false;
        public static Dictionary<int, Action> Moves { get; set; } = new Dictionary<int, Action>() 
        {
            {1, new Action(() => Move1_CreateAndFillElementsDocument()) },
            {2, new Action(() => Move2_ManualElementsValidation()) },
            {3, new Action(() => Move3_CreateAndFillGroupsDocument()) },
            {4, new Action(() => Move4_GroupsValidation()) },
            {5, new Action(() => Move5_GroupsCreation()) },
            {6, new Action(() => Move6_ElementsCreation()) }
        };
        public static void RunAlgorithm(bool workMode)
        {
            UseImportFile = workMode;
            bool firstRun = true;

            if(CommonSettings.CursorPosition > 1)
            {
                Console.WriteLine($"\t Вы остановились на шаге {CommonSettings.CursorPosition}. Продолжаем?\n");
                if (!CommonCode.UserValidationPlusOrMinus("продолжить", "начать c первого шага"))
                    CommonSettings.CursorPosition = 1;
            }
            do
            {
                if (!firstRun)
                    CommonSettings.CursorPosition += 1;
                firstRun = false;
                Action move = Moves[CommonSettings.CursorPosition];
                move.Invoke();
            }
            while (CommonSettings.CursorPosition != Moves.Count());
            CommonSettings.CursorPosition = 1;
        }
        private static void Move1_CreateAndFillElementsDocument()
        {
            PolynomBase.ElementsActualisation.Initialize();

            string name = "Наполение документа элементами";
            int cursorPosition = 1;

            Console.WriteLine($"\tШаг {cursorPosition} - \"{name}\"");
            ElementsActualisation.CreateAndFillElementsDocument();

            Console.WriteLine($"Наполнили");
            Move1IsDone = true;
        }
        private static void Move2_ManualElementsValidation()
        {
            string name = "Ручная валидация элементов";
            int cursorPosition = 2;

            Console.WriteLine($"\tШаг {cursorPosition} - \"{name}\"");
            Validation();

            Move2IsDone = true;

            void Validation()
            {
                pointer1: Console.WriteLine
                (
                "Excel файл \"Актуализация элементов\" создан и ожидает ручной валидации.\n" +
                "После окончания работы не забудьте в \"Настройки.xlsx\", на странице \"Актуализатор элeментов\" задать статус \"Актуализирован\"."
                );
                Console.WriteLine("Если закончили, введите +. Или закройте программу.");
                CommonCode.UserValidationPlus();

                if (ElementsActualisationSettings.Status == ActualisationStatus.Не_актуализирован)
                {
                    Console.WriteLine("Статус актуализации элементов: \"Не_актуализирован\".Хотите продложить без актуализации?\n");
                    if (!CommonCode.UserValidationPlusOrMinus("да", "нет"))
                        goto pointer1;
                }
                if (ElementsActualisationSettings.Status == ActualisationStatus.Актуализирован)
                    Console.WriteLine("Статус актуализации элементов: \"Актуализирован\".\n");
            }
        }
        private static void Move3_CreateAndFillGroupsDocument()
        {
            string name = "Наполнение документа группами";
            int cursorPosition = 3;

            Console.WriteLine($"\tШаг {cursorPosition} - \"{name}\"");
            Validation();
            GroupsActualisation.CreateAndFillGroupsDocument();

            Move3IsDone = true;

            void Validation()
            {
                if (Move1IsDone == false)
                {
                    Console.WriteLine("Вы не запускали шаг 1 - \"Наполение документа элементами\". Возможно, информация в файле \"Актуализация элементов\" устарела.");
                    if (!CommonCode.UserValidationPlusOrMinus("продолжить", "завершить"))
                        throw new Exception();
                }
            }
        }
        private static void Move4_GroupsValidation()
        {
            string name = "Ручная валидация групп";
            int cursorPosition = 4;

            Console.WriteLine($"\tШаг {cursorPosition} - \"{name}\"");
            Validation();

            Move4IsDone = true;

            void Validation()
            {
                pointer2: Console.WriteLine
                (
                    "Excel файл \"Актуализация групп\" создан и ожидает ручной валидации.\n" +
                    "После окончания работы не забудьте в \"Настройки.xlsx\", на странице \"Актуализатор групп\" задать статус \"Актуализирован\"."
                );
                Console.WriteLine("Если закончили, введите +. Или закройте программу.");
                CommonCode.UserValidationPlus();

                if (GroupsActualisationSettings.Status == ActualisationStatus.Не_актуализирован)
                {
                    Console.WriteLine("Статус актуализации групп: \"Не_актуализирован\".Хотите продложить без актуализации?\n");
                    if (!CommonCode.UserValidationPlusOrMinus("да", "нет"))
                        goto pointer2;
                }
                if (GroupsActualisationSettings.Status == ActualisationStatus.Актуализирован)
                    Console.WriteLine("Статус актуализации элементов: \"Актуализирован\".\n");
            }
        }
        private static void Move5_GroupsCreation()
        {
            string name = "Загрузка групп";
            int cursorPosition = 5;

            Console.WriteLine($"\tШаг {cursorPosition} - \"{name}\"");
            PolynomObjectsCreation.GroupsCreation();

            Console.WriteLine("Загрузили группы");
            Move5IsDone = true;
        }
        private static void Move6_ElementsCreation()
        {
            string name = "Загрузка элементов в группы";
            int cursorPosition = 6;

            if (!Move5IsDone)
            {
                Console.WriteLine($"{name} невозможно без шага 5 - \"Загрузка групп\".");
                CommonSettings.CursorPosition = 4;
                return;
            }
            Console.WriteLine($"\tШаг {cursorPosition} - \"{name}\"");

            if (!UseImportFile)
                PolynomObjectsCreation.ElementsCreation();
            if(UseImportFile)
                PolynomObjectsCreation.ElementsCreationImportFile();

            Console.WriteLine("Загрузили элементы");
            Move5IsDone = true;
        }
    }
}
