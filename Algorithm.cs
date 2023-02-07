using DocumentFormat.OpenXml.Office2013.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using static TCS_Polynom_data_actualiser.AppSettings;
using static TCS_Polynom_data_actualiser.GroupsActualisation;
using static TCS_Polynom_data_actualiser.ElementsActualisation;

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
            {1, new Action(() => Move1_ElementsActualisation()) },
            {2, new Action(() => Move2_ElementsValidation()) },
            {3, new Action(() => Move3_GroupsActualisation()) },
            {4, new Action(() => Move4_GroupsValidation()) },
            {5, new Action(() => Move5_GroupsCreation()) },
            {6, new Action(() => Move6_ElementsCreation()) }
        };
        public static void RunAlgorithm(int workMode)
        {
            switch (workMode)
            {
                case 1:
                    UseImportFile = false;
                    break;
                case 2:
                    UseImportFile = true;
                    break;
            }
            bool firstRun = true;
            if(firstRun && CommonSettings.CursorPosition > 1)
            {
                Console.WriteLine($"\t Вы остановились на шаге {CommonSettings.CursorPosition}. Продолжаем?\n" +
                    $"\"+\" - продолжить\n" +
                    $"\"-\" - начать заново\n");
                if (CommonCode.UserValidation() == false)
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
        private static void Move1_ElementsActualisation()
        {
            int cursorPosition = 1;
            string name = "Наполнение документа элeментами";

            Console.WriteLine($"\tШаг {cursorPosition} - \"{name}\"");

            Console.WriteLine("Получили элементы Tcs");
            PolynomBase.FillElementsPropertyes();
            Console.WriteLine("Получили элементы Polynom");
            ElementsActualisation.FillDocumentData();

            Move1IsDone = true;
        }
        private static void Move2_ElementsValidation()
        {
            string userAction;
            int cursorPosition = 2;
            string name = "Ручная валидация элементов";

            Console.WriteLine($"\tШаг {cursorPosition} - \"{name}\"");

            pointer1: Console.WriteLine
            (
                "Excel файл \"Актуализация эллементов\" создан и ожидает ручной валидации.\n" +
                "После окончания работы не забудьте в \"Настройки.xlsx\", на странице \"Актуализатор элментов\" задать статус \"Актуализирован\""
            );
            Console.WriteLine();
            Console.WriteLine
            (
                "*Ручная валидация элментов нужна для более точного переноса, так как в Полиноме уже могут сущесвовать элменты, но иметь другоие имена.\n" +
                "Если пропустить этот этап, то новые элменты запишутся поверх существующих, но с другими именами.*"
            );
            do
            {
                Console.WriteLine("Если закончили, введите +. Или закройте программу.");
                userAction = Console.ReadLine();
            }
            while (userAction != "+");

            if (ElementsActualisationSettings.Status == ActualisationStatus.Не_актуализирован)
            {
                Console.WriteLine("Статус актуализации элементов: \"Не_актуализирован\".Хотите продложить без актуализации?\n");
                if (CommonCode.UserValidation()){}
                else
                    goto pointer1;
            }
            if (ElementsActualisationSettings.Status == ActualisationStatus.Актуализирован)
                Console.WriteLine("Статус актуализации элементов: \"Актуализирован\".\n");

            Move2IsDone = true;
        }
        private static void Move3_GroupsActualisation()
        {
            int cursorPosition = 3;
            string name = "Наполнение документа группами";

            Console.WriteLine($"\tШаг {cursorPosition} - \"{name}\"");

            GroupsActualisation.FillDocumentData();

            Move3IsDone = true;
        }
        private static void Move4_GroupsValidation()
        {
            string userAction;
            int cursorPosition = 4;
            string name = "Ручная валидация групп";

            Console.WriteLine($"\tШаг {cursorPosition} - \"{name}\"");

            pointer2: Console.WriteLine
            (
                "Excel файл \"Актуализация групп\" создан и ожидает ручной валидации.\n" +
                "После окончания работы не забудьте в \"Настройки.xlsx\", на странице \"Актуализатор групп\" задать статус \"Актуализирован\""
            );
            Console.WriteLine();
            Console.WriteLine
            (
                "***\n" +
                "Ручная валидация групп нужна для более точно распределения элементов по группам.\n" +
                "Программа уже проверила совпадающие с tcs(столбец \"A\") группы в полиноме и внесла наиманования групп в колонку \"Polynom\"(Столбец \"B\").\n" +
                "Пустые поля колокни \"B\" не нашли соответсвия с tcs, значит в дальнейшем в полиноме будут созданы группы колоки \"A\", или заполните поля вручную.\n" +
                "***"
            );
            do
            {
                Console.WriteLine("Если закончили, введите +. Или закройте программу.");
                userAction = Console.ReadLine();
            }
            while (userAction != "+");

            if (GroupsActualisationSettings.Status == ActualisationStatus.Не_актуализирован)
            {
                Console.WriteLine("Статус актуализации групп: \"Не_актуализирован\".Хотите продложить без актуализации?\n");
                if (CommonCode.UserValidation()) { }
                else
                    goto pointer2;
            }
            if (GroupsActualisationSettings.Status == ActualisationStatus.Актуализирован)
                Console.WriteLine("Статус актуализации элементов: \"Актуализирован\".\n");

            Move4IsDone = true;
        }
        private static void Move5_GroupsCreation()
        {
            string userAction;
            int cursorPosition = 5;
            string name = "Создание групп";

            Console.WriteLine($"\tШаг {cursorPosition} - \"{name}\"");

            PolynomObjectsCreation.GroupsCreation();

            Console.WriteLine("Заполнили справочники группами");

            Move5IsDone = true;
        }
        private static void Move6_ElementsCreation()
        {
            string userAction;
            int cursorPosition = 6;
            string name = "Создание элементов в группах";
            if (Move5IsDone)
            {
                Console.WriteLine($"\tШаг {cursorPosition} - \"{name}\"");

                PolynomObjectsCreation.ElementsCreation(UseImportFile);

                Console.WriteLine("Изменения применены");
                Console.WriteLine("Создали все элементы.");

                Move5IsDone = true;
            }
            else
            {
                Console.WriteLine($"{name} невозможно без шага 5 - Создание групп. Создаем группы?");
                if (CommonCode.UserValidation())
                    CommonSettings.CursorPosition = CommonSettings.CursorPosition - 2;
                else
                    CommonSettings.CursorPosition = 1;
            }
        }
    }
}
