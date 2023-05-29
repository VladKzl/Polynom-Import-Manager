using DocumentFormat.OpenXml.Office2013.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using static Polynom_Import_Manager.AppBase;

namespace Polynom_Import_Manager
{
    public static class Algorithm
    {
        public enum RunModule
        {
            Elements,
            Propertyes,
            Concepts
        }
        public enum RunMode
        {
            API,
            ImportFile
        }
        public static RunMode runMode;
        public static RunModule targetModule;
        public static void RunAlgorithm(RunModule _targetModule, RunMode _runMode)
        {
            targetModule = _targetModule;
            runMode = _runMode;
            bool firstRun = true;
            RunAlgorithmMoves();

            void RunAlgorithmMoves()
            {
                if (CommonSettings.CursorPosition > 1)
                {
                    Console.WriteLine($"\t Вы остановились на шаге {CommonSettings.CursorPosition}. Продолжаем?\n");
                    if (!CommonCode.UserValidationPlusOrMinus("продолжить", "начать c первого шага"))
                        CommonSettings.CursorPosition = 1;
                }
                if (targetModule == RunModule.Elements)
                    RunMoves(ElemetnsAlgorithm.Moves);
                if (targetModule == RunModule.Propertyes)
                    RunMoves(PropertyesAlgorithm.Moves);
                /*if (_targetModule == RunModule.Concepts)
                    RunMoves(ElemetnsAlgorithm.Moves);*/
            }
            void RunMoves(Dictionary<int, Action> moves)
            {
                do
                {
                    if (!firstRun)
                        CommonSettings.CursorPosition += 1;
                    firstRun = false;
                    Action move = moves[CommonSettings.CursorPosition];
                    move.Invoke();
                }
                while (CommonSettings.CursorPosition != moves.Count());
                CommonSettings.CursorPosition = 1;
            }
        }
        public static class ElemetnsAlgorithm
        {
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
                {5, new Action(() => Move5_GroupsAndElementsCreation()) }
            };
            private static void Move1_CreateAndFillElementsDocument()
            {
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
                    "После окончания работы c файлом на странице \"Статус\" задайте статус элементов \"Актуализирован\"."
                    );
                    Console.WriteLine("Если закончили, введите +. Или закройте программу.");
                    CommonCode.UserValidationPlus();

                    ElementsFile.ReconnectWorkBook(false);

                    if (ElementsFile.StatusElements == ActualisationStatus.Не_актуализирован)
                    {
                        Console.WriteLine("Статус актуализации элементов: \"Не_актуализирован\".Хотите продложить без актуализации?\n");
                        if (!CommonCode.UserValidationPlusOrMinus("да", "нет"))
                            goto pointer1;
                    }
                    if (ElementsFile.StatusElements == ActualisationStatus.Актуализирован)
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
                        "Группы в excel файле \"Актуализация элементов\" ожидают ручной валидации.\n" +
                        "После окончания работы c файлом на странице \"Статус\" задайте статус групп \"Актуализирован\"."
                    );
                    Console.WriteLine("Если закончили, введите +. Или закройте программу.");
                    CommonCode.UserValidationPlus();

                    if (ElementsFile.StatusGroups == ActualisationStatus.Не_актуализирован)
                    {
                        Console.WriteLine("Статус актуализации групп: \"Не_актуализирован\".Хотите продложить без актуализации?\n");
                        if (!CommonCode.UserValidationPlusOrMinus("да", "нет"))
                            goto pointer2;
                    }
                    if (ElementsFile.StatusGroups == ActualisationStatus.Актуализирован)
                        Console.WriteLine("Статус актуализации элементов: \"Актуализирован\".\n");
                }
            }
            private static void Move5_GroupsAndElementsCreation()
            {
                string name = "Загрузка групп";
                int cursorPosition = 5;
                Console.WriteLine($"\tШаг {cursorPosition} - \"{name}\"");

                if (runMode == RunMode.API)
                    PolynomElementsCreation.UseApi();
                if (runMode == RunMode.ImportFile)
                    PolynomElementsCreation.UseImportFile();

                Console.WriteLine("Загрузили группы");
                Move5IsDone = true;
            }
        }
        public static class PropertyesAlgorithm
        {
            private static bool Move1IsDone = false;
            private static bool Move2IsDone = false;
            private static bool Move3IsDone = false;
            private static bool Move4IsDone = false;
            private static bool Move5IsDone = false;
            private static bool Move6IsDone = false;
            public static Dictionary<int, Action> Moves { get; set; } = new Dictionary<int, Action>()
            {
                {1, new Action(() => Move1_CreateAndFillPropertyesDocument()) },
                {2, new Action(() => Move2_ManualPropertyesValidation()) },
                {3, new Action(() => Move3_FillGroups()) },
                {4, new Action(() => Move4_GroupsValidation()) },
                {5, new Action(() => Move5_GroupsAndPropertiesCreation()) }
            };
            private static void Move1_CreateAndFillPropertyesDocument()
            {
                string name = "Наполение документа свойствами";
                int cursorPosition = 1;

                Console.WriteLine($"\tШаг {cursorPosition} - \"{name}\"");
                PropertiesActualisation.CreateAndFillElementsDocument();

                Console.WriteLine($"Наполнили");
                Move1IsDone = true;
            }
            private static void Move2_ManualPropertyesValidation()
            {
                string name = "Ручная валидация свойств";
                int cursorPosition = 2;

                Console.WriteLine($"\tШаг {cursorPosition} - \"{name}\"");
                PropertiesFile.ReconnectWorkBook();
                Validation();

                Move2IsDone = true;

                void Validation()
                {
                    pointer1: Console.WriteLine
                    (
                    "Excel файл \"Актуализация свойств\" создан и ожидает ручной валидации.\n" +
                    "После окончания работы не забудьте на странице \"Статус\" задать статус элментов \"Актуализирован\"."
                    );
                    Console.WriteLine("Если закончили, введите +. Или закройте программу.");
                    CommonCode.UserValidationPlus();

                    if (PropertiesFile.StatusProperties == ActualisationStatus.Не_актуализирован)
                    {
                        Console.WriteLine("Статус актуализации свойств: \"Не актуализирован\".Хотите продложить без актуализации?\n");
                        if (!CommonCode.UserValidationPlusOrMinus("да", "нет"))
                            goto pointer1;
                    }
                    if (PropertiesFile.StatusProperties == ActualisationStatus.Актуализирован)
                        Console.WriteLine("Статус актуализации свойств: \"Актуализирован\".\n");
                }
            }
            private static void Move3_FillGroups()
            {
                string name = "Наполнение документа группами свойств";
                int cursorPosition = 3;

                Console.WriteLine($"\tШаг {cursorPosition} - \"{name}\"");
                Validation();
                PropertyesGroupsActualisation.ActualiseGroups();

                Move3IsDone = true;

                void Validation()
                {
                    if (Move1IsDone == false)
                    {
                        Console.WriteLine("Вы не запускали шаг 1 - \"Наполение документа свойствами\". Возможно, информация в файле \"Актуализация свойств\" устарела.");
                        if (!CommonCode.UserValidationPlusOrMinus("продолжить", "завершить"))
                            throw new Exception();
                    }
                }
            }
            private static void Move4_GroupsValidation()
            {
                string name = "Ручная валидация групп свойств";
                int cursorPosition = 4;

                Console.WriteLine($"\tШаг {cursorPosition} - \"{name}\"");
                Validation();

                Move4IsDone = true;

                void Validation()
                {
                    pointer2: Console.WriteLine
                    (
                        "Группы для свойств подготовлены и ожидают ручной валидации.\n" +
                        "После окончания работы не забудьте на странице \"Статус\" задать статус групп \"Актуализирован\"."
                    );
                    Console.WriteLine("Если закончили, введите +. Или закройте программу.");
                    CommonCode.UserValidationPlus();

                    if (PropertiesFile.StatusGroups == ActualisationStatus.Не_актуализирован)
                    {
                        Console.WriteLine("Статус актуализации групп свойств: \"Не актуализирован\".Хотите продложить без актуализации?\n");
                        if (!CommonCode.UserValidationPlusOrMinus("да", "нет"))
                            goto pointer2;
                    }
                    if (PropertiesFile.StatusGroups == ActualisationStatus.Актуализирован)
                        Console.WriteLine("Статус актуализации групп свойств: \"Актуализирован\".\n");
                }
            }
            private static void Move5_GroupsAndPropertiesCreation()
            {
                string name = "Загрузка групп и свойств";
                int cursorPosition = 5;

                Console.WriteLine($"\tШаг {cursorPosition} - \"{name}\"");

                if (runMode == RunMode.API)
                    PolynomPropertiesCreation.UseApi();
                if (runMode == RunMode.ImportFile)
                    PolynomPropertiesCreation.UseImportFile();

                Console.WriteLine("Загрузили группы и свойства");
                Move5IsDone = true;
            }
        }
    }
}
