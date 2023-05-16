using Ascon.Polynom.Api;
using Ascon.Vertical.Application.Configuration;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using static TCS_Polynom_data_actualiser.AppBase;

namespace TCS_Polynom_data_actualiser
{
    public class ChangesRollback
    {
        public static void AllRolbacks()
        {
            Console.WriteLine($"Выберите изменения какого документа откатываем. Ведите 1 или 2. \n" +
                $"1 - Откат изменений текущего документа \n" +
                $"2 - Откат изменений документа из архива");
            switch(Console.ReadLine())
            {
                case "1":
                    CurrentDocumentRollback();
                    break;
                case "2":
                    ArchiveDocumentRollback();
                    break;
                default:
                    Console.WriteLine("Не вернный выбор.");
                    break;
            }
        }
        public static void CurrentDocumentRollback()
        {
            Console.WriteLine("Начали откат");
            ITransaction transaction = PolynomBase.Session.Objects.StartTransaction();
            foreach (string type in ElementsFileSettings.Types)
            {
                var workSheet = ElementsActualisationWorkBook.Value.Worksheet(type);
                // Так как исапольузется Range то буквы колонок меняются на A,B,C и тд.
                var elementsRows = workSheet.Range(workSheet.Cell(3, "B"), workSheet.Column("C").LastCellUsed()).Rows();
                
                foreach (var elementRow in elementsRows)
                {
                    string elementName = elementRow.Cell("A").Value.ToString();
                    ElementsDeletion(elementName);
                    CommonCode.GetPercent(elementsRows.Count(), elementRow.RowNumber(), type + "-");
                }
                string groupName = $"Нераспределенные {type}";
                RootGroupsDeletion(groupName);
            }
            transaction.Commit();
        }
        public static void ArchiveDocumentRollback()
        {
            Console.WriteLine($"Введите путь до файла элементов из архива\n" +
                @"К примеру: D:\ascon_obmen\kozlov_vi\Полином\Приложения\TCS_Polynom_data_actualiser\Архив Актуализация элементов\Актуализация элементов-1191953373.xlsx");
            string path = Console.ReadLine();

            Console.WriteLine("Начали откат");
            ITransaction transaction = PolynomBase.Session.Objects.StartTransaction();
            foreach (string type in ElementsFileSettings.Types)
            {
                XLWorkbook archiveBook = new XLWorkbook(path);
                var workSheet = archiveBook.Worksheet(type);
                if(workSheet != null)
                {
                    // Так как исапольузется Range то буквы колонок меняются на A,B,C и тд.
                    var elementsRows = workSheet.Range(workSheet.Cell(3, "B"), workSheet.Column("C").LastCellUsed()).Rows();

                    foreach (var elementRow in elementsRows)
                    {
                        string elementName = elementRow.Cell("A").Value.ToString();
                        ElementsDeletion(elementName);
                    }
                    string groupName = $"Нераспределенные {type}";
                    RootGroupsDeletion(groupName);
                }
            }
            transaction.Commit();
        }
        private static void ElementsDeletion(string elementName)
        {
            List<IElement> elements;
            if (PolynomBase.TrySearchElementsInAllReferences(elementName, out elements))
            {
                foreach(var element in elements)
                    element.Delete();
            }
            else
            {
                Console.WriteLine($"Элементы {elementName} не удалось удалить, так как они не были найдены.");
            }
        }
        private static void RootGroupsDeletion(string groupName)
        {
            List<IGroup> groups;
            if (PolynomBase.TrySearchGroupsInAllReferences(groupName, out groups))
            {
                foreach (var group in groups)
                    group.Delete();
            }
            else
            { 
                Console.WriteLine($"Root группу {groupName} не удалось удалить, так как она не была найдена.");
            }
        }
    }
}
