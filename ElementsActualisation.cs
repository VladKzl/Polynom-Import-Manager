using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using static TCS_Polynom_data_actualiser.AppSettings;

namespace TCS_Polynom_data_actualiser
{
    public class ElementsActualisation
    {
        public static XLWorkbook NewElementsActualisationWorkBook = new XLWorkbook();
        public static Lazy<XLWorkbook> CurrentElementsActualisationWorkBook = new Lazy<XLWorkbook>(new Func<XLWorkbook>(() => new XLWorkbook(ElementsActualisationSettings.FilePath)));
        public static void FillDocumentData()
        {
            if (File.Exists(ElementsActualisationSettings.FilePath))
                File.Move(ElementsActualisationSettings.FilePath, ElementsActualisationSettings.ArchivePath);

            foreach (string type in CommonSettings.Types)
            {
                var sheet = NewElementsActualisationWorkBook.AddWorksheet(type);
                sheet.Cell(1, "A").Value = "В TCS есть в Полиноме есть";
                sheet.Cell(1, "B").Value = "В TCS есть в Полиноме нет";
                sheet.Cell(1, "C").Value = "Группы TCS(типы)";
                sheet.Cell(1, "D").Value = "Что делать с \"B\"";
                sheet.Cell(1, "E").Value = "Текущее имя в Полиноме";
                sheet.Cell(1, "G").Value = "В TCS нет в Полиноме есть";
                sheet.Cell(1, "H").Value = $"Все {type} в Полином";
                sheet.Cell(1, "I").Value = $"Все {type} в TCS";

                sheet.Cell(2, "A").FormulaA1 = "=СЧЁТЗ(A3:A100000)";
                sheet.Cell(2, "B").FormulaA1 = "=СЧЁТЗ(B3:B100000)";
                sheet.Cell(2, "D").FormulaA1 = "Создать/Переименовать";
                sheet.Cell(2, "G").FormulaA1 = "=СЧЁТЗ(G3:G100000)";
                sheet.Cell(2, "H").FormulaA1 = "=СЧЁТЗ(H3:H100000)";
                sheet.Cell(2, "I").FormulaA1 = "=СЧЁТЗ(I3:I100000)";

                sheet.Row(1).Style.Font.FontColor = XLColor.CarrotOrange;
                sheet.Row(2).Style.Font.FontColor = XLColor.Gray;
            }
            foreach (string type in CommonSettings.Types)
            {
                var workSheet = NewElementsActualisationWorkBook.Worksheets.Worksheet(type);

                var tcsIntersectPolynom = TCSIntersectPolynom(type).GroupBy(x => x).Select(g => g.First()).ToList(); // Проверить!!!
                var tCSExceptPolynom = TCSExceptPolynom(type).GroupBy(x => x).Select(g => g.First()).ToList();
                var polynomExceptTCS = PolynomExceptTCS(type).GroupBy(x => x).Select(g => g.First()).ToList();
                var allTCS = AllTCS(type);
                var allPolynom = AllPolynom(type);

                for (int i = 0; i < tcsIntersectPolynom.Count(); i++)
                {
                    int _i = i + 3;
                    workSheet.Cell(_i, "A").Value = tcsIntersectPolynom[i];
                }
                int _rowNum = 3;
                for (int i = 0; i < tCSExceptPolynom.Count(); i++)
                {
                    int rowNum = i + _rowNum;

                    if(TCSBase.ElementsNamesAndGroup.Where(x => x.First() == tCSExceptPolynom[i]).Count() > 1)
                    {
                        foreach(var groupByName in TCSBase.ElementsNamesAndGroup.Where(x => x.First() == tCSExceptPolynom[i]).Select(x => x[1]))
                        {
                            workSheet.Cell(rowNum, "B").Value = tCSExceptPolynom[i];
                            workSheet.Cell(rowNum, "C").Value = groupByName;
                            workSheet.Cell(rowNum, "D").Value = "Создать";

                            _rowNum++;
                            rowNum = i + _rowNum;
                        }
                        _rowNum--;
                    }
                    else
                    {
                        string group = TCSBase.ElementsNamesAndGroup.Where(x => x.First() == tCSExceptPolynom[i]).Select(x => x[1]).First();
                        workSheet.Cell(rowNum, "B").Value = tCSExceptPolynom[i];
                        workSheet.Cell(rowNum, "C").Value = group;
                        workSheet.Cell(rowNum, "D").Value = "Создать";
                    }
                }
                for (int i = 0; i < polynomExceptTCS.Count(); i++)
                {
                    int _i = i + 3;
                    workSheet.Cell(_i, "G").Value = polynomExceptTCS[i];
                }
                for (int i = 0; i < allPolynom.Count(); i++)
                {
                    int _i = i + 3;
                    workSheet.Cell(_i, "H").Value = allPolynom[i];
                }
                for (int i = 0; i < allTCS.Count(); i++)
                {
                    int _i = i + 3;
                    workSheet.Cell(_i, "I").Value = allTCS[i];
                }
            }
            NewElementsActualisationWorkBook.SaveAs(ElementsActualisationSettings.FilePath);
        }
        private static List<string> TCSIntersectPolynom(string type)
        {
            return TCSBase.NamesByType[type].Intersect(PolynomBase.ElementsNamesByType[type]).ToList();
        }
        private static List<string> TCSExceptPolynom(string type)
        {
            return TCSBase.NamesByType[type].Except(PolynomBase.ElementsNamesByType[type]).ToList();
        }
        private static List<string> PolynomExceptTCS(string type)
        {
            return PolynomBase.ElementsNamesByType[type].Except(TCSBase.NamesByType[type]).ToList();
        }
        private static List<string> AllTCS(string type)
        {
            return TCSBase.NamesByType[type];
        }
        private static List<string> AllPolynom(string type)
        {
            return PolynomBase.ElementsNamesByType[type];
        }
    }
}
