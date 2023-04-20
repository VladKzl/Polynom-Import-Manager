using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using static TCS_Polynom_data_actualiser.AppBase;

namespace TCS_Polynom_data_actualiser
{
    public class ElementsActualisation
    {
        public static XLWorkbook NewElementsActualisationWorkBook = new XLWorkbook();
        public static void CreateAndFillElementsDocument()
        {
            if (File.Exists(ElementsActualisationSettings.FilePath))
                File.Move(ElementsActualisationSettings.FilePath, ElementsActualisationSettings.ArchivePath);
            CreateDocAndSheets();
            FillSheets();
            using (NewElementsActualisationWorkBook)
            {
                NewElementsActualisationWorkBook.SaveAs(ElementsActualisationSettings.FilePath);
            }
            void CreateDocAndSheets()
            {
                foreach (string type in CommonSettings.Types)
                {
                    var sheet = NewElementsActualisationWorkBook.AddWorksheet(type);
                    sheet.Cell(1, "A").Value = "В TCS есть в Полиноме есть";
                    sheet.Cell(1, "A").CreateComment().AddText("");
                    sheet.Cell(1, "B").Value = "В TCS есть в Полиноме нет";
                    sheet.Cell(1, "B").CreateComment().AddText("");
                    sheet.Cell(1, "C").Value = "Группы TCS(типы)";
                    sheet.Cell(1, "C").CreateComment().AddText("");
                    sheet.Cell(1, "D").Value = "Ручная валидация";
                    sheet.Cell(1, "D").CreateComment().AddText
                        (
                            "Если, элемент не найден в полиноме, но существует с другим наименованием, впишите имена в ячейки. " +
                            "Элементы столбца \"D\" будут удалены из полинома, и заменены элементами столбца \"B\""
                        );
                    sheet.Cell(1, "E").Value = "В TCS нет в Полиноме есть";
                    sheet.Cell(1, "E").CreateComment().AddText("");
                    sheet.Cell(1, "F").Value = $"Все {type} в Полином";
                    sheet.Cell(1, "F").CreateComment().AddText("");
                    sheet.Cell(1, "G").Value = $"Все {type} в TCS";
                    sheet.Cell(1, "G").CreateComment().AddText("");

                    sheet.Cell(2, "A").FormulaA1 = "=СЧЁТЗ(A3:A100000)";
                    sheet.Cell(2, "B").FormulaA1 = "=СЧЁТЗ(B3:B100000)";
                    sheet.Cell(2, "D").FormulaA1 = "=СЧЁТЗ(D3:D100000)";
                    sheet.Cell(2, "E").FormulaA1 = "=СЧЁТЗ(E3:E100000)";
                    sheet.Cell(2, "F").FormulaA1 = "=СЧЁТЗ(F3:F100000)";

                    sheet.Row(1).Style.Font.FontColor = XLColor.CarrotOrange;
                    sheet.Row(2).Style.Font.FontColor = XLColor.Gray;
                }
            }
            void FillSheets()
            {
                foreach (string type in CommonSettings.Types)
                {
                    var workSheet = NewElementsActualisationWorkBook.Worksheets.Worksheet(type);

                    List<string> tcsIntersectPolynom = TCSIntersectPolynom(type).GroupBy(x => x).Select(g => g.First()).ToList(); // Проверить!!!
                    List<string> tcsExceptPolynom = TCSExceptPolynom(type).GroupBy(x => x).Select(g => g.First()).ToList();
                    List<string> polynomExceptTCS = PolynomExceptTCS(type).GroupBy(x => x).Select(g => g.First()).ToList();
                    List<string> allTcs = AllTCS(type);
                    List<string> allPolynom = AllPolynom(type);

                    FillTcsIntersectPolynomColumn();
                    FillTcsExceptPolynomColumn();
                    FillPolynomExceptTCSColumn();
                    FillAllPolynomColumn();
                    FillAllTCSColumn();

                    void FillTcsIntersectPolynomColumn()
                    {
                        for (int i = 0; i < tcsIntersectPolynom.Count(); i++)
                        {
                            int rowNum = i + 3;
                            workSheet.Cell(rowNum, "A").Value = tcsIntersectPolynom[i];
                        }
                    }
                    void FillTcsExceptPolynomColumn()
                    {
                        int _rowNum = 3;
                        for (int i = 0; i < tcsExceptPolynom.Count(); i++)
                        {
                            int rowNum = i + _rowNum;

                            var nameAndGroupPairs = TCSBase.ElementsActualisation.ElementsNameAndGroupForAllTypes.FindAll(x => x.elementName == tcsExceptPolynom[i]);
                            if(nameAndGroupPairs.Count == 1)
                            {
                                workSheet.Cell(rowNum, "B").Value = nameAndGroupPairs.First().elementName;
                                workSheet.Cell(rowNum, "C").Value = nameAndGroupPairs.First().groupName;
                                continue;
                            }
                            foreach(var nameAndGroupPair in nameAndGroupPairs)
                            {
                                workSheet.Cell(rowNum, "B").Value = nameAndGroupPair.elementName;
                                workSheet.Cell(rowNum, "C").Value = nameAndGroupPair.groupName;
                                _rowNum++;
                                rowNum++;
                            }
                            _rowNum--;
                        }
                    }
                    void FillPolynomExceptTCSColumn()
                    {
                        for (int i = 0; i < polynomExceptTCS.Count(); i++)
                        {
                            int _i = i + 3;
                            workSheet.Cell(_i, "E").Value = polynomExceptTCS[i];
                        }
                    }
                    void FillAllPolynomColumn()
                    {
                        for (int i = 0; i < allPolynom.Count(); i++)
                        {
                            int _i = i + 3;
                            workSheet.Cell(_i, "F").Value = allPolynom[i];
                        }
                    }
                    void FillAllTCSColumn()
                    {
                        for (int i = 0; i < allTcs.Count(); i++)
                        {
                            int _i = i + 3;
                            workSheet.Cell(_i, "G").Value = allTcs[i];
                        }
                    }
                }
            }
        }
        private static List<string> TCSIntersectPolynom(string type)
        {
            return TCSBase.ElementsActualisation.ElementsNameByTcsType[type].Intersect(PolynomBase.ElementsActualisation.ElementsNameByTcsType[type]).ToList();
        }
        private static List<string> TCSExceptPolynom(string type)
        {
            return TCSBase.ElementsActualisation.ElementsNameByTcsType[type].Except(PolynomBase.ElementsActualisation.ElementsNameByTcsType[type]).ToList();
        }
        private static List<string> PolynomExceptTCS(string type)
        {
            return PolynomBase.ElementsActualisation.ElementsNameByTcsType[type].Except(TCSBase.ElementsActualisation.ElementsNameByTcsType[type]).ToList();
        }
        private static List<string> AllTCS(string type)
        {
            return TCSBase.ElementsActualisation.ElementsNameByTcsType[type];
        }
        private static List<string> AllPolynom(string type)
        {
            return PolynomBase.ElementsActualisation.ElementsNameByTcsType[type];
        }
    }
}
