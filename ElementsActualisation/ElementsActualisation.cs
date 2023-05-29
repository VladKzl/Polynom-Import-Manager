using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using static Polynom_Import_Manager.AppBase;
using Ascon.Polynom.Api;

namespace Polynom_Import_Manager
{
    public class ElementsActualisation
    {
        public static void CreateAndFillElementsDocument()
        {
            if (File.Exists(ElementsFile.FilePath))
                File.Move(ElementsFile.FilePath, ElementsFile.ArchivePath);
            ElementsFile.ReconnectWorkBook(true);

            CreateStatusSheet();
            CreateSheets();
            FillSheets();
            ElementsFile.WorkBook.SaveAs(ElementsFile.FilePath);

            void CreateStatusSheet()
            {
                var statusSheet = ElementsFile.WorkBook.AddWorksheet("Статус");
                statusSheet.Cell(1, "A").Value = "Статус элементов";
                statusSheet.Cell(2, "A").Value = "Не актуализирован";
                statusSheet.Cell(1, "A").Style.Font.FontColor = XLColor.Green;

                statusSheet.Cell(1, "C").Value = "Статус групп";
                statusSheet.Cell(2, "C").Value = "Не актуализирован";
                statusSheet.Cell(1, "C").Style.Font.FontColor = XLColor.Green;
            }
            void CreateSheets()
            {
                foreach (string type in ElementsSettings.Types)
                {
                    var propertiesSheet = ElementsFile.WorkBook.AddWorksheet(type);
                    propertiesSheet.Cell(1, "A").Value = "В TCS есть в Полиноме есть";
                    propertiesSheet.Cell(1, "A").CreateComment().AddText("");
                    propertiesSheet.Cell(1, "B").Value = "В TCS есть в Полиноме нет";
                    propertiesSheet.Cell(1, "B").CreateComment().AddText("");
                    propertiesSheet.Cell(1, "C").Value = "Расположение TCS";
                    propertiesSheet.Cell(1, "C").CreateComment().AddText("");
                    propertiesSheet.Cell(1, "D").Value = "Предлагаемые расположения в Полином";
                    propertiesSheet.Cell(1, "D").CreateComment().AddText("");
                    propertiesSheet.Cell(1, "E").Value = "Номер выбранного расположения";
                    propertiesSheet.Cell(1, "E").CreateComment().AddText("");
                    propertiesSheet.Cell(1, "F").Value = "В TCS нет в Полиноме есть";
                    propertiesSheet.Cell(1, "F").CreateComment().AddText("");
                    propertiesSheet.Cell(1, "G").Value = $"Все {type} в Полином";
                    propertiesSheet.Cell(1, "G").CreateComment().AddText("");
                    propertiesSheet.Cell(1, "H").Value = $"Все {type} в TCS";
                    propertiesSheet.Cell(1, "H").CreateComment().AddText("");

                    propertiesSheet.Cell(2, "A").FormulaA1 = "=СЧЁТЗ(A3:A100000)";
                    propertiesSheet.Cell(2, "B").FormulaA1 = "=СЧЁТЗ(B3:B100000)";
                    propertiesSheet.Cell(2, "D").FormulaA1 = "=СЧЁТЗ(B3:B100000)";
                    propertiesSheet.Cell(2, "F").FormulaA1 = "=СЧЁТЗ(F3:F100000)";
                    propertiesSheet.Cell(2, "G").FormulaA1 = "=СЧЁТЗ(F3:F100000)";
                    propertiesSheet.Cell(2, "H").FormulaA1 = "=СЧЁТЗ(F3:F100000)";

                    propertiesSheet.Row(1).Style.Font.FontColor = XLColor.Green;
                    propertiesSheet.Row(2).Style.Font.FontColor = XLColor.Gray;

                    propertiesSheet.Column("B").Style.NumberFormat.SetNumberFormatId((int)XLPredefinedFormat.Number.Text);
                }
            }
            void FillSheets()
            {
                foreach (string type in ElementsSettings.Types)
                {
                    var workSheet = ElementsFile.WorkBook.Worksheets.Worksheet(type);

                    List<string> tcsTypeElements = TCSBase.Elements.ElementAndPathByTcsType[type].Select(x => x.element).ToList();
                    List<string> polynomTypeElements = PolynomBase.Elements.ElementsByTcsType.Value[type].Select(x => x.Name).ToList();

                    List<string> tcsIntersectPolynom = tcsTypeElements.Intersect(polynomTypeElements).Distinct().ToList(); //Проверить!!!
                    List<string> tcsExceptPolynom = tcsTypeElements.Except(polynomTypeElements).Distinct().ToList();
                    List<string> polynomExceptTCS = polynomTypeElements.Except(tcsTypeElements).Distinct().ToList();
                    List<string> allTcs = tcsTypeElements;
                    List<string> allPolynom = polynomTypeElements;

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

                            var nameAndGroupPairs = TCSBase.Elements.ElementAndPathByTcsType[type].FindAll(x => x.element == tcsExceptPolynom[i]);
                            if(nameAndGroupPairs.Count == 1)
                            {
                                workSheet.Cell(rowNum, "B").Value = nameAndGroupPairs.First().element;
                                workSheet.Cell(rowNum, "C").Value = nameAndGroupPairs.First().path;
                                continue;
                            }
                            foreach(var nameAndGroupPair in nameAndGroupPairs)
                            {
                                workSheet.Cell(rowNum, "B").Value = nameAndGroupPair.element;
                                workSheet.Cell(rowNum, "C").Value = nameAndGroupPair.path;
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
                            workSheet.Cell(_i, "F").Value = polynomExceptTCS[i];
                        }
                    }
                    void FillAllPolynomColumn()
                    {
                        for (int i = 0; i < allPolynom.Count(); i++)
                        {
                            int _i = i + 3;
                            workSheet.Cell(_i, "G").Value = allPolynom[i];
                        }
                    }
                    void FillAllTCSColumn()
                    {
                        for (int i = 0; i < allTcs.Count(); i++)
                        {
                            int _i = i + 3;
                            workSheet.Cell(_i, "H").Value = allTcs[i];
                        }
                    }
                }
            }
        }
    }
}
