using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Polynom_Import_Manager.AppBase;

namespace Polynom_Import_Manager
{
    internal class PropertiesActualisation
    {
        private static XLWorkbook NewPropertiesActualisationWorkBook { get; set; } = new XLWorkbook();
        private static IXLWorksheet Worksheet { get; set; }
        public static void CreateAndFillElementsDocument()
        {
            if (File.Exists(PropertiesFile.FilePath))
                File.Move(PropertiesFile.FilePath, PropertiesFile.ArchivePath);
            CreateDocAndSheets();

            Worksheet = NewPropertiesActualisationWorkBook.Worksheet("Свойства");
            FillSheets();

            using (NewPropertiesActualisationWorkBook)
            {
                NewPropertiesActualisationWorkBook.SaveAs(PropertiesFile.FilePath);
            }

            void CreateDocAndSheets()
            {
                var statusSheet = NewPropertiesActualisationWorkBook.AddWorksheet("Статус");
                statusSheet.Cell(1,"A").Value = "Статус свойств";
                statusSheet.Cell(2, "A").Value = "Не актуализирован";
                statusSheet.Cell(1, "A").Style.Font.FontColor = XLColor.Green;

                statusSheet.Cell(1, "C").Value = "Статус групп";
                statusSheet.Cell(2, "C").Value = "Не актуализирован";
                statusSheet.Cell(1, "C").Style.Font.FontColor = XLColor.Green;

                var propertiesSheet = NewPropertiesActualisationWorkBook.AddWorksheet("Свойства");
                propertiesSheet.Cell(1, "A").Value = "В TCS есть в Полиноме есть";
                propertiesSheet.Cell(1, "A").CreateComment().AddText("");
                propertiesSheet.Cell(1, "B").Value = "В TCS есть в Полиноме нет";
                propertiesSheet.Cell(1, "B").CreateComment().AddText("");
                propertiesSheet.Cell(1, "C").Value = "Расположение TCS";
                propertiesSheet.Cell(1, "C").CreateComment().AddText("");
                propertiesSheet.Cell(1, "D").Value = "Предлагаемые расположения в Полином";
                propertiesSheet.Cell(1, "D").CreateComment().AddText("");
                propertiesSheet.Cell(1, "E").Value = "Номер выбранного расположения"; ;
                propertiesSheet.Cell(1, "E").CreateComment().AddText("");
                propertiesSheet.Cell(1, "F").Value = "В TCS нет в Полиноме есть";
                propertiesSheet.Cell(1, "F").CreateComment().AddText("");
                propertiesSheet.Cell(1, "G").Value = $"Все свойства в Полином";
                propertiesSheet.Cell(1, "G").CreateComment().AddText("");
                propertiesSheet.Cell(1, "H").Value = $"Все свойства в TCS";
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
            void FillSheets()
            {
                List<string> tcsIntersectPolynom = TCSIntersectPolynom().Distinct().ToList();
                List<string> tcsExceptPolynom = TCSExceptPolynom().Distinct().ToList();
                List<string> polynomExceptTCS = PolynomExceptTCS().Distinct().ToList();
                List<string> allTcs = AllTCS();
                List<string> allPolynom = AllPolynom();

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
                        Worksheet.Cell(rowNum, "A").Value = tcsIntersectPolynom[i];
                    }
                }
                void FillTcsExceptPolynomColumn()
                {
                    int _rowNum = 3;
                    for (int i = 0; i < tcsExceptPolynom.Count(); i++)
                    {
                        int rowNum = i + _rowNum;

                        var propertyAndPath = TCSBase.Propertyes.PropertiesAndPath.FindAll(x => x.prop == tcsExceptPolynom[i]);

                        if (propertyAndPath.Count == 1)
                        {
                            Worksheet.Cell(rowNum, "B").Value = propertyAndPath.First().prop;
                            Worksheet.Cell(rowNum, "C").Value = propertyAndPath.First().path;
                            continue;
                        }
                        for(int _i = 0; _i < propertyAndPath.Count; _i++)
                        {
                            Worksheet.Cell(rowNum, "B").Value = propertyAndPath[_i].prop;
                            Worksheet.Cell(rowNum, "C").Value = propertyAndPath[_i].path;
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
                        Worksheet.Cell(_i, "F").Value = polynomExceptTCS[i];
                    }
                }
                void FillAllPolynomColumn()
                {
                    for (int i = 0; i < allPolynom.Count(); i++)
                    {
                        int _i = i + 3;
                        Worksheet.Cell(_i, "G").Value = allPolynom[i];
                    }
                }
                void FillAllTCSColumn()
                {
                    for (int i = 0; i < allTcs.Count(); i++)
                    {
                        int _i = i + 3;
                        Worksheet.Cell(_i, "H").Value = allTcs[i];
                    }
                }
            }
        }
        private static List<string> TCSIntersectPolynom()
        {
            return TCSBase.Propertyes.PropertiesNames.Intersect(PolynomBase.PropertyesActualisation.PropertiesNames).ToList();
        }
        private static List<string> TCSExceptPolynom()
        {
            return TCSBase.Propertyes.PropertiesNames.Except(PolynomBase.PropertyesActualisation.PropertiesNames).ToList();
        }
        private static List<string> PolynomExceptTCS()
        {
            return PolynomBase.PropertyesActualisation.PropertiesNames.Except(TCSBase.Propertyes.PropertiesNames).ToList();
        }
        private static List<string> AllTCS()
        {
            return TCSBase.Propertyes.PropertiesNames;
        }
        private static List<string> AllPolynom()
        {
            return PolynomBase.PropertyesActualisation.PropertiesNames;
        }

    }
}
