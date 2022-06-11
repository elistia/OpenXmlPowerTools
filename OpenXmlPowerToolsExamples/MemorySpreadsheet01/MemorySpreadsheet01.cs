using DocumentFormat.OpenXml.Packaging;
using Codeuctivity.OpenXmlPowerTools;
using System;
using System.Data;
using System.IO;

namespace MemorySpreadsheet01
{
    internal class MemorySpreadsheetExample
    {
        private static void Main(string[] args)
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            DataTable d = new DataTable();
            d.Columns.Add("String", typeof(String));
            d.Columns.Add("DateTime", typeof(DateTime));
            d.Columns.Add("Bool", typeof(bool));
            d.Columns.Add("Integer", typeof(Int32));
            d.Columns.Add("Float", typeof(float));
            d.Columns.Add("Double", typeof(double));
            d.Columns.Add("Decimal", typeof(decimal));

            d.Rows.Add("Date", new DateTime(2021, 6, 29), true, (Int32)42, 5.93F, Math.PI, (decimal)4.99);
            d.Rows.Add("Date and Time", new DateTime(2020, 10, 01, 13, 55, 1), false);

            MemoryStream stream = new MemoryStream();
            using (OpenXmlMemoryStreamDocument streamDoc = OpenXmlMemoryStreamDocument.CreateSpreadsheetDocument())
            {
                using (SpreadsheetDocument doc = streamDoc.GetSpreadsheetDocument())
                {
                    string sheetName = "Test";
                    WorksheetAccessor.CreateDefaultStyles(doc);
                    WorksheetPart sheet = WorksheetAccessor.AddWorksheet(doc, sheetName);
                    MemorySpreadsheet ms = new MemorySpreadsheet();

                    for (int h = 0; h < d.Columns.Count; h++)
                    {
                        ms.SetCellValue(1, h + 1, d.Columns[h].ColumnName, WorksheetAccessor.GetStyleIndex(doc, "Total"));
                    }

                    int rowIndex = 2;
                    foreach (DataRow row in d.Rows)
                    {
                        for (int c = 0; c < d.Columns.Count; c++)
                        {
                            ms.SetCellValue(rowIndex, c + 1, row[c]);
                        }
                        rowIndex++;
                    }

                    rowIndex++;

                    int dateStyle = WorksheetAccessor.GetStyleIndex(doc, "Normal", (int)WorksheetAccessor.NumberFormats.Date);

                    ms.SetCellValue(rowIndex, 1, "Date Only");
                    ms.SetCellValue(rowIndex, 2, new DateTime(2022, 06, 11, 19, 42, 42), dateStyle);

                    rowIndex++;

                    int dateTimeStyle = WorksheetAccessor.GetStyleIndex(doc, "Normal", (int)WorksheetAccessor.NumberFormats.DateTime);

                    ms.SetCellValue(rowIndex, 1, "Datetime");
                    ms.SetCellValue(rowIndex, 2, new DateTime(2022, 06, 11, 19, 42, 42), dateTimeStyle);

                    rowIndex++;

                    int currencyStyle = WorksheetAccessor.GetStyleIndex(doc, "Normal", (int)WorksheetAccessor.NumberFormats.CurrencyDecimal);

                    ms.SetCellValue(rowIndex, 1, "Currency");
                    ms.SetCellValue(rowIndex, 2, 102.66666F, currencyStyle);

                    rowIndex++;

                    WorksheetAccessor.CellAlignment centreAlignment = new WorksheetAccessor.CellAlignment();
                    centreAlignment.HorizontalAlignment = WorksheetAccessor.CellAlignment.Horizontal.Center;

                    WorksheetAccessor.PatternFill fillYellow = new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid, new WorksheetAccessor.ColorInfo("FFFF00"), new WorksheetAccessor.ColorInfo("FFFF00"));
                    int yellowFillIndex = WorksheetAccessor.GetFillIndex(doc, fillYellow);

                    int percentStyle = WorksheetAccessor.GetStyleIndex(doc, (int)WorksheetAccessor.NumberFormats.PercentDecimal, 0, yellowFillIndex, 0, centreAlignment, false, false);

                    ms.SetCellValue(rowIndex, 1, "Percent");
                    ms.SetCellValue(rowIndex, 2, 0.66666F, percentStyle);


                    WorksheetAccessor.SetSheetContents(sheet, ms);
                }

                streamDoc.GetModifiedSmlDocument().WriteByteArray(stream);
                stream.Position = 0;
            }

            FileStream file = new FileStream(Path.Combine(tempDi.FullName, "Test1.xlsx"), FileMode.Create, FileAccess.Write);
            stream.WriteTo(file);
            file.Close();
            stream.Close();
        }
    }
}
