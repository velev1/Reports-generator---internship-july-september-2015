using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;

namespace SS_Reports.File_Managers
{
    class NewFile
    {
        BackgroundWorker newFileWorker;
        //Worker events
        DoWorkEventArgs workerArgs;
        private string FileName { get; set; }

        private ExcelPackage newExcelFile;

        /// <summary>
        /// Creates a new (.xlsx) file, formatting and saving it.
        /// </summary>
        /// <param name="fileName">Name of the new file.</param>
        public NewFile(string fileName, object sender, DoWorkEventArgs e)
        {
            FileName = fileName;
            newFileWorker = (BackgroundWorker)sender;
            workerArgs = e;
            newExcelFile = new ExcelPackage();
            FormatAndSave();
        }
        /// <summary>
        /// Calls the format method and then saves the new file.
        /// </summary>
        private void FormatAndSave()
        {
            FileInfo file = new FileInfo(FileName);
            FormatFile();
            if (newFileWorker.CancellationPending == false)
                newExcelFile.SaveAs(file);
            else
                workerArgs.Cancel = true;
        }
        /// <summary>
        /// Formats the new file.
        /// </summary>
        private void FormatFile()
        {
            ExcelWorksheet sheet = newExcelFile.Workbook.Worksheets.Add("Review");
            ExcelWorksheet sheet1 = newExcelFile.Workbook.Worksheets.Add("Top 5 statistics");
            ExcelWorksheet sheet2 = newExcelFile.Workbook.Worksheets.Add("Overall sales by platform");
            ExcelWorksheet sheet3 = newExcelFile.Workbook.Worksheets.Add("Overall sales by game");
            FormatFirstSheet(sheet);
            FormatSecondSheet(sheet1);
            FormatThirdSheet(sheet2);
            FormatFourthSheet(sheet3);
        }

        /// <summary>
        /// Formats the first sheet
        /// </summary>
        /// <param name="ws"></param>
        private void FormatFirstSheet(ExcelWorksheet ws)
        {
            FirstSheetSelloutLabel(ws);
            FirstSheetStockLabel(ws);
            FirstSheetTotalLabel(ws);
            FirstSheetSelloutNamcoLabel(ws);
            FirstSheetStockNamcoLabel(ws);
            FirstSheetTotalNamcoLabel(ws);
            FirstSheetTableSelloutFormatting(ws);
            FirstSheetTableStockFormatting(ws);
            FirstSheetTableTotalFormatting(ws);
            FirstSheetFormatChartSellout(ws);
            FirstSheetFormatChartTotal(ws);
        }

        /// <summary>
        /// Formats the second sheet
        /// </summary>
        /// <param name="ws"></param>
        private void FormatSecondSheet(ExcelWorksheet ws)
        {
            SecondSheetTableTopFiveFormatting(ws);
        }

        /// <summary>
        /// Formats the third sheet
        /// </summary>
        /// <param name="ws"></param>
        private void FormatThirdSheet(ExcelWorksheet ws)
        {
            ThirdSheetTableStatisticsFormatting(ws);
        }

        /// <summary>
        /// Formats the fourth sheet.
        /// </summary>
        /// <param name="ws"></param>
        private void FormatFourthSheet(ExcelWorksheet ws)
        {
            FourthSheetTableStatisticsFormatting(ws);
        }
        private void FirstSheetSelloutLabel(ExcelWorksheet ws)
        {
            ApplyBorders(ws.Cells["A1:BE1"]);
            ws.Cells["A1:BE1"].Merge = true;
            ws.Cells["A1:BE1"].Value = "Sell out";
            ws.Cells["A1:BE1"].Style.Font.Bold = true;
            ws.Cells["A1:BE1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#DAEEF3");
            ws.Cells["A1:BE1"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            ws.Cells["A1:BE1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }
        
        private void FirstSheetStockLabel(ExcelWorksheet ws)
        {
            ApplyBorders(ws.Cells["BG1:BJ1"]);
            ws.Cells["BG1:BJ1"].Merge = true;
            ws.Cells["BG1:BJ1"].Value = "Stock";
            ws.Cells["BG1:BJ1"].Style.Font.Bold = true;
            ws.Cells["BG1:BJ1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#DAEEF3");
            ws.Cells["BG1:BJ1"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            ws.Cells["BG1:BJ1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }

        private void FirstSheetTotalLabel(ExcelWorksheet ws)
        {
            ApplyBorders(ws.Cells["BL1:BO1"]);
            Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#DAEEF3");
            ws.Cells["BL1:BO1"].Merge = true;
            ws.Cells["BL1:BO1"].Value = "Total";
            ws.Cells["BL1:BO1"].Style.Font.Bold = true;
            ws.Cells["BL1:BO1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["BL1:BO1"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            ws.Cells["BL1:BO1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }

        private void FirstSheetSelloutNamcoLabel(ExcelWorksheet ws)
        {
            ApplyBorders(ws.Cells["A4:A16"]);
            ws.Cells["A4:A16"].Merge = true;
            ws.Cells["A4:A16"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#DA9694");
            ws.Cells["A4:A16"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            ws.Cells["A4:A16"].Style.TextRotation = 90;
            ws.Cells["A4:A16"].Value = "Namco";
            ws.Cells["A4:A16"].Style.Font.Bold = true;
            ws.Cells["A4:A16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }

        private void FirstSheetStockNamcoLabel(ExcelWorksheet ws)
        {
            ApplyBorders(ws.Cells["BG4:BG16"]);
            ws.Cells["BG4:BG16"].Merge = true;
            ws.Cells["BG4:BG16"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#DA9694");
            ws.Cells["BG4:BG16"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            ws.Cells["BG4:BG16"].Style.TextRotation = 90;
            ws.Cells["BG4:BG16"].Value = "Namco";
            ws.Cells["BG4:BG16"].Style.Font.Bold = true;
            ws.Cells["BG4:BG16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }

        private void FirstSheetTotalNamcoLabel(ExcelWorksheet ws)
        {
            ApplyBorders(ws.Cells["BL4:BL16"]);
            ws.Cells["BL4:BL16"].Merge = true;
            ws.Cells["BL4:BL16"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#DA9694");
            ws.Cells["BL4:BL16"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            ws.Cells["BL4:BL16"].Style.TextRotation = 90;
            ws.Cells["BL4:BL16"].Value = "Namco";
            ws.Cells["BL4:BL16"].Style.Font.Bold = true;
            ws.Cells["BL4:BL16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }

        private void FirstSheetTableSelloutFormatting(ExcelWorksheet ws)
        {
            ApplyBorders(ws.Cells["B3:BE16"]);
            //"Total" row formatting
            ws.Cells["B16:BC16"].Style.Font.Bold = true;
            ws.Cells["B16:BC16"].Style.Font.Size = 12;
            //Week one is labelebed by default, otherwise the chart gets buggy.
            ws.Cells["C3"].Value = "w0";
            //Align all week labels on row 3
            ws.Cells["C3:BB3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            //Console and total labels
            int row = 4;
            foreach (Enums.OutputAbbreviations abbreviation in Enum.GetValues(typeof(Enums.OutputAbbreviations)))
            {
                ws.Cells["B" + row].Value = ws.Cells["BC" + row].Value = Enums.EnumHelper.GetDescription(abbreviation);
                row++;
            }
            ws.Cells["B16"].Value = ws.Cells["BC16"].Value = "Total";
            //"Total" row formulas
            ws.Cells["C16"].Formula = "=SUM(C4:C15)";
            ws.Cells["D16"].Formula = "=SUM(D4:D15)";
            ws.Cells["E16"].Formula = "=SUM(E4:E15)";
            ws.Cells["F16"].Formula = "=SUM(F4:F15)";
            ws.Cells["G16"].Formula = "=SUM(G4:G15)";
            ws.Cells["H16"].Formula = "=SUM(H4:H15)";
            ws.Cells["I16"].Formula = "=SUM(I4:I15)";
            ws.Cells["J16"].Formula = "=SUM(J4:J15)";
            ws.Cells["K16"].Formula = "=SUM(K4:K15)";
            ws.Cells["L16"].Formula = "=SUM(L4:L15)";
            ws.Cells["M16"].Formula = "=SUM(M4:M15)";
            ws.Cells["N16"].Formula = "=SUM(N4:N15)";
            ws.Cells["O16"].Formula = "=SUM(O4:O15)";
            ws.Cells["P16"].Formula = "=SUM(P4:P15)";
            ws.Cells["Q16"].Formula = "=SUM(Q4:Q15)";
            ws.Cells["R16"].Formula = "=SUM(R4:R15)";
            ws.Cells["S16"].Formula = "=SUM(S4:S15)";
            ws.Cells["T16"].Formula = "=SUM(T4:T15)";
            ws.Cells["U16"].Formula = "=SUM(U4:U15)";
            ws.Cells["V16"].Formula = "=SUM(V4:V15)";
            ws.Cells["W16"].Formula = "=SUM(W4:W15)";
            ws.Cells["X16"].Formula = "=SUM(X4:X15)";
            ws.Cells["Y16"].Formula = "=SUM(Y4:Y15)";
            ws.Cells["Z16"].Formula = "=SUM(Z4:Z15)";
            ws.Cells["AA16"].Formula = "=SUM(AA4:AA15)";
            ws.Cells["AB16"].Formula = "=SUM(AB4:AB15)";
            ws.Cells["AC16"].Formula = "=SUM(AC4:AC15)";
            ws.Cells["AD16"].Formula = "=SUM(AD4:AD15)";
            ws.Cells["AE16"].Formula = "=SUM(AE4:AE15)";
            ws.Cells["AF16"].Formula = "=SUM(AF4:AF15)";
            ws.Cells["AG16"].Formula = "=SUM(AG4:AG15)";
            ws.Cells["AH16"].Formula = "=SUM(AH4:AH15)";
            ws.Cells["AI16"].Formula = "=SUM(AI4:AI15)";
            ws.Cells["AJ16"].Formula = "=SUM(AJ4:AJ15)";
            ws.Cells["AK16"].Formula = "=SUM(AK4:AK15)";
            ws.Cells["AL16"].Formula = "=SUM(AL4:AL15)";
            ws.Cells["AM16"].Formula = "=SUM(AM4:AM15)";
            ws.Cells["AN16"].Formula = "=SUM(AN4:AN15)";
            ws.Cells["AO16"].Formula = "=SUM(AO4:AO15)";
            ws.Cells["AP16"].Formula = "=SUM(AP4:AP15)";
            ws.Cells["AQ16"].Formula = "=SUM(AQ4:AQ15)";
            ws.Cells["AR16"].Formula = "=SUM(AR4:AR15)";
            ws.Cells["AS16"].Formula = "=SUM(AS4:AS15)";
            ws.Cells["AT16"].Formula = "=SUM(AT4:AT15)";
            ws.Cells["AU16"].Formula = "=SUM(AU4:AU15)";
            ws.Cells["AV16"].Formula = "=SUM(AV4:AV15)";
            ws.Cells["AW16"].Formula = "=SUM(AW4:AW15)";
            ws.Cells["AX16"].Formula = "=SUM(AX4:AX15)";
            ws.Cells["AY16"].Formula = "=SUM(AY4:AY15)";
            ws.Cells["AZ16"].Formula = "=SUM(AZ4:AZ15)";
            ws.Cells["BA16"].Formula = "=SUM(BA4:BA15)";
            ws.Cells["BB16"].Formula = "=SUM(BB4:BB15)";
            //"Total pcs" cell formatting
            ws.Cells["BD3"].Style.Font.Bold = true;
            ws.Cells["BD3"].Style.Font.Size = 12;
            ws.Cells["BD3"].Value = "Total pcs";
            //"Total pcs" column formatting
            Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#92D050");
            ws.Cells["BD4:BD16"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["BD4:BD16"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            //"Total pcs" column formulas
            int firstRow = 4;
            int lastRow = 16;
            for (row = firstRow; row <=lastRow; row++)
            {
                ws.Cells["BD" + row].Formula = "SUM(C" + row + ":BB" + row + ")";
            }
            //"Total pcs"percentage formatting
            ws.Cells["BE4:BE16"].Style.Numberformat.Format = "0.00%";
            //"Total pcs" percentage formulas
            firstRow = 4;
            lastRow = 15;
            for (row = firstRow; row <=lastRow; row++)
                ws.Cells["BE" + row].Formula = "=BD" + row + "/BD16";
            ws.Cells["BE16"].Formula = "=SUM(BE4:BE15)";
        }

        private void FirstSheetTableStockFormatting(ExcelWorksheet ws)
        {
            ApplyBorders(ws.Cells["BH4:BJ16"]);
            int row = 4;
            foreach (Enums.OutputAbbreviations abbreviation in Enum.GetValues(typeof(Enums.OutputAbbreviations)))
            {
                ws.Cells["BH" + row].Value = Enums.EnumHelper.GetDescription(abbreviation);
                row++;
            }
            ws.Cells["BH16"].Value = "Total";
            ws.Cells["BI16"].Formula = "=SUM(BI4:BI15)";
            //Days in stock cell
            ws.Cells["BJ3"].Value = "Days in stock";
            ws.Cells["BJ4:BJ16"].Style.Numberformat.Format = "0";
        }

        private void FirstSheetTableTotalFormatting(ExcelWorksheet ws)
        {
            ApplyBorders(ws.Cells["BM4:BO16"]);
            int row = 4;
            foreach (Enums.OutputAbbreviations abbreviation in Enum.GetValues(typeof(Enums.OutputAbbreviations)))
            {
                ws.Cells["BM" + row].Value = Enums.EnumHelper.GetDescription(abbreviation);
                row++;
            }
            ws.Cells["BM16"].Value = "Total";
            ws.Cells["BN3"].Value = "Sales";
            ws.Cells["BN4:BN16"].Style.Numberformat.Format = "0.00%";
            int firstRow = 4;
            int lastRow = 16;
            for (row = firstRow; row <=lastRow; row++)
            {
                ws.Cells["BN" + row].Formula = "=BE" + row;
            }
            ws.Cells["BO3"].Value = "Stock";
            ws.Cells["BO4:BO16"].Style.Numberformat.Format = "0.00%";
            firstRow = 4;
            lastRow = 15;
            for (row = firstRow; row <= lastRow; row++)
            {
                ws.Cells["BO" + row].Formula = "=BI" + row + "/BI16";
            }
            ws.Cells["BO16"].Formula = "=SUM(BO4:BO15)";
        }
        private void FirstSheetFormatChartSellout(ExcelWorksheet ws)
        {
            var chart = ws.Drawings.AddChart("Sellout", eChartType.ColumnClustered);
            chart.SetPosition(16, 0, 1, 0);
            chart.Series.Add("C4:BD4", "C3:BD3").HeaderAddress = ws.Cells[4, 2];
            chart.Series.Add("C5:BD5", "C3:BD3").HeaderAddress = ws.Cells[5, 2];
            chart.Series.Add("C6:BD6", "C3:BD3").HeaderAddress = ws.Cells[6, 2];
            chart.Series.Add("C7:BD7", "C3:BD3").HeaderAddress = ws.Cells[7, 2];
            chart.Series.Add("C8:BD8", "C3:BD3").HeaderAddress = ws.Cells[8, 2];
            chart.Series.Add("C9:BD9", "C3:BD3").HeaderAddress = ws.Cells[9, 2];
            chart.Series.Add("C10:BD10", "C3:BD3").HeaderAddress = ws.Cells[10, 2];
            chart.Series.Add("C11:BD11", "C3:BD3").HeaderAddress = ws.Cells[11, 2];
            chart.Series.Add("C12:BD12", "C3:BD3").HeaderAddress = ws.Cells[12, 2];
            chart.Series.Add("C13:BD13", "C3:BD3").HeaderAddress = ws.Cells[13, 2];
            chart.Series.Add("C14:BD14", "C3:BD3").HeaderAddress = ws.Cells[14, 2];
            chart.Series.Add("C15:BD15", "C3:BD3").HeaderAddress = ws.Cells[15, 2];
            chart.Series.Add("C16:BD16", "C3:BD3").HeaderAddress = ws.Cells[16, 2];
            chart.SetPosition(340, 0);
            chart.SetSize(3648, 450);
        }

        private void FirstSheetFormatChartTotal(ExcelWorksheet ws)
        {
            var chart2 = ws.Drawings.AddChart("Total", eChartType.ColumnClustered);
            chart2.Series.Add("BN4:BN15", "BM4:BM15").HeaderAddress = ws.Cells["BN3"];
            chart2.Series.Add("BO4:BO15", "BM4:BM15").HeaderAddress = ws.Cells["BO3"];
            chart2.SetPosition(340, 3712);
            chart2.SetSize(576, 450);
        }

        private void ApplyBorders(ExcelRange range)
        {
            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
        }

        private void SecondSheetTableTopFiveFormatting(ExcelWorksheet ws)
        {
            SecondSheetTableTopLabel(ws);
            SecondSheetTableTopFiveBestSellingShopsOverall(ws);
            SecondSheetTableTopFiveBestSellingShopsLatestWeek(ws);
            SecondSheetTableTopFiveBestSellingGamesOverall(ws);
            SecondSheetTableTopFiveBestSellingGamesLatestWeek(ws);
        }
        private void SecondSheetTableTopLabel(ExcelWorksheet ws)
        {
            ApplyBorders(ws.Cells["A1:R1"]);
            ws.Cells["A1:R1"].Merge = true;
            ws.Cells["A1:R1"].Value = "Top 5 statistics";
            ws.Cells["A1:R1"].Style.Font.Bold = true;
            ws.Cells["A1:R1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#DAEEF3");
            ws.Cells["A1:R1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A1:R1"].Style.Fill.BackgroundColor.SetColor(colFromHex);
        }
        private void SecondSheetTableTopFiveBestSellingShopsOverall(ExcelWorksheet ws)
        {
            ApplyBorders(ws.Cells["C5:H11"]);
            ws.Cells["C5:H6"].Style.Font.Bold = true;
            ws.Cells["C5:H5"].Value = "Top 5 shops by sales (Overall)";
            ws.Cells["C5:H6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["C5:H5"].Merge = true;
            Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFA500");
            ws.Cells["C5:H5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["C5:H5"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            const int firstRow = 6;
            const int lastRow = 11;
            for (int row = firstRow; row <= lastRow; row++)
                ws.Cells["C" + row + ":G" + row].Merge = true;
            ws.Cells["C6:G6"].Value = "Shop";
            ws.Cells["H6"].Value = "Sales";
        }
        private void SecondSheetTableTopFiveBestSellingShopsLatestWeek(ExcelWorksheet ws)
        {
            ApplyBorders(ws.Cells["K5:P11"]);
            ws.Cells["K5:P6"].Style.Font.Bold = true;
            ws.Cells["K5:P5"].Value = "Top 5 shops by sales (Latest week)";
            ws.Cells["K5:P6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["K5:P5"].Merge = true;
            Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFA500");
            ws.Cells["K5:P5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["K5:P5"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            const int firstRow = 6;
            const int lastRow = 11;
            for (int row = firstRow; row <= lastRow; row++)
                ws.Cells["K" + row + ":O" + row].Merge = true;
            ws.Cells["K6:O6"].Value = "Shop";
            ws.Cells["P6"].Value = "Sales";
        }
        private void SecondSheetTableTopFiveBestSellingGamesOverall(ExcelWorksheet ws)
        {
            ApplyBorders(ws.Cells["C14:H20"]);
            ws.Cells["C14:H15"].Style.Font.Bold = true;
            ws.Cells["C14:H14"].Value = "Top 5 games by sales (Overall)";
            ws.Cells["C14:H15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["C14:H14"].Merge = true;
            Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFA500");
            ws.Cells["C14:H14"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["C14:H14"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            const int firstRow=15;
            const int lastRow=20;
            for (int row =firstRow; row <=lastRow; row++)
                ws.Cells["C" + row + ":G" + row].Merge = true;
            ws.Cells["C15:G15"].Value = "Game";
            ws.Cells["H15"].Value = "Sales";
        }
        private void SecondSheetTableTopFiveBestSellingGamesLatestWeek(ExcelWorksheet ws)
        {
            ApplyBorders(ws.Cells["K14:P20"]);
            ws.Cells["K14:P15"].Style.Font.Bold = true;
            ws.Cells["K14:P14"].Value = "Top 5 games by sales (Latest week)";
            ws.Cells["K14:P15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["K14:P14"].Merge = true;
            Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFA500");
            ws.Cells["K14:P14"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["K14:P14"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            int firstRow = 15;
            int lastRow = 20;
            for (int row = firstRow; row <= lastRow; row++)
                ws.Cells["K" + row + ":O" + row].Merge = true;
            ws.Cells["K15:O15"].Value = "Game";
            ws.Cells["P15"].Value = "Sales";
        }
        private void ThirdSheetTableStatisticsFormatting(ExcelWorksheet ws)
        {
            ApplyBorders(ws.Cells["A1:R1"]);
            ws.Cells["A1:R1"].Merge = true;
            ws.Cells["A1:R1"].Value = "Overall sales by platform";
            ws.Cells["A1:R1"].Style.Font.Bold = true;
            ws.Cells["A1:R1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#DAEEF3");
            ws.Cells["A1:R1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A1:R1"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            int column = 4;
            foreach (Enums.OutputAbbreviations abbreviation in Enum.GetValues(typeof(Enums.OutputAbbreviations)))
            {
                ws.Cells[3, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[3, column].Value = Enums.EnumHelper.GetDescription(abbreviation);
                column++;
            }
            ws.Cells["P3"].Value = "Total";
        }
        private void FourthSheetTableStatisticsFormatting(ExcelWorksheet ws)
        {
            ApplyBorders(ws.Cells["A1:R1"]);
            ws.Cells["A1:R1"].Merge = true;
            ws.Cells["A1:R1"].Value = "Overall sales by game";
            ws.Cells["A1:R1"].Style.Font.Bold = true;
            ws.Cells["A1:R1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#DAEEF3");
            ws.Cells["A1:R1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A1:R1"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            ws.Cells["A3"].Value = "Platform";
            ws.Cells["B3:D3"].Merge = true;
            ws.Cells["B3:D3"].Value = "Game";
            ws.Cells["E3"].Value = "Sales";
        }
    }
}
