using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace SS_Reports
{
    /// <summary>
    /// The set of common methods.
    /// </summary>
    ///  Methods like write-to-file reside here, because they are common for all shops. 
    /// Read-source-file is store-specific method, so it resides into the specified class
    internal abstract class StoreCore
    {
        /// <summary>
        /// Groups the stock/sales into a single object.
        /// </summary>
        protected class StockSales
        {
            internal int Stock { get; set; }
            internal int Sales { get; set; }
        }

        private XLWorkbook sourceDataWorkbook;
        private XLWorkbook outputDataWorkbook;

        //Source data worksheet
        protected IXLWorksheet sourceDataSheet;

        //Output sheets
        protected IXLWorksheet outputReviewSheet;
        protected IXLWorksheet outputTopFiveSheet;
        protected IXLWorksheet outputSalesByPlatformSheet;
        protected IXLWorksheet outputSalesByGameSheet;

        //Flag, indicating whether the user is adding or subtracting data.
        protected bool subtractData;

        internal bool CancelProcess { get; set; }

        /// <summary>
        /// Shop; platform; game; stock/sales
        /// </summary>
        protected Dictionary<string, Dictionary<string, Dictionary<string, StockSales>>> newDataDictionary = new Dictionary<string, Dictionary<string, Dictionary<string, StockSales>>>();

        internal StoreCore(string sourceFile, string destinationFile, bool subtractData)
        {
            sourceDataWorkbook = new XLWorkbook(sourceFile);
            outputDataWorkbook = new XLWorkbook(destinationFile);
            this.subtractData = subtractData;
            sourceDataSheet = sourceDataWorkbook.Worksheet(1);
            outputReviewSheet = outputDataWorkbook.Worksheet(1);
            outputTopFiveSheet = outputDataWorkbook.Worksheet(2);
            outputSalesByPlatformSheet = outputDataWorkbook.Worksheet(3);
            outputSalesByGameSheet = outputDataWorkbook.Worksheet(4);
        }
        protected Dictionary<string, Dictionary<string, Dictionary<string, StockSales>>> NewestSourceData
        {
            get { return newDataDictionary; }
            set { newDataDictionary = value; }
        }
        protected XLWorkbook OutputDataWorkbook
        {
            get { return outputDataWorkbook; }
            set { outputDataWorkbook = value; }
        }

        /// <summary>
        /// Writes the new data to the output file.
        /// </summary>
        protected void WriteData()
        {
            if (subtractData == false)
            {
                WriteReview();
            }
            else
            {
                SubtractFromReview();
                ClearTopFiveStatistics();
            }
            WriteOverallSalesByPlatform();
            WriteOverallSalesByGame();
            WriteTopFiveShopsAndGamesBySalesOverall();
        }

        /// <summary>
        /// Writes to the review sheet (sheet 1) into the output file.
        /// </summary>
        private void WriteReview()
        {
            var currentCulture = CultureInfo.CurrentCulture;
            //Report week
            var weekNo = currentCulture.Calendar.GetWeekOfYear(
                            SourceFileDate(),
                            currentCulture.DateTimeFormat.CalendarWeekRule,
                            currentCulture.DateTimeFormat.FirstDayOfWeek);
            //Bounds of the sellout table.
            int selloutTableFirstWeek = 3;//Column C, therefore 3.
            int selloutTableLastWeek = 54;//Column BB, therefore 54.
            for (int column = selloutTableFirstWeek; column <= selloutTableLastWeek; column++)
            {
                //Find an empty column in the sellout table.
                if (outputReviewSheet.Cell(3, column).GetString() != "" && outputReviewSheet.Cell(3, column).GetString() != "w0")
                {
                    if (column == selloutTableLastWeek)
                    {
                        throw new Exceptions.OutputFileIsFullException("The selected source file is full. Please select another one.");
                    }
                    continue;
                }
                //Extract the sales for each platform.
                Dictionary<string, StockSales> stockAndSalesByPlatform = GetSourceStockSalesPerPlatform();
                outputReviewSheet.Cell("BI3").Value = "Stock w" + weekNo;
                outputReviewSheet.Cell(3, column).Value = "w" + weekNo;
                //The sellout table begins at the 4th row
                int firstRow = 4;
                for (int row = firstRow; row < Enum.GetNames(typeof(Enums.OutputAbbreviations)).Length + firstRow; row++)
                {
                    string currentOutputPlatformAbbreviation = outputReviewSheet.Cell("B" + row).GetString();
                    if (stockAndSalesByPlatform[currentOutputPlatformAbbreviation].Sales != int.MinValue)
                    {
                        outputReviewSheet.Cell(row, column).Value = stockAndSalesByPlatform[currentOutputPlatformAbbreviation].Sales;
                    }
                    outputReviewSheet.Cell("BI" + row).Value = "";
                    outputReviewSheet.Cell("BJ" + row).Value = "";
                    if (stockAndSalesByPlatform[currentOutputPlatformAbbreviation].Stock != int.MinValue)
                    {
                        outputReviewSheet.Cell("BI" + row).Value = stockAndSalesByPlatform[currentOutputPlatformAbbreviation].Stock;
                        if (outputReviewSheet.Cell("BD" + row).GetString() != "0")
                            outputReviewSheet.Cell("BJ" + row).FormulaA1 = "=(BI" + row + "/BD" + row + ")*7*" + PlatformWeeksInStock(row);
                    }
                }
                outputReviewSheet.Cell("BJ16").FormulaA1 = "=BI16/BD16*7*" + TotalWeeksAddedToFile();
                break;
            }
        }

        /// <summary>
        /// Subtracts the given data from the review sheet (sheet 1) off the output file.
        /// </summary>
        private void SubtractFromReview()
        {
            var currentCulture = CultureInfo.CurrentCulture;
            var weekNo = currentCulture.Calendar.GetWeekOfYear(
                            SourceFileDate(),
                            currentCulture.DateTimeFormat.CalendarWeekRule,
                            currentCulture.DateTimeFormat.FirstDayOfWeek);
            //Column indexes, these are the bounds of the sellout table.
            Dictionary<string, StockSales> stockAndSalesPerPlatform = GetSourceStockSalesPerPlatform();
            int columnToSubtract = GetColumnWeekToDelete(stockAndSalesPerPlatform, weekNo);
            outputReviewSheet.Cell(3, columnToSubtract).Value = "w0";
            //The sellout table begins at the 4th row
            int firstRow = 4;
            for (int row = firstRow; row < Enum.GetNames(typeof(Enums.OutputAbbreviations)).Length + firstRow; row++)
            {
                outputReviewSheet.Cell(row, columnToSubtract).Value = "";
            }
            for (int row = firstRow; row < Enum.GetNames(typeof(Enums.OutputAbbreviations)).Length + firstRow; row++)
            {
                if (outputReviewSheet.Cell("BD" + row).GetString() != "0" && outputReviewSheet.Cell("BI" + row).GetString() != "" && outputReviewSheet.Cell("BI" + row).GetString() != "No data")
                    outputReviewSheet.Cell("BJ" + row).FormulaA1 = "=(BI" + row + "/BD" + row + ")*7*" + PlatformWeeksInStock(row);
                else if (outputReviewSheet.Cell("BI" + row).GetString() == "No data")
                    outputReviewSheet.Cell("BJ" + row).Value = "No data";
                else
                    outputReviewSheet.Cell("BJ" + row).Value = "";
            }
            outputReviewSheet.Cell("BJ16").FormulaA1 = "=BI16/BD16*7*" + TotalWeeksAddedToFile();

            if (Regex.IsMatch(outputReviewSheet.Cell(3, "BI").GetString(), "w" + weekNo, RegexOptions.IgnoreCase) == true)
            {
                for (int row = firstRow; row < Enum.GetNames(typeof(Enums.OutputAbbreviations)).Length + firstRow; row++)
                {
                    if (outputReviewSheet.Cell(row, "BI").GetString() == "")
                    {
                        if (stockAndSalesPerPlatform[outputReviewSheet.Cell(row, "B").GetString()].Stock == 0)
                        {
                            continue;
                        }
                    }
                    else
                    {
                        int parser;
                        if (int.TryParse(outputReviewSheet.Cell(row, "BI").GetString(), out parser) && parser == stockAndSalesPerPlatform[outputReviewSheet.Cell(row, "B").GetString()].Stock)
                        {
                            continue;
                        }
                        break;
                    }
                    for (int rowToBeCleared = firstRow; rowToBeCleared < Enum.GetNames(typeof(Enums.OutputAbbreviations)).Length + firstRow; rowToBeCleared++)
                    {
                        outputReviewSheet.Cell("BI" + rowToBeCleared).Value = "No data";
                        outputReviewSheet.Cell("BJ" + rowToBeCleared).Value = "No data";
                    }
                }
            }
        }

        /// <summary>
        /// Returns the stock/sales per platform.
        /// </summary>
        /// <returns></returns>
        private Dictionary<string, StockSales> GetSourceStockSalesPerPlatform()
        {
            Dictionary<string, StockSales> stockAndSalesPerPlatform = new Dictionary<string, StockSales>();
            foreach (KeyValuePair<string, Dictionary<string, Dictionary<string, StockSales>>> shop in NewestSourceData)
            {
                foreach (KeyValuePair<string, Dictionary<string, StockSales>> platform in shop.Value)
                {
                    if (stockAndSalesPerPlatform.ContainsKey(platform.Key) == false)
                    {
                        stockAndSalesPerPlatform.Add(platform.Key, new StockSales { Stock = int.MinValue, Sales = int.MinValue });
                    }
                    foreach (KeyValuePair<string, StockSales> game in platform.Value)
                    {
                        if (stockAndSalesPerPlatform[platform.Key].Stock == int.MinValue || stockAndSalesPerPlatform[platform.Key].Sales == int.MinValue)
                        {
                            stockAndSalesPerPlatform[platform.Key].Stock = 0;
                            stockAndSalesPerPlatform[platform.Key].Sales = 0;
                        }
                        stockAndSalesPerPlatform[platform.Key].Stock += game.Value.Stock;
                        stockAndSalesPerPlatform[platform.Key].Sales += game.Value.Sales;
                    }
                }
            }
            return stockAndSalesPerPlatform;
        }

        /// <summary>
        /// Returns all columns from the output file, which correspond to the given week. Used to find a column to delete.
        /// </summary>
        /// <param name="sourceData"></param>
        /// <param name="weekNo"></param>
        /// <returns></returns>
        private int GetColumnWeekToDelete(Dictionary<string, StockSales> sourceData, int weekNo)
        {
            List<int> allColumnsRepresentingTheWeek = new List<int>();
            //Column indexes, these are the bounds of the sellout table.
            int selloutTableFirstWeek = 3;//Column C, therefore 3.
            int selloutTableLastWeek = 54;//Column BB, therefore 54.
            bool recordExists = false;
            for (int icolumn = selloutTableFirstWeek; icolumn <= selloutTableLastWeek; icolumn++)
            {
                if (outputReviewSheet.Cell(3, icolumn).GetString() == "w0" || outputReviewSheet.Cell(3, icolumn).GetString() != "w" + weekNo)
                {
                    if (icolumn == selloutTableLastWeek && recordExists == false)
                    {
                        throw new Exceptions.OutputFileNoRecordsFoundException("The output file does not contain any record for week " + weekNo + ".");
                    }
                    continue;
                }
                else if (outputReviewSheet.Cell(3, icolumn).GetString() == "w" + weekNo)
                {
                    recordExists = true;
                    allColumnsRepresentingTheWeek.Add(icolumn);
                }
            }
            for (int icolumn = allColumnsRepresentingTheWeek.Count - 1; icolumn >= 0; icolumn--)
            {
                //Indicates whether to continue searching after the i-th pass.
                bool continueSearching = false;
                //The sellout table begins at the 4th row
                int firstRow = 4;
                for (int row = firstRow; row < Enum.GetNames(typeof(Enums.OutputAbbreviations)).Length + firstRow; row++)
                {
                    string currentPlatform = outputReviewSheet.Cell("B" + row).GetString();
                    if (outputReviewSheet.Cell(row, allColumnsRepresentingTheWeek[icolumn]).GetString() == "")
                    {
                        if (sourceData[currentPlatform].Sales == int.MinValue)
                            continue;
                        continueSearching = true;
                        break;
                    }
                    else if (Convert.ToInt32(outputReviewSheet.Cell(row, allColumnsRepresentingTheWeek[icolumn]).GetString()) != sourceData[outputReviewSheet.Cell("B" + row).GetString()].Sales)
                    {
                        continueSearching = true;
                        break;
                    }
                }
                if (icolumn == 0 && continueSearching == true)
                    throw new Exceptions.OutputFileNoRecordsFoundException("Cannot find the column to contain the data to be subtracted.");
                if (continueSearching == true)
                    continue;
                return allColumnsRepresentingTheWeek[icolumn];
            }
            return 0;
        }

        /// <summary>
        /// Clears the top 5 statistics on the second sheet in the output file.
        /// </summary>
        /// Used when the user subtracts the latest data added.
        private void ClearTopFiveStatistics()
        {
            //Rows from 7 to 11 in columns K and P.
            int firstRow = 7;
            int lastRow = 11;
            for (int row = firstRow; row <= lastRow; row++)
            {
                outputTopFiveSheet.Cell("K" + row).Value = "";
                outputTopFiveSheet.Cell("P" + row).Value = "";
            }
            //Rows from 16 to 20 in columns K and P.
            firstRow = 16;
            lastRow = 20;
            for (int row = firstRow; row <= lastRow; row++)
            {
                outputTopFiveSheet.Cell("K" + row).Value = "";
                outputTopFiveSheet.Cell("P" + row).Value = "";
            }
        }

        /// <summary>
        /// Used for the estimates. Returns a number of weeks, during which there's been at least one game from a specific platform present for sale in any shop.
        /// </summary>
        /// <param name="currentRow">The current row from the review sheet.</param>
        /// <returns>Total number of weeks.</returns>
        protected int PlatformWeeksInStock(int currentRow)
        {
            int weeksInStock = 0;
            for (int column = 3; column <= 55; column++)
            {
                if (outputReviewSheet.Cell(currentRow, column).GetString() != "" && outputReviewSheet.Cell(3, column).GetString() != "w0" && outputReviewSheet.Cell(3, column).GetString() != "")
                    weeksInStock++;
            }
            return weeksInStock;
        }
        /// <summary>
        /// Returns the number of weeks for which a report is present in the output file.
        /// </summary>
        /// <returns>Number of weeks.</returns>
        private int TotalWeeksAddedToFile()
        {
            int totalWeeks = 0;
            for (int column = 3; column <= 55; column++)
            {
                //w0 also indicates empty column.
                if (outputReviewSheet.Cell(3, column).GetString() != "w0" && outputReviewSheet.Cell(3, column).GetString() != "")
                    totalWeeks++;
            }
            return totalWeeks;
        }
        /// <summary>
        /// Returns true if the output file selected is preformatted using the "Create new" button from the interface.
        /// </summary>
        /// <returns>True if the output file is correct, false if not.</returns>
        protected bool DestinationFileSignature()
        {
            if (OutputDataWorkbook.Worksheets.Count != 4)
                return false;
            if (outputReviewSheet.Cell("AT1").GetString() != "Sell out")
                return false;
            else if (outputReviewSheet.Cell("BH1").GetString() != "Stock")
                return false;
            else
                return true;
        }
        /// <summary>
        /// Writes data on the top-five sheet (sheet 2) in the output file.
        /// </summary>
        private void WriteTopFiveShopsAndGamesBySalesOverall()
        {
            TopFiveGamesBySalesOverall();
            TopFiveShopsBySalesOverall();
            if (subtractData == false)
            {
                TopFiveShopsBySalesLatestWeek();
                TopFiveGamesBySalesLatestWeek();
            }
        }

        /// <summary>
        /// Writes the top-five shops by sales (overall) statistics.
        /// </summary>
        private void TopFiveShopsBySalesOverall()
        {
            Dictionary<string, int> gamesAndSales = new Dictionary<string, int>();
            for (int row = 4; row <= outputSalesByPlatformSheet.LastRowUsed().RowNumber(); row++)
            {
                int n;
                if (int.TryParse(outputSalesByPlatformSheet.Cell(row, 16).GetString(), out n) == false)
                    continue;
                gamesAndSales.Add(outputSalesByPlatformSheet.Cell(row, 1).GetString(), n);
            }
            List<KeyValuePair<string, int>> gamesAndSalesSorted = gamesAndSales.ToList();
            //Sort by sales
            gamesAndSalesSorted.Sort((current, next) => { return next.Value.CompareTo(current.Value); });
            //Rows from 7 to 11 in columns C and H.
            const int firstRow = 7;
            const int lastRow = 11;
            for (int row = firstRow; row <= lastRow; row++)
            {
                if (gamesAndSalesSorted.Count > row - firstRow)
                {
                    outputTopFiveSheet.Cell("C" + row).Value = gamesAndSalesSorted[row - firstRow].Key;
                    outputTopFiveSheet.Cell("H" + row).Value = gamesAndSalesSorted[row - firstRow].Value;
                }
            }
        }
        /// <summary>
        /// Writes the top-five games by sales (overall) statistics.
        /// </summary>
        private void TopFiveGamesBySalesOverall()
        {
            Dictionary<string, Dictionary<string, int>> gamesAndSales = new Dictionary<string, Dictionary<string, int>>();
            for (int row = 4; row <= outputSalesByGameSheet.LastRowUsed().RowNumber(); row++)
            {
                if (gamesAndSales.ContainsKey(outputSalesByGameSheet.Cell(row, "A").GetString()))
                {
                    if (gamesAndSales[outputSalesByGameSheet.Cell(row, "A").GetString()].ContainsKey(outputSalesByGameSheet.Cell(row, "B").GetString()) == false)
                    {
                        int n;
                        if (int.TryParse(outputSalesByGameSheet.Cell(row, "E").GetString(), out n) == false)
                            continue;
                        gamesAndSales[outputSalesByGameSheet.Cell(row, "A").GetString()].Add(outputSalesByGameSheet.Cell(row, "B").GetString(), n);
                    }
                }
                else
                {
                    gamesAndSales.Add(outputSalesByGameSheet.Cell(row, "A").GetString(), new Dictionary<string, int>());
                    gamesAndSales[outputSalesByGameSheet.Cell(row, "A").GetString()].Add(outputSalesByGameSheet.Cell(row, "B").GetString(), Convert.ToInt32(outputSalesByGameSheet.Cell(row, "E").GetString()));
                }
            }
            List<KeyValuePair<string, Dictionary<string, int>>> platformsGamesAndSalesList = gamesAndSales.ToList();
            Dictionary<string, int> combinePlatformAndGame = new Dictionary<string, int>();
            foreach (KeyValuePair<string, Dictionary<string, int>> platform in platformsGamesAndSalesList)
            {
                foreach (KeyValuePair<string, int> game in platform.Value)
                {
                    combinePlatformAndGame.Add(platform.Key + " " + game.Key, game.Value);
                }
            }
            List<KeyValuePair<string, int>> gamesAndSalesListSorted = combinePlatformAndGame.ToList();
            //Sort by sales
            gamesAndSalesListSorted.Sort((current, next) => { return next.Value.CompareTo(current.Value); });
            //Rows from 16 to 20 in columns C and H.
            const int firstRow = 16;
            const int lastRow = 20;
            for (int row = firstRow; row <= lastRow; row++)
            {
                if (gamesAndSalesListSorted.Count > row - firstRow)
                {
                    outputTopFiveSheet.Cell("C" + row).Value = gamesAndSalesListSorted[row - firstRow].Key;
                    outputTopFiveSheet.Cell("H" + row).Value = gamesAndSalesListSorted[row - firstRow].Value;
                }
            }
        }

        /// <summary>
        /// Writes the top five shops by sales (latest week) statistics.
        /// </summary>
        private void TopFiveShopsBySalesLatestWeek()
        {
            Dictionary<string, int> shopsAndSales = new Dictionary<string, int>();
            foreach (KeyValuePair<string, Dictionary<string, Dictionary<string, StockSales>>> shop in newDataDictionary)
            {
                if (shopsAndSales.ContainsKey(shop.Key) == false)
                    shopsAndSales.Add(shop.Key, 0);
                foreach (KeyValuePair<string, Dictionary<string, StockSales>> platform in shop.Value)
                {
                    foreach (KeyValuePair<string, StockSales> game in platform.Value)
                    {
                        shopsAndSales[shop.Key] += newDataDictionary[shop.Key][platform.Key][game.Key].Sales;
                    }
                }
            }
            List<KeyValuePair<string, int>> shopsAndSalesList = shopsAndSales.ToList();
            //Sort by sales
            shopsAndSalesList.Sort((current, next) => { return next.Value.CompareTo(current.Value); });
            //Rows from 7 to 11 in columns K and P.
            const int firstRow = 7;
            const int lastRow = 11;
            for (int row = firstRow; row <= lastRow; row++)
            {
                if (shopsAndSalesList.Count > row - firstRow)
                {
                    outputTopFiveSheet.Cell("K" + row).Value = shopsAndSalesList[row - firstRow].Key;
                    outputTopFiveSheet.Cell("P" + row).Value = shopsAndSalesList[row - firstRow].Value;
                }
            }
        }

        /// <summary>
        /// Writes the top five games by sales (latest week) statistics.
        /// </summary>
        private void TopFiveGamesBySalesLatestWeek()
        {
            Dictionary<string, int> platformsGamesAndSales = new Dictionary<string, int>();
            foreach (KeyValuePair<string, Dictionary<string, Dictionary<string, StockSales>>> shop in NewestSourceData)
            {
                foreach (KeyValuePair<string, Dictionary<string, StockSales>> platform in shop.Value)
                {
                    foreach (KeyValuePair<string, StockSales> game in platform.Value)
                    {
                        if (platformsGamesAndSales.ContainsKey(platform.Key + " " + game.Key) == true)
                        {
                            platformsGamesAndSales[platform.Key + " " + game.Key] += game.Value.Sales;
                        }
                        else
                        {
                            platformsGamesAndSales.Add(platform.Key + " " + game.Key, game.Value.Sales);
                        }
                    }
                }
            }
            List<KeyValuePair<string, int>> platformsGamesAndSalesToList = platformsGamesAndSales.ToList();
            //Sort by sales
            platformsGamesAndSalesToList.Sort((current, next) => { return next.Value.CompareTo(current.Value); });
            //Rows from 16 to 20 in columns K and P.
            const int firstRow = 16;
            const int lastRow = 20;
            for (int row = firstRow; row <= lastRow; row++)
            {
                if (platformsGamesAndSalesToList.Count > row - firstRow)
                {
                    outputTopFiveSheet.Cell("K" + row).Value = platformsGamesAndSalesToList[row - firstRow].Key;
                    outputTopFiveSheet.Cell("P" + row).Value = platformsGamesAndSalesToList[row - firstRow].Value;
                }
            }
        }

        /// <summary>
        /// Writes data to the overall sales by platform (sheet 3). 
        /// </summary>
        private void WriteOverallSalesByPlatform()
        {
            if (outputSalesByPlatformSheet.LastRowUsed().RowNumber() > 3)
            {
                OverallSalesByPlatformRecordsExisting();
            }
            else
            {
                OverallSalesByPlatformRecordsNotExisting();
            }
        }
        /// <summary>
        /// Writes to the overall sales sheet, adding any previously contained data to the new data.
        /// Also used when subtracting. In that case it subtracts the new data from the previously contained data.
        /// </summary>
        private void OverallSalesByPlatformRecordsExisting()
        {
            Dictionary<string, Dictionary<string, int>> currentOverallSales = GetCurrentOverallSalesPerPlatform();
            foreach (KeyValuePair<string, Dictionary<string, Dictionary<string, StockSales>>> shop in NewestSourceData)
            {
                if (currentOverallSales.ContainsKey(shop.Key) == false)
                    currentOverallSales.Add(shop.Key, new Dictionary<string, int>());
                foreach (KeyValuePair<string, Dictionary<string, StockSales>> platform in shop.Value)
                {
                    if (currentOverallSales[shop.Key].ContainsKey(platform.Key) == false)
                        currentOverallSales[shop.Key].Add(platform.Key, 0);
                    foreach (KeyValuePair<string, StockSales> game in platform.Value)
                    {
                        if (subtractData == false)
                            currentOverallSales[shop.Key][platform.Key] += game.Value.Sales;
                        else
                            currentOverallSales[shop.Key][platform.Key] -= game.Value.Sales;
                    }
                }
            }
            int firstRow = 4;
            var newFileStatistics = currentOverallSales.ToArray();
            for (int row = firstRow; row < newFileStatistics.Length + firstRow; row++)
            {
                outputSalesByPlatformSheet.Range("A" + row + ":C" + row).Merge();
                outputSalesByPlatformSheet.Cell(row, 1).Value = newFileStatistics[row - firstRow].Key.ToString();
                outputSalesByPlatformSheet.Cell(row, 16).FormulaA1 = "=SUM(D" + row + ":O" + row + ")";
                foreach (KeyValuePair<string, int> platform in newFileStatistics[row - 4].Value)
                {
                    for (int column = 3; column < 16; column++)
                    {
                        if (outputSalesByPlatformSheet.Cell(3, column).GetString() == platform.Key)
                        {
                            outputSalesByPlatformSheet.Cell(row, column).Value = platform.Value;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Writes to the overall sales sheet.
        /// </summary>
        private void OverallSalesByPlatformRecordsNotExisting()
        {
            const int firstRow = 4;
            const int firstColumn = 3;
            const int lastColumn = 15;
            var currentFileStatistics = NewestSourceData.ToArray();
            for (int row = firstRow; row < currentFileStatistics.Length + firstRow; row++)
            {
                outputSalesByPlatformSheet.Range("A" + row + ":C" + row).Merge();
                outputSalesByPlatformSheet.Cell(row, 1).Value = currentFileStatistics[row - firstRow].Key.ToString();
                outputSalesByPlatformSheet.Cell(row, 16).FormulaA1 = "=SUM(D" + row + ":O" + row + ")";
                foreach (KeyValuePair<string, Dictionary<string, StockSales>> platform in currentFileStatistics[row - 4].Value)
                {
                    int sumSales = 0;
                    foreach (KeyValuePair<string, StockSales> game in platform.Value)
                    {
                        sumSales += currentFileStatistics[row - firstRow].Value[platform.Key][game.Key].Sales;
                    }
                    for (int column = firstColumn; column <= lastColumn; column++)
                    {
                        if (outputSalesByPlatformSheet.Cell(3, column).GetString() == platform.Key)
                        {
                            outputSalesByPlatformSheet.Cell(row, column).Value = sumSales;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Returns a dictionary, containing the overall sales per platform.
        /// </summary>
        /// <returns></returns>
        private Dictionary<string, Dictionary<string, int>> GetCurrentOverallSalesPerPlatform()
        {
            //The first row from sheet 3 that is possible to read from.
            int firstRow = 4;
            Dictionary<string, Dictionary<string, int>> currentSales = new Dictionary<string, Dictionary<string, int>>();
            for (int row = firstRow; row <= outputSalesByPlatformSheet.LastRowUsed().RowNumber(); row++)
            {
                if (outputSalesByPlatformSheet.Cell(row, 1).GetString() == "")
                    continue;
                if (currentSales.ContainsKey(outputSalesByPlatformSheet.Cell(row, 1).GetString()) == false)
                {
                    currentSales.Add(outputSalesByPlatformSheet.Cell(row, 1).GetString(), new Dictionary<string, int>());
                    int parser;
                    //first column containing platform name.
                    int firstColumn = 4;
                    for (int column = firstColumn; column < outputSalesByPlatformSheet.LastColumnUsed().ColumnNumber(); column++)
                    {
                        if (outputSalesByPlatformSheet.Cell(3, column).GetString() == "Total")
                            break;
                        if (currentSales[outputSalesByPlatformSheet.Cell(row, 1).GetString()].ContainsKey(outputSalesByPlatformSheet.Cell(3, column).GetString()) == false)
                        {
                            if (outputSalesByPlatformSheet.Cell(row, column).GetString() != "" && int.TryParse(outputSalesByPlatformSheet.Cell(row, column).GetString(), out parser) == true)
                            {
                                currentSales[outputSalesByPlatformSheet.Cell(row, 1).GetString()].Add(outputSalesByPlatformSheet.Cell(3, column).GetString(), 0);
                                currentSales[outputSalesByPlatformSheet.Cell(row, 1).GetString()][outputSalesByPlatformSheet.Cell(3, column).GetString()] += parser;
                            }
                        }
                    }
                }
            }
            return currentSales;
        }

        /// <summary>
        /// Writes the overall sales by game.
        /// </summary>
        private void WriteOverallSalesByGame()
        {
            //If there are more than 3 rows in use on the 4th sheet then some records exist.
            if (outputSalesByGameSheet.LastRowUsed().RowNumber() > 3)
            {
                SalesByGameRecordsExisting();
            }
            else
            {
                SalesByGameRecordsNotExisting();
            }
        }
        /// <summary>
        /// Writes to the sales by game sheet, adding any previously contained data to the new data.
        /// </summary>
        private void SalesByGameRecordsExisting()
        {
            int firstRow = 4;
            int currentLastRow;
            foreach (KeyValuePair<string, Dictionary<string, Dictionary<string, StockSales>>> shop in NewestSourceData)
            {
                foreach (KeyValuePair<string, Dictionary<string, StockSales>> platform in shop.Value)
                {
                    foreach (KeyValuePair<string, StockSales> game in platform.Value)
                    {
                        currentLastRow = outputSalesByGameSheet.LastRowUsed().RowNumber();
                        for (int row = firstRow; row <= currentLastRow; row++)
                        {
                            if (outputSalesByGameSheet.Cell("A" + row).GetString() == platform.Key)
                            {
                                if (outputSalesByGameSheet.Cell("B" + row).GetString() == game.Key)
                                {
                                    if (subtractData == false)
                                    {
                                        outputSalesByGameSheet.Cell("E" + row).Value = Convert.ToInt32(outputSalesByGameSheet.Cell("E" + row).GetString()) + game.Value.Sales;
                                    }
                                    else
                                    {
                                        outputSalesByGameSheet.Cell("E" + row).Value = Convert.ToInt32(outputSalesByGameSheet.Cell("E" + row).GetString()) - game.Value.Sales;
                                    }
                                    break;
                                }
                            }
                            if (subtractData == false && row == currentLastRow)
                            {
                                outputSalesByGameSheet.Cell("A" + (currentLastRow + 1)).Value = platform.Key;
                                outputSalesByGameSheet.Cell("B" + (currentLastRow + 1)).Value = game.Key;
                                outputSalesByGameSheet.Range("B" + (currentLastRow + 1) + ":D" + (currentLastRow + 1)).Merge();
                                outputSalesByGameSheet.Cell("E" + (currentLastRow + 1)).Value = game.Value.Sales;
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Writes to the sales by game sheet.
        /// </summary>
        private void SalesByGameRecordsNotExisting()
        {
            int firstRow = 4;
            //Varies
            int currentLastRow;
            foreach (KeyValuePair<string, Dictionary<string, Dictionary<string, StockSales>>> shop in NewestSourceData)
            {
                foreach (KeyValuePair<string, Dictionary<string, StockSales>> platform in shop.Value)
                {
                    foreach (KeyValuePair<string, StockSales> game in platform.Value)
                    {
                        currentLastRow = outputSalesByGameSheet.LastRowUsed().RowNumber();
                        if (currentLastRow < firstRow)
                        {
                            outputSalesByGameSheet.Cell("A" + (currentLastRow + 1)).Value = platform.Key;
                            outputSalesByGameSheet.Cell("B" + (currentLastRow + 1)).Value = game.Key;
                            outputSalesByGameSheet.Range("B" + (currentLastRow + 1) + ":D" + (currentLastRow + 1)).Merge();
                            outputSalesByGameSheet.Cell("E" + (currentLastRow + 1)).Value = game.Value.Sales;
                            continue;
                        }
                        for (int row = firstRow; row <= currentLastRow; row++)
                        {
                            if (outputSalesByGameSheet.Cell("A" + row).GetString() == platform.Key)
                            {
                                if (outputSalesByGameSheet.Cell("B" + row).GetString() == game.Key)
                                {
                                    outputSalesByGameSheet.Cell("E" + row).Value = Convert.ToInt32(outputSalesByGameSheet.Cell("E" + row).GetString()) + game.Value.Sales;
                                    break;
                                }
                            }
                            if (row == currentLastRow)
                            {
                                outputSalesByGameSheet.Cell("A" + (currentLastRow + 1)).Value = platform.Key;
                                outputSalesByGameSheet.Cell("B" + (currentLastRow + 1)).Value = game.Key;
                                outputSalesByGameSheet.Range("B" + (currentLastRow + 1) + ":D" + (currentLastRow + 1)).Merge();
                                outputSalesByGameSheet.Cell("E" + (currentLastRow + 1)).Value = game.Value.Sales;
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Fills in the output platform abbreviations.
        /// </summary>
        protected virtual void PrepareDataDictionary()
        {
            foreach (Enums.OutputAbbreviations abbreviation in Enum.GetValues(typeof(Enums.OutputAbbreviations)))
            {
                foreach (KeyValuePair<string, Dictionary<string, Dictionary<string, StockSales>>> store in NewestSourceData)
                {
                    if (store.Value.ContainsKey(Enums.EnumHelper.GetDescription(abbreviation)) == false)
                        store.Value.Add(Enums.EnumHelper.GetDescription(abbreviation), new Dictionary<string, StockSales>());
                }
            }
        }
        //Methods to be implemented by the children.
        /// <summary>
        /// Returns true if the report was not cancelled (and was successful), otherwise false.
        /// </summary>
        /// <param name="cancellationPending"></param>
        /// <returns></returns>
        internal abstract bool Report(bool cancellationPending);

        /// <summary>
        /// Reads the source data and stores it into a dictionary. Overriden by the child-classes.
        /// </summary>
        protected abstract void ReadSourceData();

        /// <summary>
        /// Check if the source file is responding to the file, selected from the user in the menu.
        /// Overriden by the child classes.
        /// </summary>
        /// <returns></returns>
        protected abstract bool SourceFileSignature();

        /// <summary>
        /// Returns the shop name on a specified row. 
        /// Overriden by the child classes
        /// </summary>
        /// <param name="row">Row to read on.</param>
        /// <returns></returns>
        protected abstract string GetShop(int row);

        /// <summary>
        /// Returns the report date from the source file.
        /// Overriden by the child classes.
        /// </summary>
        /// <returns></returns>
        protected abstract DateTime SourceFileDate();
    }
}
