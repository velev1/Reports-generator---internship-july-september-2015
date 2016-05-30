using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace SS_Reports.Stores
{
    class StoreTechnomarket : StoreCore
    {
        internal StoreTechnomarket(string sourceFile, string outputFile, bool subtractData) : base(sourceFile, outputFile, subtractData)
        { }
        /// <summary>
        /// Prepares the source data dictionary, reads the source data, writes the data to the output file and saves it.
        /// </summary>
        /// <param name="cancellationPending"></param>
        /// <returns></returns>
        internal override bool Report(bool cancellationPending)
        {
            if (SourceFileSignature() == false)
            {
                throw new Exceptions.SourceFileNotMatchingSelectedFileException("The source file isn't from the selected retailer.");
            }
            if (DestinationFileSignature() == false)
            {
                throw new Exceptions.OutputFileNotCorrectException("The destionation file isn't correct. Select another file or create new.");
            }
            PrepareDataDictionary();
            ReadSourceData();
            WriteData();
            if (cancellationPending)
                return false;
            OutputDataWorkbook.Save();
            return true;
        }

        /// <summary>
        /// Return shop on column.
        /// </summary>
        /// <param name="column"></param>
        /// <returns></returns>
        protected override string GetShop(int column)
        {
            string currentShop = sourceDataSheet.Cell(5, column).GetString().Trim();
            currentShop = Regex.Replace(currentShop, @"^(?:\d*)?", string.Empty).Trim();
            return currentShop;
        }

        /// <summary>
        /// Fills in the file-specific shops, then calls the base preparation method.
        /// </summary>
        protected override void PrepareDataDictionary()
        {
            string subStore;
            const int firstColumn = 8;
            int lastColumn = sourceDataSheet.LastColumnUsed().ColumnNumber();
            //Adding every shop from the source file to the statistics dictionary.
            for (int column = firstColumn; column <= lastColumn; column++)
            {
                subStore = GetShop(column);
                if (subStore != "")
                {
                    if (NewestSourceData.ContainsKey(subStore) == false)
                    {
                        newDataDictionary.Add(subStore, new Dictionary<string, Dictionary<string, StockSales>>());
                    }
                }
            }
            base.PrepareDataDictionary();
        }

        /// <summary>
        /// Reads the source data, storing it into the data dictionary.
        /// </summary>
        protected override void ReadSourceData()
        {
            const int firstRow = 1;
            int lastRow = sourceDataSheet.LastRowUsed().RowNumber();
            //Iterate through every row in the source file.
            for (int row = firstRow; row <= lastRow; row++)
            {
                ulong n;
                if (sourceDataSheet.Cell("C" + row).Value == null || ulong.TryParse(sourceDataSheet.Cell("C" + row).GetString(), out n) == false || (n.ToString().Length != 12 && n.ToString().Length != 13))
                {
                    //If current row doesn't contain data move onto the next row.
                    continue;
                }
                string currentGamePlatform = "";
                string currentGameTitle = "";
                foreach (Enums.TechnomarketAbbreviations abbreviation in Enum.GetValues(typeof(Enums.TechnomarketAbbreviations)))
                {
                    Regex extractPlatformFromRow = new Regex("^" + Enums.EnumHelper.GetDescription(abbreviation), RegexOptions.IgnoreCase);
                    string dataRow = sourceDataSheet.Cell("B" + row).GetString();
                    //the index of the first space within the data row, used to remove the space in XBOX 360 to make it XBOX360
                    int indexOfFirstSpace = sourceDataSheet.Cell("B" + row).GetString().IndexOf(" ");
                    if (indexOfFirstSpace > 0)
                    {
                        dataRow = dataRow.Substring(0, indexOfFirstSpace) + string.Empty + dataRow.Substring(indexOfFirstSpace + 1);
                    }
                    var m = extractPlatformFromRow.Matches(dataRow);
                    if (m.Count != 0)
                    {
                        currentGamePlatform = Enums.EnumHelper.GetDescription(abbreviation);
                        currentGameTitle = Regex.Replace(dataRow, "^" + currentGamePlatform, string.Empty).Trim();
                        break;
                    }
                }
                if (currentGamePlatform == "")
                {
                    currentGamePlatform = Enums.EnumHelper.GetDescription(Enums.TechnomarketAbbreviations.Other);
                    currentGameTitle = sourceDataSheet.Cell("B" + row).Value.ToString().Trim();
                }
                int parser;
                //The first column containing shop name is H, therefore 8.
                int firstColumn = 8;
                for (int column = firstColumn; column <= sourceDataSheet.LastColumnUsed().ColumnNumber(); column++)
                {
                    string shop = GetShop(column);
                    if (shop == "")
                    {
                        //TODO appropriate exception
                    }
                    if (NewestSourceData[shop][currentGamePlatform].ContainsKey(currentGameTitle) == false)
                    {
                        //adding each game, along with the stock and sales
                        //1d stock, 2d sales
                        NewestSourceData[shop][currentGamePlatform].Add(currentGameTitle, new StockSales());
                    }
                    //stock
                    if (column % 2 == 0)
                    {
                        if (int.TryParse(sourceDataSheet.Cell(row, column).GetString(), out parser) == true)
                            NewestSourceData[shop][currentGamePlatform][currentGameTitle].Stock += parser;
                    }

                    //sales
                    else
                    {
                        if (int.TryParse(sourceDataSheet.Cell(row, column).GetString(), out parser) == true)
                            NewestSourceData[shop][currentGamePlatform][currentGameTitle].Sales += parser;
                    }
                }
            }
        }

        /// <summary>
        /// Returns the report date from the source file.
        /// </summary>
        /// <returns></returns>
        protected override DateTime SourceFileDate()
        {
            DateTime[] fromToDates = new DateTime[2];
            Regex reg = new Regex(@"\d{2}\.\d{2}\.\d{4}");
            int datesCount = 0;
            foreach (Match m in reg.Matches(sourceDataSheet.Cell("A3").Value.ToString()))
            {
                if (datesCount >= 2)
                    throw new Exception("There are more than two dates in the date scope of the file. Please, leave only first and last week in cell A3 of the file.");
                fromToDates[datesCount] = DateTime.ParseExact(m.Value, "dd.MM.yyyy", CultureInfo.InvariantCulture);
                datesCount++;
            }
            if (DateTime.Compare(fromToDates[0], fromToDates[1]) > 0)
            {
                DateTime swap = fromToDates[0];
                fromToDates[0] = fromToDates[1];
                fromToDates[1] = swap;
            }
            return fromToDates[1];
        }

        /// <summary>
        /// Check if the source file is responding to the file, selected from the user in the menu.
        /// </summary>
        /// <returns></returns>
        protected override bool SourceFileSignature()
        {
            Regex rx = new Regex("Technomarket", RegexOptions.IgnoreCase);
            if (rx.Match(sourceDataSheet.Cell(3, 1).GetString()).ToString() != "")
                return true;
            return false;
        }
    }
}