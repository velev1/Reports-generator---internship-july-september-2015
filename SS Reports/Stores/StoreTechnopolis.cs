using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace SS_Reports.Stores
{
    class StoreTechnopolis : StoreCore
    {

        internal StoreTechnopolis(string sourceFile, string outputFile, bool subtractData) : base(sourceFile, outputFile, subtractData)
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
        /// Fills in the file-specific shops, then calls the base preparation method.
        /// </summary>
        protected override void PrepareDataDictionary()
        {
            //Filling in the shops
            string subStore;
            for (int row = 1; row <= sourceDataSheet.LastRowUsed().RowNumber(); row++)
            {
                subStore = GetShop(row);
                if (subStore != "")
                {
                    if (NewestSourceData.ContainsKey(subStore) == false)
                    {
                        NewestSourceData.Add(subStore, new Dictionary<string, Dictionary<string, StockSales>>());
                    }
                }
            }
            base.PrepareDataDictionary();
        }

        /// <summary>
        /// Returns the shop name on a specified row.
        /// </summary>
        /// <param name="row">Row number.</param>
        /// <returns>Shop name.</returns>
        protected override string GetShop(int row)
        {
            return sourceDataSheet.Cell(row, 3).GetString().Trim();
        }

        /// <summary>
        /// Reads the source data, storing it into the data dictionary.
        /// </summary>
        protected override void ReadSourceData()
        {
            //Iterate through every row in the source file.
            for (int row = 1; row <= sourceDataSheet.LastRowUsed().RowNumber(); row++)
            {
                int n;
                if (sourceDataSheet.Cell("A" + row).Value == null || int.TryParse(sourceDataSheet.Cell("A" + row).Value.ToString(), out n) == false)
                {
                    //If current row doesn't contain data move onto the next row.
                    continue;
                }
                string gamePlatform = "";
                string gameTitle = "";
                var maxEnumValue = Enum.GetValues(typeof(Enums.TechnopolisAbbreviations)).Cast<Enums.TechnopolisAbbreviations>().Max();
                foreach (Enums.TechnopolisAbbreviations abbreviation in Enum.GetValues(typeof(Enums.TechnopolisAbbreviations)))
                {
                    Regex extractPlatformFromCurrentRecord = new Regex("^" + Enums.EnumHelper.GetDescription(abbreviation), RegexOptions.IgnoreCase);
                    if (sourceDataSheet.Cell("B" + row).GetString() == "" || extractPlatformFromCurrentRecord.IsMatch(sourceDataSheet.Cell("B" + row).GetString()) == false)
                    {
                        if ((int)abbreviation == (int)maxEnumValue)
                        {
                            gamePlatform = Enums.EnumHelper.GetDescription(Enums.OutputAbbreviations.Other);
                            gameTitle = sourceDataSheet.Cell("B" + row).GetString().Trim().TrimEnd('>');
                            break;
                        }
                        continue;
                    }
                    foreach (Enums.OutputAbbreviations outputAbbreviation in Enum.GetValues(typeof(Enums.OutputAbbreviations)))
                    {
                        if (abbreviation.ToString() == outputAbbreviation.ToString())
                            gamePlatform = Enums.EnumHelper.GetDescription(outputAbbreviation);
                    }
                    gameTitle = Regex.Replace(sourceDataSheet.Cell("B" + row).GetString().Trim().TrimEnd('>'), "^" + Enums.EnumHelper.GetDescription(abbreviation), string.Empty).Trim();
                    break;
                }
                do
                {
                    string subStore = GetShop(row);
                    //TODO throw appropriate exception
                    if (NewestSourceData[subStore][gamePlatform].ContainsKey(gameTitle) == false)
                    {
                        NewestSourceData[subStore][gamePlatform].Add(gameTitle, new StockSales());
                    }
                    int num;
                    if (int.TryParse(sourceDataSheet.Cell("E" + row).GetString(), out num) == true)
                        NewestSourceData[subStore][gamePlatform][gameTitle].Stock += num;
                    if (int.TryParse(sourceDataSheet.Cell("D" + row).GetString(), out num) == true)
                        NewestSourceData[subStore][gamePlatform][gameTitle].Sales += num;
                    row++;
                }
                while (row != sourceDataSheet.LastRowUsed().RowNumber() && sourceDataSheet.Cell("A" + row).Value.ToString() == "" && sourceDataSheet.Cell("A" + (row + 1)).Value.ToString() == "");
            }
        }


        /// <summary>
        /// Returns the report date from the source file.
        /// </summary>
        /// <returns></returns>
        protected override DateTime SourceFileDate()
        {
            DateTime toDate = default(DateTime);
            int columnContainingDate = 1;
            for (int column = columnContainingDate; columnContainingDate < 4; columnContainingDate++, column++)
            {
                if (sourceDataSheet.Cell(1, column).Value.ToString() == "")
                {
                    if (column == 3)
                    {
                        throw new Exception("Please add \"From-To\" dates in any of the following cells: A1, B1, C1. The date format should be DD.MM-DD.MM.YY or DD.MM-DD.MM.YYYY.");
                    }
                    continue;
                }
                break;
            }
            string dateRowWithoutSpaces = sourceDataSheet.Cell(1, columnContainingDate).Value.ToString().Replace(" ", String.Empty);
            Regex reg = new Regex(@"-\d{2}\.\d{2}\.\d{2}(?:\d{2})?");
            if (reg.Matches(dateRowWithoutSpaces).Count != 1)
                throw new Exception("The date format must be DD.MM-DD.MM.YY or DD.MM-DD.MM.YYYY. The date must be located in any of the following cells: A1, B1, C1.");
            toDate = DateTime.Parse(reg.Matches(dateRowWithoutSpaces)[0].ToString().Substring(1));
            return toDate;
        }

        /// <summary>
        /// Check if the source file is responding to the file, selected from the user in the menu.
        /// </summary>
        /// <returns></returns>
        protected override bool SourceFileSignature()
        {
            Regex rx = new Regex(@"^(Технополис|Видеолукс|WEB|GSM)");
            for (int row = 1; row <= sourceDataSheet.LastRowUsed().RowNumber(); row++)
            {
                if (rx.Match(sourceDataSheet.Cell("C" + row).GetString()).ToString() != "")
                    return true;
            }
            return false;
        }
    }
}
