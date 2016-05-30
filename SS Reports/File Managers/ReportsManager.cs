using System;
using System.IO;

using SS_Reports.Stores;
using System.ComponentModel;

namespace SS_Reports.File_Managers
{
    /// <summary>
    /// Reports.
    /// </summary>
    class ReportsManager : IDisposable
    {
        BackgroundWorker reportsWorker;
        //Worker events
        DoWorkEventArgs workerArgs;

        private string sourceFile;
        private string outputFile;
        private bool subtractData;
        private Enums.Stores store;

        /// <summary>
        /// Creates the reports manager.
        /// </summary>
        /// <param name="sourceFile">Source file.</param>
        /// <param name="outputFile">Output file</param>
        public ReportsManager(string sourceFile, string outputFile, Enums.Stores store, bool subtractData, BackgroundWorker reportsWorker, DoWorkEventArgs workerArgs)
        {
            if (Path.GetExtension(sourceFile) == ".xls")
            {
                XLS_to_XLSX_Converter.ExcelConverter converter = new XLS_to_XLSX_Converter.ExcelConverter(sourceFile);
                this.sourceFile = converter.XLStoXLSX();
            }
            else
                this.sourceFile = sourceFile;

            this.outputFile = outputFile;
            this.store = store;
            this.subtractData = subtractData;
            this.reportsWorker = reportsWorker;
            this.workerArgs = workerArgs;
        }

        /// <summary>
        /// Generates the report based on the selected retailer and the input file.
        /// </summary>
        public void GenerateReport()
        {
            StoreCore storeReports = null;
            switch (store)
            {
                case Enums.Stores.Technopolis:
                    storeReports = new StoreTechnopolis(sourceFile, outputFile, subtractData);
                    break;
                case Enums.Stores.Technomarket:
                    storeReports = new StoreTechnomarket(sourceFile, outputFile, subtractData);
                    break;
            }
            bool notCancelled = storeReports.Report(reportsWorker.CancellationPending);
            if (notCancelled == false)
                workerArgs.Cancel = true;
            Form1.allowFormClosingResetEvent.Set();
        }
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (Path.GetDirectoryName(Path.GetDirectoryName(sourceFile)) + "\\" == Path.GetTempPath())
                    Directory.Delete(Path.GetDirectoryName(sourceFile), true);
            }
        }
    }
}
