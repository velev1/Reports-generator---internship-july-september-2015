using System;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;

namespace SS_Reports.XLS_to_XLSX_Converter
{
    /// <summary>
    /// The converter used to make a temporary Excel(.xlsx) file whenever the source file selected has an (.xls) extension.
    /// </summary>
    /// Because ClosedXML cannot operate with (.xls) files.
    class ExcelConverter
    {
        private Application excelApp;
        private Workbook workbook;
        private int pID = -1;
        private string pathToXLSFile;
        internal ExcelConverter(string xlsFile)
        {
            this.pathToXLSFile = xlsFile;
        }
        /// <summary>
        /// The converting function.
        /// </summary>
        /// <returns>Path to the new (.xlsx) file.</returns>
        internal string XLStoXLSX()
        {
            string folderPath = Path.GetTempPath();
            folderPath += Path.GetRandomFileName() + @"\";
            string fileName = Path.GetRandomFileName() + ".xlsx";
            try
            {
                DirectoryInfo newDirectory = Directory.CreateDirectory(folderPath);
                newDirectory.Attributes = FileAttributes.Directory | FileAttributes.Hidden;
                excelApp = new Application();
                HandleRef hwnd = new HandleRef(excelApp, (IntPtr)excelApp.Hwnd);
                GetWindowThreadProcessId(hwnd, out pID);
                excelApp.Visible = false;
                workbook = excelApp.Workbooks.Open(pathToXLSFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                workbook.SaveAs(folderPath + fileName, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                return folderPath + fileName;
            }
            finally
            {
                workbook.Close(false, Type.Missing, Type.Missing);
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
                excelApp = null;
                GC.Collect();
                KillProcess(pID, "EXCEL");
            }
        }

        /// <summary>
        /// Kills a specific process defined by ID and name.
        /// </summary>
        /// <param name="pID">Process ID.</param>
        /// <param name="processName">Process name.</param>
        private void KillProcess(int pID, string processName)
        {
            System.Diagnostics.Process[] AllProcesses = System.Diagnostics.Process.GetProcessesByName(processName);
            foreach (System.Diagnostics.Process process in AllProcesses)
            {
                if (process.Id == pID)
                {
                    process.Kill();
                }
            }
            AllProcesses = null;
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetWindowThreadProcessId(HandleRef handle, out int processId);
    }
}
