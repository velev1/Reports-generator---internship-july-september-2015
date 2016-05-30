using System;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using SS_Reports.File_Managers;
using SS_Reports.XLS_to_XLSX_Converter;

namespace SS_Reports
{
    public partial class Form1 : Form
    {
        /// <summary>
        /// If the form is being closed while the Report worker is still running this blocks the form's thread until the background thread finishes working.
        /// </summary>
        internal static readonly ManualResetEvent allowFormClosingResetEvent = new ManualResetEvent(true);
        public Form1()
        {
            InitializeComponent();
            sourceFilePathBox.GotFocus += HideCaret;
            destinationFileTextBox.GotFocus += HideCaret;
            progressTextBox.GotFocus += HideCaret;
            technopolisRadioButton.Tag = Enums.Stores.Technopolis;
            technomarketRadioButton.Tag = Enums.Stores.Technomarket;
        }

        //Hiding the caret within the text boxes.
        [DllImport("user32.dll")]
        static extern bool HideCaret(IntPtr hWnd);
        public void HideCaret(object sender, EventArgs e)
        {
            HideCaret(((TextBox)sender).Handle);
        }

        /// <summary>
        /// Associated with the source file browse button.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void sourceFileBrowseButton_Click(object sender, EventArgs e)
        {
            if (ReportsManagerWorker.IsBusy == false)
            {
                browseSourceFileDialog.Filter = "Excel Workbook(*.xlsx, *.xls)|*.xlsx;*.xls";
                var sourceFile = browseSourceFileDialog.ShowDialog();
                if (sourceFile == DialogResult.OK)
                {
                    if (sourceFilePathBox.Text != browseSourceFileDialog.FileName)
                    {
                        sourceFilePathBox.Text = browseSourceFileDialog.FileName;
                        progressTextBox.Text = "";
                    }
                }
            }
        }

        /// <summary>
        /// Associated with the destionation file browse button.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void destinationFileBrowseButton_Click(object sender, EventArgs e)
        {
            if (ReportsManagerWorker.IsBusy == false && ReportsManagerWorker.IsBusy == false)
            {
                openDestinationFileDialog.Filter = "Excel Workbook(*.xlsx)|*.xlsx";
                var destFile = openDestinationFileDialog.ShowDialog();
                if (destFile == DialogResult.OK)
                {
                    destinationFileTextBox.Text = openDestinationFileDialog.FileName;
                }
            }
        }

        /// <summary>
        /// Associated with the destination file create new button.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void destinationFileCreateButton_Click(object sender, EventArgs e)
        {
            if (CreateNewFileWorker.IsBusy == false)
            {
                createNewFileDialog.Filter = "Excel Workbook(*.xlsx)|*.xlsx";
                var newFile = createNewFileDialog.ShowDialog();
                if (newFile == DialogResult.OK)
                {
                    openDestinationFileDialog.FileName = createNewFileDialog.FileName;
                    CreateNewFileWorker.RunWorkerAsync(openDestinationFileDialog.FileName);
                    progressTextBox.Text = "Creating new file.";
                    destinationFileTextBox.Text = createNewFileDialog.FileName;
                }
            }
        }
        /// <summary>
        ///  Associated with the source file clear button.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void sourceFileClearButton_Click(object sender, EventArgs e)
        {
            sourceFilePathBox.Clear();
            browseSourceFileDialog.Reset();
        }

        /// <summary>
        /// Associated with the destination file clear button.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void destinationFileClearButton_Click(object sender, EventArgs e)
        {
            destinationFileTextBox.Clear();
            openDestinationFileDialog.Reset();
        }

        /// <summary>
        /// This is the worker's main method.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void NewFileWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            NewFile newFile = new NewFile((string)e.Argument, sender, e);
        }

        /// <summary>
        /// Executes when the worker finishes working.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void NewFileWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                progressTextBox.Text = "Canceled!";
            }
            else if (e.Error != null)
            {
                progressTextBox.Text = "Error: " + e.Error.Message;
            }
            else
            {
                progressTextBox.Text = "New file created!";
            }
        }

        /// <summary>
        /// This is the worker's main method.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ReportsManagerWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            Tuple<string, string, Enums.Stores, bool> @params = e.Argument as Tuple<string, string, Enums.Stores, bool>;
            using (ReportsManager report = new ReportsManager(@params.Item1, @params.Item2, @params.Item3, @params.Item4, (BackgroundWorker)sender, e))
            {
                report.GenerateReport();
            }
        }

        /// <summary>
        /// Executes when the worker finishes working.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ReportsManagerWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                progressTextBox.Text = "Canceled!";
            }
            else if (e.Error != null)
            {
                progressTextBox.Text = "Error: " + e.Error.Message;
            }
            else if (e.Result != null)
            {
                progressTextBox.Text = e.Result.ToString();
            }
            else
            {
                progressTextBox.Text = "Successfully finished!";
            }
        }

        /// <summary>
        /// Associated with the cancel button.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cancelProcessButton_Click(object sender, EventArgs e)
        {
            if (CreateNewFileWorker.IsBusy == true)
            {
                CreateNewFileWorker.CancelAsync();
            }
            else if (ReportsManagerWorker.IsBusy == true)
            {
                ReportsManagerWorker.CancelAsync();
            }
        }

        /// <summary>
        /// Associated with the generate report button.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void generateReportButton_Click(object sender, EventArgs e)
        {
            if (ReportsManagerWorker.IsBusy == false && CreateNewFileWorker.IsBusy == false)
            {
                string errorMessage;
                if (InputRequirementsMet(out errorMessage) == false)
                {
                    progressTextBox.Text = errorMessage;
                    return;
                }
                //Source and output files and selected retailer.
                Enums.Stores selected = (Enums.Stores)retailersGroupBox.Controls.OfType<RadioButton>().FirstOrDefault(r => r.Checked).Tag;
                var @params = Tuple.Create<string, string, Enums.Stores, bool>(browseSourceFileDialog.FileName, openDestinationFileDialog.FileName, selected, subtractCheckBox.Checked);
                ReportsManagerWorker.RunWorkerAsync(@params);
                if (subtractCheckBox.Checked == false)
                    progressTextBox.Text = "Generating report...";
                else
                    progressTextBox.Text = "Subtracting...";
            }
        }

        /// <summary>
        /// Invoked when the form is being closed. 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.TaskManagerClosing || e.CloseReason == CloseReason.UserClosing || e.CloseReason == CloseReason.WindowsShutDown ||
                e.CloseReason == CloseReason.ApplicationExitCall || e.CloseReason == CloseReason.FormOwnerClosing || e.CloseReason == CloseReason.MdiFormClosing)
            {
                if (ReportsManagerWorker.IsBusy == true)
                {
                    ReportsManagerWorker.CancelAsync();
                    allowFormClosingResetEvent.Reset();
                    allowFormClosingResetEvent.WaitOne();
                }
            }
        }

        /// <summary>
        /// Checks if the user has selected a source file, destination file and the source retailer
        /// </summary>
        /// <param name="message"></param>
        /// <returns></returns>
        private bool InputRequirementsMet(out string message)
        {
            message = "";
            if (browseSourceFileDialog.FileName == "")
            {
                message = "Source file not selected.";
                return false;
            }
            if (openDestinationFileDialog.FileName == "")
            {
                message = "Destination file not selected.";
                return false;
            }
            try
            {
                Enums.Stores selected = (Enums.Stores)retailersGroupBox.Controls.OfType<RadioButton>().FirstOrDefault(r => r.Checked).Tag;
            }
            catch (Exception ex)
            {
                message = "Retailer not selected.";
                return false;
            }
            return true;
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
