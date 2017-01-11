using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

/* Josh Beck, CRC Technologies 2017
 * Comparison Tool
 * 
 * The purpose of this tool is to take two reports from Cobra and create a comparison report.
 * The two reports should be nearly identical with the difference being the value (whether that is in $ or otherwise).
 * The tool will take the two reports and create a report with the deltas.
 */

namespace ComparisonTool
{
    public partial class Form1 : Form
    {
        // the current panel we are on.
        int currentPanel = 0;

        // Array of panels that will have the different pages of the app
        Panel[] panelsArray = new Panel[2];
        // will have both file names
        string[] filenamesArray = new string[2];
        // the names of the columns used to match
        string[] uniqueNames = new string[4];
        // the indices of the unique names, names array may not be necessary
        int[] uniqueNameIndices = new int[4];
        // the two xlRanges
        Excel.Range[] xlRanges = new Excel.Range[2];
        // the excel sheets as arrays of objects
        object[,] xlString1, xlString2;
        // the index of the value, what column it is 
        int valueIndex = 0;
        int dateIndex = 0;
        // size of first xlSheet
        int xl1Size = 0;

        public Form1()
        {
            InitializeComponent();
            panelsArray[0] = filesPanel;
            panelsArray[1] = columnsPanel;
            filesPanel.Show();
            columnsPanel.Hide();
            //loadingPanel.Hide();
            backButton.Enabled = false;
            nextButton.Enabled = false;
        }

        public void readExcelFile(string filename, int fileNumber)
        {
            // set up to read in the xl file
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(filename);
            Excel.Worksheet xlWorkSheet = xlWorkBook.Sheets[4];
            Excel.Range xlRange = xlWorkSheet.UsedRange;

            string[] colNames;
            object[,] xlString;
            string[] Fields;
            object[,] values;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            // save the row count globally
            xl1Size = rowCount;

            // need to put column names in the selection boxes, so clear them out
            column1Box.Items.Clear();
            column2Box.Items.Clear();
            column3Box.Items.Clear();
            column4Box.Items.Clear();

            colNames = new string[colCount];
            xlString = new object[rowCount, colCount];
            Fields = new string[colCount];
            values = (object[,])xlRange.Value2;


            int NumRow = 1;
            while (NumRow <= values.GetLength(0))
            {
                for (int c = 1; c <= colCount; c++)
                {
                    // had to do this to store dates as a date object
                    if (c - 1 == dateIndex && NumRow != 1)
                    {
                        double d = Convert.ToDouble(values[NumRow, c]);

                        DateTime conv = DateTime.FromOADate(d);

                        xlString[NumRow - 1, c - 1] = conv.ToString(@"MM\/dd\/yyyy");
                    }
                    else
                    {
                        // if it's the first row, it's the column names
                        if (NumRow == 1)
                        {
                            colNames[c - 1] = values[NumRow, c].ToString();
                            if (colNames[c-1] == "Value")
                            {
                                valueIndex = c-1;
                            }
                            else if (colNames[c-1] == "Date")
                            {
                                dateIndex = c-1;
                            }
                        }
                        xlString[NumRow - 1, c - 1] = values[NumRow, c];
                    }
                }
                NumRow++;
            }
            xlString1 = xlString;


            /*System.Array myvalues;
            string[] colNames;
  
            myvalues = (System.Array)xlRange.Rows[1].Cells.Value;
            colNames = myvalues.OfType<object>().Select(o => o.ToString()).ToArray();

            for (int i = 0; i < colNames.Length; i++)
            {
                if (colNames[i] == "Value")
                {
                    valueIndex = i;
                }
                else if(colNames[i] == "Date")
                {
                    dateIndex = i;
                }
            }*/

            column1Box.Items.AddRange(colNames);
            column2Box.Items.AddRange(colNames);
            column3Box.Items.AddRange(colNames);
            column4Box.Items.AddRange(colNames);

            GC.Collect();
            GC.WaitForPendingFinalizers();

            xlWorkBook.Close(SaveChanges: false);
            xlApp.Quit();
        }

        private void changePanel(int currentIndex, int previousIndex)
        {
            if (currentIndex < panelsArray.Length && currentIndex > -1)
            {
                Panel curr = panelsArray[currentIndex];

                Panel prev = panelsArray[previousIndex];
               
                currentPanel = currentIndex;
                if (currentPanel == 0)
                {
                    curr.Visible = true;
                    prev.Visible = false;
                    backButton.Enabled = false;
                    if (file1Label.Text != "None" && file2Label.Text != "None")
                    {
                        nextButton.Enabled = true;
                    }
                    else
                    {
                        nextButton.Enabled = false;
                    }
                }
                else
                {
                    nextButton.Enabled = false;
                    curr.Visible = true;
                    prev.Visible = false;
                    backButton.Enabled = true;
                }
            }
        }

        private void openHelpButton_Click(object sender, EventArgs e)
        {
        
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void columnsPanel_Paint(object sender, PaintEventArgs e)
        {

        }

        private void nextButton_Click(object sender, EventArgs e)
        {
            readExcelFile(openFile1Dialog.FileName, 1);
            changePanel(currentPanel + 1, currentPanel);
        }

        private void filesParent_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void backButton_Click(object sender, EventArgs e)
        {
            changePanel(currentPanel - 1, currentPanel);
        }

        private void columnsPanel_Paint_1(object sender, PaintEventArgs e)
        {

        }

        private void file1SelectButton_Click(object sender, EventArgs e)
        {
            if (openFile1Dialog.ShowDialog() == DialogResult.OK)
            {
                /* add the filename to the array of excel filenames, 
                 * grab the last bit of the filename to set the name on the label for filename */
                filenamesArray[0] = openFile1Dialog.FileName;
                string[] filenameArray = openFile1Dialog.FileName.Split('\\');
                string tmp = filenameArray[filenameArray.Length - 1];
                file1Label.Text = tmp;
                if (file1Label.Text != "None" && file2Label.Text != "None")
                {
                    nextButton.Enabled = true;
                }
                else
                {
                    nextButton.Enabled = false;
                }
            }
        }

        private void file2SelectButton_Click(object sender, EventArgs e)
        {
            if (openFile2Dialog.ShowDialog() == DialogResult.OK)
            {
                /* add the filename to the array of excel filenames, 
                 * grab the last bit of the filename to set the name on the label for filename */
                filenamesArray[1] = openFile2Dialog.FileName;
                string[] filenameArray = openFile2Dialog.FileName.Split('\\');
                string tmp = filenameArray[filenameArray.Length - 1];
                file2Label.Text = tmp;
                if (file1Label.Text != "None" && file2Label.Text != "None")
                {
                    nextButton.Enabled = true;
                }
                else
                {
                    nextButton.Enabled = false;
                }
            }
        }

        private void checkUniqueNames()
        {
            //TODO: Enable and disable create report buttons
            bool good = true;
            for (int i = 0; i < uniqueNames.Length; i++)
            {
                if (uniqueNames[i] == null)
                {
                    good = false;
                }
            }

            if (good)
            {
                nextButton.Enabled = true;
            }
            else
            {
                //MessageBox.Show("not good");
                nextButton.Enabled = false;
            }
        }

        private int[] readExcelFile2()
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(filenamesArray[1]);
            Excel.Worksheet xlWorkSheet = xlWorkBook.Sheets[4];
            Excel.Range xlRange = xlWorkSheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            string[] filenameArray = filenamesArray[1].Split('\\');
            string tmp = filenameArray[filenameArray.Length - 1];

            object[,] xlString = new object[rowCount, colCount];
            string[] Fields = new string[colCount];
            object[,] values = (object[,])xlRange.Value2;
            int NumRow = 1;
            while (NumRow <= values.GetLength(0))
            {
                for (int c = 1; c <= colCount; c++)
                {
                    if (c - 1 == dateIndex && NumRow != 1)
                    {
                        double d = Convert.ToDouble(values[NumRow, c]);

                        DateTime conv = DateTime.FromOADate(d);

                        xlString[NumRow - 1, c - 1] = conv.ToString(@"MM\/dd\/yyyy");
                    }
                    else
                    {
                        xlString[NumRow - 1, c - 1] = values[NumRow, c];
                    }
                }
                NumRow++;
            }
            xlString2 = xlString;

            xlWorkBook.Close(SaveChanges: false);
            xlApp.Quit();

            int[] rv = new int[2];
            rv[0] = rowCount;
            rv[1] = colCount;
            return rv;
        }

        private void createReportButton_Click(object sender, EventArgs e)
        {
            int[] counts = readExcelFile2();
            createReportByResult(counts[0], counts[1]);

            this.Close();
        }

        private void createReportByResult(int rowCount, int colCount)
        {
            Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = true;

            // the delta in value
            double difference = 0.0;
            object[,] finalResults = new object[rowCount, colCount];
            for (int i = 1; i < rowCount; i++)
            {
                string[] match = new string[colCount];
                //MessageBox.Show(valueIndex.ToString() + " " + xlString2[i, valueIndex]);
                difference = Convert.ToDouble(xlString2[i, valueIndex]);
                for (int j = 1; j < xl1Size; j++)
                {
                    /*MessageBox.Show(j + " " + xlString1[j, 0] + " " + xlString1[j, 1] + " " + xlString1[j, 2] + " " + xlString1[j, 3] + " " + xlString1[j, 4]
                        + "\n" + i + " " + xlString2[i, 0] + " " + xlString2[i, 1] + " " + xlString2[i, 2] + " " + xlString2[i, 3] + " " + xlString2[i, 4]);
                    */
                    if (xlString1[j, uniqueNameIndices[0]].ToString() == xlString2[i, uniqueNameIndices[0]].ToString() && xlString1[j, uniqueNameIndices[1]].ToString() == xlString2[i, uniqueNameIndices[1]].ToString()
                        && xlString1[j, uniqueNameIndices[2]].ToString() == xlString2[i, uniqueNameIndices[2]].ToString() && xlString1[j, uniqueNameIndices[3]].ToString() == xlString2[i, uniqueNameIndices[3]].ToString())
                    {
                        difference = difference - Convert.ToDouble(xlString1[j, valueIndex]);
                        break;
                    }
                }
                for (int j = 0; j < colCount; j++)
                {
                    if (valueIndex == j)
                    {
                        finalResults[i - 1, j] = difference;
                    }
                    else
                    {
                        finalResults[i - 1, j] = xlString2[i, j];
                    }
                }
            }

            Excel._Workbook oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
            Excel._Worksheet oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
            var startCell = (Excel.Range)oSheet.Cells[1, 1];
            var endCell = (Excel.Range)oSheet.Cells[rowCount, colCount];
            var writeRange = oSheet.Range[startCell, endCell];

            writeRange.Value2 = finalResults;

            //MessageBox.Show("Success");

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void column4Box_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            uniqueNames[3] = column4Box.Text;
            uniqueNameIndices[3] = column4Box.SelectedIndex;
            checkUniqueNames();
        }

        private void column2Box_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            uniqueNames[1] = column2Box.Text;
            uniqueNameIndices[1] = column2Box.SelectedIndex;
            checkUniqueNames();
        }

        private void column3Box_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            uniqueNames[2] = column3Box.Text;
            uniqueNameIndices[2] = column3Box.SelectedIndex;
            checkUniqueNames();
        }

        private void column1Box_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            uniqueNames[0] = column1Box.Text;
            uniqueNameIndices[0] = column1Box.SelectedIndex;
            checkUniqueNames();
        }

        private void byYearButton_Click(object sender, EventArgs e)
        {
            int[] counts = readExcelFile2();
            createReportByYear(counts[0], counts[1]);

            this.Close();
        }

        private void createReportByYear(int rowCount, int colCount)
        {
            Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = true;

            // the delta in value
            double difference = 0.0;
            object[,] finalResults = new object[rowCount, colCount];
            for (int i = 1; i < rowCount; i++)
            {
                string[] match = new string[colCount];
                //MessageBox.Show(valueIndex.ToString() + " " + xlString2[i, valueIndex]);
                difference = Convert.ToDouble(xlString2[i, valueIndex]);
                for (int j = 1; j < xl1Size; j++)
                {
                    /*MessageBox.Show(j + " " + xlString1[j, 0] + " " + xlString1[j, 1] + " " + xlString1[j, 2] + " " + xlString1[j, 3] + " " + xlString1[j, 4]
                        + "\n" + i + " " + xlString2[i, 0] + " " + xlString2[i, 1] + " " + xlString2[i, 2] + " " + xlString2[i, 3] + " " + xlString2[i, 4]);
                    */
                    if (xlString1[j, uniqueNameIndices[0]].ToString() == xlString2[i, uniqueNameIndices[0]].ToString() && xlString1[j, uniqueNameIndices[1]].ToString() == xlString2[i, uniqueNameIndices[1]].ToString()
                        && xlString1[j, uniqueNameIndices[2]].ToString() == xlString2[i, uniqueNameIndices[2]].ToString() && xlString1[j, uniqueNameIndices[3]].ToString() == xlString2[i, uniqueNameIndices[3]].ToString())
                    {
                        difference = difference - Convert.ToDouble(xlString1[j, valueIndex]);
                        break;
                    }
                }
                for (int j = 0; j < colCount; j++)
                {
                    if (valueIndex == j)
                    {
                        finalResults[i - 1, j] = difference;
                    }
                    else
                    {
                        finalResults[i - 1, j] = xlString2[i, j];
                    }
                }
            }

            Excel._Workbook oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
            Excel._Worksheet oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
            var startCell = (Excel.Range)oSheet.Cells[1, 1];
            var endCell = (Excel.Range)oSheet.Cells[rowCount, colCount];
            var writeRange = oSheet.Range[startCell, endCell];

            writeRange.Value2 = finalResults;

            //MessageBox.Show("Success");

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void helpButton_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Select the titles of the rows to compare against. The values of the rows at these column names should be the same.\n\nThis will be how the program decides which values to get a delta from. It is also how we create the report.");
        }
    }
}
