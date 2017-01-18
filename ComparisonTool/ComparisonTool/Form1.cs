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
        // the excel sheets as arrays of objects
        object[,] xlString1, xlString2;
        // the index of the value, what column it is 
        int valueIndex = 0;
        int dateIndex = 0;
        // size of first xlSheet
        int xl1Size = 0;

        //Dictionary<string, DateTime[]> datesDict = new Dictionary<string, DateTime[]>();
        Dictionary<string, DateTime> datesDict2 = new Dictionary<string, DateTime>();

        int xl1RowCount;
        int xl2RowCount;

        SortedDictionary<string, DateTime[]> datesDict = new SortedDictionary<string, DateTime[]>();

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

        public void readExcelFile(string filename, bool getYears)
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
            xlString1 = values;

            for (int i = 1; i <= colCount; i++)
            {
                colNames[i - 1] = values[1, i].ToString();
                if (colNames[i - 1] == "Value")
                {
                    valueIndex = i - 1;
                }
                else if (colNames[i - 1] == "Date")
                {
                    dateIndex = i - 1;
                }
            }

            xlWorkSheet = xlWorkBook.Sheets[3];
            xlRange = xlWorkSheet.UsedRange;

            int calRowCount = xlRange.Rows.Count;
            int calColCount = xlRange.Columns.Count;

            xl1RowCount = calRowCount;
            values = (object[,])xlRange.Value2;

            if (getYears)
            {
                int NumRow = 1;
                int correction = 0;

                while (NumRow <= values.GetLength(0))
                {
                    if (NumRow != 1)
                    {
                        double d = Convert.ToDouble(values[NumRow, 1]);

                        DateTime[] dates = new DateTime[2];

                        DateTime conv = DateTime.FromOADate(d);

                        dates[0] = conv;
                        dates[1] = new DateTime();
                        dates[1] = DateTime.FromOADate(99999.99);
                        datesDict.Add(values[NumRow, 2].ToString(), dates);

                        correction++;
                    }
                    NumRow++;
                }
            }

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
            /*TODO: Enable and disable create report buttons
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
                nextButton.Enabled = false;
            }*/
        }

        private int[] readExcelFile2(bool getYears)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(filenamesArray[1]);
            Excel.Worksheet xlWorkSheet = xlWorkBook.Sheets[4];
            Excel.Range xlRange = xlWorkSheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            object[,] xlString = new object[rowCount, colCount];
            string[] Fields = new string[colCount];
            object[,] values = (object[,])xlRange.Value2;

            xlString2 = values;

            // grab the fiscal year calendar
            if (getYears)
            {
                // TODO write the code to grab fiscal year ranges
                xlWorkSheet = xlWorkBook.Sheets[3];
                xlRange = xlWorkSheet.UsedRange;

                int calRowCount = xlRange.Rows.Count;
                int calColCount = xlRange.Columns.Count;

                xl2RowCount = calRowCount;
                values = (object[,])xlRange.Value2;
                int NumRow = 1;
                int correction = 0;

                while (NumRow <= values.GetLength(0))
                {
                    if (NumRow != 1)
                    {
                        DateTime[] dates = new DateTime[2];
                        double d = Convert.ToDouble(values[NumRow, 1]);

                        DateTime conv = DateTime.FromOADate(d);

                        if (datesDict.TryGetValue(values[NumRow, 2].ToString(), out dates))
                        {
                            dates[1] = conv;
                            datesDict[values[NumRow, 2].ToString()] = dates;
                        }
                        else
                        {
                            dates[1] = conv;
                            dates[0] = new DateTime();
                            datesDict[values[NumRow, 2].ToString()] = dates;
                        }

                        correction++;
                    }
                    NumRow++;
                }
            }

            xlWorkBook.Close(SaveChanges: false);
            xlApp.Quit();

            int[] rv = new int[2];
            rv[0] = rowCount;
            rv[1] = colCount;
            return rv;
        }

        private void createReportByResult(int rowCount, int colCount)
        {
            Dictionary<string, double[]> resultsDict = new Dictionary<string, double[]>();
            // the delta in value
            double[] resultsVal;
            string name;
            for (int i = 2; i <= rowCount; i++)
            {
                resultsVal = new double[3];
                name = "";
                for (int j = 1; j < colCount - 1; j++)
                {
                    // because of the weird way excel file reading works every index is +1 what it should be
                    if (j != dateIndex + 1)
                    {
                        name += xlString2[i, j].ToString();
                    }

                    if (j < colCount - 2 && j != dateIndex + 1)
                    {
                        name += "_";
                    }
                }
                double[] tryArr;
                double d = Convert.ToDouble(xlString2[i, dateIndex + 1]);
               
                if (!resultsDict.TryGetValue(name, out tryArr))
                {
                    resultsVal[0] = 0.00;
                    resultsVal[1] = Convert.ToDouble(xlString2[i, (valueIndex + 1)]);
                    resultsVal[2] = resultsVal[1];
                    resultsDict.Add(name, resultsVal);
                }
                else
                {
                    tryArr[1] = Convert.ToDouble(xlString2[i, (valueIndex + 1)]);
                    tryArr[2] = tryArr[1];
                    resultsDict[name] = tryArr;
                }
            }

            //count of rows in xl sheet 1 that aren't in xl sheet 2
            int count = 0;
            for (int i = 2; i <= xl1Size; i++)
            {
                name = "";
                for (int j = 1; j < colCount - 1; j++)
                {
                    // because of the weird way excel file reading works every index is +1 what it should be
                    if (j != dateIndex + 1)
                    {
                        name += xlString1[i, j].ToString();
                    }

                    if (j < colCount - 2)
                    {
                        name += "_";
                    }
                }

                double d = Convert.ToDouble(xlString1[i, dateIndex + 1]);
                
                if (resultsDict.TryGetValue(name, out resultsVal))
                {
                    resultsVal[0] = Convert.ToDouble(xlString1[i, valueIndex + 1]);
                    resultsVal[2] = resultsVal[1] - resultsVal[0];
                    resultsDict[name] = resultsVal;
                }
                else
                {
                    resultsVal = new double[3 * (xl1RowCount - 1)];
                    resultsVal[0] = Convert.ToDouble(xlString1[i, valueIndex + 1]);
                    resultsVal[1] = 0;
                    resultsVal[2] = -resultsVal[0];
                    resultsDict.Add(name, resultsVal);
                    count++;
                }
            }

            int finalColSize = colCount + 3;
            object[,] finalResults = new object[rowCount + count, finalColSize];
            string current = "";
            string[] splitString;
            int currRow = 1;
            int currCol = 0;
            foreach (var entry in resultsDict)
            {
                //MessageBox.Show(entry.Key);
                splitString = entry.Key.Split('_');
                currCol = 0;

                for (int i = 0; i < splitString.Length; i++)
                {
                    finalResults[currRow, i] = splitString[i];
                    currCol++;
                }

                current = "";
                for (int i = 0; i < entry.Value.Length; i++)
                {
                    current += entry.Value[i].ToString() + " ";
                    finalResults[currRow, currCol - 1] = entry.Value[i];
                    currCol++;
                }

                currRow++;
            }

            for (int i = 0; i < dateIndex; i++)
            {
                finalResults[0, i] = xlString2[1, i + 1];
            }

            finalResults[0, dateIndex] = "Before";
            finalResults[0, dateIndex + 1] = "After";
            finalResults[0, dateIndex + 2] = "Delta";

            Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = true;
            Excel._Workbook oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
            Excel._Worksheet oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
            var startCell = (Excel.Range)oSheet.Cells[1, 1];
            var endCell = (Excel.Range)oSheet.Cells[rowCount + count, finalColSize];
            var writeRange = oSheet.Range[startCell, endCell];

            writeRange.NumberFormat = "#,##0.00";
            //writeRange.


            writeRange.Value2 = finalResults;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void createReportByYear(int rowCount, int colCount)
        {   
            Dictionary<string, double[]> resultsDict = new Dictionary<string, double[]>();
            // the delta in value
            double[] resultsVal;
            string name;
            for (int i = 2; i <= rowCount; i++)
            {
                resultsVal = new double[3 * (xl2RowCount)];
                name = "";
                for (int j = 1; j < colCount - 1; j++)
                {
                    // because of the weird way excel file reading works every index is +1 what it should be
                    if (j != dateIndex+1)
                    {
                        name += xlString2[i, j].ToString();
                    }

                    if (j < colCount - 2 && j != dateIndex+1)
                    {
                        name += "_";
                    }
                }
                double[] tryArr;
                double d = Convert.ToDouble(xlString2[i, dateIndex+1]);
                int offset = 0;
                int k = 0;
                string[] dateNames = datesDict.Keys.ToArray();

                DateTime conv = DateTime.FromOADate(d);
                DateTime curr = new DateTime();
                curr = DateTime.FromOADate(99999.99);
                foreach (var date in datesDict)
                {
                    if (conv <= date.Value[1])
                    {
                        if (date.Value[1] < curr)
                        {
                            curr = date.Value[1];
                            offset = k;
                        }
                    }
                    /*if (k == 0 && conv <= date.Value[1])
                    {
                        offset = k;
                    }
                    else if (conv <= date.Value[1] && conv > datesDict[dateNames[k - 1]][1])
                    {
                        offset = k;
                    }*/

                    k++;
                }
                //MessageBox.Show(offset.ToString() + " " + datesDict.Count.ToString());
                
                if (! resultsDict.TryGetValue(name, out tryArr))
                {
                    resultsVal[offset*3 + 1] = Convert.ToDouble(xlString2[i, (valueIndex + 1)]);
                    resultsVal[offset*3 + 2] = resultsVal[offset*3 + 1];
                    resultsDict.Add(name, resultsVal);
                }
                else
                {
                    tryArr[offset*3 + 1] = Convert.ToDouble(xlString2[i, (valueIndex + 1)]);
                    tryArr[offset*3 + 2] = tryArr[offset*3 + 1];
                    resultsDict[name] = tryArr;
                }
            }

            //count of rows in xl sheet 1 that aren't in xl sheet 2
            int count = 0;
            for (int i = 2; i <= xl1Size; i++)
            {
                name = "";
                for (int j = 1; j < colCount - 1; j++)
                {
                    // because of the weird way excel file reading works every index is +1 what it should be
                    if (j != dateIndex + 1)
                    {
                        name += xlString1[i, j].ToString();
                    }

                    if (j < colCount - 2)
                    {
                        name += "_";
                    }
                }

                double d = Convert.ToDouble(xlString1[i, dateIndex + 1]);
                int offset = 0;
                int k = 0;
                string[] dateNames = datesDict.Keys.ToArray();

                DateTime conv = DateTime.FromOADate(d);
                DateTime curr = new DateTime();
                curr = DateTime.FromOADate(99999.99);

                foreach (var date in datesDict)
                {
                    if (conv <= date.Value[0])
                    {
                        if (date.Value[0] < curr)
                        {
                            curr = date.Value[0];
                            offset = k;
                        }
                    }

                    k++;
                }


                if (resultsDict.TryGetValue(name, out resultsVal))
                {
                    resultsVal[offset*3] = Convert.ToDouble(xlString1[i, valueIndex + 1]);
                    resultsVal[offset*3 + 2] = resultsVal[offset*3 + 1] - resultsVal[offset*3];
                    resultsDict[name] = resultsVal;                  
                }
                else
                {
                    resultsVal = new double[3 * (xl1RowCount-1)];
                    resultsVal[offset*3] = Convert.ToDouble(xlString1[i, valueIndex + 1]);
                    resultsVal[offset*3 + 1] = 0;
                    resultsVal[offset*3 + 2] = -resultsVal[offset];
                    resultsDict.Add(name, resultsVal);
                    count++;
                }
            }

            int finalColSize = xl1RowCount * colCount * 3 + colCount;
            object[,] finalResults = new object[rowCount+count, finalColSize];
            string current = "";
            string[] splitString;
            int currRow = 1;
            int currCol = 0;
            foreach (var entry in resultsDict)
            {

                splitString = entry.Key.Split('_');
                currCol = 0;

                for (int i = 0; i < splitString.Length; i++)
                {
                    finalResults[currRow, i] = splitString[i];
                    currCol++;
                }

                current = "";
                for (int i = 0; i < entry.Value.Length; i++)
                {
                    current += entry.Value[i].ToString() + " ";
                    finalResults[currRow, currCol - 1] = entry.Value[i];
                    currCol++;
                }

                currRow++;
            }

            for (int i = 0; i < dateIndex; i++)
            {
                finalResults[0, i] = xlString2[1, i+1];
            }

            int c = 0;
            foreach (var date in datesDict)
            {
                finalResults[0, dateIndex + (c * 3)] = date.Key + " Before";
                finalResults[0, (dateIndex) + 1 + c * 3] = date.Key + " After";
                finalResults[0, dateIndex + (c * 3) + 2] = date.Key + " Delta";
                c++;
            }

            Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = true;
            Excel._Workbook oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
            Excel._Worksheet oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
            var startCell = (Excel.Range)oSheet.Cells[1, 1];
            var endCell = (Excel.Range)oSheet.Cells[rowCount + count, finalColSize];
            var writeRange = oSheet.Range[startCell, endCell];

            writeRange.NumberFormat = "#,##0.00";
            //writeRange.


            writeRange.Value2 = finalResults;
           
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void createReportButton_Click(object sender, EventArgs e)
        {
            readExcelFile(openFile1Dialog.FileName, false);
            int[] counts = readExcelFile2(false);
            createReportByResult(counts[0], counts[1]);

            this.Close();
        }

        private void byYearButton_Click(object sender, EventArgs e)
        {
            readExcelFile(openFile1Dialog.FileName, true);
            int[] counts = readExcelFile2(true);
            createReportByYear(counts[0], counts[1]);

            this.Close();
        }

        private void column4Box_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            //uniqueNames[3] = column4Box.Text;
            //uniqueNameIndices[3] = column4Box.SelectedIndex;
            checkUniqueNames();
        }

        private void column2Box_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            //uniqueNames[1] = column2Box.Text;
            //uniqueNameIndices[1] = column2Box.SelectedIndex;
            checkUniqueNames();
        }

        private void column3Box_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            //uniqueNames[2] = column3Box.Text;
            //uniqueNameIndices[2] = column3Box.SelectedIndex;
            checkUniqueNames();
        }

        private void column1Box_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            //uniqueNames[0] = column1Box.Text;
            //uniqueNameIndices[0] = column1Box.SelectedIndex;
            checkUniqueNames();
        }

        private void helpButton_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Select the titles of the rows to compare against. The values of the rows at these column names should be the same.\n\nThis will be how the program decides which values to get a delta from. It is also how we create the report.");
        }
    }
}
