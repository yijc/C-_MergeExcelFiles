using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

/// <summary>
/// Merge multiple Microsoft Office Excel files
/// refer to http://stackoverflow.com/questions/7271771/merging-multiple-excel-files-into-one
/// http://stackoverflow.com/questions/27285615/how-do-i-merge-multiple-excel-files-to-a-single-excel-file
/// http://stackoverflow.com/questions/7568613/how-to-merge-two-excel-files-into-one-with-their-sheet-names
/// </summary>
namespace MergeExcel
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();

            // the default button is OK
            this.AcceptButton = OKBtn;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void SourceFileBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "All Excel Files(*.xlsx;*.xls)|*.xlsx;*.xls";

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                string fileName = dlg.FileName;
                if (m_srcFileCount < 3 && fileName.Length > 0)
                {
                    m_srcExcelFiles[m_srcFileCount] = fileName;
                    m_srcFileCount++;
                    UpdateSourceFileList();
                }
            }
        }

        private void TargetFileBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "All Excel Files(*.xlsx;*.xls)|*.xlsx;*.xls";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                string fileName = dlg.FileName;
                if (fileName.Length > 0)
                {
                    m_trgExcelFile = fileName;
                    TargFileTextbox.Text = fileName;
                }
            }

        }

        private void TargFileTextbox_TextChanged(object sender, EventArgs e)
        {
            m_trgExcelFile = TargFileTextbox.Text;
        }

        private void KeyColumnTextbox_TextChanged(object sender, EventArgs e)
        {
            m_keyColumnName1 = textBox_KeyColumn1.Text;
            m_keyColumnName2 = textBox_KeyColumn2.Text;
        }

        private void OKBtn_Click(object sender, EventArgs e)
        {
            UInt16 nSourceFileCount = 0;
            foreach(string fileName in m_srcExcelFiles)
            {
                if (fileName.Length == 0)
                    break;
                nSourceFileCount++;
            }
            if (nSourceFileCount < 2)
            {
                MessageBox.Show("Please select at least two source Excel files!", "Miss source Excel file(s)",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (m_trgExcelFile.Length == 0)
            {
                MessageBox.Show("Please provide a target Excel file!", "Miss target Excel file",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (m_keyColumnName1.Length == 0 || m_keyColumnName2.Length == 0)
            {
                MessageBox.Show("Please fill the key column name!", "key column name is empty",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            bool bSucceed = false;
            if (m_bDoMergeColumn && DoMergeColumnExcelFiles())
            {
                bSucceed = true;
                MessageBox.Show("Merge Excel files Rows succeed!", "Success");
            }
            else if (!m_bDoMergeColumn && DoMergeRowExcelFiles())
            {
                bSucceed = true;
                MessageBox.Show("Merge Excel files columns succeed!", "Success");
            }
            else
            {
                bSucceed = false;
                MessageBox.Show(m_errorMsg, "Merge Failed");
            }

            // clear Excel files
            if (bSucceed)
            {
                m_srcExcelFiles[0] = "";
                m_srcExcelFiles[1] = "";
                m_srcExcelFiles[2] = "";
                m_trgExcelFile = "";
                m_srcFileCount = 0;
                UpdateSourceFileList();
                TargFileTextbox.Text = m_trgExcelFile;
                TargFileTextbox.Refresh();
            }
            m_errorMsg = "";
        }

        private void CancelBtn_Click(object sender, EventArgs e)
        {
            // Exit
            Application.Exit();
        }


        private void radioButton_Columns_CheckedChanged(object sender, EventArgs e)
        {
            m_bDoMergeColumn = radioButton_Columns.Checked;
        }
        private void radioButton_Rows_CheckedChanged(object sender, EventArgs e)
        {
            //m_bDoMergeColumn = radioButton_Rows.Checked;
        }

        private void UpdateSourceFileList()
        {
            SourceFileCheckedListBox.Items.Clear();
            UInt16 nSourceFileCount = 0;
            foreach (string fileName in m_srcExcelFiles)
            {
                if (fileName.Length > 0)
                {
                    SourceFileCheckedListBox.Items.Add(fileName);
                    SourceFileCheckedListBox.SetItemChecked(nSourceFileCount, true);
                    nSourceFileCount++;
                }
            }
            SourceFileCheckedListBox.Refresh();
        }

        private bool DoMergeRowExcelFiles()
        {
            {
                Exception e = new System.ArgumentException("Don't support doing merge by rows!!");
                throw e;
            }

            bool result = true;
            Excel.Application excelApp = null;
            Excel.Workbooks workBooks = null;

            Excel.Workbook wbSource = null;
            Excel.Worksheet wsSource = null;

            Excel.Workbook wbTarget = null;
            Excel.Worksheet wsTarget = null;

            string fileTemplate = @"C:\template_excel.xlsx";
            string firstSourceFile = m_srcExcelFiles[0];

            try //try to open it. If its a proper excel file
            {
                // create Excel application
                excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false;
                // create one new sheet
                excelApp.SheetsInNewWorkbook = 1;

                workBooks = excelApp.Workbooks;

                //Create target workbook
                wbTarget = workBooks.Open(fileTemplate);
                wsTarget = (Excel.Worksheet)wbTarget.Worksheets.get_Item(1);

                // open the first source file
                wbSource = workBooks.Open(firstSourceFile);
                // So far, assume each work book has only one sheet
                wsSource = (Excel.Worksheet)wbSource.Worksheets.get_Item(1);
                // copy the first source sheet to the target
                wsSource.Copy(wsTarget);
                //get the copied sheet
                wsTarget = wbTarget.Worksheets.get_Item(1);

                // looking for the "NationalID" and "PatientID"
                const string nationalID = "NationalID";
                const string patientID = "PatientID";
                int nationalIDColumn = -1;
                int patientIDColumn = -1;
                //Excel.Range firstRow = wsTarget.get_Range("1A","1K");
                //int count = columns.Count;
                for (int i = 1; i < 10; ++i)
                {
                    string strCellVal = "";
                    object cellVal = wsTarget.Cells[1, i].value;
                    if (cellVal != null)
                        strCellVal = cellVal.ToString();
                    if (nationalID.CompareTo(cellVal) == 0)
                    {
                        nationalIDColumn = i;
                    }
                    else if (patientID.CompareTo(cellVal) == 0)
                    {
                        patientIDColumn = i;
                    }
                }

                // looking for the rows in target, skip the first row
                List<string> nationalIDList = new List<string>();
                Excel.Range targetRange = wsTarget.UsedRange;
                int nRowsCount = targetRange.Rows.Count;
                for (int nIndex = 2; nIndex <= nRowsCount; ++nIndex)
                {
                    Excel.Range currentRow = targetRange.Rows[nIndex];
                    Excel.Range currentRowCells = currentRow.Cells;
                    string strID = currentRowCells[nationalIDColumn].value as string;
                    if (strID == null)
                        nationalIDList.Add(string.Empty);
                    else
                        nationalIDList.Add(strID);
                }

                //////////////////////////////
                // Never use 2 dots with com objects.
                // refer to http://stackoverflow.com/questions/158706/how-to-properly-clean-up-excel-interop-objects
                // do row merge
                for (int i = 1; i < m_srcExcelFiles.Length; ++i)
                {
                    Excel.Workbook wbTemp = null;
                    Excel.Worksheet wsTemp = null;
                    Excel.Range usedRange = null;

                    string fileName = m_srcExcelFiles[i];
                    if (fileName == null || fileName.Length == 0)
                        continue;
                    try
                    {
                        wbTemp = workBooks.Open(fileName);
                        wsTemp = (Excel.Worksheet)wbTemp.Worksheets.get_Item(1);
                        usedRange = wsTemp.UsedRange;
                        if (usedRange != null)
                        {
                            int nRows = usedRange.Rows.Count;
                            int nCols = usedRange.Columns.Count;
                            Excel.Range rows = usedRange.Rows;
                            Excel.Range firstRow = rows[1];
                            Excel.Range firstRowCells = firstRow.Cells;

                            // Get the ID column position
                            //foreach (Excel.Range cell in usedRange.Rows)
                            int nationalIDSrc = -1;
                            //int patientIDSrc = -1;
                            for (int j=1; j< nCols; ++j)
                            {
                                string value = firstRowCells[j].value as string;
                                if (nationalID.CompareTo(value) == 0)
                                {
                                    nationalIDSrc = j;
                                    break;
                                }
                                //else if (patientID.CompareTo(value) == 0)
                                //{
                                //    patientIDSrc = j;
                                //}
                            }

                            // skip the first row as a table header
                            for (int k=2; k <= nRows; ++k)
                            {
                                Excel.Range thisRow = rows[k];
                                Excel.Range thisRowCells = thisRow.Cells;
                                string nationalIDVal = thisRowCells[nationalIDSrc].value as string;
                                if (nationalIDVal == null || nationalIDVal.Length == 0)
                                    continue;

                                Excel.Range insertRow = null;
                                int targetRowIndex = 2;


                                // looking for the rows in target, skip the first row
                                //Excel.Range targetRange = wsTarget.UsedRange;
                                //int nRowsCount = targetRange.Rows.Count;
                                //bool bFound = false;
                                //for(targetRowIndex = 2; targetRowIndex <= nRowsCount; ++targetRowIndex)
                                //{
                                //    Excel.Range lookupRow = targetRange.Rows[targetRowIndex];
                                //    Excel.Range lookupRowCells = lookupRow.Cells;
                                //    string lookupID = lookupRowCells[nationalIDColumn].value as string;
                                //    if (nationalIDVal.CompareTo(lookupID) == 0)
                                //    {
                                //        bFound = true;
                                //    }
                                //    else if (bFound)
                                //    {
                                //        // find out the last row has the same ID
                                //        break;
                                //    }
                                //}
                                int nLastIndex = nationalIDList.FindLastIndex(id => id == nationalIDVal);
                                if (nLastIndex == -1)
                                {
                                    targetRowIndex = nationalIDList.Count + 2;
                                }
                                else
                                {
                                    // move to the next and add the table header into account.
                                    // and rows index is 1 based
                                    targetRowIndex = nLastIndex + 3;
                                }

                                /////////////////////////////////
                                // Insert a new row 
                                //insertRow = wsTarget.get_Range(wsTarget.Cells[targetRowIndex,1], wsTarget.Cells[targetRowIndex,nCols]);
                                Excel.Range tempCell = wsTarget.Cells[targetRowIndex, 1];
                                insertRow = tempCell.EntireRow;
                                if (insertRow != null)
                                {
                                    // Insert a new row
                                    insertRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                                    // Get the new row
                                    Excel.Range targetRows = wsTarget.Rows;
                                    Excel.Range newRow = targetRows[targetRowIndex];
                                    Excel.Range newRowCells = newRow.Cells;

                                    thisRowCells.Copy(newRowCells);
                                    nLastIndex++;
                                    if (nLastIndex == -1 || nLastIndex>nationalIDList.Count)
                                        nationalIDList.Add(nationalIDVal);
                                    else
                                        nationalIDList.Insert(nLastIndex, nationalIDVal);
                                    // Assign values
                                    //for (int nColumn =1; nColumn < nCols; ++nColumn)
                                    //{
                                    //    string value = thisRowCells[nColumn].Value2 as string;
                                    //    newRowCells[nColumn].Value2 = value;
                                    //}
                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        throw e;
                    }
                    finally
                    {
                        if (wbTemp != null)
                        {
                            wbTemp.Close();
                        }
                        //if (rows != null)
                        //{
                        //    rows.close();
                        //}

                        Marshal.ReleaseComObject(wbTemp);
                        Marshal.ReleaseComObject(wsTemp);
                    }
                }

                // save the result
                wbTarget.SaveCopyAs(m_trgExcelFile);
            }//end of try
            catch (Exception e)
            {
                m_errorMsg = e.Message;
                result = false;
            }
            finally
            {
                if (wbSource != null)
                {
                    wbSource.Close();
                }
                if (wbTarget != null)
                {
                    wbTarget.Close();
                }
                if (workBooks != null)
                {
                    workBooks.Close();
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                }

                Marshal.ReleaseComObject(excelApp);
                Marshal.ReleaseComObject(workBooks);
                Marshal.ReleaseComObject(wbSource);
                Marshal.ReleaseComObject(wsSource);
                Marshal.ReleaseComObject(wbTarget);
                Marshal.ReleaseComObject(wsTarget);
            }
            return result;
        }

        private bool DoMergeColumnExcelFiles()
        {
            bool result = true;
            Excel.Application excelApp = null;
            Excel.Workbooks workBooks = null;

            Excel.Workbook wbSource = null;
            Excel.Worksheet wsSource = null;

            Excel.Workbook wbTarget = null;
            Excel.Worksheet wsTarget = null;

            string fileTemplate = @"C:\template_excel.xlsx";
            string firstSourceFile = m_srcExcelFiles[0];

            //MessageBox.Show("Do Merge Columns!", "Success");
            try //try to open it. If its a proper excel file
            {
                // create Excel application
                excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false;
                // create one new sheet
                excelApp.SheetsInNewWorkbook = 1;

                workBooks = excelApp.Workbooks;

                //MessageBox.Show("Start  Excel Application succeed!", "Success");

                //Create target workbook
                wbTarget = workBooks.Open(fileTemplate);
                wsTarget = (Excel.Worksheet)wbTarget.Worksheets.get_Item(1);

                //MessageBox.Show("Open source Excel file succeed!", "Success");

                // open the first source file
                wbSource = workBooks.Open(firstSourceFile);
                // So far, assume each work book has only one sheet
                wsSource = (Excel.Worksheet)wbSource.Worksheets.get_Item(1);
                // copy the first source sheet to the target
                wsSource.Copy(wsTarget);
                //get the copied sheet
                wsTarget = wbTarget.Worksheets.get_Item(1);

                // looking for the "Name" and "Receiver1"
                string patientName = m_keyColumnName1;
                int patientNameColumn = -1;
                //Excel.Range firstRow = wsTarget.get_Range("1A","1K");
                //int count = columns.Count;
                for (int i = 1; i < 10; ++i)
                {
                    string strCellVal = "";
                    object cellVal = wsTarget.Cells[1, i].value;
                    if (cellVal != null)
                        strCellVal = cellVal.ToString();
                    if (patientName.CompareTo(cellVal) == 0)
                    {
                        patientNameColumn = i;
                        break;
                    }
                }
                if (patientNameColumn == -1)
                {
                    Exception e = new System.ArgumentException("Invalid key column's name 1 : " + patientName);
                    throw e;
                }
                
                // looking for the rows in target, skip the first row
                List<string> patientNameList = new List<string>();
                Excel.Range targetRange = wsTarget.UsedRange;
                int nTargetRowsCount = targetRange.Rows.Count;
                int ntargetColumnCount = targetRange.Columns.Count;
                for (int nIndex = 2; nIndex <= nTargetRowsCount; ++nIndex)
                {
                    Excel.Range currentRow = targetRange.Rows[nIndex];
                    Excel.Range currentRowCells = currentRow.Cells;
                    string strName = currentRowCells[patientNameColumn].value as string;
                    if (strName != null && strName.Length > 0)
                        patientNameList.Add(strName);
                }

                // the last row to insert the row from source which is not in the first Excel.
                int nTargetLastRowCount = nTargetRowsCount + 1;

                //////////////////////////////
                // Never use 2 dots with com objects.
                // refer to http://stackoverflow.com/questions/158706/how-to-properly-clean-up-excel-interop-objects
                // do columns merge
                for (int i = 1; i < m_srcExcelFiles.Length; ++i)
                {
                    Excel.Workbook wbTemp = null;
                    Excel.Worksheet wsTemp = null;
                    Excel.Range usedRange = null;

                    string fileName = m_srcExcelFiles[i];
                    if (fileName == null || fileName.Length == 0)
                        continue;
                    try
                    {
                        wbTemp = workBooks.Open(fileName);

                        //MessageBox.Show("Open the second source Excel file succeed!", "Success");

                        wsTemp = (Excel.Worksheet)wbTemp.Worksheets.get_Item(1);
                        usedRange = wsTemp.UsedRange;
                        if (usedRange != null)
                        {
                            int nRows = usedRange.Rows.Count;
                            int nCols = usedRange.Columns.Count;
                            Excel.Range rows = usedRange.Rows;
                            Excel.Range firstRow = rows[1];
                            Excel.Range firstRowCells = firstRow.Cells;

                            // copy header
                            Excel.Range headerRow = targetRange.Rows[1];
                            if (headerRow != null)
                            {
                                // Insert a new cells
                                Excel.Range newColumns = null;
                                //insertRow.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
                                newColumns = wsTarget.Cells[1, ntargetColumnCount + 1];

                                firstRowCells.Copy(newColumns);
                            }


                            // looking for the "Receiver1"
                            string strReceiverName = m_keyColumnName2;
                            int receiverNameColumn = -1;

                            //foreach (Excel.Range cell in usedRange.Rows)
                            for (int j = 1; j < nCols; ++j)
                            {
                                string value = firstRowCells[j].value as string;
                                if (strReceiverName.CompareTo(value) == 0)
                                {
                                    receiverNameColumn = j;
                                    break;
                                }
                            }

                            if (receiverNameColumn == -1)
                            {
                                Exception e = new System.ArgumentException("Invalid key column's name 2: " + m_keyColumnName2);
                                throw e;
                            }

                            // skip the first row as a table header
                            for (int k = 2; k <= nRows; ++k)
                            {
                                Excel.Range thisRow = rows[k];
                                Excel.Range thisRowCells = thisRow.Cells;
                                string receiverName = thisRowCells[receiverNameColumn].value as string;
                                if (receiverName == null || receiverName.Length == 0)
                                    continue;

                                int targetRowIndex = patientNameList.FindIndex(id => id == receiverName);
                                if (targetRowIndex == -1)
                                {
                                    // Don't find the receiver in the source Excel table.
                                    //// This record will be lost...
                                    // copy to the end...
                                    {
                                        Excel.Range newColumns = null;
                                        newColumns = wsTarget.Cells[nTargetLastRowCount, ntargetColumnCount + 1];
                                        thisRowCells.Copy(newColumns);
                                        nTargetLastRowCount++;
                                    }
                                    continue;
                                }

                                // The zero-based index, plus the title row
                                targetRowIndex += 2;
                                

                                Excel.Range insertRow = null;
                                int startColumnIndex = ntargetColumnCount + 1;
                                /////////////////////////////////
                                // Insert columns 
                                //insertRow = wsTarget.get_Range(wsTarget.Cells[targetRowIndex, startColumnIndex], wsTarget.Cells[targetRowIndex, startColumnIndex+nCols]);
                                //insertRow = wsTarget.Cells[targetRowIndex, ntargetColumnCount+1];
                                //insertRow = tempCell.EntireRow;

                                insertRow = targetRange.Rows[targetRowIndex];
                                if (insertRow != null)
                                {
                                    // Insert a new cells
                                    Excel.Range newColumns = null;
                                    //insertRow.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
                                    // Get the new cells
                                    //Excel.Range targetRows = wsTarget.Rows;
                                    //Excel.Range newRow = targetRows[targetRowIndex];
                                    //Excel.Range newRowCells = newRow.Cells;
                                    newColumns = wsTarget.Cells[targetRowIndex, ntargetColumnCount + 1];

                                    thisRowCells.Copy(newColumns);
                                    // Assign values
                                    //for (int nColumn =1; nColumn < nCols; ++nColumn)
                                    //{
                                    //    string value = thisRowCells[nColumn].Value2 as string;
                                    //    newRowCells[nColumn].Value2 = value;
                                    //}
                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        throw e;
                    }
                    finally
                    {
                        if (wbTemp != null)
                        {
                            wbTemp.Close();
                        }
                        //if (rows != null)
                        //{
                        //    rows.close();
                        //}

                        Marshal.ReleaseComObject(wbTemp);
                        Marshal.ReleaseComObject(wsTemp);
                    }
                }

                // save the result
                wbTarget.SaveCopyAs(m_trgExcelFile);
            }//end of try
            catch (Exception e)
            {
                m_errorMsg = e.Message;
                result = false;
            }
            finally
            {
                if (wbSource != null)
                {
                    wbSource.Close();
                }
                if (wbTarget != null)
                {
                    wbTarget.Close();
                }
                if (workBooks != null)
                {
                    workBooks.Close();
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                }

                Marshal.ReleaseComObject(excelApp);
                Marshal.ReleaseComObject(workBooks);
                Marshal.ReleaseComObject(wbSource);
                Marshal.ReleaseComObject(wsSource);
                Marshal.ReleaseComObject(wbTarget);
                Marshal.ReleaseComObject(wsTarget);
            }
            return result;
        }

        // member variables
        private string m_trgExcelFile = "";
        private string[] m_srcExcelFiles = { "","","" };
        private UInt16 m_srcFileCount = 0;
        private string m_errorMsg = "";
        private bool m_bDoMergeColumn = true;
        private string m_keyColumnName1 = "NationalID";
        private string m_keyColumnName2 = "NationalID";
    }
}
