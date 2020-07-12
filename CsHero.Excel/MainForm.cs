using ClosedXML.Excel;
using Gs.Excel.Comparer.Structures.Classes;
using Gs.Utility;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Gs.Excel.Comparer
{
    public partial class MainForm : Form
    {
        private ExcelFileItem sourceItem;
        private ExcelFileItem destinationItem;

        public MainForm()
        {
            InitializeComponent();
        }

        private void btnSelectSourceFiles_Click(object sender, EventArgs e)
        {
            var result = openFileDialog1.ShowDialog();
            if (result != DialogResult.OK)
                return;
            txtSource.Text = openFileDialog1.FileName;
            sourceItem = new ExcelFileItem()
            {
                FilePath = openFileDialog1.FileName
            };
            sourceItemBindingSource.DataSource = sourceItem;
            FillSheets(sourceItem);
        }

        private void clbSourceFiles_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (clbSourceFiles.SelectedIndex == -1)
            //    return;
            //SourceItem sourceItem = (sourceItemBindingSource.Current as SourceItem);
            //FillSheets(sourceItem);
        }

        private void FillColumns(ExcelFileItem sourceItem)
        {
            var columns = ExcelHelper.GetColumnsNames(sourceItem.FilePath, cmbSheetName.Text).ToList();
            if (columns == null)
                return;
            excelColnfoBindingSource.DataSource = columns.ToList();

        }

        private void FillSheets(ExcelFileItem sourceItem)
        {
            cmbSheetName.Items.Clear();
            var sheetsList = ExcelHelper.GetSheetsName(sourceItem.FilePath);
            if (sheetsList == null)
            {
                MessageBox.Show("Error, invalid format, make sure you are using an Excel File.");
                return;
            }
            cmbSheetName.Items.AddRange(sheetsList.ToArray());
            if (!string.IsNullOrWhiteSpace(sourceItem.SheetName))
                cmbSheetName.Text = sourceItem.SheetName;
        }

        private void btnSelectDest_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            if (result != DialogResult.OK)
                return;
            destinationItem = new ExcelFileItem()
            {
                FilePath = openFileDialog1.FileName
            };
            destItemBindingSource.DataSource = destinationItem;
            txtDestFile.Text = destinationItem.FilePath;
            FillDestinationSheets(txtDestFile.Text);
        }

        private void FillDestinationSheets(string fileName)
        {
            cmbDestSheetName.Items.Clear();
            var data = ExcelHelper.GetSheetsName(fileName);
            if (data == null)
            {
                MessageBox.Show("Invalid Format");
                return;
            }
            cmbDestSheetName.Items.AddRange(data.ToArray());
        }

        private void FillDestinationColumns(string sheetName)
        {
            var columns = ExcelHelper.GetColumnsNames(txtDestFile.Text, sheetName).ToList();
            if (columns == null)
                return;
            excelColnfoBindingSource1.DataSource = columns.ToList();
        }

        private void btnExecute_Click(object sender, EventArgs e)
        {
            List<ItemRow> sourceData = new List<ItemRow>();
            DataTable sourceDt = null, destDt;
            OleDbConnection sourceConn = null, destConn = null;
            try
            {
                sourceConn = ExcelHelper.CreateConnection(sourceItem.FilePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Error in Block 1, {0}", ex.Message));//Block 1
                return;
            }
            try
            {
                sourceDt = ExcelHelper.ReadDataFromSheet(sourceConn, sourceItem.SheetName);
                ExcelHelper.NormalizeExcelDataTable(sourceDt);
                sourceData = sourceDt.Rows.Cast<DataRow>().Select(row => new ItemRow(row[sourceItem.PhoneNumberIndex].ToString())).ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Error in Block 2, {0}", ex.Message));//Block 2
                return;
            }

            if (sourceData.Count == 0)
                return;

            try
            {
                destConn = ExcelHelper.CreateConnection(destinationItem.FilePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Error in Block 3, {0}", ex.Message));//Block 3
                return;
            }

            try
            {
                destDt = ExcelHelper.ReadDataFromSheet(destConn, cmbDestSheetName.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Error in Block 3, {0}", ex.Message));//Block 3
                return;
            }
            finally
            {
                destConn.Close();
            }
            var destData = destDt.Rows.Cast<DataRow>().Select(row => new ItemRow(row[destinationItem.PhoneNumberIndex].ToString())).ToList();

            var exceptData = sourceData.Where(item => !destData.Any(item2 => item2.PhoneNumber == item.PhoneNumber)).ToList();


            sourceData = sourceData.Where(item => exceptData.Any(item2 => item2.PhoneNumber == item.PhoneNumber)).ToList();



            DataTable exceptDt = sourceDt.Clone();
            foreach (DataRow sourceRow in sourceDt.Rows)
            {
                if (exceptData.Any(item => item.PhoneNumber == sourceRow[sourceItem.PhoneNumberIndex].ToString()))
                {
                    exceptDt.Rows.Add(sourceRow.ItemArray);
                }
            }

            if (exceptDt.Rows.Count == 0)
            {
                MessageBox.Show("No Difference Between Columns Records!");
                return;
            }

            string newFileName = Path.Combine(Path.GetDirectoryName(txtDestFile.Text), "DIFF_" + Path.GetFileName(txtDestFile.Text));

            XLWorkbook wb = new XLWorkbook(txtDestFile.Text);

            string newSheetName = cmbDestSheetName.Text + "_NEW";
            int position = wb.Worksheet(cmbDestSheetName.Text).Position + 1;
            //if (cbReplace.Checked)
            //{
            //    newSheetName = cmbDestSheetName.Text;
            //    wb.Worksheet(cmbDestSheetName.Text).Delete();
            //    position--;
            //}

            ExcelHelper.NormalizeExcelDataTable(exceptDt);
            for (int i = wb.Worksheets.Count; i > 0; i--)
            {
                wb.Worksheets.Delete(i);
            }
            var ws = wb.Worksheets.Add(exceptDt, newSheetName);
            ws.Position = position;
            ws.AutoFilter.Enabled = false;
            ws.Row(1).Delete();
            wb.SaveAs(newFileName);
        }

        private void cmbSheetName_SelectedValueChanged(object sender, EventArgs e)
        {
            if (cmbSheetName.SelectedIndex == -1)
                return;
            FillColumns(sourceItem);
        }

        private void cmbDestSheetName_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbDestSheetName.SelectedIndex == -1)
                return;
            FillDestinationColumns(cmbDestSheetName.SelectedItem.ToString());
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
