using Gs.Utility.Excel.Structure.Classes;
using Gs.Utility.Excel.Structure.Enums;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gs.Utility
{
    public static class ExcelHelper
    {
        public static bool IsOleDbDriverInstalled(OledDbDriverName oleDbDriverName)
        {
            string driverName = GetDriverName(oleDbDriverName);
            var reader = OleDbEnumerator.GetRootEnumerator();
            var list = new List<String>();
            while (reader.Read())
            {
                for (var i = 0; i < reader.FieldCount; i++)
                {
                    if (reader.GetName(i) == "SOURCES_NAME")
                    {
                        if (reader.GetValue(i).ToString() == driverName)
                        {
                            reader.Close();
                            return true;
                        }
                    }
                }
            }
            reader.Close();
            return false;
        }

        public static void NormalizeExcelDataTable(DataTable dt)
        {
            string[] column = dt.Columns.OfType<DataColumn>()
                                         .ToList().Where(item => !string.IsNullOrWhiteSpace(item.ColumnName))
                                         .Select(item => item.ColumnName).ToArray();
            int columnCount = 0;

            //if (column[0] != "F1")
            //{
            //    columnCount = 1;
            //    for (int i = 1; i < dt.Columns.Count; i++)
            //    {
            //        if (dt.Columns[i].ColumnName != string.Format("F{0}", i + 1))
            //        {
            //            columnCount++;
            //        }
            //    }
            //}
            int rowIndex = 0;
            int firstColumnIndex = -1;
            while (columnCount == 0)
            {
                DataRow row = dt.Rows[rowIndex];
                for (int i = 0; i < column.Length; i++)
                {
                    if (!string.IsNullOrWhiteSpace(row[i].ToString()))
                    {
                        columnCount++;
                        if (firstColumnIndex == -1)
                        {
                            firstColumnIndex = i;
                        }
                    }
                }
                rowIndex++;
            }
            ShrinkColumnsTo(dt, columnCount, firstColumnIndex);
        }

        public static void ExportToTextFile(DataTable dt, string columnSeperator, string filePath, bool overwrite = false)
        {
            string[] column = dt.Columns.OfType<DataColumn>()
                                         .ToList().Where(item => !string.IsNullOrWhiteSpace(item.ColumnName))
                                         .Select(item => item.ColumnName).ToArray();



            if (overwrite)
                if (File.Exists(filePath))
                    File.Delete(filePath);
            using (StreamWriter writer = File.CreateText(filePath))
            {
                //if (column[0] != "F1")
                //    writer.WriteLine(string.Join(columnSeperator, column));
                foreach (DataRow item in dt.Rows)
                {
                    var itemArray = item.ItemArray;
                    writer.WriteLine(string.Join(columnSeperator, itemArray));
                }
                writer.Flush();
                writer.Close();
            }
        }

        public static void ShrinkColumnHasNoHeader(OleDbConnection conn, DataTable dt, int headerRow)
        {
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                if (string.IsNullOrWhiteSpace(dt.Rows[headerRow][dt.Columns[i]].ToString()))
                {

                }
            }
        }

        public static void ShrinkColumnsTo(DataTable dt, int columnCount, int offset = 0)
        {

            for (int i = 0; i < offset; i++)
            {
                dt.Columns.RemoveAt(0);
            }
            while (dt.Columns.Count > columnCount)
            {

                dt.Columns.RemoveAt(columnCount);
            }
        }

        public static void EncolseDataWith(DataTable dt, string encloseExp)
        {
            foreach (DataRow item in dt.Rows)
            {
                foreach (DataColumn col in dt.Columns)
                {
                    if (col.DataType == typeof(string))
                    {
                        item[col] = string.Format("{0}{1}{0}", encloseExp, item[col].ToString());
                    }
                }
            }
        }

        public static DataTable ReadDataFromSheet(OleDbConnection conn, string sheetName)
        {
            DataTable dt = new DataTable();
            string source = string.Format("{0}$", sheetName);
            if (sheetName.IsNumeric())
            {
                source = string.Format("'{0}$'", sheetName);
            }
            OleDbCommand oleDbCommand = new OleDbCommand(source, conn);
            oleDbCommand.CommandType = CommandType.TableDirect;
            var reader = oleDbCommand.ExecuteReader();
            dt.Load(reader);
            reader.Close();
            return dt;
        }

        private static string GetDriverName(OledDbDriverName oleDbDriverName)
        {
            switch (oleDbDriverName)
            {
                case OledDbDriverName.MicrosoftACEOLEDB12:
                    return "Microsoft.ACE.OLEDB.12.0";
                case OledDbDriverName.MicrosoftJetOLEDB40:
                    return "Microsoft.Jet.OLEDB.4.0";
                default:
                    throw new Exception("Driver Not Supported!");
                    break;
            }
        }

        public static string GetConnectionString(OledDbDriverName driverName, ExtendedProperties extenProps, string filePath)
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            // XLSX - Excel 2007, 2010, 2012, 2013
            props["Provider"] = GetDriverName(driverName);
            props["Extended Properties"] = extenProps.GetStringValue();
            props["Data Source"] = filePath;

            // XLS - Excel 2003 and Older
            //props["Provider"] = "Microsoft.Jet.OLEDB.4.0";
            //props["Extended Properties"] = "Excel 8.0";
            //props["Data Source"] = "C:\\MyExcel.xls";

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }


        public static bool HasSheetNamed(OleDbConnection conn, string sheetName)
        {
            using (var dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null))
            {
                var filtred = dt.Select(string.Format("TABLE_NAME ='''{0}$'''", sheetName));
                return filtred.Length > 0;
            }
        }


        public static IEnumerable<string> GetSheetsName(OleDbConnection conn)
        {
            using (var dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null))
            {
                var sheets = dt.Rows.Cast<DataRow>().Select(item => item[2].ToString());
                List<string> cleanList = new List<string>();
                foreach (var sheet in sheets)
                {
                    string temp = sheet;
                    if (temp.StartsWith("'"))
                    {
                        temp = temp.Substring(1, temp.Length - 2);
                    }
                    if (temp.EndsWith("$"))
                    {
                        temp = temp.Substring(0, temp.Length - 1);
                    }
                    cleanList.Add(temp);
                }
                return cleanList;
            }
        }



        public static IEnumerable<string> GetSheetsName(string fileName)
        {
            IEnumerable<OledDbDriverName> drivers = Enum.GetValues(typeof(OledDbDriverName)).Cast<OledDbDriverName>();
            foreach (var item in drivers)
            {
                String connStr = ExcelHelper.GetConnectionString(item, ExtendedProperties.Excel_12 | ExtendedProperties.HDR_NO | ExtendedProperties.IMEX_YES, fileName);
                OleDbConnection conn = null;
                IEnumerable<string> sheetsList;
                try
                {
                    using (conn = new OleDbConnection(connStr))
                    {
                        conn.Open();
                        List<string> cleanList = new List<string>();
                        using (var dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null))
                        {
                            sheetsList = dt.Rows.Cast<DataRow>().Select(row => row[2].ToString());
                            foreach (var sheet in sheetsList)
                            {
                                string temp = sheet;
                                if (temp.StartsWith("'"))
                                {
                                    temp = temp.Substring(1, temp.Length - 2);
                                }
                                if (temp.EndsWith("$"))
                                {
                                    temp = temp.Substring(0, temp.Length - 1);
                                }
                                cleanList.Add(temp);
                            }
                        }
                        conn.Close();
                        return cleanList;
                    }
                }
                catch (Exception ex)
                {

                }
                finally
                {
                    if (conn != null)
                    {
                        if (conn.State != ConnectionState.Closed)
                            conn.Close();
                        conn.Dispose();
                    }
                }
            }
            return null;
        }



        public static IEnumerable<ExcelColnfo> GetColumnsNames(string fileName, string sheetName)
        {
            IEnumerable<OledDbDriverName> drivers = Enum.GetValues(typeof(OledDbDriverName)).Cast<OledDbDriverName>();
            List<ExcelColnfo> columnsNames = new List<ExcelColnfo>();
            foreach (var item in drivers)
            {
                String connStr = ExcelHelper.GetConnectionString(item, ExtendedProperties.Excel_12 | ExtendedProperties.HDR_NO | ExtendedProperties.IMEX_YES, fileName);
                OleDbConnection conn = null;
                try
                {
                    using (conn = new OleDbConnection(connStr))
                    {
                        conn.Open();
                        var dataTable = ReadDataFromSheet(conn, sheetName);
                        for (int i = 0; i < dataTable.Columns.Count; i++)
                        {
                            columnsNames.Add(new ExcelColnfo() { ColumnIndex = i, ColumnName = dataTable.Rows[0][i].ToString() });
                        }
                        conn.Close();
                        return columnsNames;
                    }
                }
                catch (Exception ex)
                {

                }
                finally
                {
                    if (conn != null)
                    {
                        if (conn.State != ConnectionState.Closed)
                            conn.Close();
                        conn.Dispose();
                    }
                }
            }
            return null;
        }




        public static OleDbConnection CreateConnection(string fileName)
        {
            IEnumerable<OledDbDriverName> drivers = Enum.GetValues(typeof(OledDbDriverName)).Cast<OledDbDriverName>();
            foreach (var item in drivers)
            {
                String connStr = ExcelHelper.GetConnectionString(item, ExtendedProperties.Excel_12 | ExtendedProperties.HDR_NO | ExtendedProperties.IMEX_YES, fileName);
                OleDbConnection conn = null;
                try
                {
                    conn = new OleDbConnection(connStr);
                    conn.Open();
                    return conn;
                }
                catch (Exception)
                {
                    if (conn != null)
                    {
                        if (conn.State != ConnectionState.Closed)
                            conn.Close();
                        conn.Dispose();
                    }
                }
            }
            return null;
        }


    }
}
