using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using RGiesecke.DllExport;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace simpleOne
{
    public class ExcelHandler
    {
        private static void LogError(string message)
        {
            string logFilePath = "error_log.txt";
            try
            {
                using (StreamWriter writer = new StreamWriter(logFilePath, true))
                {
                    writer.WriteLine($"{DateTime.Now}: {message}");
                }
            }
            catch (Exception)
            {
                // Handle any errors that might occur during logging
            }
        }

        [DllExport("WriteToXlsx", CallingConvention = CallingConvention.StdCall)]
        public static bool WriteToXlsx(string filename, string sheetName, string data)
        {
            try
            {
                // Validate filename
                if (string.IsNullOrWhiteSpace(filename))
                {
                    throw new ArgumentException("Filename cannot be empty or whitespace.");
                }

                string[] dataArray = data.Split(',');
                Application excelApp = new Application();
                Workbook workbook = null;
                Worksheet sheet;

                // Handle file creation
                if (File.Exists(filename))
                {
                    workbook = excelApp.Workbooks.Open(filename);
                }
                else
                {
                    // Ensure directory exists
                    string directoryPath = Path.GetDirectoryName(filename);
                    if (!Directory.Exists(directoryPath))
                    {
                        Directory.CreateDirectory(directoryPath);
                    }

                    workbook = excelApp.Workbooks.Add();
                    workbook.SaveAs(filename);
                }

                // Handle sheet creation or access
                sheet = workbook.Sheets.Cast<Worksheet>().FirstOrDefault(ws => ws.Name == sheetName);
                if (sheet == null)
                {
                    sheet = workbook.Sheets.Add();
                    sheet.Name = sheetName;
                }

                // Write data to the sheet
                int row = sheet.Cells[sheet.Rows.Count, 1].End(XlDirection.xlUp).Row + 1;
                for (int i = 0; i < dataArray.Length; i++)
                {
                    sheet.Cells[row, i + 1].Value = dataArray[i];
                }

                // Save and close workbook
                workbook.Save();
                workbook.Close();
                excelApp.Quit();

                return true;
            }
            catch (Exception e)
            {
                LogError($"An error occurred in WriteToXlsx: {e.Message}");
                return false;
            }
        }
        [DllExport("ReadRowCount", CallingConvention = CallingConvention.StdCall)]
        public static int ReadRowCount(string filename, string sheetName)
        {
            try
            {
                if (!File.Exists(filename))
                {
                    LogError($"File '{filename}' does not exist.");
                    return 0;
                }

                Application excelApp = new Application();
                Workbook workbook = excelApp.Workbooks.Open(filename);
                Worksheet sheet = workbook.Sheets.Cast<Worksheet>().FirstOrDefault(ws => ws.Name == sheetName);

                if (sheet == null)
                {
                    LogError($"Sheet '{sheetName}' does not exist in the file.");
                    workbook.Close();
                    excelApp.Quit();
                    return 0;
                }

                int rowCount = sheet.Cells[sheet.Rows.Count, 1].End(XlDirection.xlUp).Row;
                workbook.Close();
                excelApp.Quit();

                return rowCount;
            }
            catch (Exception e)
            {
                LogError($"An error occurred in ReadRowCount: {e.Message}");
                return 0;
            }
        }
        [DllExport("ReadRow", CallingConvention = CallingConvention.StdCall)]
        public static void ReadRow(string filename, string sheetName, int row, IntPtr result, int resultSize)
        {
            try
            {
                if (!File.Exists(filename))
                {
                    LogError($"File '{filename}' does not exist.");
                    Marshal.Copy(Encoding.UTF8.GetBytes(""), 0, result, 0); // Return an empty string
                    return;
                }

                Application excelApp = new Application();
                Workbook workbook = excelApp.Workbooks.Open(filename);
                Worksheet sheet = workbook.Sheets.Cast<Worksheet>().FirstOrDefault(ws => ws.Name == sheetName);

                if (sheet == null)
                {
                    LogError($"Sheet '{sheetName}' does not exist in the file.");
                    workbook.Close();
                    excelApp.Quit();
                    Marshal.Copy(Encoding.UTF8.GetBytes(""), 0, result, 0); // Return an empty string
                    return;
                }

                if (row < 1 || row > sheet.Cells[sheet.Rows.Count, 1].End(XlDirection.xlUp).Row)
                {
                    LogError($"Row {row} does not exist in the sheet.");
                    workbook.Close();
                    excelApp.Quit();
                    Marshal.Copy(Encoding.UTF8.GetBytes(""), 0, result, 0); // Return an empty string
                    return;
                }

                int cols = sheet.Cells[row, sheet.Columns.Count].End(XlDirection.xlToLeft).Column;
                string[] rowData = new string[cols];
                for (int i = 0; i < cols; i++)
                {
                    rowData[i] = sheet.Cells[row, i + 1].Text;
                }

                string resultString = string.Join(",", rowData);
                byte[] resultBytes = Encoding.UTF8.GetBytes(resultString);

                if (resultBytes.Length > resultSize)
                {
                    LogError($"Result buffer size is too small.");
                    Marshal.Copy(Encoding.UTF8.GetBytes(""), 0, result, 0); // Return an empty string
                }
                else
                {
                    Marshal.Copy(resultBytes, 0, result, resultBytes.Length);
                    Marshal.Copy(new byte[] { 0 }, 0, result + resultBytes.Length, 1); // Null-terminate the string
                }

                workbook.Close();
                excelApp.Quit();
            }
            catch (Exception e)
            {
                LogError($"An error occurred in ReadRow: {e.Message}");
                Marshal.Copy(Encoding.UTF8.GetBytes(""), 0, result, 0); // Return an empty string
            }
        }

    }
}
