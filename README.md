here is a simple use case, 
Example 1: This runs Excell in bg meaning Excell must be installed in the machine. So you can pre_fill cells with any logic once data is written it will auto update cells based on your formulars

    
//
//    

//#import "mt5ExcelInterop.dll"
//   bool WriteToXlsx(const char &filename[], const char &sheetName[], const char &data[]);
//   int ReadRowCount(const char &filename[], const char &sheetName[]);
//   void ReadRow(const char &filename[], const char &sheetName[], int row, char &result[], int result_size);
//#import
//
//void OnStart()
//{ 
//    string filename = "test.xlsx";
//    string fileString=TerminalInfoString(TERMINAL_DATA_PATH) + "\\MQL5\\Libraries\\" + filename;  
//    string sheetString = "Sheet1";
//    string dataString = "Hello,World,123";
//    
//    char file[];
//    StringToCharArray(fileString, file);
//    char sheet[];
//    StringToCharArray(sheetString, sheet);
//    char data[];
//    StringToCharArray(dataString, data);
//
//    // Write to Excel
//    bool writeSuccess = WriteToXlsx(file, sheet, data);
//    Print("Write Success: ", writeSuccess);
//    
//    // Read row count
//    int rowCount = ReadRowCount(file, sheet);
//    Print("Row Count: ", rowCount);
//    
//    // Read a row
//    if(rowCount > 0)
//    {
//        char rowData[256]; // Adjust size as needed
//        ReadRow(file, sheet, rowCount, rowData, ArraySize(rowData));
//        Print("Row Data: ", CharArrayToString(rowData));
//    }
//    else
//    {
//        Print("No rows to read.");
//    }
//}

Example 2: Does not require Excell installed limited to r/w
////
//#import "mt5Excel.dll"
//   bool WriteToXlsx(const char &filename[], const char &sheetName[], const char &data[]);
//   int ReadRowCount(const char &filename[], const char &sheetName[]);
//   void ReadRow(const char &filename[], const char &sheetName[], int row, char &result[], int result_size);
//#import
//
//void OnStart()
//{
//    string filename = "test.xlsx";
//    string fileString=TerminalInfoString(TERMINAL_DATA_PATH) + "\\MQL5\\Libraries\\" + filename;  
//    string sheetString = "Sheet1";
//    string dataString = "Hello,World,123";
//
//    char file[];
//    StringToCharArray(fileString, file);
//    char sheet[];
//    StringToCharArray(sheetString, sheet);
//    char data[];
//    StringToCharArray(dataString, data);
//
//    // Write to Excel
//    bool writeSuccess = WriteToXlsx(file, sheet, data);
//    Print("Write Success: ", writeSuccess);
//
//    // Read row count
//    int rowCount = ReadRowCount(file, sheet);
//    Print("Row Count: ", rowCount);
//
//    // Read a row
//    if(rowCount > 0)
//    {
//        char rowData[]; // Adjust size as needed
//        ReadRow(file, sheet, rowCount, rowData, ArraySize(rowData));
//        string rowString = CharArrayToString(rowData);
//        Print("Row Data: ", rowString);
//    }
//    else
//    {
//        Print("No rows to read.");
//    }
//}
