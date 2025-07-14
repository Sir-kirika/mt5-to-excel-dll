
# 📊 Export MT5 (MetaTrader 5) Data to Excel Using a C# DLL

This project demonstrates how to export real-time or historical data from **MetaTrader 5 (MT5)** to a `.xlsx` **Excel file** using a **custom C# DLL** and the lightweight `NanoXLSX` library — **without needing to install Excel or use COM automation**.

✅ **No Python required**  
✅ **No Excel installation needed**  
✅ **Works with MQL5 `import` and allows low-latency data writing**  
✅ **100% offline and portable**

---

## 🔧 Features

- 📁 Write to `.xlsx` files without opening Excel
- 🧩 Exported DLL functions callable directly from MQL5
- ⚡ Fast execution: avoids COM latency or subprocess overhead
- 🧪 Includes a working MT5 script for testing
- 💼 Useful for trading logs, backtesting, and data collection

---

## 📦 Technologies Used

| Tool / Language | Purpose |
|-----------------|---------|
| **C# (.NET Framework)** | Native DLL logic |
| **[NanoXLSX](https://github.com/ricoSuter/NanoXLSX)** | Write Excel `.xlsx` files |
| **[DllExport](https://github.com/3F/DllExport)** | Expose C# methods to MQL5 |
| **MQL5 (MetaTrader 5)** | Calling the DLL from expert/script/indicator |

---

## 🧠 How It Works

1. A C# DLL is compiled with `DllExport` to expose native functions.
2. The DLL uses `NanoXLSX` to write values to specified Excel cells.
3. MQL5 imports the DLL functions using `import`.
4. MT5 calls the DLL with parameters like file path, sheet name, cell, and value.

---

## 🚀 Example MQL5 Usage


```mql5
//here is a simple use case, 
//Example 1: This runs Excell in bg meaning Excell must be installed in the machine. So you can pre_fill cells with any logic once data is written it will auto update cells based on your formulars   

#import "mt5ExcelInterop.dll"
   bool WriteToXlsx(const char &filename[], const char &sheetName[], const char &data[]);
   int ReadRowCount(const char &filename[], const char &sheetName[]);
   void ReadRow(const char &filename[], const char &sheetName[], int row, char &result[], int result_size);
#import

void OnStart()
{ 
    string filename = "test.xlsx";
    string fileString=TerminalInfoString(TERMINAL_DATA_PATH) + "\\MQL5\\Libraries\\" + filename;  
    string sheetString = "Sheet1";
    string dataString = "Hello,World,123";
    
    char file[];
    StringToCharArray(fileString, file);
    char sheet[];
    StringToCharArray(sheetString, sheet);
    char data[];
    StringToCharArray(dataString, data);

    // Write to Excel
    bool writeSuccess = WriteToXlsx(file, sheet, data);
    Print("Write Success: ", writeSuccess);
    
    // Read row count
    int rowCount = ReadRowCount(file, sheet);
    Print("Row Count: ", rowCount);
    
    // Read a row
    if(rowCount > 0)
    {
        char rowData[256]; // Adjust size as needed
        ReadRow(file, sheet, rowCount, rowData, ArraySize(rowData));
        Print("Row Data: ", CharArrayToString(rowData));
    }
    else
    {
        Print("No rows to read.");
    }
}

//Example 2: Does not require Excell installed limited to r/w
//
#import "mt5Excel.dll"
   bool WriteToXlsx(const char &filename[], const char &sheetName[], const char &data[]);
   int ReadRowCount(const char &filename[], const char &sheetName[]);
   void ReadRow(const char &filename[], const char &sheetName[], int row, char &result[], int result_size);
#import

void OnStart()
{
    string filename = "test.xlsx";
    string fileString=TerminalInfoString(TERMINAL_DATA_PATH) + "\\MQL5\\Libraries\\" + filename;  
    string sheetString = "Sheet1";
    string dataString = "Hello,World,123";

    char file[];
    StringToCharArray(fileString, file);
    char sheet[];
    StringToCharArray(sheetString, sheet);
    char data[];
    StringToCharArray(dataString, data);

    // Write to Excel
    bool writeSuccess = WriteToXlsx(file, sheet, data);
    Print("Write Success: ", writeSuccess);

    // Read row count
    int rowCount = ReadRowCount(file, sheet);
    Print("Row Count: ", rowCount);

    // Read a row
    if(rowCount > 0)
    {
        char rowData[]; // Adjust size as needed
        ReadRow(file, sheet, rowCount, rowData, ArraySize(rowData));
        string rowString = CharArrayToString(rowData);
        Print("Row Data: ", rowString);
    }
    else
    {
        Print("No rows to read.");
    }
}
```

## 📂 Project Structure

```bash
mt5-to-excel-dll/
├── ExcelExporterDll/       # ✅ C# DLL with exported functions (usable by MT5)
├── ExportToExcel/          # 🔬 Console app for testing NanoXLSX (not used in MT5)
├── ExcelTester.mq5         # 🧪 MQL5 script to test the DLL
```

> 🔥 Use `ExcelExporterDll.dll` from MetaTrader. `ExportToExcel` is for standalone testing only.

---

## 🔨 Build Instructions

1. Open `ExcelExporterDll` in Visual Studio.
2. Install `DllExport` (via NuGet or [manual setup](https://github.com/3F/DllExport)).
3. Set build target to **x64**, **Release**, and **.NET Framework 4.7+**.
4. Compile → use the resulting `.dll` in your MQL5 script.

---

## 🔗 Related Projects and Resources

- [NanoXLSX Library](https://github.com/ricoSuter/NanoXLSX)
- [DllExport (by 3F)](https://github.com/3F/DllExport)
- [MetaTrader 5 Documentation](https://www.metatrader5.com/en/terminal/help)

---

## 💬 Questions or Suggestions?

Feel free to open an [issue](https://github.com/Sir-kirika/mt5-to-excel-dll/issues) or create a pull request if you have improvements or questions.

---

## 📢 Spread the Word!

If this project helped you, give it a ⭐ star, share it on forums like [MQL5 Community](https://www.mql5.com/en/forum) or [Reddit](https://reddit.com/r/Forex), or fork it to expand functionality.

---

## 📛 License

This project is open-source and free to use. See the `LICENSE` file for more details.

---

## 📍 Suggested Repository Name and Topics

### ✅ Rename Repo to:
```
mt5-to-excel-dll
```

### ✅ Suggested GitHub Topics:
```
mql5 mt5 excel dll csharp unmanaged-exports nanoxlsx export-to-excel
```
