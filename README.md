
# 📊 Export MT5 (MetaTrader 5) Data to Excel Using a C# DLL

This project demonstrates how to export real-time or historical data from **MetaTrader 5 (MT5)** to a `.xlsx` **Excel file** using a **custom C# DLL** — with two options:
- Using **NanoXLSX** (lightweight and fast)
- Using **Excel Interop** (slower but with full Excel features like auto-fill and formulas)

✅ **Supports both Interop and non-Interop modes**  
✅ **Works with MQL5 `import` and allows low-latency data writing**  
✅ **100% offline and portable (NanoXLSX version)**

---

## ⚠️ Two Versions: Interop vs NanoXLSX (xlint)

This repo includes **two different libraries** for exporting MT5 data to Excel:

### 1. **Excel Interop Version** (Slower but Full Excel Features)
- Uses `Microsoft.Office.Interop.Excel`
- Opens Excel in the background
- Supports full Excel features like:
  - Auto-fill
  - Native formulas (e.g., `=SUM(A1:A10)`)
  - Formatting and more
- ✅ Use this if you need full Excel functionality
- ❌ Downside: **Slower and requires Excel installed**

### 2. **NanoXLSX (xlint) Version** (Faster and Portable)
- Uses the lightweight `NanoXLSX` library
- Writes `.xlsx` files without needing Excel
- ⚡ Fast, portable, and Excel-independent
- ❌ Excel-specific functions like auto-fill, formulas will not work

> 👉 Choose based on your needs: **features (Interop)** vs **speed (NanoXLSX)**

---

## 🔧 Features

- 📁 Write to `.xlsx` files from MetaTrader 5
- 🧩 Exported DLL functions callable from MQL5
- ⚡ Low-latency function calls (no subprocesses)
- 🧪 Includes working MQL5 test script
- 💼 Useful for trade logging, backtests, or analytics

---

## 📦 Technologies Used

| Tool / Language | Purpose |
|-----------------|---------|
| **C# (.NET Framework)** | Native DLL logic |
| **[NanoXLSX](https://github.com/ricoSuter/NanoXLSX)** | Write Excel `.xlsx` files (no Excel needed) |
| **Excel Interop** | Excel automation (slow but full feature) |
| **[DllExport](https://github.com/3F/DllExport)** | Export C# methods to MQL5 |
| **MQL5 (MetaTrader 5)** | Calling the DLL |

---

## 🧠 How It Works

1. A C# DLL is compiled with `DllExport` to expose native functions.
2. Two options:
   - **NanoXLSX version** writes Excel files directly
   - **Interop version** launches Excel in the background
3. MQL5 imports the DLL functions using `import`.
4. You call the DLL with file path, sheet, cell, and value.

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
```
```mql5
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

---

## 📂 Project Structure

```bash
mt5-to-excel-dll/
├── ExcelExporterDll/       # ✅ DLL with NanoXLSX (use this for speed)
├── ExportToExcel/          # 🐢 Interop version (Excel automation)
├── ExcelTester.mq5         # 🧪 MQL5 test script
```

> 🔥 Use `ExcelExporterDll.dll` in your MQL5 project. Interop version is optional for full Excel logic.

---

## 🔨 Build Instructions
### NB// you can use the dll as is, but if you want to modify/improve it follow the below steps:
1. Open `mt5ExcelInterop.sln` or `xlnt_excell_dll.sln` in Visual Studio.
2. Install `DllExport` (via NuGet or [manual setup](https://github.com/3F/DllExport)).
3. Build as **x64**(mt5) or **x86**(mt4), **Release**, and **.NET Framework 4.+**
4. Compile → use resulting `.dll` in your MQL5 project.

---

## 🔗 Related Projects and Resources

- [NanoXLSX](https://github.com/ricoSuter/NanoXLSX)
- [DllExport by 3F](https://github.com/3F/DllExport)
- [MetaTrader 5 Documentation](https://www.metatrader5.com/en/terminal/help)

---

## 💬 Questions or Suggestions?

Open an [issue](https://github.com/Sir-kirika/mt5-to-excel-dll/issues) or send a pull request.

---

## 📢 Spread the Word!

If this helped you:
- ⭐ Star the repo
- 🗣️ Share it on MQL5.com forums or Reddit
- 🔁 Fork and improve it

---

## 📛 License

This project is open-source.

---

