// ExcelHandler.cpp

// Include Windows headers for DLL exporting and path handling.
#include <windows.h>

// Include standard headers.
#include <string>
#include <fstream>
#include <sstream>
#include <vector>
#include <ctime>
#include <cstring>
#include <stdexcept>

// Include the xlnt library header (ensure that xlnt is installed and the include path is set).
#include <xlnt/xlnt.hpp>

// -----------------------------------------------------------------------------
// Helper: LogError
// -----------------------------------------------------------------------------
static void LogError(const std::string& message)
{
    try
    {
        // Get the DLL’s directory.
        char modulePath[MAX_PATH] = { 0 };
        if (GetModuleFileNameA((HMODULE)&__ImageBase, modulePath, MAX_PATH) == 0)
        {
            // Fallback: use current directory if GetModuleFileName fails.
            strcpy_s(modulePath, ".");
        }
        // Remove the executable name to obtain the directory.
        std::string fullPath(modulePath);
        size_t pos = fullPath.find_last_of("\\/");
        std::string directory = (pos != std::string::npos) ? fullPath.substr(0, pos) : ".";

        std::string logFilePath = directory + "\\error_log.txt";

        // Open the log file in append mode.
        std::ofstream logFile(logFilePath, std::ios::out | std::ios::app);
        if (logFile.is_open())
        {
            // Get current time.
            std::time_t now = std::time(nullptr);
            char timeStr[64];
            ctime_s(timeStr, sizeof(timeStr), &now);
            // Remove trailing newline from ctime_s output.
            std::string timeString(timeStr);
            if (!timeString.empty() && timeString[timeString.size() - 1] == '\n')
                timeString.erase(timeString.size() - 1);

            logFile << timeString << ": " << message << std::endl;
            logFile.close();
        }
    }
    catch (...)
    {
        // If logging fails, there is not much we can do.
    }
}

// The following symbol is used by GetModuleFileName to obtain the DLL module handle.
extern "C" IMAGE_DOS_HEADER __ImageBase;

// -----------------------------------------------------------------------------
// Utility: Split a string by comma into a vector of strings.
// -----------------------------------------------------------------------------
static std::vector<std::string> SplitString(const std::string& str)
{
    std::vector<std::string> tokens;
    std::istringstream stream(str);
    std::string token;
    while (std::getline(stream, token, ','))
    {
        tokens.push_back(token);
    }
    return tokens;
}

// -----------------------------------------------------------------------------
// Exported Function: WriteToXlsx
// Parameters:
//   filename - path to the XLSX file
//   sheetName - name of the worksheet
//   data - a comma‐separated string of cell values to be written on a new row
// Returns: true on success, false on error.
// -----------------------------------------------------------------------------
extern "C" __declspec(dllexport) bool __stdcall WriteToXlsx(const char* filename, const char* sheetName, const char* data)
{
    try
    {
        if (!filename || !sheetName || !data)
            throw std::invalid_argument("Null pointer passed as parameter.");

        // Convert input parameters to std::string.
        std::string fileStr(filename);
        std::string sheetStr(sheetName);
        std::string dataStr(data);

        // Split the data string by commas.
        std::vector<std::string> dataArray = SplitString(dataStr);

        xlnt::workbook wb;
        // Check if the file exists. If so, load it; if not, create a new workbook.
        std::ifstream infile(fileStr);
        if (infile.good())
        {
            wb.load(fileStr);
        }
        else
        {
            wb = xlnt::workbook();
        }
        infile.close();

        xlnt::worksheet ws;
        // If the workbook already has the sheet, use it; otherwise create it.
        if (wb.contains(sheetStr))
        {
            ws = wb.sheet_by_title(sheetStr);
        }
        else
        {
            ws = wb.create_sheet(sheetStr);
        }

        // Determine the next row to write to.
        // xlnt::worksheet::highest_row() returns the index of the last row with data.
        // We write on the next row.
        auto startRow = ws.highest_row() + 1;

        // Write each data element into successive columns (starting at column 1).
        for (std::size_t i = 0; i < dataArray.size(); ++i)
        {
            ws.cell(xlnt::cell_reference(static_cast<unsigned>(i + 1), startRow)).value(dataArray[i]);
        }

        // Save the workbook.
        wb.save(fileStr);
        return true;
    }
    catch (const std::exception& ex)
    {
        LogError(std::string("An error occurred in WriteToXlsx: ") + ex.what());
        return false;
    }
    catch (...)
    {
        LogError("An unknown error occurred in WriteToXlsx.");
        return false;
    }
}

// -----------------------------------------------------------------------------
// Exported Function: ReadRowCount
// Parameters:
//   filename - path to the XLSX file
//   sheetName - name of the worksheet
// Returns: the number of rows in the sheet; returns 0 on error.
// -----------------------------------------------------------------------------
extern "C" __declspec(dllexport) int __stdcall ReadRowCount(const char* filename, const char* sheetName)
{
    try
    {
        if (!filename || !sheetName)
            throw std::invalid_argument("Null pointer passed as parameter.");

        std::string fileStr(filename);
        std::string sheetStr(sheetName);

        xlnt::workbook wb;
        std::ifstream infile(fileStr);
        if (!infile.good())
        {
            LogError("File does not exist in ReadRowCount.");
            return 0;
        }
        wb.load(fileStr);
        infile.close();

        if (!wb.contains(sheetStr))
        {
            LogError("Sheet '" + sheetStr + "' does not exist in the file in ReadRowCount.");
            return 0;
        }

        xlnt::worksheet ws = wb.sheet_by_title(sheetStr);
        // Return the highest row with data.
        return static_cast<int>(ws.highest_row());
    }
    catch (const std::exception& ex)
    {
        LogError(std::string("An error occurred in ReadRowCount: ") + ex.what());
        return 0;
    }
    catch (...)
    {
        LogError("An unknown error occurred in ReadRowCount.");
        return 0;
    }
}

// -----------------------------------------------------------------------------
// Exported Function: ReadRow
// Parameters:
//   filename - path to the XLSX file
//   sheetName - name of the worksheet
//   row - the row number to read (1-indexed)
//   result - pointer to a caller-provided buffer where the CSV string will be written
//   resultSize - size of the result buffer in bytes
// Behavior: Writes a CSV (comma-separated) string of the row’s cell values to the result buffer.
//           If an error occurs or the buffer is too small, an empty string is written.
// -----------------------------------------------------------------------------
extern "C" __declspec(dllexport) void __stdcall ReadRow(const char* filename, const char* sheetName, int row, char* result, int resultSize)
{
    try
    {
        if (!filename || !sheetName || !result)
            throw std::invalid_argument("Null pointer passed as parameter.");

        std::string fileStr(filename);
        std::string sheetStr(sheetName);

        xlnt::workbook wb;
        std::ifstream infile(fileStr);
        if (!infile.good())
        {
            LogError("File does not exist in ReadRow.");
            if (resultSize > 0) result[0] = '\0';
            return;
        }
        wb.load(fileStr);
        infile.close();

        if (!wb.contains(sheetStr))
        {
            LogError("Sheet '" + sheetStr + "' does not exist in the file in ReadRow.");
            if (resultSize > 0) result[0] = '\0';
            return;
        }

        xlnt::worksheet ws = wb.sheet_by_title(sheetStr);
        // Check that the requested row exists.
        if (row < 1 || row > static_cast<int>(ws.highest_row()))
        {
            LogError("Row " + std::to_string(row) + " does not exist in the sheet in ReadRow.");
            if (resultSize > 0) result[0] = '\0';
            return;
        }

        // Determine the highest column number in the row.
        // (Note: xlnt does not provide a row-wise “highest_column”, so we use the worksheet’s highest column.)
        auto highestCol = ws.highest_column();
        // Build CSV string.
        std::ostringstream oss;
        bool first = true;
        for (unsigned col = 1; col <= highestCol.index; ++col)
        {
            if (!first)
                oss << ",";
            else
                first = false;

            std::string cellText = ws.cell(xlnt::cell_reference(col, row)).to_string();
            oss << cellText;
        }
        std::string rowData = oss.str();

        // Check if the result buffer is large enough.
        // We add one for the null terminator.
        if (static_cast<int>(rowData.size() + 1) > resultSize)
        {
            LogError("Result buffer size is too small in ReadRow.");
            if (resultSize > 0) result[0] = '\0';
            return;
        }

        // Copy the string into the result buffer and null-terminate.
        std::memcpy(result, rowData.c_str(), rowData.size() + 1);
    }
    catch (const std::exception& ex)
    {
        LogError(std::string("An error occurred in ReadRow: ") + ex.what());
        if (resultSize > 0) result[0] = '\0';
    }
    catch (...)
    {
        LogError("An unknown error occurred in ReadRow.");
        if (resultSize > 0) result[0] = '\0';
    }
}
