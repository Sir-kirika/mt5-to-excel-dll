// The following symbol is used by GetModuleFileName to obtain the DLL module handle.
#include "pch.h"

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
#include <iterator> // Include iterator header for std::advance

// Include the xlnt library header.
#include <xlnt/xlnt.hpp>

// The following symbol is used by GetModuleFileName to obtain the DLL module handle.
extern "C" IMAGE_DOS_HEADER __ImageBase;

// ... Rest of your code ...

// ----------------------------------------------------------------------------
// Helper function to log errors.
// ----------------------------------------------------------------------------
static void LogError(const std::string& message)
{
    try
    {
        // Get the DLL's directory.
        char modulePath[MAX_PATH] = { 0 };
        if (GetModuleFileNameA((HMODULE)&__ImageBase, modulePath, MAX_PATH) == 0)
        {
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
            if (!timeString.empty() && timeString.back() == '\n')
                timeString.pop_back();

            logFile << timeString << ": " << message << std::endl;
            logFile.close();
        }
    }
    catch (...)
    {
        // If logging fails, there is not much we can do.
    }
}

// ----------------------------------------------------------------------------
// Utility function to split a string by comma into a vector of strings.
// ----------------------------------------------------------------------------
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

// ----------------------------------------------------------------------------
// Exported Function: WriteToXlsx
// ----------------------------------------------------------------------------
extern "C" __declspec(dllexport) bool __stdcall WriteToXlsx(const char* filename, const char* sheetName, const char* data)
{
    try
    {
        if (!filename || !sheetName || !data)
            throw std::invalid_argument("Null pointer passed as parameter.");

        std::string fileStr(filename);
        std::string sheetStr(sheetName);
        std::string dataStr(data);

        // Split the data string by commas.
        std::vector<std::string> dataArray = SplitString(dataStr);

        xlnt::workbook wb;
        std::ifstream infile(fileStr);
        if (infile.good())
        {
            wb.load(fileStr);
        }
        infile.close();

        xlnt::worksheet ws;
        bool sheetExists = false;

        // Check if the sheet exists
        for (const auto& title : wb.sheet_titles())
        {
            if (title == sheetStr)
            {
                sheetExists = true;
                break;
            }
        }

        if (sheetExists)
        {
            ws = wb.sheet_by_title(sheetStr);
        }
        else
        {
            ws = wb.create_sheet();
            ws.title(sheetStr);
        }

        // Determine the next row to write to.
        auto startRow = ws.highest_row();
        if (ws.cell("A1").value<std::string>().empty() && startRow == 1)
        {
            startRow = 1;
        }
        else
        {
            startRow += 1;
        }

        // Write each data element into successive columns (starting at column 1).
       // for (std::size_t i = 0; i < dataArray.size(); ++i)
        //{
         //   ws.cell(startRow, static_cast<std::uint32_t>(i + 1)).value(dataArray[i]);
        //}

        // Instead of writing across columns, write down rows
        for (std::size_t i = 0; i < dataArray.size(); ++i)
        {
            ws.cell(static_cast<std::uint32_t>(1 + i), startRow).value(dataArray[i]); // Always column 1, increase row
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

// ----------------------------------------------------------------------------
// Exported Function: ReadRowCount
// ----------------------------------------------------------------------------
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

        xlnt::worksheet ws;
        bool sheetExists = false;

        // Check if the sheet exists
        for (const auto& title : wb.sheet_titles())
        {
            if (title == sheetStr)
            {
                sheetExists = true;
                break;
            }
        }

        if (!sheetExists)
        {
            LogError("Sheet '" + sheetStr + "' does not exist in the file in ReadRowCount.");
            return 0;
        }

        ws = wb.sheet_by_title(sheetStr);

        // Return the highest row with data.
        auto highestRow = ws.highest_row();
        if (ws.cell("A1").value<std::string>().empty() && highestRow == 1)
        {
            return 0;
        }
        else
        {
            return static_cast<int>(highestRow);
        }
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

// ----------------------------------------------------------------------------
// Exported Function: ReadRow
// ----------------------------------------------------------------------------
extern "C" __declspec(dllexport) void __stdcall ReadRow(const char* filename, const char* sheetName, int rowNumber, char* result, int resultSize)
{
    try
    {
        // Check for null pointers
        if (!filename || !sheetName || !result)
            throw std::invalid_argument("Null pointer passed as parameter.");

        std::string fileStr(filename);
        std::string sheetStr(sheetName);

        xlnt::workbook wb;
        // Load the workbook
        wb.load(fileStr);

        if (!wb.contains(sheetStr))
        {
            LogError("Sheet '" + sheetStr + "' does not exist in the file.");
            if (result && resultSize > 0)
                result[0] = '\0'; // Ensure result is empty
            return;
        }

        xlnt::worksheet ws = wb.sheet_by_title(sheetStr);

        // Verify that the requested row exists
        if (rowNumber < 1 || rowNumber > static_cast<int>(ws.highest_row()))
        {
            LogError("Row " + std::to_string(rowNumber) + " does not exist in the sheet.");
            if (result && resultSize > 0)
                result[0] = '\0'; // Ensure result is empty
            return;
        }

        // Find the last column with data in the specified row
        unsigned int lastColumnWithData = 0;
        unsigned int highestColumnIndex = ws.highest_column().index;

        for (unsigned int col = 1; col <= highestColumnIndex; ++col)
        {
            xlnt::cell cell = ws.cell(xlnt::cell_reference(col, rowNumber));
            if (cell.has_value())
            {
                lastColumnWithData = col;
            }
        }

        if (lastColumnWithData == 0)
        {
            // The row has no data
            if (result && resultSize > 0)
                result[0] = '\0'; // Ensure result is empty
            return;
        }

        // Build the CSV string up to the last column with data
        std::ostringstream oss;
        for (unsigned int col = 1; col <= lastColumnWithData; ++col)
        {
            if (col > 1)
                oss << ",";

            xlnt::cell cell = ws.cell(xlnt::cell_reference(col, rowNumber));
            if (cell.has_value())
            {
                oss << cell.to_string();
            }
            else
            {
                // Include empty strings for empty cells between data
                oss << "";
            }
        }

        std::string rowData = oss.str();

        // Check if the result buffer is large enough
        if (static_cast<int>(rowData.size() + 1) > resultSize)
        {
            LogError("Result buffer size is too small in ReadRow.");
            if (result && resultSize > 0)
                result[0] = '\0'; // Ensure result is empty
            return;
        }

        // Copy the CSV string into the result buffer
        std::memcpy(result, rowData.c_str(), rowData.size() + 1);
    }
    catch (const std::exception& ex)
    {
        LogError("An error occurred in ReadRow: " + std::string(ex.what()));
        if (result && resultSize > 0)
            result[0] = '\0'; // Ensure result is empty
    }
    catch (...)
    {
        LogError("An unknown error occurred in ReadRow.");
        if (result && resultSize > 0)
            result[0] = '\0'; // Ensure result is empty
    }
}
