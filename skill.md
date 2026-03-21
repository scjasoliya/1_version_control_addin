---
name: VBA Code Automation and Testing
description: Instructions for using the VBA Version Control Addin and Python win32com to inject, export, and test VBA code automatically.
---

# VBA Code Automation and Testing Skill

This skill provides instructions on how an agent can use the VBA Version Control Addin in conjunction with the Python `win32com.client` library to automate the injection, extraction, and testing of VBA code in Excel workbooks.

## Prerequisites

1.  **VBA Version Control Addin**: Must be installed and active in Excel. This add-in provides the ribbon tools and macros needed to export and import VBA project components seamlessly.
2.  **Python Environment**: Requires Python installed on the system.
3.  **pywin32 Package**: Requires the `pywin32` package for Python (`pip install pywin32`).

## Python Snippets for VBA Automation

### 1. Connecting to Excel using win32com

To interact with Excel via Python, use the following template to grab an existing instance or start a new one:

```python
import win32com.client
import os
import time

# Connect to Excel (starts a new instance if one isn't open)
try:
    excel = win32com.client.GetActiveObject("Excel.Application")
except Exception:
    excel = win32com.client.Dispatch("Excel.Application")

excel.Visible = True  # Make Excel visible for debugging
excel.DisplayAlerts = False

# Open a specific workbook
workbook_path = os.path.abspath("your_workbook.xlsm")
if os.path.exists(workbook_path):
    wb = excel.Workbooks.Open(workbook_path)
else:
    wb = excel.Workbooks.Add()
    wb.SaveAs(workbook_path, FileFormat=52) # 52 = xlOpenXMLWorkbookMacroEnabled
```

### 2. Triggering VBA Macros

You can trigger any public VBA subroutine or function using the `Application.Run` method. This is highly useful for running automated tests on newly generated VBA code.

```python
# Assuming 'wb' is your workbook object and 'excel' is the Application object
try:
    # Run a macro named 'MyTestMacro' located in 'Module1'
    # Format: "'WorkbookName.xlsm'!ModuleName.MacroName"
    wb_name = os.path.basename(workbook_path)
    macro_path = f"'{wb_name}'!Module1.MyTestMacro"

    result = excel.Application.Run(macro_path)
    print(f"Macro executed successfully. Result: {result}")
except Exception as e:
    print(f"Error running macro: {e}")
```

### 3. Using the VBA Version Control Addin Tools

If the VBA Version Control Addin is active, you can trigger its Export and Import routines directly from Python to synchronize the workbook's VBA project with the file system.

```python
# Triggering the Export Code routine from the add-in
try:
    # The add-in macros might be in a specific add-in file (e.g., 'VBA_Version_Control.xlam')
    # Or if running from the workbook itself, just name the module:
    excel.Application.Run("'VBA_Version_Control.xlam'!modVbaSync.ExportCode")
    print("VBA Code Export triggered successfully.")
except Exception as e:
    print(f"Failed to export code: {e}")

# Triggering the Import Code routine from the add-in
try:
    print("Triggering VBA Code Import...")
    excel.Application.Run("'VBA_Version_Control.xlam'!modVbaSync.ImportCode")

    # Note: If the VBA project is locked with a password, the add-in uses a RAM-level
    # bypass that leverages Application.OnTime to asynchronously hook and unlock it.
    # Therefore, we MUST wait for the background operation to complete before continuing.
    time.sleep(3)
    print("VBA Code Import completed.")
except Exception as e:
    print(f"Failed to import code: {e}")
```

## Agent Workflow for Automated Testing

When generating new VBA code and tasked with testing it, follow this workflow:

1.  **Generate the Code**: Write the `.bas`, `.cls`, or `.frm` file containing the new VBA code to the appropriate directory structure used by the VBA Version Control Addin (e.g., `<TargetFolder>/VBA_<WorkbookName>/Modules/`).
2.  **Generate a Test Macro**: Create a specific VBA subroutine designed to assert conditions and return results (or log them to a cell) based on the newly generated code. Ensure this is saved alongside the other files.
3.  **Import via Python**: Use a Python script with `win32com` to open the target workbook and call the Add-in's `ImportCode` macro to pull the new files into the workbook. Wait for the import to finish (e.g., `time.sleep(5)` if it was locked).
4.  **Execute the Test via Python**: Use `excel.Application.Run("'YourWorkbook.xlsm'!YourTestMacro")` to execute the test.
5.  **Evaluate Results**: Read the outputs (either returned directly or written to specific worksheet cells using `excel.Range("A1").Value`) to determine if the test passed.
6.  **Cleanup**: Close the workbook (saving if the test passed, or discarding changes if it failed).

```python
# Cleanup example:
wb.Close(SaveChanges=False) # Or True
excel.Quit()
```

## Important Considerations

- **Trust Access to the VBA Project Object Model**: This setting MUST be enabled in Excel's Trust Center (`File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings`) for any programmatic modification or export/import of VBA projects to succeed.
- **Asynchronous Actions**: The `modVBAPasswordBypass` mechanism in the add-in relies on `Application.OnTime`. When scripting with Python, be aware that you MUST implement a delay (`time.sleep`) to wait for the asynchronous unlocking and code importing/exporting to complete before trying to run the imported macros.
- **Optional Parameters**: The Ribbon macros `ExportCode` and `ImportCode` take an `Optional ByVal control As IRibbonControl`. Because they are optional, Python's `Application.Run` can call them without passing the Ribbon control argument.
