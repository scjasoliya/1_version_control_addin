Attribute VB_Name = "modSheetProtection"
'@Folder("Security")
' ===============================================================
' Module : modSheetProtection
' Purpose: Remove worksheet and workbook protection from Excel
'          files by manipulating the underlying OOXML package.
'
' Technique:
'   Excel .xlsx/.xlsm/.xlam files are ZIP archives containing XML.
'   Sheet protection is stored as <sheetProtection .../> nodes in
'   xl\worksheets\sheet*.xml files. Workbook-level protection is
'   stored as <workbookProtection .../> in xl\workbook.xml.
'
'   This module:
'     1. Saves a copy of the target workbook as a temp ZIP
'     2. Extracts the ZIP via PowerShell (reuses modArchiveTools)
'     3. Scans and strips all protection XML nodes
'     4. Repacks the folder into a new ZIP
'     5. Renames to the original Excel extension
'
'   The original file is NEVER modified. A new timestamped copy
'   is created alongside it.
'
' Dependencies:
'   - modArchiveTools (GetTempFolder, ExpandArchiveWithPowerShell,
'     CompressFolderWithPowerShell, GetBaseName, GetFolderFromPath,
'     DeleteFolderRecursive, FileExists, FolderExists,
'     EnsureFolderExists)
'
' Requirements:
'   - PowerShell available on the system
'   - File must be saved to disk (not a new unsaved workbook)
' ===============================================================
Option Explicit

' Debug flag for this module
Private Const DEBUG_MODE As Boolean = False

' ====== RIBBON ENTRY POINTS ======

' RemoveSheetProtection
' Ribbon callback: removes all sheet and workbook protection from
' the active workbook, producing a new timestamped copy.
'@EntryPoint
'@Ignore ParameterNotUsed
Public Sub RemoveSheetProtection(ByVal control As IRibbonControl)
    On Error GoTo EH

    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook.", vbExclamation, "No Workbook"
        Exit Sub
    End If
    If Len(ActiveWorkbook.Path) = 0 Then
        MsgBox "Please save the workbook to disk first.", vbExclamation, "Workbook Not Saved"
        Exit Sub
    End If

    Dim resultPath As String
    resultPath = RemoveAllProtectionFromFile(ActiveWorkbook)

    If Len(resultPath) > 0 Then
        MsgBox "Protection removed. Unprotected copy saved to:" & vbCrLf & vbCrLf & _
               resultPath, vbInformation, "Protection Removed"
    End If
    Exit Sub

EH:
    MsgBox "Failed to remove protection: " & Err.Description, vbCritical, "Error"
End Sub

' ====== PUBLIC API ======

' RemoveAllProtectionFromFile
' Strips all <sheetProtection .../> and <workbookProtection .../>
' nodes from the given Excel workbook. Returns the path to the new
' unprotected copy, or vbNullString on failure.
'
' Parameters:
'   wb — the open Workbook object (SaveCopyAs used to avoid file-lock errors)
Private Function RemoveAllProtectionFromFile(ByVal wb As Workbook) As String
    On Error GoTo EH

    Dim filePath    As String
    Dim baseName    As String
    Dim parentDir   As String
    Dim tempFolder  As String
    Dim tempZip     As String
    Dim ext         As String
    Dim timeStamp   As String
    Dim outputFile  As String
    Dim sheetsFound As Long
    Dim wbFound     As Long
    Dim copyOk      As Boolean

    filePath  = wb.FullName
    baseName  = modArchiveTools.GetBaseName(filePath)
    parentDir = modArchiveTools.GetFolderFromPath(filePath)
    tempFolder = modArchiveTools.GetTempFolder() & baseName & "_unprotect"
    tempZip    = modArchiveTools.GetTempFolder() & baseName & "_unprotect.zip"

    ext = LCase$(Mid$(filePath, InStrRev(filePath, ".") + 1))
    If Len(ext) = 0 Then ext = "xlsx"

    ' Clean up any previous temp artifacts (fully suppressed — safe to ignore failures)
    On Error Resume Next
    Kill tempZip
    modArchiveTools.DeleteFolderRecursive tempFolder
    Err.Clear
    On Error GoTo EH

    ' Step 1: Use SaveCopyAs — works even when the workbook is open/locked
    On Error Resume Next
    wb.SaveCopyAs tempZip
    copyOk = (Err.Number = 0)
    On Error GoTo EH

    If Not copyOk Or Not modArchiveTools.FileExists(tempZip) Then
        MsgBox "Failed to create temporary ZIP copy." & vbCrLf & _
               "Target path: " & tempZip & vbCrLf & vbCrLf & _
               "Err: " & Err.Number & " - " & Err.Description, vbCritical, "Copy Failed"
        RemoveAllProtectionFromFile = vbNullString
        Exit Function
    End If

    ' Step 2: Extract
    If Not modArchiveTools.EnsureFolderExists(tempFolder) Then
        MsgBox "Cannot create temp folder: " & vbCrLf & tempFolder, vbCritical, "Folder Failed"
        GoTo Cleanup
    End If

    If Not modArchiveTools.ExpandArchiveWithPowerShell(tempZip, tempFolder) Then
        MsgBox "Failed to extract ZIP archive.", vbCritical, "Extract Failed"
        GoTo Cleanup
    End If

    ' Step 3: Strip sheet protection from all worksheet XML files
    sheetsFound = StripSheetProtection(tempFolder)
    If DEBUG_MODE Then Debug.Print "Stripped sheetProtection from " & sheetsFound & " file(s)"

    ' Step 4: Strip workbook protection from workbook.xml
    wbFound = StripWorkbookProtection(tempFolder)
    If DEBUG_MODE Then Debug.Print "Stripped workbookProtection from " & wbFound & " file(s)"

    If sheetsFound = 0 And wbFound = 0 Then
        MsgBox "No protection nodes found in the workbook.", vbInformation, "Nothing to Remove"
        GoTo Cleanup
    End If

    ' Step 5: Repack
    On Error Resume Next: Kill tempZip: On Error GoTo EH

    If Not modArchiveTools.CompressFolderWithPowerShell(tempFolder & "\", tempZip) Then
        MsgBox "Failed to repack the archive.", vbCritical, "Repack Failed"
        GoTo Cleanup
    End If

    ' Step 6: Rename to final output
    timeStamp = Format$(Now, "yyyy-mm-dd_HHmm")
    outputFile = parentDir & baseName & "_unprotected_" & timeStamp & "." & ext

    On Error Resume Next: Kill outputFile: On Error GoTo EH
    If modArchiveTools.CopyFileRobust(tempZip, outputFile, True) Then
        On Error Resume Next: Kill tempZip: On Error GoTo EH
    End If

    If modArchiveTools.FileExists(outputFile) Then
        RemoveAllProtectionFromFile = outputFile
    Else
        MsgBox "Failed to rename output file.", vbCritical, "Rename Failed"
        RemoveAllProtectionFromFile = vbNullString
    End If

Cleanup:
    On Error Resume Next
    Kill tempZip
    modArchiveTools.DeleteFolderRecursive tempFolder
    On Error GoTo 0
    Exit Function

EH:
    MsgBox "Error removing protection: " & Err.Description & vbCrLf & vbCrLf & _
           "Paths in use:" & vbCrLf & _
           "  tempZip   = " & tempZip & vbCrLf & _
           "  tempFolder= " & tempFolder, vbCritical, "Error"
    RemoveAllProtectionFromFile = vbNullString
    Resume Cleanup
End Function

' ====== PRIVATE HELPERS ======

' StripSheetProtection
' Scans all xl\worksheets\sheet*.xml files and removes
' <sheetProtection .../> nodes. Returns count of files modified.
Private Function StripSheetProtection(ByVal extractFolder As String) As Long
    Dim sheetsDir As String
    Dim fileName  As String
    Dim filePath  As String
    Dim content   As String
    Dim modified  As String
    Dim count     As Long

    sheetsDir = extractFolder & "\xl\worksheets"
    If Not modArchiveTools.FolderExists(sheetsDir) Then
        sheetsDir = extractFolder & "\xl\worksheets\"
        If Not modArchiveTools.FolderExists(Left$(sheetsDir, Len(sheetsDir) - 1)) Then
            StripSheetProtection = 0
            Exit Function
        End If
    End If

    ' Iterate all .xml files in the worksheets folder
    fileName = Dir$(sheetsDir & "\*.xml", vbNormal)
    Do While Len(fileName) > 0
        filePath = sheetsDir & "\" & fileName

        content = ReadFileContent(filePath)
        modified = RemoveXmlNode(content, "sheetProtection")

        If modified <> content Then
            WriteFileContent filePath, modified
            count = count + 1
            If DEBUG_MODE Then Debug.Print "  Stripped sheetProtection from: " & fileName
        End If

        fileName = Dir$
    Loop

    StripSheetProtection = count
End Function

' StripWorkbookProtection
' Removes <workbookProtection .../> from xl\workbook.xml.
' Returns 1 if modified, 0 if not found.
Private Function StripWorkbookProtection(ByVal extractFolder As String) As Long
    Dim wbXmlPath As String
    Dim content   As String
    Dim modified  As String

    wbXmlPath = extractFolder & "\xl\workbook.xml"
    If Not modArchiveTools.FileExists(wbXmlPath) Then
        StripWorkbookProtection = 0
        Exit Function
    End If

    content = ReadFileContent(wbXmlPath)
    modified = RemoveXmlNode(content, "workbookProtection")

    If modified <> content Then
        WriteFileContent wbXmlPath, modified
        StripWorkbookProtection = 1
        If DEBUG_MODE Then Debug.Print "  Stripped workbookProtection from workbook.xml"
    Else
        StripWorkbookProtection = 0
    End If
End Function

' RemoveXmlNode
' Removes all occurrences of a self-closing XML node by tag name.
' Handles both <tagName .../> (self-closing) and <tagName ...>...</tagName>.
'
' Parameters:
'   xmlContent — the full XML string
'   tagName    — the node name to remove (e.g. "sheetProtection")
'
' Returns the modified string with all matching nodes removed.
Private Function RemoveXmlNode(ByVal xmlContent As String, ByVal tagName As String) As String
    Dim result  As String
    Dim posOpen As Long
    Dim posEnd  As Long
    Dim closeTag As String

    result = xmlContent

    ' Remove self-closing nodes: <tagName ... />
    Do
        posOpen = InStr(1, result, "<" & tagName, vbTextCompare)
        If posOpen = 0 Then Exit Do

        ' Find the closing > or />
        posEnd = InStr(posOpen, result, "/>")
        If posEnd > 0 And posEnd < InStr(posOpen, result, ">") + 1 Then
            ' Self-closing: <tagName .../>
            result = Left$(result, posOpen - 1) & Mid$(result, posEnd + 2)
        Else
            ' Could be <tagName ...>...</tagName>
            posEnd = InStr(posOpen, result, ">")
            If posEnd = 0 Then Exit Do

            closeTag = "</" & tagName & ">"
            Dim posClose As Long
            posClose = InStr(posEnd, result, closeTag, vbTextCompare)
            If posClose > 0 Then
                result = Left$(result, posOpen - 1) & Mid$(result, posClose + Len(closeTag))
            Else
                ' Self-closing with space before >: <tagName ... >
                result = Left$(result, posOpen - 1) & Mid$(result, posEnd + 1)
            End If
        End If
    Loop

    RemoveXmlNode = result
End Function

' ReadFileContent
' Reads the entire content of a text file as a string.
Private Function ReadFileContent(ByVal filePath As String) As String
    Dim f As Integer
    On Error GoTo EH
    f = FreeFile
    Open filePath For Binary Access Read Shared As #f
    If LOF(f) > 0 Then
        ReadFileContent = Input$(LOF(f), f)
    Else
        ReadFileContent = vbNullString
    End If
    Close #f
    Exit Function
EH:
    On Error Resume Next
    If f <> 0 Then Close #f
    ReadFileContent = vbNullString
End Function

' WriteFileContent
' Writes a string to a file, overwriting any existing content.
Private Sub WriteFileContent(ByVal filePath As String, ByVal content As String)
    Dim f As Integer
    On Error GoTo EH
    f = FreeFile
    ' Delete existing file first to avoid stale bytes
    On Error Resume Next: Kill filePath: On Error GoTo EH
    Open filePath For Binary Access Write Shared As #f
    Put #f, 1, content
    Close #f
    Exit Sub
EH:
    On Error Resume Next
    If f <> 0 Then Close #f
End Sub
