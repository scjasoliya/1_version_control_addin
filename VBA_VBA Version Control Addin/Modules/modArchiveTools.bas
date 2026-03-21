Attribute VB_Name = "modArchiveTools"
'@Folder("ZIP UNZIP")
' Module: modArchiveTools
' Purpose: Extract and repackage Excel .xlsm/.xlam files as ZIP archives for inspection or modification.
Option Explicit

' Debug flag for this module (set True to enable logging)
Private Const DEBUG_MODE As Boolean = False

' Enable PowerShell fallback for extraction if Shell fails
Private Const ALLOW_POWERSHELL_FALLBACK As Boolean = True

' Shell CopyHere flags (FOF_* from SHFileOperation)
Private Const FOF_NOCONFIRMATION As Long = &H10
Private Const FOF_NOCONFIRMMKDIR As Long = &H200
Private Const FOF_NOERRORUI As Long = &H400

'@EntryPoint
'@Ignore ParameterNotUsed
Public Sub ExtractExcelArchive(ByVal control As IRibbonControl)
    ExtractExcelFile
End Sub

'@EntryPoint
'@Ignore ParameterNotUsed
Public Sub RepackExcelArchive(ByVal control As IRibbonControl)
    RepackExcelFolder
End Sub

'===========================
' Extract the contents of a selected .xlsm/.xlam file into a folder (named after the file).
'===========================
Private Sub ExtractExcelFile()
    Dim filePath As String
    Dim baseName As String
    Dim targetFolder As String

    Dim tempZip As String
    Dim shellZipPath As String
    Dim altZipCopy As String

    Dim srcFolder As Object
    Dim destFolder As Object

    Dim ShellApp As Object
    Dim parentPath As String
    Dim copyOk As Boolean

    On Error GoTo CleanFail

    ' Act on the active workbook
    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook open.", vbExclamation, "No Active Workbook"
        Exit Sub
    End If
    If ActiveWorkbook.path = vbNullString Then
        MsgBox "Please save the active workbook to a file before extracting.", vbExclamation, "Workbook Not Saved"
        Exit Sub
    End If

    filePath = ActiveWorkbook.FullName

    ' Derive base name and target folder path
    baseName = GetBaseName(filePath)
    parentPath = GetFolderFromPath(filePath)
    targetFolder = parentPath & baseName

    ' Ensure parent exists; create destination folder (clean if already present)
    If Not EnsureFolderExists(parentPath) Then
        MsgBox "Parent folder does not exist or cannot be accessed: " & parentPath, vbExclamation, "Invalid Path"
        Exit Sub
    End If

    If FolderExists(targetFolder) Then
        If MsgBox("Folder '" & baseName & "' already exists. Overwrite its contents?", _
                  vbYesNo + vbQuestion, "Folder Exists") = vbNo Then
            Exit Sub
        End If
        On Error Resume Next
        DeleteFolderRecursive targetFolder
        On Error GoTo 0
    End If

    If Not EnsureFolderExists(targetFolder) Then
        MsgBox "Could not create destination folder: " & targetFolder, vbCritical, "Create Folder Failed"
        Exit Sub
    End If

    If DEBUG_MODE Then Debug.Print "Extracting file: " & filePath & " -> Folder: " & targetFolder

    ' ----- TEMP ZIP PATH (robust) -----
    tempZip = GetTempFolder() & baseName & "_tmp.zip"

    ' Clean any previous temp zip
    On Error Resume Next: Kill tempZip: On Error GoTo 0

    ' ----- COPY SOURCE TO TEMP ZIP (with diagnostics & fallback) -----
    ' Use SaveCopyAs to avoid file locking issues with the active workbook
    On Error Resume Next
    ActiveWorkbook.SaveCopyAs tempZip
    copyOk = (Err.Number = 0)
    On Error GoTo CleanFail

    If Not copyOk Then
        MsgBox "Failed to copy to temporary zip." & vbCrLf & _
               "From: " & filePath & vbCrLf & _
               "To:   " & tempZip, vbCritical, "Copy Failed"
        GoTo CleanFail
    End If

    If DEBUG_MODE Then
        Debug.Print "Temp ZIP path: " & tempZip & "  (exists=" & FileExists(tempZip) & ", len=" & Len(tempZip) & ")"
    End If

    If Not FileExists(tempZip) Then
        MsgBox "Temporary zip was not created: " & tempZip, vbCritical, "Copy Failed"
        GoTo CleanFail
    End If

    ' Build a Shell-friendly path to the temp zip (short path or simple alt copy)
    shellZipPath = GetShellFriendlyZipPath(tempZip, altZipCopy)
    If DEBUG_MODE Then Debug.Print "Shell ZIP path used: " & shellZipPath & IIf(altZipCopy <> vbNullString, " (alt copy)", vbNullString)

    ' Retry until Shell.Namespace(shellZipPath) returns a folder (race/timing fix)
    Set ShellApp = CreateObject("Shell.Application")
    Set srcFolder = SafeShellNamespace(ShellApp, shellZipPath, 5000, 100) ' up to 5s

    ' If Shell still refuses, try PowerShell Expand-Archive fallback (optional)
    If srcFolder Is Nothing And ALLOW_POWERSHELL_FALLBACK Then
        If DEBUG_MODE Then Debug.Print "Shell.Namespace failed. Trying PowerShell Expand-Archive fallback..."
        If ExpandArchiveWithPowerShell(shellZipPath, targetFolder) Then
            GoTo ExtractDone                     ' skip Shell extraction, PowerShell already extracted
        End If
    End If

    If srcFolder Is Nothing Then
        MsgBox "The temporary ZIP exists but could not be opened by Shell." & vbCrLf & _
               "This is usually due to a long/UNC/non-ASCII path issue." & vbCrLf & _
               "Temp path used: " & vbCrLf & shellZipPath, _
               vbCritical, "ZIP Open Failed"
        GoTo CleanFail
    End If

    Set destFolder = SafeShellNamespace(ShellApp, targetFolder, 3000, 100)
    If destFolder Is Nothing Then
        MsgBox "Destination folder cannot be opened by Shell: " & targetFolder, _
               vbCritical, "Invalid Destination"
        GoTo CleanFail
    End If

    ' Extract: copy all items from the ZIP to the target folder
    destFolder.CopyHere srcFolder.Items, FOF_NOCONFIRMATION Or FOF_NOERRORUI Or FOF_NOCONFIRMMKDIR

    ' Wait for typical OOXML marker file to appear (more robust than fixed sleep)
    If Not WaitForFile(targetFolder & Application.PathSeparator & "[Content_Types].xml", 12000, 150) Then
        ' Fallback brief wait; large files may take longer
        Application.Wait (Now + TimeValue("0:00:02"))
    End If

ExtractDone:
    On Error Resume Next
    If Len(altZipCopy) > 0 And altZipCopy <> tempZip Then Kill altZipCopy
    Kill tempZip
    On Error GoTo 0

    If DEBUG_MODE Then Debug.Print "Extraction completed for: " & baseName

    MsgBox "Extracted contents of '" & baseName & "' to folder:" & vbCrLf & targetFolder, _
           vbInformation, "Extraction Complete"
    Exit Sub

CleanFail:
    On Error Resume Next
    If Len(altZipCopy) > 0 Then Kill altZipCopy
    If Len(tempZip) > 0 Then Kill tempZip
    On Error GoTo 0
End Sub

'===========================
' Repack a folder back into .xlsm/.xlam with timestamped name.
'===========================



Private Sub RepackExcelFolder()
    Dim folderPath As String
    Dim parentDir As String
    Dim baseName As String

    Dim fdlg As FileDialog
    Dim ext As String
    Dim timeStamp As String
    Dim outputZip As String
    Dim outputFile As String


    ' Prompt for folder containing extracted OOXML content
    Set fdlg = Application.FileDialog(msoFileDialogFolderPicker)
    With fdlg
        .Title = "Select Extracted Folder to Repack into Excel File"
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With

    ' Normalize path
    If Right$(folderPath, 1) <> Application.PathSeparator Then
        folderPath = folderPath & Application.PathSeparator
    End If

    ' Validate folder exists and OOXML marker present at this level
    If Not FolderExists(folderPath) Then
        MsgBox "Selected folder does not exist or is not accessible:" & vbCrLf & folderPath, _
               vbCritical, "Folder Missing"
        Exit Sub
    End If

    If Dir$(folderPath & "[Content_Types].xml") = vbNullString Then
        MsgBox "The selected folder does not contain an OOXML package ([Content_Types].xml not found)." & vbCrLf & _
               "Folder: " & folderPath, _
               vbCritical, "Invalid Package"
        Exit Sub
    End If

    baseName = GetFolderName(Left$(folderPath, Len(folderPath) - 1))
    parentDir = GetFolderFromPath(Left$(folderPath, Len(folderPath) - 1))

    ' Determine extension from content types; default to xlsm
    ext = DetermineExcelExtension(folderPath)
    If ext = vbNullString Then
        ext = "xlsm"
        If DEBUG_MODE Then Debug.Print "Could not determine Excel type; defaulting to .xlsm"
    End If

    timeStamp = Format$(Now, "yyyy-mm-dd_HHMM")
    outputZip = parentDir & baseName & "_" & timeStamp & ".zip"
    outputFile = parentDir & baseName & "_" & timeStamp & "." & ext

    ' Clean any previous outputs
    On Error Resume Next
    Kill outputZip
    Kill outputFile
    On Error GoTo 0

    If DEBUG_MODE Then Debug.Print "Repacking (PowerShell) folder contents: " & folderPath & " -> " & outputZip

    ' *** PowerShell zips one level deeper by using folder\*
    If Not CompressFolderWithPowerShell(folderPath, outputZip) Then
        MsgBox "PowerShell Compress-Archive failed to create ZIP." & vbCrLf & _
               "Folder: " & folderPath & vbCrLf & "ZIP: " & outputZip, _
               vbCritical, "ZIP Failed"
        Exit Sub
    End If

    ' Rename .zip to target Excel extension
    On Error GoTo RenameFail
    Name outputZip As outputFile
    On Error GoTo 0

    If DEBUG_MODE Then Debug.Print "Repackaging completed: created file " & outputFile

    MsgBox "Repacked folder '" & baseName & "' into file:" & vbCrLf & outputFile, _
           vbInformation, "Repack Complete"
    Exit Sub

RenameFail:
    MsgBox "Failed to rename ZIP to final Excel file: " & vbCrLf & outputZip & " -> " & outputFile, _
           vbCritical, "Rename Failed"
End Sub

'===========================
' Helpers
'===========================

' Resolve a reliable TEMP folder with trailing backslash.
Public Function GetTempFolder() As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim p As String

    p = Environ$("TEMP")
    If Len(p) = 0 Or Not fso.FolderExists(p) Then p = Environ$("TMP")

    If Len(p) = 0 Or Not fso.FolderExists(p) Then
        On Error Resume Next
        p = fso.GetSpecialFolder(2)              ' TemporaryFolder
        On Error GoTo 0
    End If

    If Len(p) = 0 Or Not fso.FolderExists(p) Then
        p = Application.DefaultFilePath
    End If

    If Right$(p, 1) <> Application.PathSeparator Then p = p & Application.PathSeparator
    If DEBUG_MODE Then Debug.Print "Resolved TEMP folder: " & p
    GetTempFolder = p
End Function

' Try to make a Shell-friendly path to a ZIP:
' 1) Use 8.3 short path if available.
' 2) If path looks long/UNC/non-ASCII, copy to a simple local path like C:\Temp\zip_yyyymmdd_hhnnss.zip and return that.
Private Function GetShellFriendlyZipPath(ByVal zipPath As String, ByRef altCopyOut As String) As String
    On Error GoTo Fallback
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim sp As String
    altCopyOut = vbNullString

    ' Short path (8.3) if available
    sp = vbNullString
    If fso.FileExists(zipPath) Then
        On Error Resume Next
        sp = fso.GetFile(zipPath).ShortPath      ' empty if 8.3 disabled
        On Error GoTo 0
    End If
    If Len(sp) > 0 Then
        GetShellFriendlyZipPath = sp
        Exit Function
    End If

    ' If original path is short/simple ASCII, use it as-is
    If Len(zipPath) < 240 And Not IsUNCPath(zipPath) And Not ContainsNonAscii(zipPath) Then
        GetShellFriendlyZipPath = zipPath
        Exit Function
    End If

Fallback:
    ' Copy to a very simple local folder with a short ASCII name
    Dim simpleFolder As String
    simpleFolder = EnsureSimpleLocalTempFolder()
    If Len(simpleFolder) = 0 Then
        ' As last resort, return original
        GetShellFriendlyZipPath = zipPath
        Exit Function
    End If

    Dim simpleZip As String
    simpleZip = simpleFolder & "zip_" & Format$(Now, "yyyymmdd_hhnnss") & ".zip"

    If CopyFileRobust(zipPath, simpleZip, True) Then
        altCopyOut = simpleZip
        GetShellFriendlyZipPath = simpleZip
    Else
        GetShellFriendlyZipPath = zipPath
    End If
End Function

' Ensure a very simple local folder exists (e.g., C:\Temp\). Returns path with trailing backslash or "" if cannot create.
Private Function EnsureSimpleLocalTempFolder() As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim p As String

    p = "C:\" & "Temp" & Application.PathSeparator
    On Error Resume Next
    If Not fso.FolderExists(p) Then fso.CreateFolder p
    If Err.Number <> 0 Then
        Err.Clear
        ' Fallback to the root of C: (not ideal, but short)
        p = "C:\"
        If Right$(p, 1) <> Application.PathSeparator Then p = p & Application.PathSeparator
    End If
    On Error GoTo 0

    If Right$(p, 1) <> Application.PathSeparator Then p = p & Application.PathSeparator
    EnsureSimpleLocalTempFolder = p
End Function

Private Function IsUNCPath(ByVal p As String) As Boolean
    IsUNCPath = (Left$(p, 2) = "\\")
End Function

Private Function ContainsNonAscii(ByVal p As String) As Boolean
    Dim i As Long
    Dim ch As Integer

    For i = 1 To Len(p)
        ch = AscW(Mid$(p, i, 1))
        If ch < 0 Or ch > 127 Then
            ContainsNonAscii = True
            Exit Function
        End If
    Next
    ContainsNonAscii = False
End Function

' Robust copy using FileCopy then fallback to FSO.CopyFile
Public Function CopyFileRobust(ByVal src As String, ByVal dest As String, ByVal overwrite As Boolean) As Boolean
    Dim ok As Boolean
    Dim err1 As Long
    Dim desc1 As String

    On Error Resume Next
    If overwrite Then
        Kill dest
    End If
    On Error GoTo 0

    On Error Resume Next
    FileCopy src, dest
    ok = (Err.Number = 0)
    If Not ok Then
        err1 = Err.Number: desc1 = Err.Description
    End If
    On Error GoTo 0

    If Not ok Then
        On Error Resume Next
        CreateObject("Scripting.FileSystemObject").CopyFile src, dest, True
        ok = (Err.Number = 0)
        On Error GoTo 0
        If Not ok And DEBUG_MODE Then
            Debug.Print "CopyFileRobust failed. FileCopy err=" & err1 & " (" & desc1 & ") ; FSO err=" & Err.Number & " (" & Err.Description & ")"
        End If
    End If

    CopyFileRobust = ok
End Function

' Get the file name without extension from a full file path.
Public Function GetBaseName(ByVal filePath As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    GetBaseName = fso.GetBaseName(filePath)
End Function

' Get the folder path (with trailing backslash) from a full file path or folder path.
Public Function GetFolderFromPath(ByVal fileOrFolderPath As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folderPath As String
    On Error Resume Next
    folderPath = fso.GetParentFolderName(fileOrFolderPath)
    On Error GoTo 0
    If folderPath = vbNullString Then
        folderPath = Application.DefaultFilePath
    End If
    If Right$(folderPath, 1) <> Application.PathSeparator Then
        folderPath = folderPath & Application.PathSeparator
    End If
    GetFolderFromPath = folderPath
End Function

' Get just the name of a folder from a full folder path.
Private Function GetFolderName(ByVal folderPath As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    GetFolderName = fso.GetFileName(folderPath)
End Function

' Recursively delete a folder and all its contents.
' Errors are fully suppressed — never propagates to caller.
Public Sub DeleteFolderRecursive(ByVal folderPath As String)
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderPath) Then
        fso.DeleteFolder folderPath, True
    End If
    On Error GoTo 0
End Sub

' Ensure a folder exists; create it if needed (single level).
Public Function EnsureFolderExists(ByVal folderPath As String) As Boolean
    On Error GoTo Fail
    If Len(folderPath) = 0 Then
        EnsureFolderExists = False
        Exit Function
    End If
    Dim checkPath As String
    If Right$(folderPath, 1) = Application.PathSeparator Then
        checkPath = Left$(folderPath, Len(folderPath) - 1)
    Else
        checkPath = folderPath
    End If
    If Dir$(checkPath, vbDirectory) = vbNullString Then
        MkDir checkPath
    End If
    EnsureFolderExists = True
    Exit Function
Fail:
    EnsureFolderExists = False
End Function

' Check if a file exists.
Public Function FileExists(ByVal path As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir$(path)) > 0)
    On Error GoTo 0
End Function

' Check if a folder exists.
Public Function FolderExists(ByVal path As String) As Boolean
    On Error Resume Next
    FolderExists = (Len(Dir$(path, vbDirectory)) > 0)
    On Error GoTo 0
End Function

' Retry opening a Shell namespace for up to timeoutMs; returns Nothing if it never opens.
Private Function SafeShellNamespace(ByVal ShellApp As Object, ByVal path As String, _
                                    ByVal timeoutMs As Long, ByVal sleepMs As Long) As Object
    Dim startT As Single: startT = Timer
    Dim ns As Object
    Do
        Set ns = ShellApp.Namespace(path)
        If Not ns Is Nothing Then
            Set SafeShellNamespace = ns
            Exit Function
        End If
        DoEvents
        PauseMs sleepMs
        If (Timer - startT) * 1000# >= timeoutMs Then Exit Do
    Loop
    Set SafeShellNamespace = Nothing
End Function

' Busy-wait pause using Timer/DoEvents (keeps UI responsive).
Private Sub PauseMs(ByVal ms As Long)
    Dim T As Single: T = Timer
    Do While (Timer - T) * 1000# < ms
        DoEvents
    Loop
End Sub

' Wait until a file appears (exists) with retry; returns True if found.
Private Function WaitForFile(ByVal path As String, ByVal timeoutMs As Long, ByVal sleepMs As Long) As Boolean
    Dim startT As Single: startT = Timer
    Do
        If FileExists(path) Then
            WaitForFile = True
            Exit Function
        End If
        DoEvents
        PauseMs sleepMs
        If (Timer - startT) * 1000# >= timeoutMs Then Exit Do
    Loop
    WaitForFile = False
End Function

' Determine correct Excel extension by inspecting [Content_Types].xml in an extracted folder.
' Returns "xlsm", "xlsb", "xlsx", etc., or "" if not identified.
Private Function DetermineExcelExtension(ByVal folderPath As String) As String
    Dim contentFile As String
    Dim fnum As Integer
    Dim content As String
    Dim ext As String


    contentFile = folderPath & "[Content_Types].xml"
    ext = vbNullString

    If Dir$(contentFile) = vbNullString Then
        DetermineExcelExtension = ext
        Exit Function
    End If

    On Error GoTo Done
    fnum = FreeFile
    Open contentFile For Input As #fnum
    content = Input$(LOF(fnum), #fnum)
    Close #fnum

    content = LCase$(content)
    If InStr(content, "application/vnd.ms-excel.addin.macroenabled") > 0 Then
        ext = "xlam"                             ' custom macro-enabled add-in
    ElseIf InStr(content, "application/vnd.ms-excel.sheet.macroenabled") > 0 Or _
           InStr(content, "application/vnd.ms-excel.sheet.macroenabled.12") > 0 Then
        ext = "xlsm"                             ' macro-enabled workbook
    ElseIf InStr(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") > 0 Then
        ext = "xlsx"                             ' standard workbook
    ElseIf InStr(content, "application/vnd.ms-excel.sheet.binary.macroenabled") > 0 Then
        ext = "xlsb"                             ' binary workbook
    End If
Done:
    DetermineExcelExtension = ext
End Function

'===========================
' PowerShell fallback (optional)
'===========================
Public Function ExpandArchiveWithPowerShell(ByVal zipPath As String, ByVal destFolder As String) As Boolean
    On Error GoTo Fail
    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    Dim cmd As String

    ' Quote for PowerShell -LiteralPath (single quotes, escape embedded quotes)
    cmd = "powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command " & _
          """Expand-Archive -LiteralPath         '" & PSQuoteLiteral(zipPath) & "' -DestinationPath '" & PSQuoteLiteral(destFolder) & "' -Force"""

    If DEBUG_MODE Then Debug.Print "PS Cmd: " & cmd
    Dim rc As Long
    rc = wsh.Run(cmd, 0, True)                   ' hidden window, wait

    ExpandArchiveWithPowerShell = (rc = 0)
    Exit Function
Fail:
    ExpandArchiveWithPowerShell = False
End Function

Private Function PSQuoteLiteral(ByVal s As String) As String
    ' Escape single quotes for PowerShell single-quoted literal
    PSQuoteLiteral = Replace(s, "'", "''")
End Function

Public Function CompressFolderWithPowerShell(ByVal srcFolder As String, ByVal outZip As String) As Boolean
    On Error GoTo Fail
    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    Dim cmd As String
    Dim srcNoSlash As String
    Dim srcWildcard As String

    Dim quotedSrc As String
    Dim quotedOut As String


    ' Normalize: remove trailing slash for consistent wildcard
    If Right$(srcFolder, 1) = Application.PathSeparator Then
        srcNoSlash = Left$(srcFolder, Len(srcFolder) - 1)
    Else
        srcNoSlash = srcFolder
    End If

    ' *** Key change: zip the CONTENTS (one level deeper), not the folder itself
    srcWildcard = srcNoSlash & Application.PathSeparator & "*"

    ' Use -Path (supports wildcards); do NOT use -LiteralPath here
    quotedSrc = "'" & Replace(srcWildcard, "'", "''") & "'"
    quotedOut = "'" & Replace(outZip, "'", "''") & "'"

    cmd = "powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command " & _
          """Compress-Archive -Path " & quotedSrc & " -DestinationPath " & quotedOut & " -Force -CompressionLevel Optimal"""

    If DEBUG_MODE Then Debug.Print "PS Compress Cmd: " & cmd

    Dim rc As Long
    rc = wsh.Run(cmd, 0, True)                   ' hidden, wait
    CompressFolderWithPowerShell = (rc = 0)
    Exit Function
Fail:
    CompressFolderWithPowerShell = False
End Function


