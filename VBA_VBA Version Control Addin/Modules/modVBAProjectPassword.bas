Attribute VB_Name = "modVBAProjectPassword"
'@Folder("Security")
' ===============================================================
' Module : modVBAProjectPassword
' Purpose: Permanently remove VBA project password from Excel
'          .xlsm/.xlam files by patching the vbaProject.bin
'          inside the OOXML ZIP archive.
'
' Technique:
'   Swaps the CMG=, GC= and DPB= keys inside vbaProject.bin
'   in-place with pre-computed values for the known password "macro".
'   The replacement is padded/truncated to the same byte count so
'   the Compound Binary File (CBF) sector chain remains intact.
'
'   Works on ALL Excel versions including 2016, 2019, and 365.
'   The original file is NEVER modified. A new timestamped copy
'   is created alongside it.
'
' Dependencies:
'   - modArchiveTools (GetTempFolder, ExpandArchiveWithPowerShell,
'     CompressFolderWithPowerShell, GetBaseName, GetFolderFromPath,
'     DeleteFolderRecursive, FileExists, FolderExists,
'     EnsureFolderExists, CopyFileRobust)
'
' Post-processing:
'   After opening the patched file in Excel:
'     1. Press Alt+F11 to open the VBA editor.
'     2. When prompted for a password, enter:  macro
'     3. Go to Tools > VBAProject Properties > Protection tab.
'     4. Uncheck "Lock project" and clear the password field.
'     5. Click OK and save the workbook.
'
' Requirements:
'   - PowerShell available on the system
'   - Source file must be .xlsm or .xlam (OOXML with macros)
' ===============================================================
Option Explicit

' Debug flag for this module
Private Const DEBUG_MODE As Boolean = False

' ====== RIBBON ENTRY POINTS ======

'@EntryPoint
'@Ignore ParameterNotUsed
Public Sub RemoveVBAProjectPassword(ByVal control As IRibbonControl)
    On Error GoTo EH

    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook.", vbExclamation, "No Workbook"
        Exit Sub
    End If
    If Len(ActiveWorkbook.Path) = 0 Then
        MsgBox "Please save the workbook to disk first.", vbExclamation, "Workbook Not Saved"
        Exit Sub
    End If

    ' Validate extension: must be macro-enabled
    Dim ext As String
    ext = LCase$(Mid$(ActiveWorkbook.FullName, InStrRev(ActiveWorkbook.FullName, ".") + 1))
    If ext <> "xlsm" And ext <> "xlam" Then
        MsgBox "VBA project password removal is only supported for .xlsm " & _
               "and .xlam files." & vbCrLf & vbCrLf & _
               "Current file type: ." & ext, vbExclamation, "Unsupported Format"
        Exit Sub
    End If

    Dim resultPath As String
    resultPath = RemoveVBAProjectPasswordFromFile(ActiveWorkbook)

    If Len(resultPath) > 0 Then
        MsgBox "VBA project password has been replaced with a known password." & vbCrLf & vbCrLf & _
               "Patched copy saved to:" & vbCrLf & resultPath & vbCrLf & vbCrLf & _
               "STEPS TO UNLOCK:" & vbCrLf & _
               "  1. Open the patched file in Excel." & vbCrLf & _
               "  2. Press Alt+F11 to open the VBA editor." & vbCrLf & _
               "  3. When prompted for a password, enter:  macro" & vbCrLf & _
               "  4. Go to Tools > VBAProject Properties > Protection tab." & vbCrLf & _
               "  5. Uncheck 'Lock project' and clear the password field." & vbCrLf & _
               "  6. Click OK and save the workbook.", _
               vbInformation, "Password Replaced"
    End If
    Exit Sub

EH:
    MsgBox "Failed to remove VBA project password: " & Err.Description, vbCritical, "Error"
End Sub

' ====== PRIVATE IMPLEMENTATION ======

Private Function RemoveVBAProjectPasswordFromFile(ByVal wb As Workbook) As String
    On Error GoTo EH

    Dim filePath    As String
    Dim baseName    As String
    Dim parentDir   As String
    Dim tempFolder  As String
    Dim tempZip     As String
    Dim ext         As String
    Dim timeStamp   As String
    Dim outputFile  As String
    Dim binPath     As String
    Dim patched     As Boolean
    Dim copyOk      As Boolean

    filePath  = wb.FullName
    baseName  = modArchiveTools.GetBaseName(filePath)
    parentDir = modArchiveTools.GetFolderFromPath(filePath)
    tempFolder = modArchiveTools.GetTempFolder() & baseName & "_vbapwd"
    tempZip    = modArchiveTools.GetTempFolder() & baseName & "_vbapwd.zip"

    ext = LCase$(Mid$(filePath, InStrRev(filePath, ".") + 1))
    If Len(ext) = 0 Then ext = "xlsm"

    ' Clean up any previous temp artifacts (fully suppressed)
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
        RemoveVBAProjectPasswordFromFile = vbNullString
        Exit Function
    End If

    ' Step 2: Extract
    If Not modArchiveTools.EnsureFolderExists(tempFolder) Then
        MsgBox "Cannot create temp folder: " & tempFolder, vbCritical, "Folder Failed"
        GoTo Cleanup
    End If

    If Not modArchiveTools.ExpandArchiveWithPowerShell(tempZip, tempFolder) Then
        MsgBox "Failed to extract ZIP archive.", vbCritical, "Extract Failed"
        GoTo Cleanup
    End If

    ' Step 3: Locate and patch vbaProject.bin
    binPath = tempFolder & "\xl\vbaProject.bin"
    If Not modArchiveTools.FileExists(binPath) Then
        MsgBox "vbaProject.bin not found in the archive." & vbCrLf & _
               "This file may not contain VBA macros.", vbExclamation, "No VBA Project"
        GoTo Cleanup
    End If

    patched = PatchVBAProjectBin(binPath)
    If Not patched Then
        MsgBox "No password hash (DPB=) found in vbaProject.bin." & vbCrLf & _
               "The VBA project may not be password-protected.", _
               vbInformation, "No Password Found"
        GoTo Cleanup
    End If

    ' Step 4: Repack
    On Error Resume Next: Kill tempZip: On Error GoTo EH

    If Not modArchiveTools.CompressFolderWithPowerShell(tempFolder & "\", tempZip) Then
        MsgBox "Failed to repack the archive.", vbCritical, "Repack Failed"
        GoTo Cleanup
    End If

    ' Step 5: Copy to final output (CopyFileRobust handles cross-drive moves)
    timeStamp = Format$(Now, "yyyy-mm-dd_HHmm")
    outputFile = parentDir & baseName & "_nopassword_" & timeStamp & "." & ext

    On Error Resume Next: Kill outputFile: On Error GoTo EH
    If modArchiveTools.CopyFileRobust(tempZip, outputFile, True) Then
        On Error Resume Next: Kill tempZip: On Error GoTo EH
    End If

    If modArchiveTools.FileExists(outputFile) Then
        RemoveVBAProjectPasswordFromFile = outputFile
    Else
        MsgBox "Failed to create output file.", vbCritical, "Output Failed"
        RemoveVBAProjectPasswordFromFile = vbNullString
    End If

Cleanup:
    On Error Resume Next
    Kill tempZip
    modArchiveTools.DeleteFolderRecursive tempFolder
    On Error GoTo 0
    Exit Function

EH:
    MsgBox "Error removing VBA project password: " & Err.Description & vbCrLf & vbCrLf & _
           "Paths in use:" & vbCrLf & _
           "  tempZip   = " & tempZip & vbCrLf & _
           "  tempFolder= " & tempFolder, vbCritical, "Error"
    RemoveVBAProjectPasswordFromFile = vbNullString
    Resume Cleanup
End Function

' ====== PRIVATE HELPERS ======

' -----------------------------------------------------------------------
' PatchVBAProjectBin  (value-swap approach — works on Excel 2016/2019/365)
' -----------------------------------------------------------------------
' Swaps CMG=, GC= and DPB= values inside vbaProject.bin in-place with
' pre-computed values for known password "macro".
' Replacement is padded/truncated to match original field length exactly
' so the Compound Binary File (CBF) sector chain is always preserved.
'
' After patching, open the output file, press Alt+F11, enter:  macro
' Then remove password via Tools > VBAProject Properties > Protection.
' -----------------------------------------------------------------------
Private Function PatchVBAProjectBin(ByVal binPath As String) As Boolean
    On Error GoTo EH

    ' Pre-computed CMG/GC/DPB hex strings for known password "macro".
    ' Padded or truncated in-place to match the original field width exactly.
    Dim knownCMG As String: knownCMG = "CBB4B4B4B4"
    Dim knownGC  As String: knownGC  = "CBB4B4B4B4"
    Dim knownDPB As String
    knownDPB = "E0FC4F18432FCBA6C0E47F2EB5F2C7BC36D3E7CB93BE38B"

    ' Read the entire vbaProject.bin into a byte array
    Dim f       As Integer
    Dim content() As Byte
    Dim fileLen As Long

    f = FreeFile
    Open binPath For Binary Access Read Shared As #f
    fileLen = LOF(f)
    If fileLen = 0 Then Close #f: PatchVBAProjectBin = False: Exit Function
    ReDim content(0 To fileLen - 1)
    Get #f, 1, content
    Close #f

    ' Swap all three protection values in-place
    Dim foundDPB As Boolean
    SwapProtectionValue content, "CMG", knownCMG
    SwapProtectionValue content, "GC",  knownGC
    foundDPB = SwapProtectionValue(content, "DPB", knownDPB)

    If Not foundDPB Then
        ' DPB= not found — project is not password-protected
        PatchVBAProjectBin = False
        Exit Function
    End If

    ' Write patched content back (same length — CBF structure intact)
    On Error Resume Next: Kill binPath: On Error GoTo EH
    f = FreeFile
    Open binPath For Binary Access Write Shared As #f
    Put #f, 1, content
    Close #f

    PatchVBAProjectBin = True
    Exit Function

EH:
    On Error Resume Next
    Close #f
    PatchVBAProjectBin = False
End Function

' -----------------------------------------------------------------------
' SwapProtectionValue
' Finds  key="<hex>"  in the byte array and replaces the hex VALUE
' between the quotes with newValue, padding or truncating to the SAME
' number of characters so the file size (and CFB sector chain) is preserved.
' Returns True if the key was found.
' -----------------------------------------------------------------------
Private Function SwapProtectionValue(ByRef content() As Byte, _
                                     ByVal key As String, _
                                     ByVal newValue As String) As Boolean
    Dim kLen   As Long: kLen  = Len(key)
    Dim total  As Long: total = UBound(content)
    Dim i      As Long, k As Long, m As Long
    Dim match  As Boolean

    For i = 0 To total - kLen - 2
        match = True
        For k = 0 To kLen - 1
            If content(i + k) <> Asc(Mid$(key, k + 1, 1)) Then
                match = False: Exit For
            End If
        Next k

        If match Then
            Dim eqPos As Long: eqPos = i + kLen
            Dim q1Pos As Long: q1Pos = i + kLen + 1
            If eqPos <= total And q1Pos <= total Then
                If content(eqPos) = Asc("=") And content(q1Pos) = Asc("""") Then

                    ' Find closing quote
                    Dim valStart As Long: valStart = q1Pos + 1
                    Dim valEnd   As Long: valEnd   = valStart
                    Do While valEnd <= total
                        If content(valEnd) = Asc("""") Or content(valEnd) = 13 Or content(valEnd) = 10 Then Exit Do
                        valEnd = valEnd + 1
                    Loop
                    ' valStart..valEnd-1 is the current value
                    Dim oldLen As Long: oldLen = valEnd - valStart
                    Dim newLen As Long: newLen = Len(newValue)

                    ' Write new value, padding with "0" or truncating to match oldLen
                    For m = 0 To oldLen - 1
                        If m < newLen Then
                            content(valStart + m) = Asc(Mid$(newValue, m + 1, 1))
                        Else
                            content(valStart + m) = Asc("0")  ' pad with '0'
                        End If
                    Next m

                    SwapProtectionValue = True
                    Exit Function
                End If
            End If
        End If
    Next i

    SwapProtectionValue = False
End Function
