Attribute VB_Name = "modVbaSync"
'@Folder("Export Import")
' ===============================================================
' modVbaSync
'  - Exports/Imports VBA components to/from folder next to workbook:
'       <WorkbookFolder>\VBA_<WorkbookNameNoExt>\
'           Modules\<Name>.bas
'           Classes\<Name>.cls
'           Forms\<Name>.frm (+ .frx managed by VBE)
'           Documents\<CodeName>.cls
'  - Unlocks protected VBProject using RAM-level bypass
'       (modVBAPasswordBypass.BypassVBAProjectPassword)
'  - Requires: Trust access to VBOM + reference to
'       Microsoft Visual Basic for Applications Extensibility 5.3
' ===============================================================
Option Explicit

' ====== CONFIGURATION ======
Public Const ENABLE_DEBUG_LOGS As Boolean = False

' ====== FILE NAMING ======
Public Enum ComponentKind
    ckStdModule = 1
    ckClassModule = 2
    ckForm = 3
    ckDocument = 4
End Enum

' ====== LOGGING ======
Private Sub LogI(ByVal Msg As String)
    If ENABLE_DEBUG_LOGS Then Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss"); " [INFO] "; Msg
End Sub

Private Sub LogW(ByVal Msg As String)
    If ENABLE_DEBUG_LOGS Then Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss"); " [WARN] "; Msg
End Sub

Private Sub LogE(ByVal Msg As String)
    If ENABLE_DEBUG_LOGS Then Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss"); " [ERROR] "; Msg
End Sub

' ====== BASIC FILE HELPERS ======
Private Function GetParentFolder(ByVal path As String) As String
    Dim i As Long
    For i = Len(path) To 1 Step -1
        If Mid$(path, i, 1) = "\" Or Mid$(path, i, 1) = "/" Then
            GetParentFolder = Left$(path, i - 1)
            Exit Function
        End If
    Next i
    GetParentFolder = vbNullString
End Function

Private Function CombinePath(ByVal a As String, ByVal b As String) As String
    If Right$(a, 1) = "\" Or Right$(a, 1) = "/" Then
        CombinePath = a & b
    Else
        CombinePath = a & "\" & b
    End If
End Function

Private Function EnsureFolderExists(ByVal folderPath As String) As Boolean
    On Error GoTo EH
    If Len(folderPath) = 0 Then Exit Function
    If Dir(folderPath, vbDirectory) <> vbNullString Then
        EnsureFolderExists = True
        Exit Function
    End If
    Dim parent As String: parent = GetParentFolder(folderPath)
    If Len(parent) > 0 And Dir(parent, vbDirectory) = vbNullString Then
        If Not EnsureFolderExists(parent) Then Exit Function
    End If
    MkDir folderPath
    EnsureFolderExists = True
    Exit Function
EH:
    EnsureFolderExists = False
End Function

Private Sub DeleteIfExists(ByVal filePath As String)
    On Error Resume Next
    If Len(filePath) > 0 Then
        If Dir(filePath, vbNormal) <> vbNullString Then Kill filePath
    End If
    On Error GoTo 0
End Sub

' Reads the entire content of a file as a string.
' Used by ImportDocumentCode to load document module source files.
Private Function ReadAllText(ByVal filePath As String) As String
    On Error GoTo EH
    Dim f As Integer: f = FreeFile
    Open filePath For Binary Access Read As #f
    If LOF(f) > 0 Then
        ReadAllText = Input$(LOF(f), f)
    Else
        ReadAllText = vbNullString
    End If
    Close #f
    Exit Function
EH:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    If f <> 0 Then Close #f
    ReadAllText = vbNullString
End Function

Private Function CleanNameWithoutExt(ByVal fileName As String) As String
    Dim base As String: base = fileName
    Dim p As Long: p = InStrRev(base, ".")
    If p > 0 Then base = Left$(base, p - 1)
    base = Replace(base, ":", "_")
    base = Replace(base, "\", "_")
    base = Replace(base, "/", "_")
    base = Replace(base, "*", "_")
    base = Replace(base, "?", "_")
    base = Replace(base, """", "_")
    base = Replace(base, "<", "_")
    base = Replace(base, ">", "_")
    base = Replace(base, "|", "_")
    CleanNameWithoutExt = base
End Function

' ====== TRUST CHECK ======

' Returns True if programmatic access to the VBE is allowed,
' i.e., Trust Center -> "Trust access to the VBA project object model" is enabled.
' ShowMessage: if True, displays a friendly guidance message when trust is not enabled.
' Diagnostic: optional string describing how the check was determined.
Public Function IsVBATrustEnabled(Optional ByVal ShowMessage As Boolean = False, _
                                  Optional ByRef Diagnostic As String = vbNullString) As Boolean
    On Error GoTo NotTrusted_VBE
    
    ' Primary check: accessing the VBE project collection throws when trust is disabled.
    '@Ignore VariableNotUsed
    Dim n As Long
    n = Application.vbe.VBProjects.count
    IsVBATrustEnabled = True
    Diagnostic = "VBE accessible via Application.VBE.VBProjects.Count."
    Exit Function

NotTrusted_VBE:
    ' Secondary check: reading ActiveWorkbook.VBProject.Protection also throws when trust is disabled.
    On Error GoTo NotTrusted_VBProject
    If Not ActiveWorkbook Is Nothing Then
        '@Ignore VariableNotUsed
        Dim prot As Long
        prot = ActiveWorkbook.VBProject.Protection
        IsVBATrustEnabled = True
        Diagnostic = "VBE accessible via ActiveWorkbook.VBProject.Protection."
        Exit Function
    End If

NotTrusted_VBProject:
    ' Fallback (Windows only): read the registry AccessVBOM setting.
    ' Office version string like "16.0" is returned by Application.Version.
    ' Value 1 = trusted, 0 or missing = not trusted.
    Dim ver As String: ver = Application.Version
    Dim regKey As String
    regKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & ver & "\Excel\Security\AccessVBOM"
    
    On Error Resume Next
    Dim sh As Object
    Dim val As Variant

    Set sh = CreateObject("WScript.Shell")
    val = sh.RegRead(regKey)
    On Error GoTo 0
    
    If IsNumeric(val) Then
        IsVBATrustEnabled = (CLng(val) = 1)
        If IsVBATrustEnabled Then
            Diagnostic = "Registry AccessVBOM=1 (" & regKey & ")."
            Exit Function
        Else
            Diagnostic = "Registry AccessVBOM?1 (" & regKey & ")."
        End If
    Else
        ' On Mac or if registry not available/missing key.
        IsVBATrustEnabled = False
        Diagnostic = "Registry check unavailable or key missing (" & regKey & ")."
    End If
    
    If Not IsVBATrustEnabled And ShowMessage Then
        MsgBox "Programmatic access to the VBA project is not trusted." & vbCrLf & vbCrLf & _
               "Enable it via:" & vbCrLf & _
               "File ? Options ? Trust Center ? Trust Center Settings ? Macro Settings ?" & vbCrLf & _
               "[?] Trust access to the VBA project object model" & vbCrLf & vbCrLf & _
               "Then rerun the action.", vbExclamation, "VBA Trust Required"
    End If
End Function

' ====== UNLOCK CHECK ======

' UnlockVBProject is no longer needed since UnlockActiveVBAProject
' handles the asynchronous bypass and callback execution.

' ====== ONEDRIVE PATH RESOLUTION ======
Private Function GetLocalPath(ByVal wbPath As String) As String
    If InStr(1, wbPath, "http://", vbTextCompare) <> 1 And InStr(1, wbPath, "https://", vbTextCompare) <> 1 Then
        GetLocalPath = wbPath
        Exit Function
    End If
    
    Dim normalizedUrl As String
    normalizedUrl = Replace(wbPath, "/", "\")
    
    Dim odConsumer As String
    Dim odCommercial As String
    odConsumer = Environ$("OneDriveConsumer")
    odCommercial = Environ$("OneDriveCommercial")
    If odConsumer = vbNullString Then odConsumer = Environ$("OneDrive")
    If odCommercial = vbNullString Then odCommercial = Environ$("OneDrive")
    
    Dim localPath As String
    localPath = vbNullString
    
    If InStr(1, normalizedUrl, "d.docs.live.net", vbTextCompare) > 0 And Len(odConsumer) > 0 Then
        Dim parts() As String
        parts = Split(normalizedUrl, "\")
        If UBound(parts) >= 4 Then
            Dim relPath As String
            relPath = vbNullString
            Dim i As Long
            For i = 4 To UBound(parts)
                relPath = relPath & "\" & parts(i)
            Next i
            localPath = odConsumer & relPath
        End If
    ElseIf InStr(1, normalizedUrl, "sharepoint.com", vbTextCompare) > 0 And Len(odCommercial) > 0 Then
        Dim docIdx As Long
        docIdx = InStr(1, normalizedUrl, "\Documents\", vbTextCompare)
        If docIdx > 0 Then
            localPath = odCommercial & Mid$(normalizedUrl, docIdx + 10)
        End If
    End If
    
    If Len(localPath) > 0 Then
        localPath = Replace(localPath, "%20", " ")
        GetLocalPath = localPath
    Else
        GetLocalPath = wbPath
    End If
End Function

' ====== FOLDER RESOLUTION NEXT TO WORKBOOK ======
Private Function GetProjectRootFolder(ByVal wb As Workbook) As String
    Dim wbPath As String
    On Error Resume Next
    wbPath = GetLocalPath(wb.FullName)
    On Error GoTo 0
    If Len(wbPath) = 0 Then
        GetProjectRootFolder = vbNullString      ' unsaved workbook
        Exit Function
    End If
    Dim wbDir As String: wbDir = GetParentFolder(wbPath)
    Dim projFolder As String: projFolder = "VBA_" & CleanNameWithoutExt(wb.name)
    GetProjectRootFolder = CombinePath(wbDir, projFolder)
End Function

Private Sub EnsureAllSubfolders(ByVal root As String)
    EnsureFolderExists root
    EnsureFolderExists CombinePath(root, "Modules")
    EnsureFolderExists CombinePath(root, "Classes")
    EnsureFolderExists CombinePath(root, "Forms")
    EnsureFolderExists CombinePath(root, "Documents")
End Sub

Private Function ComponentFilePath(ByVal root As String, ByVal vbComp As VBIDE.VBComponent) As String
    Select Case vbComp.Type
    Case vbext_ct_StdModule
        ComponentFilePath = CombinePath(CombinePath(root, "Modules"), vbNullString & vbComp.name & ".bas")
    Case vbext_ct_ClassModule
        ComponentFilePath = CombinePath(CombinePath(root, "Classes"), vbNullString & vbComp.name & ".cls")
    Case vbext_ct_MSForm
        ComponentFilePath = CombinePath(CombinePath(root, "Forms"), vbNullString & vbComp.name & ".frm")
    Case vbext_ct_Document
        ComponentFilePath = CombinePath(CombinePath(root, "Documents"), vbNullString & vbComp.name & ".cls")
    Case Else
        ComponentFilePath = CombinePath(root, vbComp.name & ".txt")
    End Select
End Function

' ====== NORMALIZATION (Document code) ======
' Strips VERSION/Attribute lines that cannot be injected into document modules
Private Function NormalizeDocumentCode(ByVal rawCode As String) As String
    If Len(rawCode) = 0 Then NormalizeDocumentCode = vbNullString: Exit Function
    Dim lines() As String
    Dim outBuf As String
    Dim i As Long
    Dim s As String
    Dim T As String
    Dim inHeaderBlock As Boolean
    
    Dim safeCode As String
    safeCode = Replace(rawCode, vbCrLf, vbLf)
    safeCode = Replace(safeCode, vbCr, vbLf)
    lines = Split(safeCode, vbLf)
    
    For i = LBound(lines) To UBound(lines)
        s = lines(i)
        T = LCase$(Trim$(s))
        
        If T = "begin" Then
            inHeaderBlock = True
        ElseIf T = "end" Then
            If inHeaderBlock Then
                inHeaderBlock = False
                ' Skip the header's "END"
            Else
                ' This is likely a legitimate "End" statement
                outBuf = outBuf & s & vbCrLf
            End If
        ElseIf inHeaderBlock Then
            ' Skip properties inside BEGIN...END
        ElseIf Left$(s, 8) = "VERSION " Then
            ' Skip VERSION line
        ElseIf Left$(T, 10) = "attribute " Then
            ' Skip Attribute lines
        Else
            ' Keep everything else (code, comments, options)
            outBuf = outBuf & s & vbCrLf
        End If
    Next i
    
    NormalizeDocumentCode = outBuf
End Function

' ===================== PUBLIC MACROS =====================
' --------- EXPORT all components from ActiveWorkbook ---------
'@EntryPoint
'@Ignore ParameterNotUsed
Public Sub ExportCode(Optional ByVal control As IRibbonControl)
    On Error GoTo EH
    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook.", vbExclamation: Exit Sub
    End If
    If Not IsVBATrustEnabled(True) Then
        Exit Sub
    End If
    If Len(ActiveWorkbook.FullName) = 0 Then
        MsgBox "Please save the workbook first.", vbExclamation: Exit Sub
    End If
    
    ' If locked, trigger async bypass with a callback to resume here
    If ActiveWorkbook.VBProject.Protection <> 0 Then
        modVBAPasswordBypass.UnlockActiveVBAProject Nothing, "modVbaSync.ExportCodeAfterUnlock_"
        Exit Sub
    End If
    
    ExportCodeAfterUnlock_
    Exit Sub
EH:
    LogE "Export error: " & Err.Number & " - " & Err.Description
    MsgBox "Export failed: " & Err.Description, vbCritical
End Sub

Public Sub ExportCodeAfterUnlock_()
    On Error GoTo EH
    If ActiveWorkbook Is Nothing Then Exit Sub
    
    Dim root As String: root = GetProjectRootFolder(ActiveWorkbook)
    If Len(root) = 0 Then
        MsgBox "Unable to resolve project folder.", vbCritical: Exit Sub
    End If
    
    ' Clean up existing folder before export to remove stale files
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(root) Then
        fso.DeleteFolder root, True
        LogI "Deleted existing project folder: " & root
    End If
    Set fso = Nothing
    On Error GoTo EH

    EnsureAllSubfolders root
    
    LogI "Export -> " & ActiveWorkbook.FullName & " -> " & root
    Dim vbComp As VBIDE.VBComponent
    For Each vbComp In ActiveWorkbook.VBProject.VBComponents
        Dim skipExport As Boolean: skipExport = False
        
        If vbComp.Type = vbext_ct_Document Then
            Dim cm As VBIDE.CodeModule: Set cm = vbComp.CodeModule
            If cm.CountOfLines = 0 Then
                skipExport = True
            Else
                Dim visibleCode As String
                visibleCode = cm.Lines(1, cm.CountOfLines)
                visibleCode = Replace(visibleCode, "Option Explicit", "", , , vbTextCompare)
                visibleCode = Replace(visibleCode, vbCr, "")
                visibleCode = Replace(visibleCode, vbLf, "")
                visibleCode = Replace(visibleCode, " ", "")
                visibleCode = Replace(visibleCode, vbTab, "")
                
                If Len(visibleCode) = 0 Then
                    skipExport = True
                End If
            End If
        End If
        
        If skipExport Then
            LogI "Skipped exporting empty document: " & vbComp.name
        Else
            Dim fpath As String: fpath = ComponentFilePath(root, vbComp)
            DeleteIfExists fpath
            On Error Resume Next
            vbComp.Export fpath
            If Err.Number <> 0 Then
                LogE "Export failed for " & vbComp.name & ": " & Err.Description
                MsgBox "Export failed for " & vbComp.name & ": " & Err.Description, vbCritical
                Err.Clear
            Else
                LogI "Exported: " & fpath
            End If
            On Error GoTo EH
        End If
    Next vbComp
    
    MsgBox "Export complete to: " & root, vbInformation
    Exit Sub
EH:
    LogE "Export error: " & Err.Number & " - " & Err.Description
    MsgBox "Export failed: " & Err.Description, vbCritical
End Sub

' --------- IMPORT all components into ActiveWorkbook ---------
'@EntryPoint
'@Ignore ParameterNotUsed
Public Sub ImportCode(Optional ByVal control As IRibbonControl)
    On Error GoTo EH
    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook.", vbExclamation: Exit Sub
    End If
    If Not IsVBATrustEnabled(True) Then
        Exit Sub
    End If
    If Len(ActiveWorkbook.FullName) = 0 Then
        MsgBox "Please save the workbook first.", vbExclamation: Exit Sub
    End If
    
    ' If locked, trigger async bypass with a callback to resume here
    If ActiveWorkbook.VBProject.Protection <> 0 Then
        modVBAPasswordBypass.UnlockActiveVBAProject Nothing, "modVbaSync.ImportCodeAfterUnlock_"
        Exit Sub
    End If
    
    ImportCodeAfterUnlock_
    Exit Sub
EH:
    LogE "Import error: " & Err.Number & " - " & Err.Description
    MsgBox "Import failed: " & Err.Description, vbCritical
End Sub

Public Sub ImportCodeAfterUnlock_()
    On Error GoTo EH
    If ActiveWorkbook Is Nothing Then Exit Sub
    
    Dim root As String: root = GetProjectRootFolder(ActiveWorkbook)
    If Dir(root, vbDirectory) = vbNullString Then
        MsgBox "Source folder not found: " & root, vbCritical: Exit Sub
    End If
    
    LogI "Import -> " & ActiveWorkbook.FullName & " -> " & root
    
    ' Import Std Modules
    ImportFromFolder ActiveWorkbook, CombinePath(root, "Modules"), "*.bas", ckStdModule
    ' Import Classes
    ImportFromFolder ActiveWorkbook, CombinePath(root, "Classes"), "*.cls", ckClassModule
    ' Import Forms
    ImportFromFolder ActiveWorkbook, CombinePath(root, "Forms"), "*.frm", ckForm
    ' Import Documents (ThisWorkbook / Worksheets)
    ImportFromFolder ActiveWorkbook, CombinePath(root, "Documents"), "*.cls", ckDocument
    
    ' Check for extras
    CheckAndRemoveExtras ActiveWorkbook, root
    
    MsgBox "VBA Import complete", vbInformation
    'MsgBox "Import complete from: " & root, vbInformation
    Exit Sub
EH:
    LogE "Import error: " & Err.Number & " - " & Err.Description
    MsgBox "Import failed: " & Err.Description, vbCritical
End Sub

' ===================== INTERNAL IMPORT HELPERS =====================

Private Sub ImportFromFolder(ByVal wb As Workbook, ByVal folder As String, ByVal pattern As String, ByVal compKind As ComponentKind)
    On Error GoTo EH
    If Dir(folder, vbDirectory) = vbNullString Then
        LogW "Missing folder, skipping: " & folder
        Exit Sub
    End If
    
    Dim f As String
    f = Dir(CombinePath(folder, pattern), vbNormal)
    Do While Len(f) > 0
        Dim path As String: path = CombinePath(folder, f)
        Dim target As String
        
        target = CleanNameWithoutExt(f)
        
        Select Case compKind
        Case ckStdModule
            ImportStdOrClass wb, path, target, vbext_ct_StdModule
        Case ckClassModule
            ImportStdOrClass wb, path, target, vbext_ct_ClassModule
        Case ckForm
            ImportForm wb, path, target
        Case ckDocument
            ImportDocumentCode wb, path, target
        End Select
        
        f = Dir
        DoEvents
    Loop
    Exit Sub
EH:
    LogE "ImportFromFolder error: " & Err.Description
End Sub

'@Ignore ParameterNotUsed
Private Sub ImportStdOrClass(ByVal wb As Workbook, ByVal filePath As String, _
                             ByVal compName As String, ByVal compType As VBIDE.vbext_ComponentType)
    On Error GoTo EH
    Dim vbProj As VBIDE.VBProject: Set vbProj = wb.VBProject
    
    Dim existing As VBIDE.VBComponent
    Set existing = Nothing
    On Error Resume Next
    Set existing = vbProj.VBComponents(compName)
    On Error GoTo EH
    
    If Not existing Is Nothing Then
        vbProj.VBComponents.Remove existing
        LogI "Removed existing component: " & compName
    End If
    
    Dim newComp As VBIDE.VBComponent
    Set newComp = vbProj.VBComponents.Import(filePath)
    
    ' Fix cross-import: if the file was missing its header (e.g. VERSION 1.0 CLASS),
    ' VBA imports it as a Standard Module instead of a Class Module.
    If newComp.Type <> compType Then
        LogW "Cross-import detected: " & compName & " imported as type " & newComp.Type & " but expected " & compType & ". Forcing correct type."
        vbProj.VBComponents.Remove newComp
        
        Set newComp = vbProj.VBComponents.Add(compType)
        On Error Resume Next
        newComp.name = compName
        On Error GoTo EH
        
        Dim code As String: code = ReadAllText(filePath)
        code = NormalizeDocumentCode(code) ' Strip out attribute headers if any
        
        Dim cm As VBIDE.CodeModule: Set cm = newComp.CodeModule
        If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
        If Len(code) > 0 Then cm.AddFromString code
    End If
    
    If newComp.name <> compName Then
        LogI "Renaming imported component " & newComp.name & " to " & compName
        On Error Resume Next
        newComp.name = compName
        If Err.Number <> 0 Then
            LogE "Failed to rename component " & newComp.name & " to " & compName & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo EH
    End If
    
    LogI "Imported " & compName & " from " & filePath
    Exit Sub
EH:
    LogE "Import std/class failed (" & compName & "): " & Err.Description
End Sub

Private Sub ImportForm(ByVal wb As Workbook, ByVal filePath As String, ByVal formName As String)
    On Error GoTo EH
    Dim vbProj As VBIDE.VBProject: Set vbProj = wb.VBProject
    
    Dim existing As VBIDE.VBComponent
    Set existing = Nothing
    On Error Resume Next
    Set existing = vbProj.VBComponents(formName)
    On Error GoTo EH
    
    If Not existing Is Nothing Then
        vbProj.VBComponents.Remove existing
        LogI "Removed existing form: " & formName
    End If
    
    Dim newComp As VBIDE.VBComponent
    Set newComp = vbProj.VBComponents.Import(filePath)
    
    If newComp.name <> formName Then
        LogI "Renaming imported form " & newComp.name & " to " & formName
        On Error Resume Next
        newComp.name = formName
        If Err.Number <> 0 Then
            LogE "Failed to rename form " & newComp.name & " to " & formName & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo EH
    End If
    
    LogI "Imported form " & formName
    Exit Sub
EH:
    LogE "Import form failed (" & formName & "): " & Err.Description
End Sub

Private Sub ImportDocumentCode(ByVal wb As Workbook, ByVal filePath As String, ByVal docCodeName As String)
    On Error GoTo EH
    Dim vbProj As VBIDE.VBProject: Set vbProj = wb.VBProject
    Dim comp As VBIDE.VBComponent
    Set comp = Nothing
    On Error Resume Next
    Set comp = vbProj.VBComponents(docCodeName)  ' Document components use CodeName as component name
    On Error GoTo EH
    
    If comp Is Nothing Then
        LogW "Document component not found (CodeName=" & docCodeName & "). Skipping: " & filePath
        Exit Sub
    End If
    
    Dim code As String: code = ReadAllText(filePath)
    code = NormalizeDocumentCode(code)           ' strip headers/attributes
    
    Dim cm As VBIDE.CodeModule: Set cm = comp.CodeModule
    Dim n As Long: n = cm.CountOfLines
    If n > 0 Then cm.DeleteLines 1, n
    If Len(code) > 0 Then cm.AddFromString code
    
    LogI "Replaced code for document component: " & docCodeName
    Exit Sub
EH:
    LogE "Import document failed (" & docCodeName & "): " & Err.Description
End Sub

' ===================== EXTRA COMPONENT CHECK =====================

Private Sub CheckAndRemoveExtras(ByVal wb As Workbook, ByVal root As String)
    On Error GoTo EH
    
    ' 1. Collect expected component names from files
    Dim expected As Object
    Set expected = CreateObject("Scripting.Dictionary")
    expected.CompareMode = 1                     ' TextCompare

    ' Helper to add from folder:
    CollectNamesFromFolder expected, CombinePath(root, "Modules"), "*.bas"
    CollectNamesFromFolder expected, CombinePath(root, "Classes"), "*.cls"
    CollectNamesFromFolder expected, CombinePath(root, "Forms"), "*.frm"
    CollectNamesFromFolder expected, CombinePath(root, "Documents"), "*.cls"
    
    ' 2. Identify extras in WB
    Dim extras As Object
    Set extras = CreateObject("Scripting.Dictionary")
    
    Dim vbc As VBIDE.VBComponent
    For Each vbc In wb.VBProject.VBComponents
        ' We only care about removable types: StdModule, ClassModule, Form
        ' Documents (Sheets) are usually tied to data, so we don't delete the component just because code is missing,
        ' unless we want to clear code. Identifying "extra sheets" is risky.
        ' User request: "extra module/vba object".
        ' Safest approach: Ignore Documents for deletion proposal.
        
        Dim isRemovable As Boolean
        isRemovable = (vbc.Type = vbext_ct_StdModule Or _
                       vbc.Type = vbext_ct_ClassModule Or _
                       vbc.Type = vbext_ct_MSForm)
                       
        If isRemovable Then
            If Not expected.Exists(vbc.name) Then
                extras.Add vbc.name, vbc
            End If
        End If
    Next vbc
    
    If extras.count = 0 Then Exit Sub
    
    ' 3. Prompt User
    Dim Msg As String
    Msg = "The following components exist in the workbook but NOT in the source folder:" & vbCrLf & vbCrLf
    
    Dim k As Variant
    Dim count As Long: count = 0
    For Each k In extras.Keys
        Msg = Msg & " - " & k & vbCrLf
        count = count + 1
        If count > 15 Then
            Msg = Msg & "... and " & (extras.count - 15) & " more."
            Exit For
        End If
    Next k
    
    Msg = Msg & vbCrLf & "Do you want to DELETE these extra components from the workbook?"
    
    If MsgBox(Msg, vbQuestion + vbYesNo, "Remove Extra Components?") = vbYes Then
        Dim deletedList As String
        For Each k In extras.Keys
            Dim cToRemove As VBIDE.VBComponent
            Set cToRemove = extras(k)
            wb.VBProject.VBComponents.Remove cToRemove
            LogI "Deleted extra component: " & k
            deletedList = deletedList & vbCrLf & " - " & k
        Next k
        MsgBox "Extra components deleted:" & deletedList, vbInformation
    End If
    
    Exit Sub
EH:
    LogE "Error checking extras: " & Err.Description
End Sub

Private Sub CollectNamesFromFolder(ByVal dict As Object, ByVal folder As String, ByVal pattern As String)
    Dim f As String
    If Dir(folder, vbDirectory) = vbNullString Then Exit Sub
    f = Dir(CombinePath(folder, pattern), vbNormal)
    Do While Len(f) > 0
        Dim name As String
        name = CleanNameWithoutExt(f)
        If Not dict.Exists(name) Then dict.Add name, True
        f = Dir
    Loop
End Sub


