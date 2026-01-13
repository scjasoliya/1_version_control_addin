Attribute VB_Name = "modVbaSync"
'@Folder("Export Import")
' ===============================================================
' modVbaSync (combined with helpers)
'  - Exports/Imports VBA components to/from folder next to workbook:
'       <WorkbookFolder>\VBA_<WorkbookNameNoExt>\
'           Modules\Module_<Name>.bas
'           Classes\Class_<Name>.cls
'           Forms\Form_<Name>.frm (+ .frx managed by VBE)
'           Documents\Document_<CodeName>.cls
'  - Unlocks protected VBProject using password from:
'       "<WorkbookFullName>.password" (preferred)
'       or "<WorkbookFullName>.passward" (supported)
'  - Requires: Trust access to VBOM + reference to
'       Microsoft Visual Basic for Applications Extensibility 5.3
'
'  NOTE: This code cannot export code if there is password and password is provided in defined the file.
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
Private Sub LogI(ByVal msg As String)
    If ENABLE_DEBUG_LOGS Then Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss"); " [INFO] "; msg
End Sub

Private Sub LogW(ByVal msg As String)
    If ENABLE_DEBUG_LOGS Then Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss"); " [WARN] "; msg
End Sub

Private Sub LogE(ByVal msg As String)
    If ENABLE_DEBUG_LOGS Then Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss"); " [ERROR] "; msg
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
    n = Application.VBE.VBProjects.Count
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

' ====== PASSWORD FILE LOGIC ======
' Preferred: "{WorkbookFullName}.password"
' Also supported: "{WorkbookFullName}.passward"
Private Function GetPasswordFilePath(ByVal wb As Workbook) As String
    Dim p1 As String
    Dim p2 As String

    p1 = wb.FullName & ".password"
    p2 = wb.FullName & ".passward"
    If Dir(p1, vbNormal) <> vbNullString Then
        GetPasswordFilePath = p1
    ElseIf Dir(p2, vbNormal) <> vbNullString Then
        GetPasswordFilePath = p2
    Else
        GetPasswordFilePath = vbNullString
    End If
End Function

Private Function GetPasswordFromFile(ByVal wb As Workbook) As String
    On Error GoTo EH
    Dim pf As String: pf = GetPasswordFilePath(wb)
    If Len(pf) = 0 Then Exit Function
    
    Dim raw As String: raw = ReadAllText(pf)
    If Len(raw) = 0 Then Exit Function
    
    Dim lines() As String
    lines = Split(raw, vbCrLf)
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim s As String: s = Trim$(Replace(lines(i), vbCr, vbNullString))
        If Len(s) > 0 Then
            GetPasswordFromFile = s              ' first non-empty line
            Exit Function
        End If
    Next i
    Exit Function
EH:
    GetPasswordFromFile = vbNullString
End Function

' ====== UNLOCK (best-effort UI automation + manual fallback) ======

Private Function TryUnlockWithSendKeys(ByVal wb As Workbook, ByVal pwd As String) As Boolean
    On Error GoTo EH
    LogI "Attempting UI unlock (SendKeys) for: " & wb.Name
    
    '@Ignore MemberNotOnInterface
    Application.Activate: DoEvents
    
    ' Show VBE and try to set active project
    Application.VBE.MainWindow.Visible = True
    DoEvents
    
    ' Open VBE
    Application.SendKeys "%{F11}", True: DoEvents
    Application.Wait Now + TimeSerial(0, 0, 1)
    
    ' Try to set active project (may error if locked; ignore)
    On Error Resume Next
    Set Application.VBE.ActiveVBProject = wb.VBProject
    On Error GoTo 0
    
    ' Show Project Explorer
    Application.SendKeys "^{R}", True: DoEvents
    Application.Wait Now + TimeSerial(0, 0, 1)
    
    ' Type first chars of project name (helps select)
    Dim projName As String: projName = wb.VBProject.Name
    If Len(projName) > 0 Then
        Application.SendKeys Left$(projName, 3), True
        DoEvents
        Application.Wait Now + TimeSerial(0, 0, 1)
    End If
    
    ' Expand to trigger prompt, then enter password
    Application.SendKeys "{RIGHT}", True: DoEvents
    Application.Wait Now + TimeSerial(0, 0, 1)
    Application.SendKeys pwd, True
    Application.SendKeys "~", True
    DoEvents
    Application.Wait Now + TimeSerial(0, 0, 1)
    
    ' Close VBE (optional)
    Application.SendKeys "%{F11}", True
    DoEvents
    
    TryUnlockWithSendKeys = True
    Exit Function
EH:
    TryUnlockWithSendKeys = False
End Function

Private Function UnlockVBProjectWithPassword(ByVal wb As Workbook, _
                                             Optional ByVal mayUseSendKeys As Boolean = True) As Boolean
    On Error GoTo SafeExit
    
    If wb Is Nothing Then Exit Function
    If wb.VBProject.Protection = 0 Then
        UnlockVBProjectWithPassword = True
        Exit Function
    End If
    
    Dim pwd As String: pwd = GetPasswordFromFile(wb)
    If Len(pwd) = 0 Then
        LogW "Password file not found or empty. Expected: " & wb.FullName & ".password (or .passward)"
        MsgBox "VBA project is locked and no password file was found." & vbCrLf & _
               "Create a text file with the password:" & vbCrLf & _
               wb.FullName & ".password", vbExclamation
        Exit Function
    End If
    
    If mayUseSendKeys Then
        If TryUnlockWithSendKeys(wb, pwd) Then
            DoEvents
            If wb.VBProject.Protection = 0 Then
                LogI "VBA project unlocked via UI automation: " & wb.Name
                UnlockVBProjectWithPassword = True
                Exit Function
            End If
        End If
    End If
    
    ' Manual fallback (show found password)
    LogW "SendKeys unlock failed or disabled; prompting manual unlock."
    MsgBox "Please unlock the VBA project manually using this password:" & vbCrLf & vbCrLf & _
           "Workbook: " & wb.FullName & vbCrLf & _
           "Password: " & pwd & vbCrLf & vbCrLf & _
           "Steps: Alt+F11 ? select the project ? enter password.", vbInformation
    
    Dim t0 As Single: t0 = Timer
    Do While wb.VBProject.Protection <> 0
        DoEvents
        If Timer - t0 > 30 Then Exit Do          ' timeout
    Loop
    
SafeExit:
    UnlockVBProjectWithPassword = (wb.VBProject.Protection = 0)
    If Not UnlockVBProjectWithPassword Then
        LogW "Project remains locked: " & wb.Name
    End If
End Function

' ====== FOLDER RESOLUTION NEXT TO WORKBOOK ======
Private Function GetProjectRootFolder(ByVal wb As Workbook) As String
    Dim wbPath As String
    On Error Resume Next
    wbPath = wb.FullName
    On Error GoTo 0
    If Len(wbPath) = 0 Then
        GetProjectRootFolder = vbNullString      ' unsaved workbook
        Exit Function
    End If
    Dim wbDir As String: wbDir = GetParentFolder(wbPath)
    Dim projFolder As String: projFolder = "VBA_" & CleanNameWithoutExt(wb.Name)
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
        ComponentFilePath = CombinePath(CombinePath(root, "Modules"), vbNullString & vbComp.Name & ".bas")
    Case vbext_ct_ClassModule
        ComponentFilePath = CombinePath(CombinePath(root, "Classes"), vbNullString & vbComp.Name & ".cls")
    Case vbext_ct_MSForm
        ComponentFilePath = CombinePath(CombinePath(root, "Forms"), vbNullString & vbComp.Name & ".frm")
    Case vbext_ct_Document
        ComponentFilePath = CombinePath(CombinePath(root, "Documents"), vbNullString & vbComp.Name & ".cls")
    Case Else
        ComponentFilePath = CombinePath(root, vbComp.Name & ".txt")
    End Select
End Function

Private Function ParseFileName(ByVal fileName As String, _
                               ByRef kind As ComponentKind, _
                               ByRef targetName As String) As Boolean
    Dim base As String: base = fileName
    Dim p As Long: p = InStrRev(base, ".")
    Dim stem As String
    Dim ext As String

    If p > 0 Then
        stem = Left$(base, p - 1)
        ext = Mid$(base, p + 1)
    Else
        stem = base
        ext = vbNullString
    End If
    
    If Left$(stem, 7) = "Module_" Then
        kind = ckStdModule: targetName = Mid$(stem, 8)
    ElseIf Left$(stem, 6) = "Class_" Then
        kind = ckClassModule: targetName = Mid$(stem, 7)
    ElseIf Left$(stem, 5) = "Form_" Then
        kind = ckForm: targetName = Mid$(stem, 6)
    ElseIf Left$(stem, 9) = "Document_" Then
        kind = ckDocument: targetName = Mid$(stem, 10)
    Else
        ParseFileName = False: Exit Function
    End If
    
    ' warn if unexpected extension (informational only)
    Select Case kind
    Case ckStdModule: If LCase$(ext) <> "bas" Then LogW "Unexpected extension for StdModule: " & fileName
    Case ckClassModule, ckDocument: If LCase$(ext) <> "cls" Then LogW "Unexpected extension for Class/Document: " & fileName
    Case ckForm: If LCase$(ext) <> "frm" Then LogW "Unexpected extension for Form: " & fileName
    End Select
    
    ParseFileName = True
End Function

' ====== NORMALIZATION (Document code) ======
' Strips VERSION/Attribute lines that cannot be injected into document modules
Private Function NormalizeDocumentCode(ByVal rawCode As String) As String
    If Len(rawCode) = 0 Then NormalizeDocumentCode = vbNullString: Exit Function
    Dim lines() As String
    Dim outBuf As String
    Dim i As Long
    Dim s As String
    Dim t As String
    lines = Split(rawCode, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        s = Replace$(lines(i), vbCr, vbNullString)
        If Len(s) = 0 Then
            outBuf = outBuf & vbCrLf
        Else
            t = LCase$(Trim$(s))
            ' Skip unwanted lines
            If Left$(s, 7) <> "VERSION" _
               And Left$(t, 9) <> "attribute" _
               And Left$(t, 12) <> "end attribute" Then
               
                outBuf = outBuf & s & vbCrLf
            End If
        End If
    Next i

    NormalizeDocumentCode = outBuf
End Function

' ===================== PUBLIC MACROS =====================

' --------- EXPORT all components from ActiveWorkbook ---------
'@EntryPoint
'@Ignore ParameterNotUsed
Public Sub ExportCode(control As IRibbonControl)
    On Error GoTo EH
    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook.", vbExclamation: Exit Sub
    End If
    If Not IsVBATrustEnabled() Then
        MsgBox "Enable 'Trust access to the VBA project object model' and retry.", vbCritical
        Exit Sub
    End If
    If Len(ActiveWorkbook.FullName) = 0 Then
        MsgBox "Please save the workbook first.", vbExclamation: Exit Sub
    End If
    If Not UnlockVBProjectWithPassword(ActiveWorkbook, True) Then Exit Sub
    
    Dim root As String: root = GetProjectRootFolder(ActiveWorkbook)
    If Len(root) = 0 Then
        MsgBox "Unable to resolve project folder.", vbCritical: Exit Sub
    End If
    EnsureAllSubfolders root
    
    LogI "Export ? " & ActiveWorkbook.FullName & " ? " & root
    Dim vbComp As VBIDE.VBComponent
    For Each vbComp In ActiveWorkbook.VBProject.VBComponents
        Dim fpath As String: fpath = ComponentFilePath(root, vbComp)
        DeleteIfExists fpath
        On Error Resume Next
        vbComp.Export fpath
        If Err.Number <> 0 Then
            LogE "Export failed for " & vbComp.Name & ": " & Err.Description
            Err.Clear
        Else
            LogI "Exported: " & fpath
        End If
        On Error GoTo EH
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
Public Sub ImportCode(control As IRibbonControl)
    On Error GoTo EH
    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook.", vbExclamation: Exit Sub
    End If
    If Not IsVBATrustEnabled() Then
        MsgBox "Enable 'Trust access to the VBA project object model' and retry.", vbCritical
        Exit Sub
    End If
    If Len(ActiveWorkbook.FullName) = 0 Then
        MsgBox "Please save the workbook first.", vbExclamation: Exit Sub
    End If
    If Not UnlockVBProjectWithPassword(ActiveWorkbook, True) Then Exit Sub
    
    Dim root As String: root = GetProjectRootFolder(ActiveWorkbook)
    If Dir(root, vbDirectory) = vbNullString Then
        MsgBox "Source folder not found: " & root, vbCritical: Exit Sub
    End If
    
    LogI "Import ? " & ActiveWorkbook.FullName & " ? " & root
    
    ' Import Std Modules
    ImportFromFolder ActiveWorkbook, CombinePath(root, "Modules"), "*.bas"
    ' Import Classes
    ImportFromFolder ActiveWorkbook, CombinePath(root, "Classes"), "*.cls"
    ' Import Forms
    ImportFromFolder ActiveWorkbook, CombinePath(root, "Forms"), "*.frm"
    ' Import Documents (ThisWorkbook / Worksheets)
    ImportFromFolder ActiveWorkbook, CombinePath(root, "Documents"), "*.cls"
    
    MsgBox "VBA Import complete", vbInformation
    'MsgBox "Import complete from: " & root, vbInformation
    Exit Sub
EH:
    LogE "Import error: " & Err.Number & " - " & Err.Description
    MsgBox "Import failed: " & Err.Description, vbCritical
End Sub

' ===================== INTERNAL IMPORT HELPERS =====================

Private Sub ImportFromFolder(ByVal wb As Workbook, ByVal folder As String, ByVal pattern As String)
    On Error GoTo EH
    If Dir(folder, vbDirectory) = vbNullString Then
        LogW "Missing folder, skipping: " & folder
        Exit Sub
    End If
    
    Dim f As String
    f = Dir(CombinePath(folder, pattern), vbNormal)
    Do While Len(f) > 0
        Dim path As String: path = CombinePath(folder, f)
        Dim kind As ComponentKind
        Dim target As String

        If Not ParseFileName(f, kind, target) Then
            LogW "Skipping unrecognized file: " & path
            GoTo NextFile
        End If
        
        Select Case kind
        Case ckStdModule
            ImportStdOrClass wb, path, target, vbext_ct_StdModule
        Case ckClassModule
            ImportStdOrClass wb, path, target, vbext_ct_ClassModule
        Case ckForm
            ImportForm wb, path, target
        Case ckDocument
            ImportDocumentCode wb, path, target
        End Select
        
NextFile:
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
    
    vbProj.VBComponents.Import filePath
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
    
    vbProj.VBComponents.Import filePath
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


