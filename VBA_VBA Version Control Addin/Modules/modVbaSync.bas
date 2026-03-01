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

' ====== UNLOCK VIA RAM BYPASS ======

' UnlockVBProject
' Unlocks a password-protected VBProject using the RAM-level
' DialogBoxParamA hook from modVBAPasswordBypass.
' Returns True if the project is accessible after the attempt.
Private Function UnlockVBProject(ByVal wb As Workbook) As Boolean
    On Error GoTo SafeExit
    
    If wb Is Nothing Then Exit Function
    
    ' If the project is not protected, nothing to do
    If wb.VBProject.Protection = 0 Then
        UnlockVBProject = True
        Exit Function
    End If
    
    ' Attempt RAM-level bypass via modVBAPasswordBypass
    LogI "Attempting RAM bypass for: " & wb.Name
    If modVBAPasswordBypass.BypassVBAProjectPassword(wb.VBProject) Then
        LogI "VBA project unlocked via RAM bypass: " & wb.Name
        UnlockVBProject = True
        Exit Function
    End If
    
    ' Bypass failed — inform the user
    LogW "RAM bypass failed for: " & wb.Name
    MsgBox "VBA project is locked and the RAM bypass could not unlock it." & vbCrLf & vbCrLf & _
           "Workbook: " & wb.FullName & vbCrLf & vbCrLf & _
           "You may need to manually unlock the project via:" & vbCrLf & _
           "Alt+F11 ? select the project ? enter password.", vbExclamation, "Unlock Failed"
    
SafeExit:
    UnlockVBProject = (wb.VBProject.Protection = 0)
    If Not UnlockVBProject Then
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
' ====== NORMALIZATION (Document code) ======
' Strips VERSION/Attribute lines that cannot be injected into document modules
Private Function NormalizeDocumentCode(ByVal rawCode As String) As String
    If Len(rawCode) = 0 Then NormalizeDocumentCode = vbNullString: Exit Function
    Dim lines() As String
    Dim outBuf As String
    Dim i As Long
    Dim s As String
    Dim t As String
    Dim inHeaderBlock As Boolean
    
    lines = Split(rawCode, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        s = Replace$(lines(i), vbCr, vbNullString)
        t = LCase$(Trim$(s))
        
        If t = "begin" Then
            inHeaderBlock = True
        ElseIf t = "end" Then
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
        ElseIf Left$(t, 10) = "attribute " Then
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
    If Not UnlockVBProject(ActiveWorkbook) Then Exit Sub
    
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
    If Not UnlockVBProject(ActiveWorkbook) Then Exit Sub
    
    Dim root As String: root = GetProjectRootFolder(ActiveWorkbook)
    If Dir(root, vbDirectory) = vbNullString Then
        MsgBox "Source folder not found: " & root, vbCritical: Exit Sub
    End If
    
    LogI "Import ? " & ActiveWorkbook.FullName & " ? " & root
    
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
    
    If newComp.Name <> compName Then
        LogI "Renaming imported component " & newComp.Name & " to " & compName
        On Error Resume Next
        newComp.Name = compName
        If Err.Number <> 0 Then
            LogE "Failed to rename component " & newComp.Name & " to " & compName & ": " & Err.Description
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
    
    If newComp.Name <> formName Then
        LogI "Renaming imported form " & newComp.Name & " to " & formName
        On Error Resume Next
        newComp.Name = formName
         If Err.Number <> 0 Then
            LogE "Failed to rename form " & newComp.Name & " to " & formName & ": " & Err.Description
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
    expected.CompareMode = 1 ' TextCompare

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
            If Not expected.Exists(vbc.Name) Then
                extras.Add vbc.Name, vbc
            End If
        End If
    Next vbc
    
    If extras.Count = 0 Then Exit Sub
    
    ' 3. Prompt User
    Dim msg As String
    msg = "The following components exist in the workbook but NOT in the source folder:" & vbCrLf & vbCrLf
    
    Dim k As Variant
    Dim count As Long: count = 0
    For Each k In extras.Keys
        msg = msg & " - " & k & vbCrLf
        count = count + 1
        If count > 15 Then
            msg = msg & "... and " & (extras.Count - 15) & " more."
            Exit For
        End If
    Next k
    
    msg = msg & vbCrLf & "Do you want to DELETE these extra components from the workbook?"
    
    If MsgBox(msg, vbQuestion + vbYesNo, "Remove Extra Components?") = vbYes Then
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


