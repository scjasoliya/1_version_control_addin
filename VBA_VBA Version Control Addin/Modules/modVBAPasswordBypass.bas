Attribute VB_Name = "modVBAPasswordBypass"
Option Explicit

' ===============================================================
' Module : modVBAPasswordBypass
' Purpose: Session-only VBA project password bypass via API hook.
'
' CONFIRMED BEHAVIOUR (Excel 2019 x64, Build 16.0.19127.20302):
'   VBComponents.Count raises error 50289 without calling
'   DialogBoxParamA(4070) at all.  The dialog is ONLY raised when
'   the VBE Project Explorer tree node is expanded/clicked in the UI.
'
' Strategy:
'   1. Install the hook (patches DialogBoxParamA entry point).
'   2. Open VBE window safely (MainWindow.Visible = True).
'   3. Use Win32 TreeView messages to programmatically expand the
'      root project node — same action as a manual double-click.
'   4. Hook intercepts the dialog and returns IDOK silently.
'   5. Unhook IMMEDIATELY on success (Protection = 0).
'      Leaving the hook active after full unlock risks a crash.
'
'   All phases run via Application.OnTime so the VBA call stack
'   is empty when the hook callback fires (no re-entrancy crash).
'
'   64-bit Office : 12-byte MOV RAX, imm64 / JMP RAX
'   32-bit Office :  6-byte PUSH imm32   / RET
'
' NOTE: Application.VBE.ActiveVBProject is never set on a locked
'       project — doing so crashes Excel 2019 x64.
' ===============================================================

' --- CONSTANTS --------------------------------------------------
Private Const PAGE_EXECUTE_READWRITE As Long = &H40
Private Const VBA_PASSWORD_DIALOG_ID As Long = 4070

' TreeView messages
Private Const TVM_GETNEXTITEM As Long = &H110A
Private Const TVM_EXPAND      As Long = &H1102
Private Const TVM_SELECTITEM  As Long = &H110B
Private Const TVE_EXPAND      As Long = &H2
Private Const TVGN_ROOT       As Long = 0
Private Const TVGN_CARET      As Long = 9

' --- WIN32 API --------------------------------------------------
Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
                            (Destination As Any, Source As Any, ByVal Length As LongPtr)

Private Declare PtrSafe Function VirtualProtect Lib "kernel32" _
                                 (lpAddress As Any, ByVal dwSize As LongPtr, _
                                  ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

Private Declare PtrSafe Function GetModuleHandleA Lib "kernel32" _
                                 (ByVal lpModuleName As String) As LongPtr

Private Declare PtrSafe Function GetProcAddress Lib "kernel32" _
                                 (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr

Private Declare PtrSafe Function DialogBoxParam Lib "user32" Alias "DialogBoxParamA" _
                                 (ByVal hInstance As LongPtr, ByVal pTemplateName As LongPtr, _
                                  ByVal hWndParent As LongPtr, ByVal lpDialogFunc As LongPtr, _
                                  ByVal dwInitParam As LongPtr) As LongPtr

' FIX: lpWindowName declared As LongPtr so we can pass 0 (true NULL).
'      Passing vbNullString to a String param passes "" not NULL,
'      so FindWindowA would find nothing (VBE title is not "").
Private Declare PtrSafe Function FindWindowA Lib "user32" _
                                 (ByVal lpClassName As String, ByVal lpWindowName As LongPtr) As LongPtr

Private Declare PtrSafe Function FindWindowExA Lib "user32" _
                                 (ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, _
                                  ByVal lpszClass As String, ByVal lpszWindow As LongPtr) As LongPtr

Private Declare PtrSafe Function EnumChildWindows Lib "user32" _
                                 (ByVal hWndParent As LongPtr, ByVal lpEnumFunc As LongPtr, _
                                  ByVal lParam As LongPtr) As Long

Private Declare PtrSafe Function GetClassNameA Lib "user32" _
                                 (ByVal hWnd As LongPtr, ByVal lpClassName As String, _
                                  ByVal nMaxCount As Long) As Long

Private Declare PtrSafe Function SendMessageA Lib "user32" _
                                 (ByVal hWnd As LongPtr, ByVal Msg As Long, _
                                  ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

Private Declare PtrSafe Function PostMessageA Lib "user32" _
                                 (ByVal hWnd As LongPtr, ByVal Msg As Long, _
                                  ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long

' --- MODULE STATE ---------------------------------------------
#If Win64 Then
    Private HookBytes(0 To 11)   As Byte
    Private OriginBytes(0 To 11) As Byte
#Else
    Private HookBytes(0 To 5)   As Byte
    Private OriginBytes(0 To 5) As Byte
#End If
Private pFunc          As LongPtr
Private IsHooked       As Boolean
Private g_TargetProj   As Object
Private g_TreeViewHwnd As LongPtr                ' found by EnumChildWindows callback
Private g_Callback     As String                 ' optional macro to run after unlock

' --- HOOK IMPLEMENTATION --------------------------------------
Private Function GetPtr(ByVal Value As LongPtr) As LongPtr
    GetPtr = Value
End Function

Private Function MyDialogBoxParam( _
        ByVal hInstance As LongPtr, _
        ByVal pTemplateName As LongPtr, _
        ByVal hWndParent As LongPtr, _
        ByVal lpDialogFunc As LongPtr, _
        ByVal dwInitParam As LongPtr) As LongPtr
    Dim ts As String: ts = Format(Now, "hh:mm:ss")
    If pTemplateName = VBA_PASSWORD_DIALOG_ID Then
        MyDialogBoxParam = 1
    Else
        UnhookDialog
        MyDialogBoxParam = DialogBoxParam(hInstance, pTemplateName, _
                                          hWndParent, lpDialogFunc, dwInitParam)
        HookDialog
    End If
End Function

Private Function HookDialog() As Boolean
    Dim p As LongPtr
    Dim OldProtect As Long

    OldProtect = 0
    pFunc = GetProcAddress(GetModuleHandleA("user32.dll"), "DialogBoxParamA")
    If pFunc = 0 Then Exit Function
    #If Win64 Then
        Dim T(0 To 11) As Byte
        If VirtualProtect(ByVal pFunc, 12, PAGE_EXECUTE_READWRITE, OldProtect) = 0 Then
            Exit Function
        End If
        MoveMemory T(0), ByVal pFunc, 12
        p = GetPtr(AddressOf MyDialogBoxParam)
        If T(0) = &H48 And T(1) = &HB8 And T(10) = &HFF And T(11) = &HE0 Then
            ' TRUE stale hook: &H48 is just a REX.W prefix, but the full 12-byte
            ' signature confirms it's our MOV RAX / JMP RAX trampoline.
            MoveMemory OriginBytes(0), ByVal pFunc, 12
            MoveMemory HookBytes(2), p, 8
            MoveMemory ByVal pFunc + 2, p, 8
            IsHooked = True: HookDialog = True: Exit Function
        End If
        MoveMemory OriginBytes(0), ByVal pFunc, 12
        HookBytes(0) = &H48: HookBytes(1) = &HB8
        MoveMemory HookBytes(2), p, 8
        HookBytes(10) = &HFF: HookBytes(11) = &HE0
        MoveMemory ByVal pFunc, HookBytes(0), 12
    #Else
        Dim T(0 To 5) As Byte
        If VirtualProtect(ByVal pFunc, 6, PAGE_EXECUTE_READWRITE, OldProtect) = 0 Then
            Exit Function
        End If
        MoveMemory T(0), ByVal pFunc, 6
        p = GetPtr(AddressOf MyDialogBoxParam)
        If T(0) = &H68 And T(5) = &HC3 Then
            MoveMemory OriginBytes(0), ByVal pFunc, 6
            MoveMemory ByVal pFunc + 1, p, 4
            IsHooked = True: HookDialog = True: Exit Function
        End If
        MoveMemory OriginBytes(0), ByVal pFunc, 6
        HookBytes(0) = &H68
        MoveMemory HookBytes(1), p, 4
        HookBytes(5) = &HC3
        MoveMemory ByVal pFunc, HookBytes(0), 6
    #End If
    IsHooked = True: HookDialog = True
End Function

Private Sub UnhookDialog()
    If Not IsHooked Then Exit Sub
    ' Defensive: if VBA was reset while hooked, OriginBytes is lost (zeros).
    ' Restoring ZEROS to user32.dll will cause an immediate crash on next call.
    If OriginBytes(0) = 0 And OriginBytes(1) = 0 And OriginBytes(2) = 0 Then
        IsHooked = False
        Exit Sub
    End If
    #If Win64 Then
        MoveMemory ByVal pFunc, OriginBytes(0), 12
    #Else
        MoveMemory ByVal pFunc, OriginBytes(0), 6
    #End If
    IsHooked = False
End Sub

' --- WIN32 TREE VIEW TRIGGER ----------------------------------
' EnumChildWindows callback.
' Extracts the class name by trimming at the null terminator (critical:
' Left$(buf, 14) = "SysTreeView32" would always fail because "SysTreeView32"
' is 13 chars and the 14th char from GetClassNameA is Chr(0), not a space).
' Also records every class name visited so we can log the VBE hierarchy.
Public Function FindTreeViewCallback(ByVal hWnd As LongPtr, _
                                     ByVal lParam As LongPtr) As Long
    Dim buf      As String
    Dim nullPos  As Long
    Dim cls      As String

    buf = String$(256, 0)
    GetClassNameA hWnd, buf, 256
    nullPos = InStr(buf, Chr(0))
    cls = IIf(nullPos > 1, Left$(buf, nullPos - 1), vbNullString)

    If InStr(1, cls, "SysTreeView", vbTextCompare) > 0 Then
        g_TreeViewHwnd = hWnd
        FindTreeViewCallback = 0                 ' stop: found it
    Else
        FindTreeViewCallback = 1                 ' continue
    End If
End Function

' ExpandVBEProjectNode — opens VBE and sends TVM_EXPAND to the root
' project node, which triggers DialogBoxParamA(4070) on the locked project.
' Returns True if the TreeView was found and the expand message was sent.
Private Function ExpandVBEProjectNode() As Boolean
    Dim hVBE  As LongPtr
    Dim hRoot As LongPtr

    ' 1. Make VBE window visible (safe — not ActiveVBProject which crashes)
    Application.vbe.MainWindow.Visible = True

    ' 2. Find VBE main window.
    '    lpWindowName = 0 (NULL) matches any window with this class,
    '    regardless of title. Passing vbNullString to a String param
    '    would pass "" and find nothing because VBE title is not "".
    hVBE = FindWindowA("wndclass_desked_gsk", 0)
    If hVBE = 0 Then Exit Function

    ' 3. Search descendants of the VBE window.
    g_TreeViewHwnd = 0
    EnumChildWindows hVBE, GetPtr(AddressOf FindTreeViewCallback), 0

    If g_TreeViewHwnd = 0 Then
        ' Project Explorer may be undocked (floating top-level window).
        ' Fall back to searching the entire desktop for a SysTreeView32
        ' that is a top-level child of the desktop (hWndParent = 0).
        Dim hCandidate As LongPtr
        hCandidate = FindWindowExA(0, 0, "SysTreeView32", 0)
        Do While hCandidate <> 0
            ' Use it (first SysTreeView32 in the system when VBE is open
            ' will normally be the Project Explorer)
            g_TreeViewHwnd = hCandidate
            hCandidate = FindWindowExA(0, hCandidate, "SysTreeView32", 0) ' next
        Loop
    End If

    If g_TreeViewHwnd = 0 Then Exit Function

    ' 4. Get root tree item (first project node).
    hRoot = SendMessageA(g_TreeViewHwnd, TVM_GETNEXTITEM, TVGN_ROOT, 0)
    If hRoot = 0 Then Exit Function

    ' 5. Post TVM_SELECTITEM asynchronously.
    '    PostMessageA returns immediately; VBE processes the message only
    '    after our VBA code (TriggerExpand_) has returned and the call
    '    stack is empty.  The dialog then fires with VBA fully idle.
    PostMessageA g_TreeViewHwnd, TVM_SELECTITEM, TVGN_CARET, hRoot

    ' 6. Post TVM_EXPAND asynchronously for the same reason.
    PostMessageA g_TreeViewHwnd, TVM_EXPAND, TVE_EXPAND, hRoot

    ExpandVBEProjectNode = True
End Function

' --- PUBLIC API -----------------------------------------------

'@EntryPoint
'@Ignore ParameterNotUsed
Public Sub UnlockActiveVBAProject(Optional ByVal control As IRibbonControl, Optional ByVal CallbackOnSuccess As String = vbNullString)
    On Error GoTo EH
    g_Callback = CallbackOnSuccess

    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook.", vbExclamation, "No Workbook": Exit Sub
    End If

    On Error Resume Next
    Dim vbProj As Object
    Set vbProj = ActiveWorkbook.VBProject
    If Err.Number <> 0 Then
        MsgBox "Enable 'Trust access to the VBA project object model':" & vbCrLf & _
               "Excel > Options > Trust Center > Macro Settings", vbCritical, "Trust Required"
        Exit Sub
    End If
    On Error GoTo EH

    If vbProj.Protection = 0 Then
        MsgBox "VBA project is already unlocked.", vbInformation, "Already Unlocked"
        Exit Sub
    End If

    If IsHooked Then
        UnhookDialog
    End If

    If Not HookDialog() Then
        MsgBox "Could not install API hook.", vbCritical, "Hook Failed"
        Exit Sub
    End If

    Set g_TargetProj = vbProj

    ' Defer to idle context so VBA call stack is clear when hook fires.
    Application.OnTime Now, "modVBAPasswordBypass.TriggerExpand_"
    Exit Sub
EH:
    UnhookDialog
    Set g_TargetProj = Nothing
    g_Callback = vbNullString
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Bypass Error"
End Sub

' --------------------------------------------------------------
' INTERNAL CALLBACKS (Must be Public for OnTime but not for user)
' --------------------------------------------------------------

' TriggerExpand_ — called via OnTime (empty VBA call stack, no re-entrancy).
' Opens VBE and sends Win32 TreeView expand to fire DialogBoxParamA(4070).
Public Sub TriggerExpand_()
    If g_TargetProj Is Nothing Then Exit Sub

    Dim expanded As Boolean
    expanded = ExpandVBEProjectNode()

    ' Wait 2 seconds for VBE to dequeue and process the posted TreeView
    ' messages.  By then VBA is idle and the dialog hook fires cleanly.
    Application.OnTime Now + TimeSerial(0, 0, 2), "modVBAPasswordBypass.CheckResult_"
End Sub

' CheckResult_ — called via OnTime, checks Protection after the expand.
Public Sub CheckResult_()
    If g_TargetProj Is Nothing Then Exit Sub

    If g_TargetProj.Protection = 0 Then
        UnhookDialog                             ' CRITICAL: remove hook immediately on full unlock
        Dim projName As String
        projName = g_TargetProj.name
        Set g_TargetProj = Nothing
        MsgBox "VBA project '" & projName & "' unlocked for this session." & vbCrLf & vbCrLf & _
               "Note: memory-only patch - file is still protected on disk.", _
               vbInformation, "RAM Bypass Successful"
               
        ' Resume queued export/import if requested
        If Len(g_Callback) > 0 Then
            Dim cb As String: cb = g_Callback
            g_Callback = vbNullString
            Application.OnTime Now, cb
        End If
    Else
        g_Callback = vbNullString
        MsgBox "Hook is active - VBE is now open." & vbCrLf & vbCrLf & _
               "Click on '" & g_TargetProj.name & "' in the Project Explorer to unlock.", _
               vbExclamation, "One Manual Step Required"
    End If
End Sub

' ForceUnhook — abort and clean up (e.g. if unlock is no longer needed).
Public Sub ForceUnhook()
    UnhookDialog
    Set g_TargetProj = Nothing
    MsgBox "Hook removed.", vbInformation, "Cleanup"
End Sub


