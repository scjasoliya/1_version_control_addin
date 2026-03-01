Attribute VB_Name = "modVBAPasswordBypass"
'@Folder("Security")
' ===============================================================
' Module : modVBAPasswordBypass
' Purpose: Temporarily bypass VBA project password protection in
'          RAM by hooking the Win32 DialogBoxParamA function.
'
' Technique:
'   The Windows API function DialogBoxParamA is intercepted via
'   a trampoline patch. When the VBA IDE raises the password
'   dialog (template ID 4070), the callback returns 1 (OK)
'   without displaying the dialog.
'
'   On 64-bit Office: 12-byte MOV RAX, addr / JMP RAX trampoline.
'   On 32-bit Office:  6-byte PUSH addr / RET trampoline.
'
'   This is a SESSION-ONLY bypass. No files are modified.
'
' Attribution:
'   Based on code by kaybee99 / Siwtom (StackOverflow, CC BY-SA 3.0)
'   https://stackoverflow.com/a/31005696
'
' Requirements:
'   - Trust access to the VBA project object model enabled in
'     Excel > Trust Center > Macro Settings
' ===============================================================
Option Explicit

' ====== CONSTANTS ======
Private Const PAGE_EXECUTE_READWRITE As Long = &H40
Private Const VBA_PASSWORD_DIALOG_ID As Long = 4070

' ====== WIN32 API DECLARATIONS ======
' NOTE: MoveMemory destination/source are declared As Any so that
'       we can pass both raw-pointer values (ByVal) and variable
'       references (ByRef) without type-mismatch errors.
Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal Length As LongPtr)

Private Declare PtrSafe Function VirtualProtect Lib "kernel32" _
    (lpAddress As Any, ByVal dwSize As LongPtr, _
     ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

Private Declare PtrSafe Function GetModuleHandleA Lib "kernel32" _
    (ByVal lpModuleName As String) As LongPtr

Private Declare PtrSafe Function GetProcAddress Lib "kernel32" _
    (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr

Private Declare PtrSafe Function DialogBoxParam Lib "user32" _
    Alias "DialogBoxParamA" _
    (ByVal hInstance As LongPtr, ByVal pTemplateName As LongPtr, _
     ByVal hWndParent As LongPtr, ByVal lpDialogFunc As LongPtr, _
     ByVal dwInitParam As LongPtr) As Integer

' ====== MODULE-LEVEL STATE ======
#If Win64 Then
Private HookBytes(0 To 11)   As Byte   ' 12-byte: MOV RAX, imm64 / JMP RAX
Private OriginBytes(0 To 11) As Byte
#Else
Private HookBytes(0 To 5)   As Byte    ' 6-byte:  PUSH imm32 / RET
Private OriginBytes(0 To 5) As Byte
#End If
Private pFunc    As LongPtr
Private IsHooked As Boolean

' ====== PRIVATE HELPERS ======

' GetPtr — identity function used with AddressOf to get a function pointer.
Private Function GetPtr(ByVal Value As LongPtr) As LongPtr
    GetPtr = Value
End Function

' MyDialogBoxParam — intercepts DialogBoxParamA calls while hook is active.
'   Template 4070 = VBA password dialog → return 1 (OK) without showing.
'   All other dialogs → unhook, call real function, re-hook.
Private Function MyDialogBoxParam( _
        ByVal hInstance    As LongPtr, _
        ByVal pTemplateName As LongPtr, _
        ByVal hWndParent   As LongPtr, _
        ByVal lpDialogFunc As LongPtr, _
        ByVal dwInitParam  As LongPtr) As Integer

    If pTemplateName = VBA_PASSWORD_DIALOG_ID Then
        MyDialogBoxParam = 1            ' Simulate clicking OK on the password dialog
    Else
        UnhookVBAPasswordDialog
        MyDialogBoxParam = DialogBoxParam(hInstance, pTemplateName, _
                                          hWndParent, lpDialogFunc, dwInitParam)
        HookVBAPasswordDialog
    End If
End Function

' HookVBAPasswordDialog — patches DialogBoxParamA entry point with a trampoline.
' Returns True if the hook was installed successfully.
Private Function HookVBAPasswordDialog() As Boolean
    Dim p             As LongPtr
    Dim OldProtect    As Long
#If Win64 Then
    Dim TmpBytes(0 To 11) As Byte
#Else
    Dim TmpBytes(0 To 5) As Byte
#End If

    HookVBAPasswordDialog = False

    ' Locate DialogBoxParamA
    pFunc = GetProcAddress(GetModuleHandleA("user32.dll"), "DialogBoxParamA")
    If pFunc = 0 Then Exit Function

    ' Make the memory page writable
#If Win64 Then
    If VirtualProtect(ByVal pFunc, 12, PAGE_EXECUTE_READWRITE, OldProtect) = 0 Then Exit Function

    ' Read current bytes; bail if already hooked (0x48 = REX.W prefix of MOV RAX)
    MoveMemory TmpBytes(0), ByVal pFunc, 12
    If TmpBytes(0) = &H48 Then
        HookVBAPasswordDialog = True    ' Already hooked — treat as success
        IsHooked = True
        Exit Function
    End If

    ' Save original 12 bytes
    MoveMemory OriginBytes(0), ByVal pFunc, 12

    ' Build: MOV RAX, imm64 (48 B8 <8-byte addr>) ; JMP RAX (FF E0)
    p = GetPtr(AddressOf MyDialogBoxParam)
    HookBytes(0) = &H48: HookBytes(1) = &HB8       ' MOV RAX, imm64
    MoveMemory HookBytes(2), p, 8                   ' 8-byte address of callback
    HookBytes(10) = &HFF: HookBytes(11) = &HE0      ' JMP RAX

    ' Write the 12-byte trampoline
    MoveMemory ByVal pFunc, HookBytes(0), 12
#Else
    If VirtualProtect(ByVal pFunc, 6, PAGE_EXECUTE_READWRITE, OldProtect) = 0 Then Exit Function

    ' Read current bytes; bail if already hooked (0x68 = PUSH)
    MoveMemory TmpBytes(0), ByVal pFunc, 6
    If TmpBytes(0) = &H68 Then
        HookVBAPasswordDialog = True
        IsHooked = True
        Exit Function
    End If

    ' Save original 6 bytes
    MoveMemory OriginBytes(0), ByVal pFunc, 6

    ' Build: PUSH imm32 (68 <4-byte addr>) ; RET (C3)
    p = GetPtr(AddressOf MyDialogBoxParam)
    HookBytes(0) = &H68                             ' PUSH imm32
    MoveMemory HookBytes(1), p, 4                   ' 4-byte address of callback
    HookBytes(5) = &HC3                              ' RET

    ' Write the 6-byte trampoline
    MoveMemory ByVal pFunc, HookBytes(0), 6
#End If

    IsHooked = True
    HookVBAPasswordDialog = True
End Function

' UnhookVBAPasswordDialog — restores original DialogBoxParamA bytes.
Private Sub UnhookVBAPasswordDialog()
    If Not IsHooked Then Exit Sub
#If Win64 Then
    MoveMemory ByVal pFunc, OriginBytes(0), 12
#Else
    MoveMemory ByVal pFunc, OriginBytes(0), 6
#End If
    IsHooked = False
End Sub

' ====== PUBLIC API ======

' BypassVBAProjectPassword
' Installs the hook then navigates to the VBE project node, which forces
' Office to call DialogBoxParamA(…, 4070, …). Our callback intercepts that
' call and returns 1 (OK), unlocking the project for this session.
'
' Returns True if the project's Protection property is 0 (unlocked) after
' the attempt.
Public Function BypassVBAProjectPassword(ByVal vbProj As Object) As Boolean
    On Error GoTo SafeExit

    ' Nothing to do if already unlocked
    If vbProj.Protection = 0 Then
        BypassVBAProjectPassword = True
        Exit Function
    End If

    ' Install the hook
    If Not HookVBAPasswordDialog() Then
        BypassVBAProjectPassword = False
        Exit Function
    End If

    ' Force the VBE to demand the password by:
    '   1. Making the VBE window visible (required for dialog trigger)
    '   2. Setting the active VBProject (triggers the password prompt)
    On Error Resume Next

    Dim vbe As Object
    Set vbe = Application.VBE

    ' Method 1: Access VBComponents directly — triggers dialog on most versions
    Dim dummy As Long
    dummy = vbProj.VBComponents.Count

    ' Method 2: If still locked, navigate to the project in the VBE tree.
    '           This is the most reliable trigger for the dialog on modern Excel.
    If vbProj.Protection <> 0 Then
        vbe.ActiveVBProject = vbProj
    End If

    On Error GoTo SafeExit

    ' Verify result
    BypassVBAProjectPassword = (vbProj.Protection = 0)

SafeExit:
    UnhookVBAPasswordDialog
End Function

' ====== RIBBON ENTRY POINT ======

'@EntryPoint
'@Ignore ParameterNotUsed
Public Sub UnlockActiveVBAProject(ByVal control As IRibbonControl)
    On Error GoTo EH

    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook.", vbExclamation, "No Workbook"
        Exit Sub
    End If

    ' Ensure trust access to VBA object model is on
    On Error Resume Next
    Dim testRef As Object
    Set testRef = ActiveWorkbook.VBProject
    If Err.Number <> 0 Then
        MsgBox "Please enable 'Trust access to the VBA project object model' in:" & vbCrLf & _
               "Excel > File > Options > Trust Center > Trust Center Settings > Macro Settings", _
               vbCritical, "Trust Access Required"
        Exit Sub
    End If
    On Error GoTo EH

    If ActiveWorkbook.VBProject.Protection = 0 Then
        MsgBox "The VBA project is already unlocked.", vbInformation, "Already Unlocked"
        Exit Sub
    End If

    If BypassVBAProjectPassword(ActiveWorkbook.VBProject) Then
        MsgBox "VBA project unlocked for this session." & vbCrLf & vbCrLf & _
               "Note: This is a memory patch. The file is still protected on disk.", _
               vbInformation, "RAM Bypass Successful"
    Else
        MsgBox "Failed to unlock the VBA project." & vbCrLf & vbCrLf & _
               "Make sure:" & vbCrLf & _
               "  1. 'Trust access to VBA project object model' is enabled." & vbCrLf & _
               "  2. The VBE is not already open with the project locked.", _
               vbCritical, "Bypass Failed"
    End If
    Exit Sub
EH:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Bypass Error"
End Sub
