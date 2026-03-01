Attribute VB_Name = "Module1"
' Module: ShellHelper.bas
Option Explicit

' API declaraties
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    ' for compatibility (unused) reserve fields
    lpIDList As Long
    lpClass As Long
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Private Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExA" ( _
    lpExecInfo As SHELLEXECUTEINFO) As Long

' Constants
Private Const SW_SHOWNORMAL As Long = 1
Private Const SEE_MASK_NOCLOSEPROCESS As Long = &H40

' Configureer hier het logbestandpad indien gewenst
Public g_ShellHelperLogPath As String

Private Function GetDefaultLogPath() As String
    If g_ShellHelperLogPath <> vbNullString Then
        GetDefaultLogPath = g_ShellHelperLogPath
    Else
        On Error Resume Next
        GetDefaultLogPath = LOCATION_MYDOCUMENTS 'App.path
        If Right$(GetDefaultLogPath, 1) <> "\" Then GetDefaultLogPath = GetDefaultLogPath & "\"
        GetDefaultLogPath = GetDefaultLogPath & "ShellHelper.log"
        If Err.Number <> 0 Then
            GetDefaultLogPath = "C:\ShellHelper.log"
        End If
        On Error GoTo 0
    End If
End Function

Private Sub SHLog(ByVal s As String)
    Dim sPath As String
    sPath = GetDefaultLogPath()
    On Error Resume Next
    Dim fnum As Integer
    fnum = FreeFile
    Open sPath For Append As #fnum
    Print #fnum, Now; " - " & s
    Close #fnum
    On Error GoTo 0
End Sub

Public Function ShellExecuteWithFallback(ByVal sTarget As String, Optional ByVal sParams As String = vbNullString, Optional ByVal sVerb As String = "open") As Boolean
    ' Probeer ShellExecute, bij falen probeer ShellExecuteEx, daarna COMSPEC / start fallback
    Dim ret As Long
    Dim op As String, file As String, params As String
    op = sVerb
    file = sTarget
    params = sParams

    SHLog "Attempting ShellExecute. CurDir=" & CurDir$ & " COMSPEC=" & Environ$("COMSPEC") & " PATH=" & Environ$("PATH")
    SHLog "ShellExecute parameters: verb=" & op & " file=" & file & " params=" & params

    On Error Resume Next
    ret = ShellExecute(0, op, file, params, vbNullString, SW_SHOWNORMAL)
    On Error GoTo 0

    SHLog "ShellExecute returned: " & CStr(ret) & " (" & ShellExecuteErrorText(ret) & ")"

    If ret > 32 Then
        ShellExecuteWithFallback = True
        Exit Function
    End If

    ' Eerste fallback: ShellExecuteEx with SEE_MASK_NOCLOSEPROCESS
    If TryShellExecuteEx(op, file, params) Then
        ShellExecuteWithFallback = True
        Exit Function
    End If

    ' Laatste fallback: gebruik COMSPEC /C start "" "target" (start gebruikt bestandstype-associatie)
    If TryCmdStart(file, params) Then
        ShellExecuteWithFallback = True
        Exit Function
    End If

    ShellExecuteWithFallback = False
End Function

Private Function TryShellExecuteEx(ByVal sVerb As String, ByVal sFile As String, ByVal sParams As String) As Boolean
    Dim sei As SHELLEXECUTEINFO
    Dim res As Long

    sei.cbSize = Len(sei)
    sei.fMask = SEE_MASK_NOCLOSEPROCESS
    sei.hwnd = 0
    sei.lpVerb = sVerb & vbNullChar
    sei.lpFile = sFile & vbNullChar
    sei.lpParameters = sParams & vbNullChar
    sei.lpDirectory = vbNullChar
    sei.nShow = SW_SHOWNORMAL

    On Error Resume Next
    res = ShellExecuteEx(sei)
    On Error GoTo 0

    SHLog "ShellExecuteEx returned: " & CStr(res) & " (hProcess=" & CStr(sei.hProcess) & ")"

    If res <> 0 Then
        TryShellExecuteEx = True
    Else
        TryShellExecuteEx = False
    End If
End Function

Private Function TryCmdStart(ByVal sFile As String, ByVal sParams As String) As Boolean
    Dim sComspec As String
    Dim sCmd As String
    Dim retShell As Long

    sComspec = Environ$("COMSPEC")
    If Len(Trim$(sComspec)) = 0 Then
        SHLog "COMSPEC empty. Cannot use cmd fallback."
        TryCmdStart = False
        Exit Function
    End If

    ' Gebruik START om file te openen via geassocieerde handler
    ' Format: cmd /C start "" "file" params
    sCmd = """" & sComspec & """ /C start """" """ & sFile & """ " & sParams

    SHLog "Trying COMSPEC fallback. Command: " & sCmd

    On Error Resume Next
    retShell = Shell(sCmd, vbNormalFocus)
    If Err.Number <> 0 Then
        SHLog "Shell(cmd) error: " & CStr(Err.Number) & " - " & Err.Description
        Err.Clear
        TryCmdStart = False
    Else
        ' Shell returns process id > 0 on success; we still cannot reliably detect failure of start
        SHLog "Shell(cmd) returned pid: " & CStr(retShell)
        TryCmdStart = True
    End If
    On Error GoTo 0
End Function

Public Function ShellExecuteErrorText(ByVal code As Long) As String
    Select Case code
        Case 0: ShellExecuteErrorText = "SE_ERR_Fail"
        Case 2: ShellExecuteErrorText = "SE_ERR_FileNotFound"
        Case 3: ShellExecuteErrorText = "SE_ERR_PathNotFound"
        Case 5: ShellExecuteErrorText = "SE_ERR_OOM or ACCESS_DENIED"
        Case 8: ShellExecuteErrorText = "SE_ERR_OOM"
        Case 26: ShellExecuteErrorText = "SE_ERR_DLLNOTFOUND"
        Case 27: ShellExecuteErrorText = "SE_ERR_NOASSOC or ASSOC_INCOMPLETE"
        Case 28: ShellExecuteErrorText = "SE_ERR_DDETIMEOUT"
        Case 29: ShellExecuteErrorText = "SE_ERR_DDEFAIL"
        Case 30: ShellExecuteErrorText = "SE_ERR_DDEBUSY"
        Case 31: ShellExecuteErrorText = "SE_ERR_NOUI"
        Case 32: ShellExecuteErrorText = "SE_ERR_DLLNOTFOUND"
        Case Else
            If code < 0 Then
                ShellExecuteErrorText = "Negative return (unknown)"
            ElseIf code > 32 Then
                ShellExecuteErrorText = "Success"
            Else
                ShellExecuteErrorText = "Unknown code"
            End If
    End Select
End Function

' Kleine diagnostische helper die je kunt aanroepen om systeemvariabelen en assoc te loggen
Public Sub ShellHelper_LogDiagnostics(Optional ByVal sExtension As String = ".pdf")
    On Error Resume Next
    SHLog "---- Diagnostics start ----"
    SHLog "CurDir: " & CurDir$
    SHLog "App.Path: " & App.path
    SHLog "User: " & Environ$("USERNAME") & " COMPUTERNAME: " & Environ$("COMPUTERNAME")
    SHLog "COMSPEC: " & Environ$("COMSPEC")
    SHLog "PATH: " & Environ$("PATH")
    SHLog "PATHEXT: " & Environ$("PATHEXT")
    ' probe assoc/ftype via cmd if available
    Dim sCom As String, pid As Long
    sCom = Environ$("COMSPEC")
    If Len(Trim$(sCom)) > 0 Then
        ' assoc .ext
        pid = Shell("""" & sCom & """ /C assoc " & sExtension, vbHide)
        SHLog "Launched assoc probe for " & sExtension & " (pid=" & pid & ")"
        ' ftype of assoc result can be probed similarly if needed by running 'assoc' then 'ftype' in a temp batch
    Else
        SHLog "COMSPEC not present, assoc probe skipped"
    End If
    SHLog "---- Diagnostics end ----"
    On Error GoTo 0
End Sub





