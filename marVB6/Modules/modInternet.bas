Attribute VB_Name = "modInternet"
Option Explicit

' ============================================================
'   Internet Connectivity Check Module (VB6)
'   - Primary: WinHttpRequest (TLS 1.2)
'   - Fallback: URLDownloadToFile
'   - Integration: ShellExecute fallback module
'   - Author: Copilot & Jos
' ============================================================

' --- API DECLARATIONS ---------------------------------------

Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
    (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, _
     ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" _
    (ByVal lpFileName As String) As Long

' --- CONFIG --------------------------------------------------

Private Const TEST_URL_1 As String = "http://www.msftconnecttest.com/connecttest.txt"
Private Const TEST_URL_2 As String = "http://www.gstatic.com/generate_204"
Private Const TIMEOUT_MS As Long = 4000

' ============================================================
'   PUBLIC ENTRY POINT
' ============================================================

Public Function Internet_IsAvailable() As Boolean
    ' Try WinHTTP first
    If Check_WinHttp(TEST_URL_1) Then
        Internet_IsAvailable = True
        Exit Function
    End If

    If Check_WinHttp(TEST_URL_2) Then
        Internet_IsAvailable = True
        Exit Function
    End If

    ' Fallback: URLDownloadToFile
    If Check_URLDownload(TEST_URL_1) Then
        Internet_IsAvailable = True
        Exit Function
    End If

    If Check_URLDownload(TEST_URL_2) Then
        Internet_IsAvailable = True
        Exit Function
    End If

    ' Final fallback: open browser via your ShellExecute module
    ' This allows the user to visually confirm connectivity
    Call ShellExecute_Fallback("http://www.msftconnecttest.com/redirect")

    Internet_IsAvailable = False
End Function

' ============================================================
'   METHOD 1: WinHTTP (Preferred)
' ============================================================

Private Function Check_WinHttp(ByVal sURL As String) As Boolean
    On Error GoTo Fail

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.setTimeouts TIMEOUT_MS, TIMEOUT_MS, TIMEOUT_MS, TIMEOUT_MS
    http.Open "GET", sURL, False
    http.Send

    If http.Status = 200 Or http.Status = 204 Then
        Check_WinHttp = True
    End If

    Exit Function

Fail:
    Check_WinHttp = False
End Function

' ============================================================
'   METHOD 2: URLDownloadToFile (Fallback)
' ============================================================

Private Function Check_URLDownload(ByVal sURL As String) As Boolean
    On Error GoTo Fail

    Dim tmp As String
    tmp = Environ$("TEMP") & "\netcheck.tmp"

    Dim ret As Long
    ret = URLDownloadToFile(0, sURL, tmp, 0, 0)

    If ret = 0 Then
        Check_URLDownload = True
    End If

    Call DeleteFile(tmp)
    Exit Function

Fail:
    Check_URLDownload = False
End Function

' ============================================================
'   SHELLEXECUTE FALLBACK INTEGRATION
' ============================================================
' This calls your existing ShellExecute diagnostic module.
' Replace the call below with your exact function name if needed.
' ============================================================

Public Sub ShellExecute_Fallback(ByVal sURL As String)
    On Error Resume Next
    ' Your existing module already handles:
    ' - browser selection
    ' - error logging
    ' - fallback strategies
    'Call OpenURL_Fallback(sURL)   ' <--- integrate here
End Sub


