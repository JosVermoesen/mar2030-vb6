VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Audit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mFileName As String
Private mFileCh As Integer
Private mFileHdr As String
Private mPWiServices As String

Public Property Get FileName() As String
    FileName = mFileName
End Property

Public Property Let FileName(ByVal vNewValue As String)
    mFileName = vNewValue
End Property

Public Property Get FileHeader() As String
    FileHeader = mFileHdr
End Property

Public Property Let FileHeader(ByVal vNewValue As String)
    mFileHdr = vNewValue
End Property

Public Property Get PWiServices() As String
    PWiServices = mPWiServices
End Property

Public Property Let PWiServices(ByVal vNewValue As String)
    mPWiServices = vNewValue
End Property

Private Sub Class_Initialize()
    '
    mFileName = App.EXEName & ".log"
    mFileCh = -1
    mFileHdr = "User Audit Log"
    mPWiServices = ""
    '
End Sub

Private Sub Class_Terminate()
    '
    On Error Resume Next
    '
    Close #mFileCh
    '
End Sub

Public Sub WriteLog(ParamArray varParams() As Variant)
    '
    On Error GoTo LocalErr
    '
    Dim intItems As Integer
    Dim intLoop As Integer
    Dim strLine As String
    Dim vbQuote As String
    Dim strUser As String
    Dim strDate As String
    '
    vbQuote = Chr(34)
    '
    If mFileCh = -1 Then
        OpenLogFile
    End If
    '
    If modShared.UserID = "" Then
        strUser = "NO_USER"
    Else
        strUser = modShared.UserID
    End If
    '
    strLine = vbQuote & Format(Now(), "general date") & vbQuote & ","
    strLine = strLine & vbQuote & strUser & vbQuote & ","
    '
    intItems = UBound(varParams)
    For intLoop = 0 To intItems
        strLine = strLine & vbQuote & varParams(intLoop) & vbQuote & ","
    Next
    strLine = Left(strLine, Len(strLine) - 1)
    '
    Print #mFileCh, strLine
    '
    Exit Sub
    '
LocalErr:
    ReportError Err.Number, "Security.Audit.WriteLog." & Err.Source, Err.Description
    '
End Sub

Private Sub OpenLogFile()
    '
    On Error Resume Next
    '
    mFileCh = FreeFile
    '
    ' make sure file exists, if not make it
    Open mFileName For Input As mFileCh
    If Err.Number <> 0 Then
        Open mFileName For Output As mFileCh
        WriteLogHeader
    End If
    Close #mFileCh
    '
    On Error GoTo LocalErr
    '
    ' open for additional logging
    Open mFileName For Append As mFileCh
    '
    Exit Sub
    '
LocalErr:
    ReportError Err.Number, "Security.Audit.OpenLogFile." & Err.Source, Err.Description
    '
End Sub

Private Sub WriteLogHeader()
    '
    On Error GoTo LocalErr
    '
    Print #mFileCh, String(60, "*")
    Print #mFileCh, mFileHdr
    Print #mFileCh, "Aangemaakt: " & Format(Now, "general date")
    Print #mFileCh, String(60, "*")
    Print #mFileCh, ""
    '
    Exit Sub
    '
LocalErr:
    ReportError Err.Number, "Security.Audit.WriteLogHeader." & Err.Source, Err.Description
    '
End Sub

