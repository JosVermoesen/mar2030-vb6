VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Users"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Enum LogUserType
    luIn = 0
    luOut = 1
End Enum

Private mAppTitle As String
Private mConnectString As String
'
Private blnUsersLoaded As Boolean
Private cnnUsers As New ADODB.Connection
Private rstUsers As ADODB.Recordset
Private rstRights As ADODB.Recordset
Private varUserRights As Variant

Public Sub Maintenance()
    '
    On Error GoTo LocalErr
    '
    With frmUsers
        .ConnectString = mConnectString
        .Show vbModal
    End With
    '
    Exit Sub
    '
LocalErr:
    ReportError Err.Number, "Security.Users.Maintenance." & Err.Source, Err.Description
    '
End Sub

Public Function Login(Optional UserID As String, Optional Password As String, Optional NationalNumber As String) As Boolean
    '
    On Error GoTo LocalErr
    '
    Static Attempts As Integer
    Dim intAnswer As Integer
    Dim blnValidUser As Boolean
    Dim strUserID As String
    Dim strPassword As String
    Dim strNationalNumber As String
        '
    If UserID <> "" Then
        strUserID = UserID
    End If
    '
    If Password <> "" Then
        strPassword = Password
    End If
    
    
    '
    If strUserID <> "" And strPassword <> "" Then
        GoTo CheckLogin
    End If
    '
AttemptLogin:
    With frmLogin
        .Title = mAppTitle
        .UserID = UserID
        .Password = Password
        .NationalNumber = NationalNumber
        .Show vbModal
        '
        If .Status = True Then
            strUserID = .txtUserID
            strPassword = .txtPassword
        End If
    End With
    Unload frmLogin
    '
CheckLogin:
    Attempts = Attempts + 1
    blnValidUser = CheckUser(strUserID, strPassword)
    '
    If blnValidUser = True Then
        LogUser strUserID, luIn
        GetRights strUserID
        modShared.UserID = strUserID
        Login = True
        Exit Function
    End If
    '
    If Attempts = 3 Then
        MsgBox "Gebruikers aanmelding zonder succes - geen toegang", vbCritical, "Gebruiker Aanmelding"
        Login = False
        Exit Function
    End If
    '
    intAnswer = MsgBox("Ongeldige aanmelding", vbRetryCancel + vbInformation, "Gebruiker Aanmelding")
    If intAnswer = vbRetry Then
        GoTo AttemptLogin
    Else
        Login = False
        Exit Function
    End If
    '
    Exit Function
    '
LocalErr:
    ReportError Err.Number, "Security.Users.Login." & Err.Source, Err.Description
    '
End Function

Public Property Get AppTitle() As String
    AppTitle = mAppTitle
End Property

Public Property Let AppTitle(ByVal vNewValue As String)
    mAppTitle = vNewValue
End Property

Private Function CheckUser(UserID As String, Password As String)
    '
    On Error GoTo LocalErr
    '
    Dim strFind As String
    '
    strFind = "FIND='" & UserID & "." & Password & "'"
    '
    If blnUsersLoaded = False Then
        LoadUsers
    End If
    '
    With rstUsers
        .MoveFirst
        .Find strFind
        If .EOF = False Then
            CheckUser = True
        Else
            CheckUser = False
        End If
    End With
    '
    Exit Function
    '
LocalErr:
    ReportError Err.Number, "Security.Users.CheckUser." & Err.Source, Err.Description
    '
End Function

Private Sub LoadUsers()
    '
    On Error GoTo LocalErr
    '
    Dim strSQL As String
    '
    strSQL = "SELECT *,UserID &'.'& Password AS Find FROM secUsers ORDER BY UserID"
    '
    cnnUsers.Open mConnectString
    cnnUsers.CursorLocation = adUseClient
    Set rstUsers = cnnUsers.Execute(strSQL, , adCmdText)
    blnUsersLoaded = True
    '
    Exit Sub
    '
LocalErr:
    ReportError Err.Number, "Security.Users.LoadUsers." & Err.Source, Err.Description
    '
End Sub

Public Property Get ConnectString() As String
    ConnectString = mConnectString
End Property

Public Property Let ConnectString(ByVal vNewValue As String)
    mConnectString = vNewValue
End Property

Private Sub LogUser(UserID As String, Action As LogUserType)
    '
    On Error GoTo LocalErr
    '
    Dim strFind As String
    '
    strFind = "UserID='" & UserID & "'"
    '
    If blnUsersLoaded = False Then
        LoadUsers
    End If
    '
    With rstUsers
        .MoveFirst
        .Find strFind
        If .EOF = False Then
            If Action = luIn Then
                cnnUsers.Execute "UPDATE secUsers SET LastLogIn=#" & Now() & "#"
            Else
                cnnUsers.Execute "UPDATE secUsers SET LastLogOut=#" & Now() & "#"
            End If
        End If
    End With
    '
    Exit Sub
    '
LocalErr:
    ReportError Err.Number, "Security.Users.LogUser." & Err.Source, Err.Description
    '
End Sub

Private Sub GetRights(UserID As String)
    '
    On Error GoTo LocalErr
    '
    Dim strSQL As String
    Dim intCount As Integer
    '
    strSQL = "SELECT * FROM secAccess WHERE UserID='" & UserID & "'"
    Set rstRights = cnnUsers.Execute(strSQL, , adCmdText)
    '
    Exit Sub
    '
LocalErr:
    ReportError Err.Number, "Security.Users.GetgRights." & Err.Source, Err.Description
    '
End Sub

Public Function CheckRights(Object As String, Level As Integer) As Boolean
    '
    On Error GoTo LocalErr
    '
    Dim strFind As String
    '
    strFind = "Object='" & Object & "'"
    '
    With rstRights
        .MoveFirst
        .Find strFind
        If .EOF = True Then
            CheckRights = False
        Else
            CheckRights = (Level <= .Fields("Level"))
            'If rstRights(0) = "DEFAULT" Then
            '    CheckRights = True
            'Else
            '    CheckRights = False
            'End If
        End If
    End With
    '
    Exit Function
    '
LocalErr:
    ReportError Err.Number, "Security.Users.CheckRights." & Err.Source, Err.Description
    '
End Function

Private Sub Class_Initialize()
    '
    mAppTitle = "User Login"
    '
End Sub

Public Sub LogOut()
    '
    On Error GoTo LocalErr
    '
    LogUser modShared.UserID, luOut
    '
    Exit Sub
    '
LocalErr:
    ReportError Err.Number, "Security.Users.LogOut." & Err.Source, Err.Description
    '
End Sub

