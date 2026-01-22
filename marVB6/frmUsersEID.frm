VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{3D588416-8F31-4516-9E89-12B4D80C9CFC}#1.0#0"; "eIDInteropUserControlLibrary1.tlb"
Begin VB.Form frmUsersEID 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gebruikers Beveiliging"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10785
   Icon            =   "frmUsersEID.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin eIDInteropUserControlLibrary1Ctl.InteropUserControl InteropUserControl1 
      Height          =   2895
      Left            =   5880
      TabIndex        =   20
      Top             =   0
      Width           =   4815
      Object.Visible         =   "True"
      Enabled         =   "True"
      ForegroundColor =   "-2147483630"
      BackgroundColor =   "-2147483633"
      BackColor       =   "Control"
      ForeColor       =   "ControlText"
      Location        =   "392, 0"
      Name            =   "InteropUserControl"
      Size            =   "321, 193"
      Object.TabIndex        =   "0"
   End
   Begin VB.CommandButton cbEID 
      Caption         =   "eID Data Invoegen"
      Height          =   615
      Left            =   3240
      TabIndex        =   19
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtUser 
      DataField       =   "NationalNumber"
      DataSource      =   "adcUsers"
      Height          =   300
      Index           =   5
      Left            =   1320
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   1200
      Width           =   3000
   End
   Begin VB.CommandButton cmdAction 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   300
      Index           =   6
      Left            =   4560
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Rechten"
      Height          =   300
      Index           =   5
      Left            =   4560
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Verwijderen"
      Height          =   300
      Index           =   4
      Left            =   4560
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1260
      Width           =   1215
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Bij&werken"
      Height          =   300
      Index           =   3
      Left            =   4560
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Nieuw"
      Height          =   300
      Index           =   2
      Left            =   4560
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   660
      Width           =   1215
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Annuleren"
      Height          =   300
      Index           =   1
      Left            =   4560
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Bewaren"
      Height          =   300
      Index           =   0
      Left            =   4560
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   60
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc adcUsers 
      Height          =   330
      Left            =   90
      Top             =   2490
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=RVDSN"
      OLEDBString     =   "DSN=RVDSN"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "secUsers"
      Caption         =   "Users"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtUser 
      DataField       =   "LastLogOut"
      DataSource      =   "adcUsers"
      Enabled         =   0   'False
      Height          =   300
      Index           =   4
      Left            =   2130
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2040
      Width           =   1800
   End
   Begin VB.TextBox txtUser 
      DataField       =   "LastLogIn"
      DataSource      =   "adcUsers"
      Enabled         =   0   'False
      Height          =   300
      Index           =   3
      Left            =   210
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2040
      Width           =   1800
   End
   Begin VB.TextBox txtUser 
      DataField       =   "Name"
      DataSource      =   "adcUsers"
      Height          =   300
      Index           =   2
      Left            =   1320
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   780
      Width           =   3000
   End
   Begin VB.TextBox txtUser 
      DataField       =   "Password"
      DataSource      =   "adcUsers"
      Height          =   300
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   420
      Width           =   1800
   End
   Begin VB.TextBox txtUser 
      DataField       =   "UserID"
      DataSource      =   "adcUsers"
      Height          =   300
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   60
      Width           =   1800
   End
   Begin VB.Label lblUser 
      Caption         =   "Nationaal Nr."
      Height          =   300
      Index           =   5
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   1080
   End
   Begin VB.Label lblUser 
      Caption         =   "Laatste Afmelding"
      Height          =   420
      Index           =   4
      Left            =   2160
      TabIndex        =   9
      Top             =   1620
      Width           =   1200
   End
   Begin VB.Label lblUser 
      Caption         =   "Laatste Aanmelding"
      Height          =   420
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   1620
      Width           =   1200
   End
   Begin VB.Label lblUser 
      Caption         =   "Naam Detail"
      Height          =   300
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1200
   End
   Begin VB.Label lblUser 
      Caption         =   "Paswoord"
      Height          =   300
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   420
      Width           =   1200
   End
   Begin VB.Label lblUser 
      Caption         =   "Naam"
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   1200
   End
End
Attribute VB_Name = "frmUsersEID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Public ConnectString As String
'
Enum FormMode
    fmAdd = 0
    fmEdit = 1
End Enum
Dim fmMode As FormMode

Private Sub cbEID_Click()

    Dim dummy As String
    Dim strReturn() As String
            

    On Local Error Resume Next
    dummy = Me.InteropUserControl1.clickon_eid()
    If dummy <> "" Then
        strReturn = Split(dummy, ";")
        Me.txtUser(0).Text = strReturn(0)
        Me.txtUser(5).Text = strReturn(0)
        Me.txtUser(2).Text = strReturn(1) & " " & strReturn(2)
        Me.txtUser(1).Text = strReturn(3)
    End If
    If Err Then
        MsgBox Err.Description
    End If
    
End Sub

Private Sub cmdAction_Click(Index As Integer)
    '
    On Error GoTo LocalErr
    '
    With adcUsers.Recordset
        Select Case Index
            Case 0 ' bewaren
                SaveUser txtUser(0).Text
                SetFormMode fmEdit
            Case 1 ' annuleren
                .Cancel
                SetFormMode fmEdit
            Case 2 ' nieuw
                .AddNew
                SetFormMode fmAdd
            Case 3 ' bijwerken
                .Update
                .MoveNext
                .MovePrevious
            Case 4 'verwijderen
                DeleteUser txtUser(0).Text
            Case 5 'rechten
                ShowUserRights txtUser(0).Text
            Case 6 'sluiten
                Unload Me
        End Select
    End With
    '
    Exit Sub
    '
LocalErr:
    ReportError Err.Number, Err.Source, Err.Description
    '
End Sub

Public Sub SetFormMode(Mode As FormMode)
    '
    Dim blnMode As Boolean
    '
    If Mode = fmAdd Then
        blnMode = False
    Else
        blnMode = True
    End If
    '
    cmdAction(0).Enabled = Not blnMode
    cmdAction(1).Enabled = Not blnMode
    cmdAction(2).Enabled = blnMode
    cmdAction(3).Enabled = blnMode
    cmdAction(4).Enabled = blnMode
    cmdAction(5).Enabled = blnMode
    '
End Sub

Private Sub Form_Activate()
    '
    On Error GoTo LocalErr
    '
    
    With adcUsers
        .ConnectionString = ConnectString
        .Refresh
    End With
    '
    
    Exit Sub
    '
LocalErr:
    
    ReportError Err.Number, Err.Source, Err.Description
    '
End Sub

Private Sub Form_Load()
    
    '
    SetFormMode fmEdit
    '
End Sub

Public Sub DeleteUser(UserID As String)
    '
    On Error GoTo LocalErr
    '
    Dim intAnswer As Integer
    Dim cnn As ADODB.Connection
    Set cnn = New ADODB.Connection
    '
    If UserID = "DEFAULT" Then
        MsgBox "De systeembeheerder kan niet verwijderd worden.", vbCritical
    Else
        intAnswer = MsgBox("Aktieve gebruiker: " + UserID + " vernietigen?", vbYesNo + vbDefaultButton2 + vbQuestion)
        If intAnswer = vbYes Then
            cnn.Open adcUsers.ConnectionString
            cnn.Execute "DELETE * FROM secAccess WHERE UserID='" & UserID & "'"
            adcUsers.Recordset.Delete
            adcUsers.Recordset.MoveFirst
        End If
    End If
    '
    Exit Sub
    '
LocalErr:
    ReportError Err.Number, Err.Source, Err.Description
    '
End Sub

Public Sub ShowUserRights(UserID As String)
    '
    With frmUserRights
        .ConnectString = adcUsers.ConnectionString
        .UserID = UserID
        .Show vbModal
    End With
    '
End Sub

Public Sub SaveUser(UserID As String)
    '
    On Error GoTo LocalErr
    '
    Dim intAnswer As Integer
    Dim cnn As ADODB.Connection
    Set cnn = New ADODB.Connection
    
    Dim rst As ADODB.Recordset
    '
    Dim strUser As String
    Dim strObject As String
    Dim strLevel As String
    Dim strSQL As String
    '
    intAnswer = MsgBox("Nieuwe gebruiker toevoegen ?", vbYesNo + vbQuestion)
    If intAnswer = vbNo Then
        adcUsers.Recordset.Cancel
        SetFormMode fmEdit
        Exit Sub
    End If
    '
    ' open connection, get defaults
    cnn.Open adcUsers.ConnectionString, "admin", ""
    Set rst = cnn.Execute("SELECT * FROM secAccess WHERE UserID='DEFAULT'", , adOpenForwardOnly)
    '
    ' use defaults to create new user rights
    Do While rst.EOF <> True
        strUser = "'" & UserID & "'"
        strObject = "'" & rst.Fields("Object") & "'"
        strLevel = CStr(rst.Fields("Level"))
        strSQL = "INSERT INTO secAccess VALUES(" & strUser & ", " & strObject & ", " & strLevel & ")"
        cnn.Execute strSQL, , adCmdText
        rst.MoveNext
    Loop
    '
    ' update new user rec
    adcUsers.Recordset.Update
    SetFormMode fmEdit
    '
    ' clean up
    rst.Close
    cnn.Close
    Set rst = Nothing
    Set cnn = Nothing
    '
    Exit Sub
    '
LocalErr:
    ReportError Err.Number, Err.Source, Err.Description
    '
End Sub

Private Sub txtUser_Change(Index As Integer)

    Select Case Index
        Case 0
            If txtUser(0).Text = "DEFAULT" Then
                Me.cbEID.Enabled = False
            Else
                Me.cbEID.Enabled = True
            End If
    End Select
            
End Sub

Private Sub txtUser_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Select Case Index
    Case 0
        If KeyCode = 46 Then
            KeyCode = 0
        End If
End Select

End Sub

Private Sub txtUser_KeyPress(Index As Integer, KeyAscii As Integer)

Select Case Index
    Case 0
        If cmdAction(2).Enabled Then
            KeyAscii = 0
        End If
End Select

End Sub
