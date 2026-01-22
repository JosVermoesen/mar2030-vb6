VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmUserRights 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   Icon            =   "frmUserRights.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   795
      Left            =   1620
      Top             =   1200
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1402
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   1
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=RVBeveiliging"
      OLEDBString     =   "DSN=RVBeveiliging"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frmUserRights.frx":0442
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2990
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdSluiten 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   300
      Left            =   720
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Menu mnuRights 
      Caption         =   "&Rechten"
      Begin VB.Menu mnuRightsAdd 
         Caption         =   "&Bijvoegen..."
      End
      Begin VB.Menu mnuRightsDelete 
         Caption         =   "&Verwijderen..."
      End
   End
End
Attribute VB_Name = "frmUserRights"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Public UserID As String
Public ConnectString As String


Private Sub Adodc1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

If ErrorNumber = 3596 Then fCancelDisplay = True


End Sub

Private Sub cmdSluiten_Click()

    Unload Me
    
End Sub

Private Sub Form_Activate()
       
    '
    On Error GoTo LocalErr
    '
    If UserID <> "DEFAULT" Then
        mnuRights.Visible = False
    End If
    '
    Me.Caption = "Rechten voor Gebruiker [" & UserID & "]"
    '
    UpdateGrid
    '
    Exit Sub
    '
LocalErr:
    ReportError Err.Number, Err.Source, Err.Description
    '

End Sub

Private Sub Form_Resize()
    '
    On Error GoTo LocalErr
    '
    With MSHFlexGrid1
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
        '
        .ColWidth(0) = .Width * 0.65
        .ColWidth(1) = .Width * 0.3
        '
    End With
    '
    Exit Sub
    '
LocalErr:
    ReportError Err.Number, Err.Source, Err.Description
    '
End Sub

Private Sub mnuRightsAdd_Click()
    '
    With frmEditRights
        .UserID = UserID
        .Object = ""
        .Level = 0
        .Action = faAdd
        .UpdateRec = False
        .Show vbModal
        '
        If .UpdateRec = True Then
            UpdateRights .UserID, .Object, .Level
            Unload Me
        End If
    End With
    '
End Sub

Private Sub mnuRightsDelete_Click()
    '
    On Error GoTo LocalErr
    '
    If MSHFlexGrid1.Col <> 0 Then
        MsgBox "Select a Right to Delete", vbExclamation, "Delete Rights"
        Exit Sub
    End If
    '
    With frmEditRights
        .UserID = UserID
        .Object = MSHFlexGrid1.Text
        .Level = 0
        .Action = faDelete
        .UpdateRec = False
        .Show vbModal
        '
        If .UpdateRec = True Then
            UpdateRights .UserID, .Object, .Level, faDelete
            Unload Me
        End If
    End With
    '
    Exit Sub
    '
LocalErr:
    ReportError Err.Number, Err.Source, Err.Description
    '
End Sub

Private Sub MSHFlexGrid1_DblClick()
    '
    On Error GoTo LocalErr
    '
    With MSHFlexGrid1
        If .Col = 0 Then
            ShowEditRights UserID, .Text
        End If
    End With
    '
    Exit Sub
    '
LocalErr:
    ReportError Err.Number, Err.Source, Err.Description
    '
End Sub

Public Sub ShowEditRights(UserID As String, Object As String)
    '
    On Error GoTo LocalErr
    '
    With frmEditRights
        .UserID = UserID
        .Object = Object
        .Level = GetLevel(UserID, Object)
        .Action = faEdit
        .UpdateRec = False
        .Show vbModal
        '
        If .UpdateRec = True Then
            UpdateRights .UserID, .Object, .Level, faEdit
            'Unload Me
        End If
    End With
    '
    Exit Sub
    '
LocalErr:
    ReportError Err.Number, Err.Source, Err.Description
    '
End Sub

Public Sub UpdateRights(UserID As String, Object As String, Level As Integer, Optional Action As FormAction)
    '
    On Error GoTo LocalErr
    '
    Dim cnn As ADODB.Connection
    Set cnn = New ADODB.Connection
    '
    cnn.Open Adodc1.ConnectionString, "admin", ""
    '
    Select Case Action
        Case faAdd
            AddNewRight cnn, Object, Level
        Case faDelete
            cnn.Execute "DELETE * FROM secAccess WHERE Object='" & Object & "'", , adCmdText
        Case faEdit
            cnn.Execute "UPDATE secAccess SET [Level]=" & CStr(Level) & " WHERE UserID='" & UserID & "' AND Object='" & Object & "'", , adCmdText
    End Select
    '
    cnn.Close
    Set cnn = Nothing
    '
    Exit Sub
    '
LocalErr:
    ReportError Err.Number, Err.Source, Err.Description
    '
End Sub

Public Function GetLevel(UserID As String, Object As String) As Integer
    '
    On Error GoTo LocalErr
    '
    Dim cnn As New ADODB.Connection
    Dim rst As ADODB.Recordset
    '
    cnn.Open Adodc1.ConnectionString, "admin", ""
    Set rst = cnn.Execute("SELECT [Level] FROM secAccess WHERE UserID='" & UserID & "' AND Object='" & Object & "'", , adCmdText)
    GetLevel = rst.Fields("Level")
    '
    rst.Close
    cnn.Close
    Set cnn = Nothing
    Set rst = Nothing
    '
    Exit Function
    '
LocalErr:
    ReportError Err.Number, Err.Source, Err.Description
    '
End Function

Public Sub UpdateGrid()
    '
    On Error GoTo LocalErr
    '
    With Adodc1
        .ConnectionString = ConnectString
        .RecordSource = "SELECT Object,[Level] FROM secAccess WHERE USerID='" & UserID & "'"
        .Refresh
    End With
    '
    Exit Sub
    '
LocalErr:
    ReportError Err.Number, Err.Source, Err.Description
    '
End Sub

Public Sub AddNewRight(Connection As ADODB.Connection, Object As String, Level As Integer)
    '
    ' add new right to all users in the database
    '
    Dim rsUsers As ADODB.Recordset
    Dim strInsertSQL As String
    '
    Set rsUsers = Connection.Execute("SELECT Userid FROM secUsers", , adCmdText)
    '
    Do While rsUsers.EOF = False
        Connection.Execute "INSERT INTO secAccess VALUES('" & rsUsers.Fields("UserID") & "', '" & Object & "', " & CStr(Level) & ")", , adCmdText
        rsUsers.MoveNext
    Loop
    '
    rsUsers.Close
    Set rsUsers = Nothing
    '
End Sub

Private Sub MSHFlexGrid1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then MSHFlexGrid1_DblClick

End Sub
