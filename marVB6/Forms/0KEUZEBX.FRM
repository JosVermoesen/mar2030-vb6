VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form KeuzeBoxEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Editor"
   ClientHeight    =   4140
   ClientLeft      =   1080
   ClientTop       =   1470
   ClientWidth     =   6030
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4140
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   2775
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4895
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Broncode"
      TabPicture(0)   =   "0KEUZEBX.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Tekstinfo(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Wijzigen"
      TabPicture(1)   =   "0KEUZEBX.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "EditTekst"
      Tab(1).ControlCount=   1
      Begin VB.TextBox EditTekst 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2355
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   360
         Width           =   5715
      End
      Begin VB.TextBox Tekstinfo 
         DataField       =   "splitDefinitie"
         DataSource      =   "KeuzeData"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2355
         Index           =   1
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   360
         Width           =   5715
      End
   End
   Begin MSAdodcLib.Adodc KeuzeData 
      Height          =   330
      Left            =   0
      Top             =   3240
      Width           =   5895
      _ExtentX        =   10398
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ComboBox CboTableBox 
      Height          =   315
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   30
      Width           =   3345
   End
   Begin VB.CommandButton CmdNieuw 
      Caption         =   "Nieuw"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   4
      Top             =   3630
      Width           =   1450
   End
   Begin VB.TextBox Tekstinfo 
      DataField       =   "BestandsNaam"
      DataSource      =   "KeuzeData"
      Height          =   315
      Index           =   0
      Left            =   30
      TabIndex        =   3
      Top             =   30
      Width           =   2505
   End
   Begin VB.CommandButton Sluiten 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4440
      TabIndex        =   0
      Top             =   3600
      Width           =   1450
   End
   Begin VB.CommandButton Wegschrijven 
      Caption         =   "Wegschrijven"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1920
      TabIndex        =   1
      Top             =   3600
      Width           =   1450
   End
End
Attribute VB_Name = "KeuzeBoxEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim RecordIsAktief As Integer

Private Sub CboTableBox_Click()

KeuzeData.RecordSource = CboTableBox.Text
KeuzeData.Refresh

End Sub

Private Sub CmdNieuw_Click()

On Local Error Resume Next

KeuzeData.Recordset.AddNew
If CboTableBox.Text = "TipVanDeDag" Then
    TekstInfo(0).Text = Format(Now, "dd/mm/yy hh:mm:ss")
    TekstInfo(1).Text = "-"
End If

End Sub


Private Sub EditTekst_Change()

    If Wegschrijven.Enabled = False Then Wegschrijven.Enabled = True
    
End Sub

Private Sub Form_Load()

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If
Top = 0
Left = 0

KeuzeData.ConnectionString = _
    ADOJET_PROVIDER & _
    "Data Source=" & _
    PROGRAM_LOCATION & "Default2022.mdb;" & _
    "Persist Security Info=False"
    
MsgBox KeuzeData.ConnectionString
    
RecordIsAktief = False
CboTableBox.AddItem "KeuzeBoxData"
CboTableBox.AddItem "TipVanDeDag"
CboTableBox.ListIndex = 0

End Sub

Private Sub Form_Resize()

TekstInfo(1).Width = KeuzeBoxEditor.Width - 100
KeuzeData.Width = KeuzeBoxEditor.Width - 100

End Sub

'Private Sub KeuzeData_Error(DataErr As Integer, Response As Integer)

'DataErr = 0

'End Sub


Private Sub sluiten_Click()

Unload Me

End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)

    'stop
    
End Sub


Private Sub TekstInfo_Change(Index As Integer)

    Dim StringHier As String
            
    Select Case Index
        Case 0
            StringHier = TekstInfo(1).Text
            StringHier = Replace(StringHier, ";", vbCrLf)
            Me.EditTekst.Text = StringHier
            Me.Wegschrijven.Enabled = False
    End Select
    
End Sub

Private Sub TekstInfo_GotFocus(Index As Integer)

If Index = 1 Then SnelHelpPrint "Ctrl+Enter om lijn tussen te voegen", BL_LOGGING

End Sub


Private Sub Wegschrijven_Click()

On Local Error Resume Next
Dim StringHier As String

Msg = EditTekst.Text & vbCrLf & "bewaren.  Bent U zeker ?"
KtrlBox = MsgBox(Msg, 292)
If KtrlBox = 6 Then
    StringHier = Replace(EditTekst.Text, vbCrLf, ";")
    TekstInfo(1).Text = StringHier
    DoEvents
    KeuzeData.Recordset.Update
End If

End Sub

