VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ntInputbox 
   ClientHeight    =   1395
   ClientLeft      =   300
   ClientTop       =   4980
   ClientWidth     =   6930
   ControlBox      =   0   'False
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
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1395
   ScaleWidth      =   6930
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdVooruitAchteruit 
      Caption         =   ">"
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   4
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton cmdVooruitAchteruit 
      Caption         =   "<"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Hernieuw 
      Caption         =   "&Hernieuw"
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
      Left            =   4080
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   1445
   End
   Begin VB.CommandButton Sluiten 
      Cancel          =   -1  'True
      Caption         =   "Negeren"
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
      Left            =   5640
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Ok 
      Caption         =   "&Ok"
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
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   1215
   End
   Begin MSMask.MaskEdBox TekstInfo 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSComctlLib.StatusBar MedeDeling 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   6
      Top             =   1095
      Width           =   6930
      _ExtentX        =   12224
      _ExtentY        =   529
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   12347
            MinWidth        =   12347
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc DefaultData 
      Height          =   330
      Left            =   120
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Enabled         =   0
      Connect         =   ""
      OLEDBString     =   ""
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblInfo 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   600
      Width           =   2655
   End
End
Attribute VB_Name = "ntInputbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub VernieuwInfo()

    Select Case Mid(MedeDeling.SimpleText, 2, 2)
        Case "00"
            lblInfo.Caption = DefaultData.Recordset("ISOLandNummer") + ", " + DefaultData.Recordset("ISOLandkode") + ", " + DefaultData.Recordset("ISOMuntKode") + ", " + DefaultData.Recordset("LandNaam")
            TekstInfo.Text = DefaultData.Recordset("ISOLandNummer") ''
            
        Case "01"
            lblInfo.Caption = DefaultData.Recordset("PostKode") + ", " + DefaultData.Recordset("Plaatsnaam")
            TekstInfo.Text = DefaultData.Recordset("PostKode")

        Case "02"
            lblInfo.Caption = DefaultData.Recordset("PostKode") + ", " + DefaultData.Recordset("Plaatsnaam")
            TekstInfo.Text = DefaultData.Recordset("PlaatsNaam")
    End Select
    
End Sub


Private Sub cmdVooruitAchteruit_Click(Index As Integer)

On Error Resume Next

Select Case Index
    Case 0
        DefaultData.Recordset.MovePrevious
        If DefaultData.Recordset.BOF Then
            DefaultData.Recordset.MoveFirst
        End If
        
    Case 1
        DefaultData.Recordset.MoveNext
        If DefaultData.Recordset.EOF Then
            DefaultData.Recordset.MoveLast
        End If
End Select
VernieuwInfo

End Sub

Private Sub Form_Activate()

    If Hernieuw.Visible Then
        Hernieuw_Click
    End If

End Sub

Private Sub Form_Initialize()

DefaultData.ConnectionString = ADOJET_PROVIDER & "Data Source=" + PROGRAM_LOCATION + "Default2022.mdb;Persist Security Info=False"

End Sub


Private Sub Hernieuw_Click()

DefaultData.RecordSource = SQLBevel + RTrim$(TekstInfo.Text) + "%';"
MedeDeling.SimpleText = GridText + ntInputbox.DefaultData.RecordSource
DefaultData.Refresh
If DefaultData.Recordset.RecordCount Then
    VernieuwInfo
End If
Ok.Default = True

End Sub

Private Sub Ok_Click()
On Local Error Resume Next

TekstInfo.SetFocus
ntInputbox.Hide

End Sub

Private Sub sluiten_Click()
On Local Error Resume Next

TekstInfo.Text = Chr$(255)
TekstInfo.SetFocus
ntInputbox.Hide

End Sub

Private Sub TekstInfo_GotFocus()

TekstInfo.SelLength = Len(TekstInfo.Text)

End Sub

Private Sub TekstInfo_KeyPress(KeyAscii As Integer)

If ntInputbox.Tag = "Edit No" Then KeyAscii = 0
If Hernieuw.Visible Then
    Hernieuw.Default = True
End If

End Sub

