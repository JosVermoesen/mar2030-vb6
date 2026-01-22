VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form mdiXlog 
   AutoRedraw      =   -1  'True
   Caption         =   "Log"
   ClientHeight    =   3195
   ClientLeft      =   2880
   ClientTop       =   375
   ClientWidth     =   6690
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   HelpContextID   =   21100
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3195
   ScaleWidth      =   6690
   Begin VB.CommandButton cmdClip 
      Caption         =   "Kopiëren naar Klipboard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid X 
      Height          =   2265
      HelpContextID   =   21100
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3995
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   12648384
      FocusRect       =   2
      HighLight       =   2
      ScrollBars      =   2
      AllowUserResizing=   1
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
End
Attribute VB_Name = "mdiXlog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub BewaarAls()
Dim Tel As Integer
Dim FlTempBewaar As Integer
Dim DeString As String
Dim TabPos As Integer
Dim LijstSep As String

On Local Error Resume Next
Err = 0
LijstSep = LaadTekst("c:\windows\win.ini;intl", "sList")
If Err Or LijstSep = "" Then
    LijstSep = ";"
    'MsgBox "Kan lijstseparator in internationale instellingen niet vinden.  ';' wordt automatisch genomen...  Kontroleer in ieder geval de systeeminstellingen !"
End If
Screen.MousePointer = vbHourglass

On Local Error GoTo CancelError
Mim.Teken.filename = ""
Mim.Teken.CancelError = True
Mim.Teken.Filter = "Alle bestanden (*.*)|*.*"
Mim.Teken.ShowSave

FlTempBewaar = FreeFile
Open Mim.Teken.filename For Output As FlTempBewaar

For Tel = 1 To X.Rows - 1
    X.Row = Tel
    X.ColSel = X.Cols - 1
    DeString = X.Clip
    TabPos = InStr(DeString, vbTab)
    Do While TabPos > 0
        Mid(DeString, TabPos, 1) = LijstSep
        TabPos = InStr(DeString, vbTab)
    Loop
    Print #FlTempBewaar, DeString; vbCrLf;
Next
Close FlTempBewaar

CancelError:
Screen.MousePointer = vbNormal

End Sub


Private Sub cmdClip_Click()

    X.RowSel = X.Rows - 1
    X.ColSel = X.Cols - 1
    Clipboard.Clear
    Clipboard.SetText X.Clip
            
End Sub

Private Sub Form_Load()

    LaadInstellingen Me

End Sub


Private Sub Form_Resize()

    On Local Error Resume Next

    'X.Top = 0
    'X.Left = 0
    X.Width = mdiXlog.Width - 100
    X.Height = mdiXlog.Height - 1200
    cmdClip.Top = mdiXlog.Height - 1000
    
    


End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    Dim X As Boolean
    X = InstellingenBewaard(Me)

End Sub

