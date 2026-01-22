VERSION 5.00
Begin VB.Form SjabloonIO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verkoopsjabloon invoegen/wegschrijven"
   ClientHeight    =   2700
   ClientLeft      =   1095
   ClientTop       =   1470
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2700
   ScaleWidth      =   5640
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton SoortSjabloon 
      Caption         =   "&Klantgericht"
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   1455
   End
   Begin VB.OptionButton SoortSjabloon 
      Caption         =   "&Algemeen"
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Value           =   -1  'True
      Width           =   1275
   End
   Begin VB.ListBox SjabloonLijst 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1260
      Left            =   30
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   5550
   End
   Begin VB.CommandButton Invoegen 
      Caption         =   "&Invoegen"
      Default         =   -1  'True
      Height          =   750
      Left            =   2940
      TabIndex        =   3
      Top             =   90
      Width           =   1080
   End
   Begin VB.CommandButton Wegschrijven 
      Caption         =   "&Wegschrijven"
      Height          =   360
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Width           =   1410
   End
   Begin VB.TextBox TekstInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   330
      Index           =   0
      Left            =   30
      TabIndex        =   2
      Top             =   2280
      Width           =   5505
   End
   Begin VB.CommandButton Annuleren 
      Cancel          =   -1  'True
      Caption         =   "&Sluiten"
      Height          =   345
      Left            =   4080
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   510
      Width           =   1410
   End
End
Attribute VB_Name = "SjabloonIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Private Sub Annuleren_Click()

GridText = ""
Unload Me

End Sub

Private Sub Form_Load()

SoortSjabloon(1).Value = True

End Sub

Private Sub Invoegen_Click()
Dim FlX As Integer
Dim Msg As String
Dim BsType As String * 7
Dim BsTekst As String
Dim GridIO As String

If Right(SjabloonLijst.Text, 40) = Space$(40) Then
    Beep
    Exit Sub
End If
      Msg = "'" + Right(SjabloonLijst.Text, 40) + "'" + vbCrLf + vbCrLf
Msg = Msg + "Kies Ja  om enkel bij te voegen." + vbCrLf
Msg = Msg + "Kies Nee om bestaande invoer te overschrijven !" + vbCrLf
Msg = Msg + "Cancel of ESC voor nieuwe selectie..."
Ktrl = MsgBox(Msg$, 3, "Sjabloon invoegen")
Select Case Ktrl
    Case 2
        Exit Sub
    Case 7
        GridText = ""
End Select

Select Case SoortSjabloon(0).Value
    Case True
        BsType = "A999999"
    Case Else
        BsType = "K" + Left(FVT(TABLE_CUSTOMERS, 0), 6)
End Select

FlX = FreeFile
Open LOCATION_COMPANYDATA + BsType + Left(SjabloonLijst.Text, 1) + ".SJB" For Input As FlX
    Line Input #FlX, BsTekst
    Do While Not EOF(FlX)
        Line Input #FlX, BsTekst
        GridText = GridText + BsTekst + Chr$(13)
    Loop
Close FlX
Unload Me

End Sub

Private Sub SjabloonLijst_Click()

TekstInfo(0).Text = Right(SjabloonLijst.Text, 40)
Invoegen.Enabled = True

End Sub

Private Sub SjabloonLijst_DblClick()

Invoegen_Click

End Sub

Private Sub SoortSjabloon_Click(Index As Integer)
Dim BsType As String * 7
Dim BsTekst As String * 40
Dim T As Integer
Dim FlX As Integer

Do While SjabloonLijst.ListCount > 0
    SjabloonLijst.RemoveItem 0
Loop

Select Case Index
    Case 0
        BsType = "A999999"
    Case Else
        BsType = "K" + Left(FVT(TABLE_CUSTOMERS, 0), 6)
End Select

For T = 0 To 9
    If Dir$(LOCATION_COMPANYDATA + BsType + Format(T, "0") + ".SJB") = "" Then
        SjabloonLijst.AddItem Format(T, "0") + ": " + Space$(40)
    Else
        FlX = FreeFile
        Open LOCATION_COMPANYDATA + BsType + Format(T, "0") + ".SJB" For Input As FlX
            Input #FlX, BsTekst
        Close FlX
        SjabloonLijst.AddItem Format(T, "0") + ": " + BsTekst
    End If
Next
SjabloonLijst.ListIndex = 0

End Sub

Private Sub TekstInfo_Change(Index As Integer)

Invoegen.Enabled = False

End Sub

Private Sub Wegschrijven_Click()
Dim FlX As Integer
Dim Msg As String
Dim BsType As String * 7
Dim BsTekst As String * 40
Dim GridIO As String

BsTekst = TekstInfo(0).Text

      Msg = "'" + BsTekst + "'" + vbCrLf + vbCrLf
Msg = Msg + "Wegschrijven, bent U zeker ?"
Ktrl = MsgBox(Msg$, 292)
If Ktrl = 6 Then
Else
    Exit Sub
End If

Select Case SoortSjabloon(0).Value
    Case True
        BsType = "A999999"
    Case Else
        BsType = "K" + Left(FVT(TABLE_CUSTOMERS, 0), 6)
End Select

FlX = FreeFile
Open LOCATION_COMPANYDATA + BsType + Left(SjabloonLijst.Text, 1) + ".SJB" For Output As FlX
    Print #FlX, BsTekst
    Do While GridText <> ""
        LeftPos = InStr(GridText, Chr$(13))
        If LeftPos <= 1 Then
            GridText = ""
            Exit Do
        End If
        Print #FlX, Left(GridText, InStr(GridText, Chr$(13)) - 1); vbCrLf;
        GridText = Right(GridText, Len(GridText) - LeftPos)
    Loop
Close FlX
GridText = ""
Unload Me

End Sub

