VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form DashBordKlant 
   Caption         =   "DashBord Klanten"
   ClientHeight    =   5505
   ClientLeft      =   135
   ClientTop       =   435
   ClientWidth     =   9360
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5505
   ScaleWidth      =   9360
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   4380
      TabIndex        =   35
      Top             =   3090
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   0
      BackColor       =   -2147483633
      Appearance      =   1
      MonthColumns    =   2
      ShowToday       =   0   'False
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   99483650
      CurrentDate     =   36101
      MinDate         =   2
   End
   Begin VB.ComboBox CmbDokumentKeuze 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3120
      Width           =   3075
   End
   Begin VB.CommandButton CmdPrinterAfdruk 
      Caption         =   "Printer Af&druk"
      Height          =   375
      Left            =   3180
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton CmdSluiten 
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
      Height          =   375
      Left            =   7860
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   60
      Width           =   1215
   End
   Begin VB.ComboBox CmbBestandsKeuze 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1950
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   90
      Width           =   2685
   End
   Begin VB.CommandButton CmdBoekhouding 
      Caption         =   "Boekhoudin&g"
      Height          =   375
      Left            =   6300
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   60
      Width           =   1455
   End
   Begin VB.CommandButton CmdWijzigen 
      Caption         =   "&Wijzigen"
      Default         =   -1  'True
      Height          =   375
      Left            =   1950
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   1125
   End
   Begin VB.CommandButton Alfa 
      Caption         =   "Klant &zoeken "
      Height          =   375
      Left            =   4860
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   60
      Width           =   1335
   End
   Begin MSMask.MaskEdBox TxtInfo 
      Height          =   285
      Index           =   21
      Left            =   120
      TabIndex        =   33
      Tag             =   "v227"
      Top             =   2760
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtInfo 
      Height          =   285
      Index           =   20
      Left            =   4800
      TabIndex        =   30
      Tag             =   "v224"
      Top             =   2760
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtInfo 
      Height          =   285
      Index           =   19
      Left            =   4800
      TabIndex        =   29
      Tag             =   "v226"
      Top             =   2400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtInfo 
      Height          =   285
      Index           =   18
      Left            =   7260
      TabIndex        =   25
      Tag             =   "A110"
      Top             =   2100
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtInfo 
      Height          =   285
      Index           =   17
      Left            =   7260
      TabIndex        =   24
      Tag             =   "vs03"
      Top             =   1800
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtInfo 
      Height          =   285
      Index           =   16
      Left            =   7260
      TabIndex        =   23
      Tag             =   "A161"
      Top             =   1500
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtInfo 
      Height          =   285
      Index           =   15
      Left            =   7260
      TabIndex        =   21
      Tag             =   "A170"
      Top             =   1200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtInfo 
      Height          =   285
      Index           =   14
      Left            =   4800
      TabIndex        =   20
      Tag             =   "vs02"
      Top             =   2100
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtInfo 
      Height          =   285
      Index           =   13
      Left            =   4800
      TabIndex        =   19
      Tag             =   "A10A"
      Top             =   1800
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtInfo 
      Height          =   285
      Index           =   12
      Left            =   4800
      TabIndex        =   18
      Tag             =   "v201"
      Top             =   1500
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtInfo 
      Height          =   285
      Index           =   11
      Left            =   4800
      TabIndex        =   17
      Tag             =   "A120"
      Top             =   1200
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtInfo 
      Height          =   285
      Index           =   10
      Left            =   1500
      TabIndex        =   14
      Tag             =   "A108"
      Top             =   2160
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtInfo 
      Height          =   285
      Index           =   9
      Left            =   600
      TabIndex        =   13
      Tag             =   "A107"
      Top             =   2160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtInfo 
      Height          =   285
      Index           =   8
      Left            =   90
      TabIndex        =   12
      Tag             =   "A109"
      Top             =   2160
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtInfo 
      Height          =   285
      Index           =   7
      Left            =   3690
      TabIndex        =   11
      Tag             =   "A106"
      Top             =   1830
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtInfo 
      Height          =   285
      Index           =   6
      Left            =   2790
      TabIndex        =   10
      Tag             =   "A105"
      Top             =   1830
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtInfo 
      Height          =   285
      Index           =   5
      Left            =   60
      TabIndex        =   9
      Tag             =   "A104"
      Top             =   1830
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtInfo 
      Height          =   285
      Index           =   4
      Left            =   2790
      TabIndex        =   8
      Tag             =   "A127"
      Top             =   1500
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtInfo 
      Height          =   285
      Index           =   3
      Left            =   60
      TabIndex        =   7
      Tag             =   "A125"
      Top             =   1500
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtInfo 
      Height          =   285
      Index           =   2
      Left            =   2790
      TabIndex        =   6
      Tag             =   "A101"
      Top             =   1170
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtInfo 
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   5
      Tag             =   "A100"
      Top             =   1170
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtInfo 
      Height          =   285
      Index           =   0
      Left            =   150
      TabIndex        =   1
      Top             =   510
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid X 
      Height          =   1965
      Left            =   60
      TabIndex        =   36
      Top             =   3450
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   3466
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      SelectionMode   =   1
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
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Datum laatste kontakt"
      Height          =   255
      Index           =   5
      Left            =   180
      TabIndex        =   32
      Top             =   2520
      Width           =   1965
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GSM, E-mail adres"
      Height          =   255
      Index           =   4
      Left            =   7320
      TabIndex        =   31
      Top             =   2460
      Width           =   1965
   End
   Begin VB.Line Line1 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Index           =   1
      X1              =   7140
      X2              =   7140
      Y1              =   1320
      Y2              =   2640
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bank, btw, munt, o.ref."
      Height          =   255
      Index           =   3
      Left            =   7230
      TabIndex        =   22
      Top             =   870
      Width           =   2025
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Geboortedata,tel.,fax"
      Height          =   255
      Index           =   2
      Left            =   4890
      TabIndex        =   16
      Top             =   870
      Width           =   1965
   End
   Begin VB.Line Line1 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Index           =   0
      X1              =   4680
      X2              =   4680
      Y1              =   1260
      Y2              =   3000
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Adres Huishouding"
      Height          =   255
      Index           =   1
      Left            =   210
      TabIndex        =   15
      Top             =   870
      Width           =   1725
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Klantnummer"
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   1245
   End
End
Attribute VB_Name = "DashBordKlant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z



Sub XSchoon()

X.Rows = 2
With X
    .Col = 0
    .Row = 1
    .Text = ""
    
    .Col = 1
    .Text = ""
    
    .Col = 2
    .Text = ""
End With

End Sub

Private Sub Alfa_Click()

SharedFl = FlKlant

aIndex = 1
GridText = TxtInfo(1).Text
SqlSearch.Show 1
If Ktrl Then
Else
    VensterVullen vBibTekst(FlKlant, "#A110 #")
End If

End Sub



Private Sub CmbdokumentKeuze_Click()

VensterVullen vBibTekst(FlKlant, "#A110 #")

End Sub


Private Sub CmdBoekhouding_Click()

BalansKontrole FlKlant

End Sub



Private Sub CmdPrinterAfdruk_Click()

'MsgBox "Grafische schermafdruk is tijdelijk niet mogelijk"
PrintForm

End Sub



Private Sub cmdSluiten_Click()

Unload Me

End Sub


Private Sub cmdWijzigen_Click()
Dim TempoFL As Integer
Dim TempoKtrl As Integer

TempoFL = FlKlant
bGet FlKlant, 0, vSet(TxtInfo(0).Text, 12)
TempoKtrl = Ktrl
If TempoKtrl = 0 Then
    RecordToVeld TempoFL
    InsertFlag(TempoFL) = 0
Else
    InsertFlag(TempoFL) = 1
    TLBRecord(TempoFL) = ""
End If
           
If TeleBibClick(TempoFL) = False Then
    VensterVullen vBibTekst(FlKlant, "#A110 #")
Else
    If InsertFlag(TempoFL) = 0 Then
        If TempoKtrl Then
            bInsert TempoFL, 0
        Else
            bUpdate TempoFL, 0
        End If
    End If
    cmdWijzigen.Default = True
    VensterVullen vBibTekst(FlKlant, "#A110 #")
End If

End Sub



Private Sub Form_Load()

MsgBox "Nog in ontwikkeling en dus BEPERKT in gebruik en mogelijkheden.", , "DashBord Klanten"
Left = 0
Top = 0

X.Cols = 3
With X
    .Col = 0
    .Row = 0
    .Text = "dokument"
    .ColWidth(0) = 900
    
    .Col = 1
    .Text = "Datum"
    .ColWidth(1) = 960
    
    .Col = 2
    .Text = "Andere relaties"
    .ColWidth(2) = 2340
End With

CmbBestandsKeuze.AddItem "01: Klantenfiche"
CmbBestandsKeuze.ListIndex = 0

CmbDokumentKeuze.AddItem "15: Verkoopfakturen, Creditnota's"
CmbDokumentKeuze.AddItem "14: Bestelbons"
CmbDokumentKeuze.AddItem "13: Offertes"
CmbDokumentKeuze.ListIndex = 0

End Sub


Private Sub VensterVullen(Klant As String)
Dim TempoString As String
Dim LaatsteWAS As String

XSchoon
KlantVerversen:
For Teller = 1 To 21
    TxtInfo(Teller).Text = vBibTekst(FlKlant, "#" + TxtInfo(Teller).Tag + " #")
Next
TxtInfo(0).Text = vBibTekst(FlKlant, "#A110 #")

TempoString = ""
T = 0
bGet FlAllerlei, 0, vSet("K" + vBibTekst(FlKlant, "#A110 #"), 13)
If Ktrl Then
    Exit Sub
Else
    Do
        GoSub VolgendeLijn
        bNext FlAllerlei
        If Ktrl Or vSet(KeyBuf(FlAllerlei), 13) <> vSet("K" + vBibTekst(FlKlant, "#A110 #"), 13) Then
            Exit Do
        End If
    Loop
End If
Exit Sub

VolgendeLijn:
RecordToVeld FlAllerlei
Select Case Left(vBibTekst(FlAllerlei, "#v005 #"), 2)
    Case Left(CmbDokumentKeuze.Text, 2)
        If LaatsteWAS = vBibTekst(FlAllerlei, "#v033 #") Then
        Else
            LaatsteWAS = vBibTekst(FlAllerlei, "#v033 #")
            TempoString = vBibTekst(FlAllerlei, "#v033 #") + vbTab
            TempoString = TempoString + fDatumText(vBibTekst(FlAllerlei, "#v035 #")) + vbTab
            TempoString = TempoString + vBibTekst(FlAllerlei, "#v147 #") + " " + vBibTekst(FlAllerlei, "#v148 #")
            X.AddItem TempoString, X.Rows - 1
        End If
End Select
Return

End Sub

Private Sub Form_Resize()

On Local Error Resume Next
' X.Width = DashBordKlant.Width - 100
' X.Height = DashBordKlant.Height - X.Top - 400

End Sub




Private Sub TxtInfo_GotFocus(Index As Integer)

cmdWijzigen.Default = True

End Sub



Private Sub X_Click()

On Local Error Resume Next
With X
    .Col = 2
    '.SelStartCol = 0
    '.SelEndCol = 4
    '.SelStartRow = .Row
    '.SelEndRow = .Row
End With

End Sub

Private Sub X_DblClick()

GridText9 = TxtInfo(0).Text + vbTab
GridText9 = GridText9 + Left(CmbDokumentKeuze.Text, 2) + vbTab
X.Col = 0
GridText9 = GridText9 + X.Text
Unload DirekteVerkoop
DirekteVerkoop.Show
DirekteVerkoop.SetFocus

End Sub


Private Sub X_GotFocus()

cmdWijzigen.Default = False

End Sub


Private Sub X_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
    X_DblClick
End If

End Sub

Private Sub X_KeyUp(KeyCode As Integer, Shift As Integer)

X_Click

End Sub

Private Sub X_LostFocus()

cmdWijzigen.Default = True

End Sub

