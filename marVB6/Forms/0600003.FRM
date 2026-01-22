VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form HistoriekPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Historiek Rekeningen (Grootboek)"
   ClientHeight    =   2100
   ClientLeft      =   3615
   ClientTop       =   2415
   ClientWidth     =   5445
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
   ForeColor       =   &H00000000&
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2100
   ScaleWidth      =   5445
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox cbPeriodiekeTotalen 
      Alignment       =   1  'Right Justify
      Caption         =   "&Periodieke Totalen"
      Height          =   255
      Left            =   3420
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1740
      Value           =   1  'Checked
      Width           =   2010
   End
   Begin VB.CheckBox chkAfdrukLiggend 
      Caption         =   "&Liggende Printerafdruk"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1740
      Width           =   2310
   End
   Begin VB.CheckBox chkAfdrukInVenster 
      Caption         =   "Afdruk in venster"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   1830
   End
   Begin VB.CommandButton Annuleren 
      Cancel          =   -1  'True
      Caption         =   "&Sluiten"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4080
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   450
      Width           =   1260
   End
   Begin VB.CommandButton Drukken 
      Caption         =   "Af&drukken"
      Height          =   345
      Left            =   4080
      TabIndex        =   4
      Top             =   60
      Width           =   1260
   End
   Begin MSMask.MaskEdBox TekstLijn 
      Height          =   255
      Index           =   3
      Left            =   1860
      TabIndex        =   3
      Top             =   420
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      AutoTab         =   -1  'True
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TekstLijn 
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   1
      Top             =   420
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TekstLijn 
      Height          =   285
      Index           =   1
      Left            =   2910
      TabIndex        =   8
      Top             =   1140
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtPeriodeVanTot 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   1140
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      AutoTab         =   -1  'True
      MaxLength       =   23
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/#### - ##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Periode Van - &Tot"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Tot"
      Height          =   255
      Index           =   3
      Left            =   1980
      TabIndex        =   2
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Vanaf Rekening"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lijstdatu&m"
      Height          =   255
      Index           =   1
      Left            =   2910
      TabIndex        =   7
      Top             =   840
      Width           =   930
   End
End
Attribute VB_Name = "HistoriekPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit

Dim psTekst(5)              As String
Dim LFontSize(20)           As Single
Dim LAantalL(20)            As Integer
Dim FontDefChanged          As Integer

Dim VeldTXT(20)             As String

Dim SubTotaalD              As Currency
Dim SubTotaalC              As Currency
Dim TotaalD                 As Currency
Dim TotaalC                 As Currency
Dim AlgTotaalD              As Currency
Dim AlgTotaalC              As Currency

Dim Lijn                    As Long
Dim SelectieVanTot          As String


Sub AlgemeenEindTotaal()

Dim T   As Integer

For T = 0 To 8:
    VeldTXT(T) = ""
Next

VeldTXT(2) = "Proef- en Saldi :"
VeldTXT(5) = Dec$((AlgTotaalD), MASK_EURBH)
VeldTXT(6) = Dec$(Abs(AlgTotaalC), MASK_EURBH)

T = 0
Do While REPORT_TAB(T) <> 0
    Printer.Print Tab(REPORT_TAB(T));
    Printer.Print VeldTXT(T);
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        Printer.Print vbCrLf;
    End If
    T = T + 1
Loop

End Sub

Private Sub Annuleren_Click()

Unload Me

End Sub




Private Sub chkAfdrukLiggend_Click()

Set Printer = Printers(LijstPrinterNr)
On Error Resume Next
Printer.PaperBin = LaadTekst(App.Title, "LIJSTPRINTER")
If chkAfdrukLiggend.Value = 1 Then
    Printer.Orientation = vbPRORLandscape
Else
    Printer.Orientation = vbPRORPortrait
End If

End Sub

Private Sub Drukken_Click()
Dim BeginSleutel As String * 15
Dim EindSleutel As String * 15
Dim SubTitelTekst As String
Dim VorigeSleutel As String * 15
Dim DCBedrag As Currency
Dim XX As Integer

Lijn = 0
VorigeSleutel = ""
BeginSleutel = vSet((TekstLijn(2).Text), 7) + Left(SelectieVanTot, 8)
EindSleutel = vSet((TekstLijn(3).Text), 7) + Right(SelectieVanTot, 8)

psTekst(2) = "Historieken " + Mid(Mim.Caption, InStr(Mim.Caption, "["))
psTekst(0) = TekstLijn(1).Text

InitVelden
bFirst TABLE_JOURNAL, 0
bGetOrGreater TABLE_JOURNAL, 0, BeginSleutel
If Ktrl Then
    Beep
    Exit Sub
End If

Screen.MousePointer = vbHourglass
    PAGE_COUNTER = 0
    If chkAfdrukInVenster = 0 Then
        If Printer.Width > 12000 Then
            Printer.FontSize = 10
            Printer.FontName = "Courier New"
            Printer.Print " "
            Printer.FontSize = 10
        Else
            Printer.FontSize = 7.2
            Printer.FontName = "Courier New"
            Printer.Print " "
            Printer.FontSize = 7.2
            Printer.FontBold = True
        End If
    End If
    VolgendePagina
        
If DATE_CHECK(Right(KEY_BUF(TABLE_JOURNAL), 8), BOOKYEAR_KEY) Then
    If Left(KEY_BUF(TABLE_JOURNAL), 7) >= Left(BeginSleutel, 7) Then
        GoSub InstalSubTitel
        GoSub PrintInfo
    End If
End If

Do
    bNext TABLE_JOURNAL
    If Ktrl Or vSet(KEY_BUF(TABLE_JOURNAL), 15) > EindSleutel Then
        Exit Do
    End If
    If Left(VorigeSleutel, 7) = Left(KEY_BUF(TABLE_JOURNAL), 7) Then
        If Not DATE_CHECK(Right(KEY_BUF(TABLE_JOURNAL), 8), BOOKYEAR_KEY) Then
        ElseIf Mid(VorigeSleutel, 12, 2) = Mid(KEY_BUF(TABLE_JOURNAL), 12, 2) Then
            GoSub PrintInfo
        Else
            PeriodiekTotaal
            If chkAfdrukInVenster Then
            Else
                Printer.Print vbCrLf;
            End If
            VorigeSleutel = vSet(KEY_BUF(TABLE_JOURNAL), 15)
            GoSub PrintInfo
        End If
    Else
        If DATE_CHECK(Right(KEY_BUF(TABLE_JOURNAL), 8), BOOKYEAR_KEY) Then
            GoSub InstalSubTitel
            GoSub PrintInfo
        End If
    End If
Loop
    PeriodiekTotaal
    EindTotaal
    AlgemeenEindTotaal
    Printer.NewPage
    Printer.EndDoc
    Screen.MousePointer = vbNormal
    Annuleren_Click
Exit Sub

InstalSubTitel:
If Lijn <> 0 Then
    PeriodiekTotaal
    EindTotaal
End If
Lijn = 0
RecordToVeld TABLE_JOURNAL
bGet TABLE_LEDGERACCOUNTS, 0, vSet(vBibTekst(TABLE_JOURNAL, "#v019 #"), 7)
If Ktrl Then
    SubTitelTekst = vSet(vBibTekst(TABLE_JOURNAL, "#v019 #"), 7) + " rekening reeds vernietigd..."
Else
    RecordToVeld TABLE_LEDGERACCOUNTS
    SubTitelTekst = vBibTekst(TABLE_LEDGERACCOUNTS, "#v019 #") + " " + vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #")
    SnelHelpPrint SubTitelTekst, BL_LOGGING
    psTekst(3) = SubTitelTekst
End If
VorigeSleutel = vSet(KEY_BUF(TABLE_JOURNAL), 15)
PrintTitel
Return

PrintInfo:
RecordToVeld TABLE_JOURNAL
Lijn = Lijn + 1
VeldTXT(0) = Format(Lijn, "00000")
VeldTXT(1) = DATE_TEXT(vBibTekst(TABLE_JOURNAL, "#v066 #"))
VeldTXT(2) = vBibTekst(TABLE_JOURNAL, "#v067 #")
VeldTXT(3) = vBibTekst(TABLE_JOURNAL, "#v033 #")
VeldTXT(4) = vBibTekst(TABLE_JOURNAL, "#v038 #")
VeldTXT(7) = vBibTekst(TABLE_JOURNAL, "#v069 #")
If RTrim$(VeldTXT(3)) <> "" Then
    bGet TABLE_INVOICES, 0, VeldTXT(3)
    If Ktrl Then
        VeldTXT(8) = VeldTXT(3)
    Else
        VeldTXT(8) = ""
    End If
Else
    VeldTXT(8) = ""
End If

DCBedrag = Val(vBibTekst(TABLE_JOURNAL, "#v068 #"))

Select Case DCBedrag
    Case Is < 0
        SubTotaalC = SubTotaalC + DCBedrag
        TotaalC = TotaalC + DCBedrag
        VeldTXT(5) = ""
        VeldTXT(6) = Dec$(Abs(DCBedrag), MASK_EURBH)
    Case Else
        SubTotaalD = SubTotaalD + DCBedrag
        TotaalD = TotaalD + DCBedrag
        VeldTXT(5) = Dec$((DCBedrag), MASK_EURBH)
        VeldTXT(6) = ""
End Select

PrintVelden
Return

End Sub

Private Sub EindTotaal()
Dim T   As Integer

For T = 0 To 8
    VeldTXT(T) = ""
Next

VeldTXT(2) = "Boekjaartotalen :"
VeldTXT(5) = Dec$((TotaalD), MASK_EURBH)
VeldTXT(6) = Dec$((Abs(TotaalC)), MASK_EURBH)

T = 0
Do While REPORT_TAB(T) <> 0
    Printer.Print Tab(REPORT_TAB(T));
    Printer.Print VeldTXT(T);
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        Printer.Print vbCrLf;
    End If
    T = T + 1
Loop
Printer.Print vbCrLf;
If Printer.CurrentY >= Printer.Height - 1250 Then
    Printer.NewPage
    Printer.FontSize = Printer.FontSize
    Printer.Print " "
    Printer.FontSize = Printer.FontSize
    VolgendePagina
    PrintTitel
End If

Lijn = 0
SubTotaalD = 0
SubTotaalC = 0
TotaalD = 0
TotaalC = 0

End Sub

Private Sub Form_Load()

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If

PAGE_COUNTER = 0
AlgTotaalD = 0
AlgTotaalC = 0
SubTotaalD = 0
SubTotaalC = 0
TotaalD = 0
TotaalC = 0

SelectieVanTot = BOOKYEAR_FROMTO
TxtPeriodeVanTot = DATE_TEXT(Left(SelectieVanTot, 8)) + " - " + DATE_TEXT(Right(SelectieVanTot, 8))
TekstLijn(1).Text = MIM_GLOBAL_DATE
TekstLijn(2).Text = 10
TekstLijn(3).Text = "7999999"

Set Printer = Printers(LijstPrinterNr)
On Error Resume Next
Printer.PaperBin = LaadTekst(App.Title, "LIJSTPRINTER")
If Printer.Orientation = vbPRORLandscape Then
    chkAfdrukLiggend.Value = 1
Else
    chkAfdrukLiggend.Value = 0
End If
chkAfdrukLiggend_Click
    
End Sub

Private Sub InitVelden()

REPORT_FIELD(0) = "Lijn"
        REPORT_TAB(0) = 2

REPORT_FIELD(1) = "Datum"
        REPORT_TAB(1) = 9

REPORT_FIELD(2) = "Omschrijving"
        REPORT_TAB(2) = 22

REPORT_FIELD(3) = "dokument"
        REPORT_TAB(3) = 53

REPORT_FIELD(4) = "Fin.Dok."
        REPORT_TAB(4) = 67

REPORT_FIELD(5) = "       Debet"
        REPORT_TAB(5) = 76

REPORT_FIELD(6) = "      Credit"
        REPORT_TAB(6) = 89

REPORT_FIELD(7) = "T.Reken"
        REPORT_TAB(7) = 103
        
REPORT_FIELD(8) = "VSF.Kode"
        REPORT_TAB(8) = 111
        
REPORT_TAB(9) = 0

End Sub

Private Sub PeriodiekTotaal()
Dim T       As Integer

On Local Error GoTo PrtHandler5

For T = 0 To 8
    VeldTXT(T) = ""
Next

VeldTXT(2) = "Periodiek totaal :"
VeldTXT(5) = Dec$((SubTotaalD), MASK_EURBH)
VeldTXT(6) = Dec$((Abs(SubTotaalC)), MASK_EURBH)

AlgTotaalD = AlgTotaalD + SubTotaalD
AlgTotaalC = AlgTotaalC + SubTotaalC
If cbPeriodiekeTotalen.Value <> 0 Then
Else
    SubTotaalD = 0
    SubTotaalC = 0
    Exit Sub
End If

T = 0
Printer.Print vbCrLf;
Do While REPORT_TAB(T) <> 0
    Printer.Print Tab(REPORT_TAB(T));
    Printer.Print VeldTXT(T);
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        Printer.Print vbCrLf;
    End If
    T = T + 1
Loop

'Printer.Print vbCrLf;
SubTotaalD = 0
SubTotaalC = 0
If Printer.CurrentY >= Printer.Height - 1250 Then
    Printer.NewPage
    Printer.FontSize = Printer.FontSize
    Printer.Print " "
    Printer.FontSize = Printer.FontSize
    VolgendePagina
    PrintTitel
End If
Exit Sub

PrtHandler5:
MsgBox "Kontroleer de printer."
Resume

End Sub

Private Sub PrintTitel()
Dim T As Integer

On Local Error GoTo PrtHandler1

If chkAfdrukInVenster Then Exit Sub

Printer.Print Tab(1); psTekst(3); vbCrLf;
Printer.Print FULL_LINE; vbCrLf;

Do While REPORT_TAB(T) <> 0
    Printer.Print Tab(REPORT_TAB(T));
    Printer.Print REPORT_FIELD(T);
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        Printer.Print vbCrLf;
    End If
    T = T + 1
Loop

Printer.Print FULL_LINE; vbCrLf;
If Printer.CurrentY >= Printer.Height - 1250 Then
    Printer.NewPage
    Printer.FontSize = Printer.FontSize
    Printer.Print " "
    Printer.FontSize = Printer.FontSize
    VolgendePagina
    PrintTitel
End If
Exit Sub

PrtHandler1:
MsgBox "Kontroleer de printer."
Resume

End Sub

Private Sub PrintVelden()
Dim T As Integer
Dim VeldTekst As String

Do While REPORT_TAB(T) <> 0
    Printer.Print Tab(REPORT_TAB(T));
    Printer.Print VeldTXT(T);
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        Printer.Print vbCrLf;
    End If
    T = T + 1
Loop
If Printer.CurrentY >= Printer.Height - 1250 Then
    Printer.NewPage
    Printer.FontSize = Printer.FontSize
    Printer.Print " "
    Printer.FontSize = Printer.FontSize
    VolgendePagina
    PrintTitel
End If
Exit Sub

End Sub

Private Sub TekstLijn_GotFocus(Index As Integer)

TekstLijn(Index).SelStart = 0
TekstLijn(Index).SelLength = Len(TekstLijn(Index).Text)

Select Case Index
    Case 2, 3
        SnelHelpPrint "[Ctrl] voor geïndexeerd zoeken", BL_LOGGING
End Select

End Sub

Private Sub TekstLijn_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

On Local Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}", True: Exit Sub
Select Case Index
    Case 2, 3
        Select Case KeyCode
            Case 17
                aIndex = 0
                SharedFl = TABLE_LEDGERACCOUNTS
                GridText = TekstLijn(Index).Text
                SqlSearch.Show 1
                If Ktrl = 0 Then
                    TekstLijn(Index).Text = vBibTekst(TABLE_LEDGERACCOUNTS, "#v019 #")
                    SnelHelpPrint vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #"), BL_LOGGING
                End If
        End Select
End Select

End Sub

Private Sub TekstLijn_LostFocus(Index As Integer)

Select Case Index
    Case 1
        If DATE_INVALID((TekstLijn(1).Text)) Then
            Beep
            TekstLijn(1).Text = MIM_GLOBAL_DATE
            TekstLijn(1).SetFocus
        End If
    Case 2
        bGetOrGreater TABLE_LEDGERACCOUNTS, 0, vSet((TekstLijn(2).Text), 7)
        If Ktrl Then
        Else
            RecordToVeld TABLE_LEDGERACCOUNTS
            TekstLijn(2).Text = vBibTekst(TABLE_LEDGERACCOUNTS, "#v019 #")
            SnelHelpPrint vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #"), BL_LOGGING
        End If
End Select

End Sub

Private Sub VolgendePagina()

PAGE_COUNTER = PAGE_COUNTER + 1
Printer.CurrentY = 400
    Printer.Print Tab(1); psTekst(2); Tab(108); "Pagina : " + Dec$((PAGE_COUNTER), "##########"); vbCrLf;
    Printer.Print Tab(108); "Datum  : "; psTekst(0); vbCrLf; vbCrLf;

End Sub

Private Sub TxtPeriodeVanTot_GotFocus()

TxtPeriodeVanTot.SelStart = 0
TxtPeriodeVanTot.SelLength = Len(TxtPeriodeVanTot.Text)

End Sub

Private Sub TxtPeriodeVanTot_LostFocus()
        
If DATE_INVALID(Right(TxtPeriodeVanTot, 10)) Then
    MsgBox "Respecteer : " + vbCrLf + vbCrLf + "DD/MM/EEJJ - DD/MM/EEJJ a.u.b. !"
    TxtPeriodeVanTot = DATE_TEXT(Left(BOOKYEAR_FROMTO, 8)) + " - " + DATE_TEXT(Right(BOOKYEAR_FROMTO, 8))
    SelectieVanTot = BOOKYEAR_FROMTO
    TxtPeriodeVanTot.SetFocus
    Exit Sub
ElseIf Len(TxtPeriodeVanTot) <> 23 Then
    MsgBox "Respecteer : " + vbCrLf + vbCrLf + "DD/MM/EEJJ - DD/MM/EEJJ a.u.b. !"
    TxtPeriodeVanTot = DATE_TEXT(Left(BOOKYEAR_FROMTO, 8)) + " - " + DATE_TEXT(Right(BOOKYEAR_FROMTO, 8))
    TxtPeriodeVanTot.SetFocus
    SelectieVanTot = BOOKYEAR_FROMTO
    Exit Sub
Else
    SelectieVanTot = DATE_KEY(Left(TxtPeriodeVanTot, 10)) + DATE_KEY(Right(TxtPeriodeVanTot, 10))
End If

End Sub
