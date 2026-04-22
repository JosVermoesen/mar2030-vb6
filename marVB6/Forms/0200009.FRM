VERSION 5.00
Begin VB.Form TopDown 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Top Down"
   ClientHeight    =   2805
   ClientLeft      =   1155
   ClientTop       =   3660
   ClientWidth     =   6285
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
   ScaleHeight     =   2805
   ScaleWidth      =   6285
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   4
      Left            =   1920
      TabIndex        =   20
      Top             =   2400
      Width           =   4320
   End
   Begin VB.OptionButton obBoekjaar 
      Caption         =   "Actieve periode"
      Height          =   195
      Index           =   3
      Left            =   60
      TabIndex        =   19
      Top             =   2520
      Width           =   1695
   End
   Begin VB.OptionButton obBoekjaar 
      Caption         =   "2 hoogste boekjaren"
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   18
      Top             =   2040
      Width           =   2415
   End
   Begin VB.OptionButton obBoekjaar 
      Caption         =   "Huidig Boekjaar"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   17
      Top             =   1800
      Value           =   -1  'True
      Width           =   1755
   End
   Begin VB.OptionButton obBoekjaar 
      Caption         =   "Alle jaren"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   16
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CheckBox chkSelektie 
      Caption         =   "&Bestand met scheidingstekens"
      Height          =   240
      Index           =   0
      Left            =   90
      TabIndex        =   15
      Top             =   690
      Width           =   2925
   End
   Begin VB.CheckBox chkSelektie 
      Caption         =   "&Topdown"
      Height          =   240
      Index           =   1
      Left            =   90
      TabIndex        =   14
      Top             =   960
      Width           =   2325
   End
   Begin VB.CheckBox chkSelektie 
      Caption         =   "&Commissie i.p.v. omzet"
      Height          =   240
      Index           =   2
      Left            =   90
      TabIndex        =   13
      Top             =   1230
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.CheckBox chkAfdrukLiggend 
      Alignment       =   1  'Right Justify
      Caption         =   "Liggende Printerafdruk"
      Height          =   255
      Left            =   3960
      TabIndex        =   12
      Top             =   2070
      Width           =   2310
   End
   Begin VB.CheckBox chkAfdrukInVenster 
      Alignment       =   1  'Right Justify
      Caption         =   "Afdruk in venster"
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   1800
      Value           =   1  'Checked
      Width           =   1830
   End
   Begin VB.CheckBox ChkAfpunten 
      Alignment       =   1  'Right Justify
      Caption         =   "&Afpunten"
      Height          =   255
      Left            =   5040
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   240
      Width           =   4320
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   4920
      TabIndex        =   1
      Top             =   240
      Width           =   1260
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   3240
      TabIndex        =   2
      Top             =   1020
      Width           =   1605
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   3240
      TabIndex        =   3
      Top             =   1320
      Width           =   1605
   End
   Begin VB.CommandButton Drukken 
      Caption         =   "Af&drukken"
      Height          =   375
      Left            =   4980
      TabIndex        =   4
      Top             =   720
      Width           =   1275
   End
   Begin VB.CommandButton Annuleren 
      Cancel          =   -1  'True
      Caption         =   "&Sluiten"
      Height          =   375
      Left            =   4980
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1140
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "&SubTitel"
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   30
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "D&atum"
      Height          =   210
      Index           =   1
      Left            =   4920
      TabIndex        =   7
      Top             =   30
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Van"
      Height          =   240
      Index           =   2
      Left            =   3270
      TabIndex        =   8
      Top             =   750
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Tot"
      Height          =   240
      Index           =   3
      Left            =   3810
      TabIndex        =   9
      Top             =   750
      Width           =   480
   End
End
Attribute VB_Name = "TopDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit

Dim psTekst(5) As String        'Koptekstinfo
Dim aa As String

Dim RapportSelektie As String
Dim RapportTitel As String
Dim RapportDefinitie As String
Dim LijstNaam As String
Dim VeldTXT(20) As String
Dim ktrlBoekjaar As String

Dim dttot As Double
Dim dtbtw As Double
Dim dvtot As Double

Private Sub Annuleren_Click()

Unload Me

End Sub

Private Sub chkAfdrukLiggend_Click()

If chkAfdrukLiggend.Value = 1 Then
    Printer.Orientation = vbPRORLandscape
Else
    Printer.Orientation = vbPRORPortrait
End If
XDoEvents = DoEvents

End Sub

Private Sub ChkAfpunten_Click()

If ChkAfpunten.Value = 1 Then
    chkAfdrukInVenster.Value = 1
End If

End Sub





Private Sub Drukken_Click()
Dim BeginSleutel As String * 13
Dim EindSleutel As String * 13
Dim KopBuf As String
Dim XX As Integer
Dim Teller As Integer
Dim FlTabBestand As Integer

Dim DatumString As String

Dim dTOT As Currency
Dim dTnt As Currency
Dim Taal As String

ktrlBoekjaar = BOOKYEAR_FROMTO
If obBoekjaar(2).Value = True Then
    Mid(ktrlBoekjaar, 1, 4) = Format(Val(Mid(ktrlBoekjaar, 1, 4)) - 1, "0000")
End If

Select Case aIndex
    Case 2
        BeginSleutel = "L" + TekstLijn(2).text
        EindSleutel = "L" + TekstLijn(3).text
    Case 1
        BeginSleutel = "K" + TekstLijn(2).text
        EindSleutel = "K" + TekstLijn(3).text
End Select

psTekst(2) = LijstNaam + " " + Mid(Mim.Caption, InStr(Mim.Caption, "["))
psTekst(0) = TekstLijn(1).text
psTekst(3) = TekstLijn(0).text
    InitVelden

bGetOrGreater TABLE_INVOICES, 1, BeginSleutel
If Ktrl Or UCase$(vSet(KEY_BUF(TABLE_INVOICES), 13)) > UCase$(EindSleutel) Then
    Beep
    Exit Sub
ElseIf Left(KEY_BUF(TABLE_INVOICES), 1) <> Left(BeginSleutel, 1) Then
    Beep
    Exit Sub
Else
    KopBuf = Chr(0)
    If chkSelektie(0).Value = 1 Then
        GoSub OpenTabBestand
    End If
    TopDown.Enabled = False
    Me.MousePointer = vbHourglass
    XDoEvents = DoEvents
    
    bClose TABLE_DUMMY
    ClearFlDummy
    Ktrl = bOpen(TABLE_DUMMY)
    
    PAGE_COUNTER = 0
    If chkAfdrukInVenster = 0 Then
        Set Printer = Printers(LijstPrinterNr)
        On Error Resume Next
        Printer.PaperBin = LaadTekst(App.Title, "LIJSTPRINTER")
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
        End If
    End If
    PrintTitel
    RecordToVeld TABLE_INVOICES
    GoSub KontroleVoorwaarden
End If

Do
    bNext TABLE_INVOICES
    DoEvents
    If Ktrl Or UCase$(vSet(KEY_BUF(TABLE_INVOICES), 13)) > UCase$(EindSleutel) Then
        Exit Do
    End If
    RecordToVeld TABLE_INVOICES
    GoSub KontroleVoorwaarden
Loop
    GoSub VoegErBij
    GoSub PrintDeLijst
    EindTotaal
    If chkSelektie(0).Value = 1 Then
        Close FlTabBestand
    End If
    TopDown.Enabled = True
    Me.MousePointer = vbNormal
    If chkAfdrukInVenster Then
    Else
        Printer.NewPage
        Printer.EndDoc
    End If
    Annuleren_Click
Exit Sub

PrintDeLijst:
bFirst TABLE_DUMMY, 0
If Ktrl Then
    Return
Else
    GoSub StelSamen
    Do
        bNext TABLE_DUMMY
        If Ktrl Then
            Exit Do
        Else
            GoSub StelSamen
        End If
    Loop
End If
Return

StelSamen:
RecordToVeld TABLE_DUMMY
bGet aIndex, 0, vBibTekst(TABLE_DUMMY, "#v034 #")
If Ktrl Then
    VeldTXT(0) = "--"
Else
    RecordToVeld aIndex
    VeldTXT(0) = vBibTekst(aIndex, "#A10A #")
End If

Taal = vBibTekst(aIndex, "#A10C #")
If Val(vBibTekst(aIndex, "#A102 #")) = 0 Then
    VeldTXT(1) = Trim$(Trim$(vBibTekst(aIndex, "#A100 #")) + " " + vBibTekst(aIndex, "#A101 #"))
Else
    VeldTXT(1) = Trim$(Mid(fmarBoxText("003", Taal, vBibTekst(aIndex, "#A102 #")), 4, 10)) + " " + Trim$(Trim$(vBibTekst(aIndex, "#A100 #")) + " " + vBibTekst(aIndex, "#A101 #"))
End If
If Val(vBibTekst(aIndex, "#vs01 #")) = 0 Then
    VeldTXT(2) = Trim$(Trim$(vBibTekst(aIndex, "#A125 #")) + " " + vBibTekst(aIndex, "#A127 #"))
Else
    VeldTXT(2) = Trim$(Mid(fmarBoxText("003", Taal, vBibTekst(aIndex, "#vs01 #")), 4, 10)) + " " + Trim$(Trim$(vBibTekst(aIndex, "#A125 #")) + " " + vBibTekst(aIndex, "#A127 #"))
End If
VeldTXT(3) = Trim$(vBibTekst(aIndex, "#A104 #") + " " + vBibTekst(aIndex, "#A105 #") + " " + vBibTekst(aIndex, "#A106 #"))
VeldTXT(4) = Trim$(vBibTekst(aIndex, "#A109 #")) + "-" + Trim$(vBibTekst(aIndex, "#A107 #")) + " " + Trim$(vBibTekst(aIndex, "#A108 #"))
VeldTXT(5) = vBibTekst(TABLE_DUMMY, "#v037 #")
PrintVelden
If chkAfdrukInVenster Then
Else
    Printer.Print vbCrLf;
End If

If chkSelektie(0).Value = 1 Then
    For Teller = 0 To 4
        Print #FlTabBestand, VeldTXT(Teller); ";";
    Next
    Print #FlTabBestand, VeldTXT(5); vbCrLf;
End If
Return

KontroleVoorwaarden:
If obBoekjaar(1).Value = True Then
    'Enkel dit JAAR
    If Left(vBibTekst(TABLE_INVOICES, "#v035 #"), 4) < Left(BOOKYEAR_FROMTO, 4) Then
        Return
    End If
ElseIf obBoekjaar(2).Value = True Then
    'Laatste 2 jaren
    If Left(vBibTekst(TABLE_INVOICES, "#v035 #"), 4) < Left(ktrlBoekjaar, 4) Then
        Return
    End If
ElseIf obBoekjaar(3).Value = True Then
    'een bepaalde maand in hoogste boekjaar!
    If Mid(vBibTekst(TABLE_INVOICES, "#v035 #"), 1, 6) >= Mid(PERIOD_FROMTO, 1, 6) And Mid(vBibTekst(TABLE_INVOICES, "#v035 #"), 1, 6) <= Mid(PERIOD_FROMTO, 9, 6) Then
        'Stop
    Else
        Return
    End If
End If

dTOT = 0
Select Case aIndex
    Case TABLE_CUSTOMERS
        Select Case Left(vBibTekst(TABLE_INVOICES, "#v033 #"), 1)
            Case "V"
                'For Teller = 55 To 63
                    'dTot = dTot + Val(vBibTekst(TABLE_INVOICES, "#v" + Format(Teller, "000") + " #"))
                'Next
                'dTot = dTot + Val(vBibTekst(TABLE_INVOICES, "#v089 #"))
                dTOT = dTOT + Val(vBibTekst(TABLE_INVOICES, "#v249 #"))
            Case "Q"
                dTOT = dTOT + Val(vBibTekst(TABLE_INVOICES, "#v249 #"))
                'If chkSelektie(2).Value = 0 Then
                '    dTot = Val(vBibTekst(TABLE_INVOICES, "#B010 #")) - Val(vBibTekst(TABLE_INVOICES, "#B090 #"))
                'Else
                '    dTot = Val(vBibTekst(TABLE_INVOICES, "#B014 #")) - Val(vBibTekst(TABLE_INVOICES, "#B094 #"))
                'End If
        End Select
    Case TABLE_SUPPLIERS
        dTOT = dTOT + Val(vBibTekst(TABLE_INVOICES, "#v249 #"))
        'dTot = 0
        'For Teller = 46 To 49
        '    dTot = dTot + Val(vBibTekst(TABLE_INVOICES, "#v" + Format(Teller, "000") + " #"))
        'Next
End Select

If Mid(vBibTekst(TABLE_INVOICES, "#v033 #"), 2, 1) = "1" Then
    dTOT = -dTOT
End If

If Trim$(vBibTekst(TABLE_INVOICES, "#v034 #")) <> Trim$(KopBuf) Then
    If KopBuf = Chr$(0) Then
    Else
        GoSub VoegErBij
    End If
    dTnt = 0
    bGet aIndex, 0, vSet(Mid(vBibTekst(TABLE_INVOICES, "#v034 #"), 2, 12), 12)
    KopBuf = vBibTekst(TABLE_INVOICES, "#v034 #")
    TLB_RECORD(TABLE_DUMMY) = ""
    '0
    vBib TABLE_DUMMY, Mid(vBibTekst(TABLE_INVOICES, "#v034 #"), 2, 12), "v034"
    If Ktrl Then
        vBib TABLE_DUMMY, "* niet meer aanwezig *", "A100"
    Else
        RecordToVeld aIndex
        '1
        vBib TABLE_DUMMY, Left(vBibTekst(aIndex, "#A100 #"), 35), "A100"
        '2
        vBib TABLE_DUMMY, Left(RTrim$(RTrim$(vBibTekst(aIndex, "#A104 #")) + " " + RTrim$(vBibTekst(aIndex, "#A105 #")) + " " + RTrim$(vBibTekst(aIndex, "#A106 #"))), 30), "A104"
        '3
        vBib TABLE_DUMMY, RTrim$(vBibTekst(aIndex, "#A109 #")) + "-" + RTrim$(vBibTekst(aIndex, "#A107 #")), "A109"
        '4
        vBib TABLE_DUMMY, Left(RTrim$(vBibTekst(aIndex, "#A108 #")), 30), "A108"
    End If
End If
dTnt = dTnt + dTOT
Return

VoegErBij:
If dTnt = 0 Then
    Return
Else
    vBib TABLE_DUMMY, Dec$((dTnt), MASK_EURBH), "v037"
    If chkSelektie(1).Value = False Then
        'v089 = sorteren op kodenummer
        vBib TABLE_DUMMY, vBibTekst(TABLE_DUMMY, "#v034 #"), "v089"
    Else
        'v089 = sorteren op jaaromzet
        vBib TABLE_DUMMY, Dec$(999999999 - dTnt, "000000000000"), "v089"
    End If
    SnelHelpPrint FVT(TABLE_INVOICES, 1), BL_LOGGING
    bInsert TABLE_DUMMY, 0
    dttot = dttot + dTnt
End If
Return

OpenTabBestand:
FlTabBestand = FreeFile
Open PROGRAM_LOCATION & "NTExport.TXT" For Output As FlTabBestand
Msg = "Bestand met scheidingstekens wordt aangemaakt als" & vbCr & vbCr
Msg = Msg & PROGRAM_LOCATION + "NTExport.TXT"
MsgBox Msg, vbInformation
Return

End Sub

Private Sub EindTotaal()
Dim T As Integer

For T = 0 To 5
    VeldTXT(T) = ""
Next
VeldTXT(0) = "Totaal"
VeldTXT(5) = Dec$(dttot, MASK_SY(0))

T = 0
aa = ""
Do While REPORT_TAB(T) <> 0
    If chkAfdrukInVenster Then
        aa = aa & VeldTXT(T) & vbTab
    Else
        Printer.Print ; Tab(REPORT_TAB(T));
        Printer.Print ; VeldTXT(T);
    End If
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        If chkAfdrukInVenster Then
        Else
            Printer.Print vbCrLf;
        End If
    End If
    T = T + 1
Loop

If chkAfdrukInVenster Then
    Xlog.X.AddItem aa, Xlog.X.Rows - 1
    Xlog.X.ColWidth(0) = 1080
    Xlog.X.ColWidth(1) = 1950
    Xlog.X.ColWidth(2) = 1860
    Xlog.X.ColWidth(3) = 1890
    Xlog.X.ColWidth(4) = 1755
    Xlog.X.ColWidth(5) = 900
    Xlog.X.ColWidth(6) = 555
    Xlog.WindowState = 2

    Xlog.X.ColAlignment(5) = flexAlignRightTop
    
    'XLog.X.GridLines = False
    Xlog.Afsluiten.Enabled = False
    Xlog.WijzigenLijn.Visible = False
    Xlog.Afsluiten.TabStop = False
    Xlog.cbAfbeelding.Visible = False
    XLogKey = ""
    Xlog.SSTab1.TabVisible(1) = False
    Xlog.Show 1
    Xlog.WindowState = 0
    Unload Xlog
End If
Exit Sub

PrtHandler4:
MsgBox "Kontroleer de printer."
Resume

End Sub

Private Sub Form_Load()

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If

Dim TempB_MODUS As Integer

dttot = 0
dtbtw = 0
dvtot = 0
TekstLijn(1).text = MIM_GLOBAL_DATE

Select Case aIndex
    Case TABLE_CUSTOMERS
        LijstNaam = "TopDown Klanten"
    Case TABLE_SUPPLIERS
        LijstNaam = "TopDown Leveranciers"
End Select

TopDown.Caption = LijstNaam
TekstLijn(2).text = "0"
TekstLijn(3).text = String$(12, "z")
If Mim.AV(2).Enabled Then
    chkSelektie(2).Visible = True
End If

Set Printer = Printers(LijstPrinterNr)
On Error Resume Next
Printer.PaperBin = LaadTekst(App.Title, "LIJSTPRINTER")
If Printer.Orientation = vbPRORLandscape Then
    chkAfdrukLiggend.Value = 1
Else
    chkAfdrukLiggend.Value = 0
End If
Me.TekstLijn(4) = BYPERDAT.PeriodeBoekjaar

End Sub

Private Sub InitVelden()
Dim T As Integer
Dim VolgTab As Integer

REPORT_FIELD(0) = "Telefoon"
    REPORT_TAB(0) = 1

REPORT_FIELD(1) = "Naam1"
    REPORT_TAB(1) = 14

REPORT_FIELD(2) = "Naam2"
    REPORT_TAB(2) = 50

REPORT_FIELD(3) = "Straat en Nr"
    REPORT_TAB(3) = 81

REPORT_FIELD(4) = "Plaats"
    REPORT_TAB(4) = 93

If chkSelektie(2).Value = 0 Then
    REPORT_FIELD(5) = "Omzet incl"
Else
    REPORT_FIELD(5) = "Commissie"
End If
    REPORT_TAB(5) = 118

REPORT_TAB(6) = 0

If chkAfdrukInVenster Then
    Me.Hide
    Unload Xlog
    Xlog.Hide
    Xlog.Caption = "Jaaromzet"
    Xlog.X.Cols = 7
    Xlog.X.Row = 0
    For T = 0 To 5
        Xlog.X.Col = T
        Xlog.X.text = REPORT_FIELD(T)
    Next
    Me.Show
End If

End Sub

Private Sub PrintTitel()
Dim T As Integer

If chkAfdrukInVenster Then Exit Sub
    
      If usrLicentieInfo <> "" Then
        Printer.CurrentX = 50
        Printer.CurrentY = 50
        Printer.Print usrLicentieInfo;
    End If
    PAGE_COUNTER = PAGE_COUNTER + 1
    Printer.CurrentY = 400
    Printer.Print Tab(1); psTekst(2);
    Printer.Print Tab(108); "Pagina : " + Dec$((PAGE_COUNTER), "##########"); vbCrLf;
    Printer.Print Tab(108); "Datum  : " + psTekst(0); vbCrLf; vbCrLf;
    Printer.Print Tab(1); UCase$(psTekst(3)); vbCrLf;

Printer.Print Tab(1); FULL_LINE; vbCrLf;

Do While REPORT_TAB(T) <> 0
    Printer.Print Tab(REPORT_TAB(T)); REPORT_FIELD(T);
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        Printer.Print vbCrLf;
    End If
    T = T + 1
Loop
Printer.Print Tab(1); FULL_LINE; vbCrLf; vbCrLf;

End Sub

Private Sub PrintVelden()
Dim T As Integer
Dim VeldTekst As String

T = 0
aa = ""
Do While REPORT_TAB(T) <> 0
    If chkAfdrukInVenster Then
        aa = aa & VeldTXT(T) & vbTab
    Else
        Printer.Print Tab(REPORT_TAB(T));
        Printer.Print VeldTXT(T);
    End If
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        If chkAfdrukInVenster Then
        Else
            Printer.Print vbCrLf;
        End If
    End If
    T = T + 1
Loop
If chkAfdrukInVenster Then
    Xlog.X.AddItem aa, Xlog.X.Rows - 1
ElseIf Printer.CurrentY >= Printer.Height - 1200 Then
    Printer.NewPage
    Printer.FontSize = Printer.FontSize
    Printer.Print " "
    Printer.FontSize = Printer.FontSize
    PrintTitel
End If

End Sub

Private Sub TekstLijn_GotFocus(Index As Integer)

TekstLijn(Index).SelLength = Len(TekstLijn(Index).text)

End Sub

Private Sub TekstLijn_LostFocus(Index As Integer)

Select Case Index
    Case 1
        If DATE_INVALID((TekstLijn(1).text)) Then
            Beep
            TekstLijn(1).text = MIM_GLOBAL_DATE
        End If
End Select

End Sub

