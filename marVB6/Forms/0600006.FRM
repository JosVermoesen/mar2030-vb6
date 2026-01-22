VERSION 5.00
Begin VB.Form FinancieelBoek 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Financiëele Boeken"
   ClientHeight    =   4980
   ClientLeft      =   1305
   ClientTop       =   1365
   ClientWidth     =   7035
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
   ScaleHeight     =   4980
   ScaleWidth      =   7035
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCKTRL 
      Caption         =   "cKTRL"
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
      Left            =   4440
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4560
      Width           =   915
   End
   Begin VB.CheckBox chkAfdrukLiggend 
      Alignment       =   1  'Right Justify
      Caption         =   "Liggende Printerafdruk"
      Height          =   255
      Left            =   1950
      TabIndex        =   11
      Top             =   4620
      Width           =   2310
   End
   Begin VB.CheckBox chkAfdrukInVenster 
      Caption         =   "Afdruk in venster"
      Height          =   255
      Left            =   90
      TabIndex        =   10
      Top             =   4620
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.CommandButton CmdJournaalManueel 
      Caption         =   "&Journalen Manueel Zoeken"
      Height          =   705
      Left            =   5520
      TabIndex        =   9
      Top             =   540
      Width           =   1440
   End
   Begin VB.ListBox UittrekselsLijst 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   90
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   6855
   End
   Begin VB.ComboBox KeuzeInfo 
      BackColor       =   &H00FFFFFF&
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
      Index           =   0
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   3885
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   5
      Top             =   360
      Width           =   3555
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   3660
      TabIndex        =   7
      Top             =   360
      Width           =   1560
   End
   Begin VB.CommandButton Drukken 
      Caption         =   "Af&drukken"
      Height          =   435
      Left            =   5520
      TabIndex        =   3
      Top             =   60
      Width           =   1440
   End
   Begin VB.CommandButton Annuleren 
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
      Left            =   5460
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4530
      Width           =   1440
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Uittreksels periode"
      Height          =   435
      Index           =   2
      Left            =   60
      TabIndex        =   0
      Top             =   840
      Width           =   1065
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Afdrukperiode Van - &Tot"
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   2190
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Datu&m Drukken"
      Height          =   270
      Index           =   1
      Left            =   3690
      TabIndex        =   6
      Top             =   60
      Width           =   1395
   End
End
Attribute VB_Name = "FinancieelBoek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit

Dim psTekst(5)          As String
Dim TotaalD             As Currency
Dim TotaalC             As Currency

Dim LFontSize(20)       As Single
Dim LAantalL(20)        As Integer
Dim FontDefChanged      As Integer

Dim VeldTXT(17)         As String

Dim RecNummer(9)        As Integer
Dim RekeningNummer(9)   As String * 7

Dim PlGrensVan          As String * 8
Dim PlGrensTot          As String * 8

Dim ErrorMsg            As String
Dim JournaalManueelVlag As Boolean

Sub DetailFinancieelStuk(DeString As String)
Dim A As String

Unload InfoScherm
Load InfoScherm
InfoScherm.Caption = InfoScherm.Caption + ", Journaaldetail voor uittreksel : " + DeString
InfoScherm.X.Rows = 2
InfoScherm.X.Cols = 8
InfoScherm.X.Col = 0
InfoScherm.X.Row = 0
InfoScherm.X.Text = "Datum (v066)"
InfoScherm.X.ColAlignment(0) = flexAlignLeftCenter
InfoScherm.X.Col = 1
InfoScherm.X.Text = "Rekening (v019)"
InfoScherm.X.ColAlignment(1) = flexAlignLeftCenter
InfoScherm.X.Col = 2
InfoScherm.X.Text = "Naam (v020)"
InfoScherm.X.ColAlignment(2) = flexAlignLeftCenter
InfoScherm.X.Col = 3
InfoScherm.X.Text = "Bedrag (dece068)"
InfoScherm.X.ColAlignment(3) = flexAlignGeneral
InfoScherm.X.Col = 4
InfoScherm.X.Text = "Boekingsomschrijving (v067)"
InfoScherm.X.ColAlignment(4) = flexAlignLeftCenter
InfoScherm.X.Col = 5
InfoScherm.X.Text = "AV Dokum. (v033)"
InfoScherm.X.ColAlignment(5) = flexAlignLeftCenter
InfoScherm.X.Col = 6
InfoScherm.X.Text = "TegenRek. (v069)"
InfoScherm.X.ColAlignment(6) = flexAlignLeftCenter
InfoScherm.X.Col = 7
InfoScherm.X.Text = "vsfRecord"

bGet TABLE_JOURNAL, 2, vSet(DeString, 8)
If Ktrl Then
    Beep
    MsgBox "Geen journaallijnen voor " + DeString
    Exit Sub
Else
    Screen.MousePointer = vbHourglass
    GoSub VolgendeLijn
    Do
        bNext TABLE_JOURNAL
        If Ktrl Or KEY_BUF(TABLE_JOURNAL) <> DeString Then
            Exit Do
        Else
            GoSub VolgendeLijn
        End If
    Loop
End If
Screen.MousePointer = vbNormal

OpHetScherm:
InfoScherm.X.Row = 1
InfoScherm.X.Col = 0

InfoScherm.X.ColWidth(0) = 1005
InfoScherm.X.ColWidth(1) = 765
InfoScherm.X.ColWidth(2) = 2160
InfoScherm.X.ColWidth(3) = 750
InfoScherm.X.ColWidth(4) = 1950
InfoScherm.X.ColWidth(5) = 1150
InfoScherm.X.ColWidth(6) = 870
InfoScherm.X.ColWidth(7) = 6000
InfoScherm.X.ColAlignment(0) = flexAlignLeftCenter
InfoScherm.X.ColAlignment(1) = flexAlignLeftCenter
InfoScherm.X.ColAlignment(2) = flexAlignLeftCenter
InfoScherm.X.ColAlignment(3) = flexAlignGeneral
InfoScherm.X.ColAlignment(4) = flexAlignLeftCenter
InfoScherm.X.ColAlignment(5) = flexAlignLeftCenter
InfoScherm.X.ColAlignment(6) = flexAlignLeftCenter
InfoScherm.X.ColAlignment(7) = flexAlignLeftCenter

XLogShow:
InfoScherm.Show 1
Unload InfoScherm
Exit Sub

VolgendeLijn:
RecordToVeld TABLE_JOURNAL
If vSet(vBibTekst(TABLE_JOURNAL, "#v019 #"), 7) = Left(KeuzeInfo(0).Text, 7) Then
    Return
ElseIf DATE_TEXT(vBibTekst(TABLE_JOURNAL, "#v066 #")) <> Mid(UittrekselsLijst.Text, 12, 10) Then
    If Not JournaalManueelVlag Then
        Return
    End If
ElseIf RTrim$(vBibTekst(TABLE_JOURNAL, "#v019 #")) <> "" Then
    If vSet(vBibTekst(TABLE_JOURNAL, "#v069 #"), 7) <> Left(KeuzeInfo(0).Text, 7) Then
        Return
    End If
End If
A = DATE_TEXT(vBibTekst(TABLE_JOURNAL, "#v066 #")) & vbTab
A = A & vBibTekst(TABLE_JOURNAL, "#v019 #") & vbTab
bGet TABLE_LEDGERACCOUNTS, 0, vSet(vBibTekst(TABLE_JOURNAL, "#v019 #"), 7)
If Ktrl Then
    A = A & "//" & vbTab
Else
    RecordToVeld TABLE_LEDGERACCOUNTS
    A = A & vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #") & vbTab
End If
A = A & rsMAR(TABLE_JOURNAL)("v068") & vbTab
A = A & rsMAR(TABLE_JOURNAL)("v067") & vbTab
A = A & rsMAR(TABLE_JOURNAL)("v033") & vbTab
A = A & rsMAR(TABLE_JOURNAL)("v069") & vbTab
A = A & TLB_RECORD(TABLE_JOURNAL)
InfoScherm.X.AddItem A, InfoScherm.X.Rows - 1
Return

End Sub

Private Sub Annuleren_Click()

Unload Me

End Sub

Private Sub CumulPrint()
Dim RekeningNaam As String * 30
Dim Tabul As Integer

On Local Error GoTo PrtHandler

If chkAfdrukInVenster Then
Else
    On Local Error Resume Next
    Printer.EndDoc
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
    Err = 0
    On Local Error Resume Next
    If Err - 480 > 0 Then Exit Sub
    PrintTitel
    PrintTotaal
End If
Printer.Print
Printer.Print Tab(2); "** CENTRALISATIE/VIERKANTSCONTROLE **"
Printer.Print ""

bClose TABLE_DUMMY
bFirst TABLE_DUMMY, 0
RecordToVeld TABLE_DUMMY
bGet TABLE_LEDGERACCOUNTS, 0, Left(FVT(TABLE_DUMMY, 0), 7)
If Ktrl Then
    RekeningNaam = "Rekening reeds vernietigd !!!"
Else
    RecordToVeld TABLE_LEDGERACCOUNTS
    RekeningNaam = vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #")
End If
Tabul = 0
Printer.Print Tab(Tabul + 2); Dec$(Val(vBibTekst(TABLE_DUMMY, "#v013 #")), "####"); " x "; vSet(vBibTekst(TABLE_DUMMY, "#v089 #"), 7); " "; RekeningNaam; " "; Dec$(Val(vBibTekst(TABLE_DUMMY, "#v068 #")), MASK_EURBH);

Do
    bNext TABLE_DUMMY
    If Ktrl Then
        Exit Do
    End If
    RecordToVeld TABLE_DUMMY
    bGet TABLE_LEDGERACCOUNTS, 0, Left(FVT(TABLE_DUMMY, 0), 7)
    If Ktrl Then
        RekeningNaam = "Rekening reeds vernietigd !!!"
    Else
        RecordToVeld TABLE_LEDGERACCOUNTS
        RekeningNaam = vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #")
    End If
    If Tabul = 0 Then
        Tabul = 59
        Printer.Print Tab(Tabul + 2); Dec$(Val(vBibTekst(TABLE_DUMMY, "#v013 #")), "####"); " x "; vSet(vBibTekst(TABLE_DUMMY, "#v089 #"), 7); " "; RekeningNaam; " "; Dec$(Val(vBibTekst(TABLE_DUMMY, "#v068 #")), MASK_EURBH); vbCrLf;
    Else
        Tabul = 0
        Printer.Print Tab(Tabul + 2); Dec$(Val(vBibTekst(TABLE_DUMMY, "#v013 #")), "####"); " x "; vSet(vBibTekst(TABLE_DUMMY, "#v089 #"), 7); " "; RekeningNaam; " "; Dec$(Val(vBibTekst(TABLE_DUMMY, "#v068 #")), MASK_EURBH);
    End If
Loop
Exit Sub

PrtHandler:
MsgBox "Kontroleer de printer."
Resume

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

   
Private Sub CmdJournaalManueel_Click()
Dim KtrlInput As String

InputVanFinancieelStuk:
KtrlInput = InputBox("Geef nummer van het financieel stuk (8 tekens !)", "Financieel journaal kontrole ")
If KtrlInput = "" Then
    Exit Sub
ElseIf Len(KtrlInput) <> 8 Then
    Msg = "dokumentnummer bestaat uit 8 tekens !" + vbCrLf + vbCrLf
    Msg = Msg + "Voorbeeld:" + vbCrLf
    Msg = Msg + "Rekening 'GB', werkelijk jaar 19'98'," + vbCrLf
    Msg = Msg + "uittreksel 124, geeft als dokumentnummer:" + vbCrLf + vbCrLf
    Msg = Msg + "GB980124"
    MsgBox Msg
    Exit Sub
End If
JournaalManueelVlag = True
DetailFinancieelStuk KtrlInput
JournaalManueelVlag = False

End Sub

Private Sub Drukken_Click()
Dim bModDummy           As Integer
Dim TelUittreksel       As Integer
Dim DCBedrag            As Currency
Dim HetBedrag           As Currency
Dim XX                  As Integer

Dim FinRekening         As String * 7
Dim DeRekening          As String * 7

On Local Error GoTo PrtHandler9

psTekst(2) = "Financiëel Boek " + Mid(Mim.Caption, InStr(Mim.Caption, "["))
psTekst(0) = TekstLijn(1).Text
psTekst(3) = Right(KeuzeInfo(0).Text, 40) + " " + TekstLijn(0).Text

ErrorMsg = ""
TotaalD = 0
TotaalC = 0
PAGE_COUNTER = 0

If UittrekselsLijst.ListCount = 0 Then
    Exit Sub
Else
    InitVelden
    bClose TABLE_DUMMY
    ClearFlDummy
    Ktrl = bOpen(TABLE_DUMMY)

    Screen.MousePointer = vbHourglass
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
    PrintTitel
End If

For TelUittreksel = 0 To UittrekselsLijst.ListCount - 1
    UittrekselsLijst.ListIndex = TelUittreksel
    VeldTXT(0) = Mid(UittrekselsLijst.Text, 12, 10)
    VeldTXT(1) = Left(KeuzeInfo(0).Text, 7)
    FinRekening = VeldTXT(1)
    VeldTXT(2) = Mid(UittrekselsLijst.Text, 25, 30)
    VeldTXT(3) = "DS/CS Saldo van het uittreksel"
    
    VeldTXT(4) = Mid(UittrekselsLijst.Text, 56, 12)
    VeldTXT(5) = Mid(UittrekselsLijst.Text, 69, 12)
    VeldTXT(6) = Left(UittrekselsLijst.Text, 8)
    DeRekening = VeldTXT(1)
    HetBedrag = -Val(VeldTXT(5)) + Val(VeldTXT(4))
    Select Case HetBedrag
        Case Is < 0
            TotaalC = TotaalC + HetBedrag
        Case Else
            TotaalD = TotaalD + HetBedrag
    End Select
    GoSub DetailCumul
    PrintVelden
    If TelUittreksel = UittrekselsLijst.ListCount - 1 Then
        GoSub DetailPrintJournaal
    ElseIf Left(UittrekselsLijst.List(TelUittreksel), 21) = Left(UittrekselsLijst.List(TelUittreksel + 1), 21) Then
        ErrorMsg = ErrorMsg + Left(UittrekselsLijst.List(TelUittreksel), 21) + " / " + Left(UittrekselsLijst.List(TelUittreksel + 1), 21) + " onlogica." + vbCrLf
    Else
        GoSub DetailPrintJournaal
    End If
Next
PrintTotaal
CumulPrint
If VeldTXT(4) = VeldTXT(5) Then
    If ErrorMsg <> "" Then
        ErrorMsg = vbCrLf + vbCrLf + "Er zijn uittreksel onjuist ingebracht.  De boekhouding blijft gelukkig correct" + vbCrLf + vbCrLf + ErrorMsg
    End If
Else
    ErrorMsg = ErrorMsg + "DEBET<>CREDIT : de boekhoudDATABASE instabiel ?  Dient geRESTORED te worden !!!" + vbCrLf + vbCrLf
End If
    
If ErrorMsg <> "" Then
    Printer.Print vbCrLf; vbCrLf;
    Printer.Print ErrorMsg; vbCrLf; vbCrLf;
    Printer.Print "De gebruiker doet er goed aan steeds te kontroleren :"; vbCrLf;
    Printer.Print "* Eén Bankuittreksels pér UITTREKSELDATUM !"; vbCrLf;
    Printer.Print "* Steeds datum kontroleren vooraleer weg te schrijven !!!"; vbCrLf; vbCrLf;
    Printer.Print "Indien DEBET=CREDIT is de boekhouding toch correct bijgewerkt en hoeft U voor huidige onregelmatigheden"
    Printer.Print "niets recht te zetten"
End If
Printer.EndDoc
Screen.MousePointer = vbNormal
Annuleren_Click
Exit Sub

DetailPrintJournaal:
bGet TABLE_JOURNAL, 2, Left(UittrekselsLijst.Text, 8)
If Ktrl Or KEY_BUF(TABLE_JOURNAL) <> Left(UittrekselsLijst.Text, 8) Then
    Beep
    MsgBox "onlogische situatie"
    Exit Sub
Else
    GoSub PrintInfo
End If
Do
    bNext TABLE_JOURNAL
    If Ktrl Or KEY_BUF(TABLE_JOURNAL) <> Left(UittrekselsLijst.Text, 8) Then
        Exit Do
    Else
        GoSub PrintInfo
    End If
Loop
Printer.Print vbCrLf;
Return

PrintInfo:
RecordToVeld TABLE_JOURNAL
If vSet(vBibTekst(TABLE_JOURNAL, "#v019 #"), 7) = Left(KeuzeInfo(0).Text, 7) Then
    Return
ElseIf DATE_TEXT(vBibTekst(TABLE_JOURNAL, "#v066 #")) <> Mid(UittrekselsLijst.Text, 12, 10) Then
    Return
ElseIf RTrim$(vBibTekst(TABLE_JOURNAL, "#v019 #")) <> "" Then
    If vSet(vBibTekst(TABLE_JOURNAL, "#v069 #"), 7) <> Left(KeuzeInfo(0).Text, 7) Then
        Return
    End If
End If

VeldTXT(0) = ""
VeldTXT(1) = vBibTekst(TABLE_JOURNAL, "#v019 #")
bGet TABLE_LEDGERACCOUNTS, 0, vSet(VeldTXT(1), 7)
If Ktrl Then
    VeldTXT(2) = "Reeds vernietigd..."
Else
    RecordToVeld TABLE_LEDGERACCOUNTS
    VeldTXT(2) = vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #")
End If
VeldTXT(3) = vBibTekst(TABLE_JOURNAL, "#v067 #")
DCBedrag = Val(vBibTekst(TABLE_JOURNAL, "#v068 #"))
Select Case DCBedrag
    Case Is < 0
        TotaalC = TotaalC + DCBedrag
        VeldTXT(4) = ""
        VeldTXT(5) = Dec$(Abs(DCBedrag), MASK_EURBH)
    Case Else
        TotaalD = TotaalD + DCBedrag
        VeldTXT(4) = Dec$((DCBedrag), MASK_EURBH)
        VeldTXT(5) = ""
End Select
VeldTXT(6) = vBibTekst(TABLE_JOURNAL, "#v033 #")
DeRekening = VeldTXT(1)
HetBedrag = DCBedrag
GoSub DetailCumul
PrintVelden
Return

PrtHandler9:
MsgBox "Kontroleer de printer."
Printer.KillDoc
Unload Me
Exit Sub
Resume

DetailCumul:
StartPunt:
bGet TABLE_DUMMY, 0, vSet(DeRekening, 20)
If Ktrl Then
    TLB_RECORD(TABLE_DUMMY) = ""
    vBib TABLE_DUMMY, DeRekening, "v089"
    vBib TABLE_DUMMY, "0", "v013"
    vBib TABLE_DUMMY, "0", "v068"
    bInsert TABLE_DUMMY, 0
    GoTo StartPunt
Else
    RecordToVeld TABLE_DUMMY
    vBib TABLE_DUMMY, Str$(Val(vBibTekst(TABLE_DUMMY, "#v013 #")) + 1), "v013"
    vBib TABLE_DUMMY, Str$(Val(vBibTekst(TABLE_DUMMY, "#v068 #")) + HetBedrag), "v068"
    bUpdate TABLE_DUMMY, 0
End If
Return

End Sub

Private Sub Form_Load()

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If

Dim T As Integer
Dim A As String
Dim TempB_MODUS As Integer

JournaalManueelVlag = False
Set Printer = Printers(LijstPrinterNr)
On Error Resume Next
Printer.PaperBin = LaadTekst(App.Title, "LIJSTPRINTER")
If Printer.Orientation = vbPRORLandscape Then
    chkAfdrukLiggend.Value = 1
Else
    chkAfdrukLiggend.Value = 0
End If
chkAfdrukLiggend_Click

TekstLijn(0).Text = DATE_TEXT(Left(BOOKYEAR_FROMTO, 8)) + " - " + DATE_TEXT(Right(BOOKYEAR_FROMTO, 8))
TekstLijn(1).Text = MIM_GLOBAL_DATE
PlGrensVan = Left(BOOKYEAR_FROMTO, 8)
PlGrensTot = Right(BOOKYEAR_FROMTO, 8)

RecNummer(0) = 31
RekeningNummer(0) = String99(READING, 41)

RecNummer(1) = 32
RekeningNummer(1) = String99(READING, 42)

RecNummer(2) = 33
RekeningNummer(2) = String99(READING, 43)

RecNummer(3) = 34
RekeningNummer(3) = String99(READING, 44)

RecNummer(4) = 35
RekeningNummer(4) = String99(READING, 45)

RecNummer(5) = 38
RekeningNummer(5) = String99(READING, 39)

RecNummer(6) = 215
RekeningNummer(6) = String99(READING, 211)

RecNummer(7) = 216
RekeningNummer(7) = String99(READING, 212)

RecNummer(8) = 217
RekeningNummer(8) = String99(READING, 213)

RecNummer(9) = 218
RekeningNummer(9) = String99(READING, 214)

bClose TABLE_JOURNAL
For T = 0 To 9
    bGet TABLE_LEDGERACCOUNTS, 0, RekeningNummer(T)
    If Ktrl Then
        A = RekeningNummer(T) + Chr$(124) + "Niet aanwezig. Setup Boekjaar!"
    Else
        RecordToVeld TABLE_LEDGERACCOUNTS
        A = RekeningNummer(T) + Chr$(124) + RTrim$(vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #"))
    End If
    bFirst TABLE_JOURNAL, 0
    bGetOrGreater TABLE_JOURNAL, 0, RekeningNummer(T) + PlGrensVan
    If Ktrl Then
    Else
        If KEY_BUF(TABLE_JOURNAL) > RekeningNummer(T) + PlGrensTot Then
        Else
            KeuzeInfo(0).AddItem A$
        End If
    End If
Next
If KeuzeInfo(0).ListCount Then KeuzeInfo(0).ListIndex = 0

End Sub


Private Sub InitVelden()

REPORT_FIELD(0) = "Datum"
    REPORT_TAB(0) = 2

REPORT_FIELD(1) = "Rek.Nm."
    REPORT_TAB(1) = 13

REPORT_FIELD(2) = "Naam/Omschrijving"
    REPORT_TAB(2) = 21

REPORT_FIELD(3) = "Betreft"
    REPORT_TAB(3) = 62

REPORT_FIELD(4) = "       Debet"
    REPORT_TAB(4) = 93

REPORT_FIELD(5) = "      Credit"
    REPORT_TAB(5) = 105

REPORT_FIELD(6) = "dokument"
    REPORT_TAB(6) = 118

REPORT_TAB(7) = 0

End Sub

Private Sub KeuzeInfo_Click(Index As Integer)
Dim DummyText   As String
Dim DCBedrag    As Double

Screen.MousePointer = vbHourglass
Select Case Index
    Case 0
        UittrekselsLijst.Visible = False
        UittrekselsLijst.Clear
        bFirst TABLE_JOURNAL, 0
        bGetOrGreater TABLE_JOURNAL, 0, Left(KeuzeInfo(0).Text, 7) + PlGrensVan
        If Ktrl Or KEY_BUF(TABLE_JOURNAL) > Left(KeuzeInfo(0).Text, 7) + PlGrensTot Then
            MsgBox "onlogische situatie"
            GoTo EindeClick
        Else
            GoSub VoegErBij
        End If

        Do
            bNext TABLE_JOURNAL
            If Ktrl Or KEY_BUF(TABLE_JOURNAL) > Left(KeuzeInfo(0).Text, 7) + PlGrensTot Then
                Exit Do
            Else
                GoSub VoegErBij
            End If
        Loop
        If UittrekselsLijst.ListCount Then UittrekselsLijst.ListIndex = 0
        UittrekselsLijst.Visible = True
    Case Else
End Select

EindeClick:
Screen.MousePointer = vbNormal
On Local Error Resume Next
'UittrekselsLijst.SetFocus
Exit Sub

VoegErBij:
RecordToVeld TABLE_JOURNAL
If Len(vBibTekst(TABLE_JOURNAL, "#v038 #")) <> 8 Then
    Return
    'DummyText = vBibTekst(TABLE_JOURNAL, "#v033 #") + " " + Chr$(124) + " "
Else
    DummyText = vBibTekst(TABLE_JOURNAL, "#v038 #") + " " + Chr$(124) + " "
End If
DummyText = DummyText + DATE_TEXT(vBibTekst(TABLE_JOURNAL, "#v066 #")) + " " + Chr$(124) + " "
DummyText = DummyText + vSet(vBibTekst(TABLE_JOURNAL, "#v067 #"), 30) + Chr$(124)
DCBedrag = Val(vBibTekst(TABLE_JOURNAL, "#v068 #"))
Select Case DCBedrag
    Case Is < 0
        DummyText = DummyText + Dec$(0, MASK_EURBH) + Chr$(124) + Dec$(Abs(DCBedrag), MASK_EURBH)
    Case Else
        DummyText = DummyText + Dec$(Abs(DCBedrag), MASK_EURBH) + Chr$(124) + Dec$(0, MASK_EURBH)
End Select
UittrekselsLijst.AddItem DummyText
Return

End Sub

Private Sub PrintTitel()
Dim T As Integer

On Local Error GoTo PrtHandler3

  If usrLicentieInfo <> "" Then
        Printer.CurrentX = 50
        Printer.CurrentY = 50
        Printer.Print usrLicentieInfo;
    End If
    PAGE_COUNTER = PAGE_COUNTER + 1
    Printer.CurrentY = 400
    Printer.Print Tab(1); psTekst(2); Tab(108); "Pagina : " + Dec$((PAGE_COUNTER), "##########");
    Printer.Print Tab(108); "Datum  : "; psTekst(0); vbCrLf$; vbCrLf$;
    Printer.Print Tab(1); UCase$(psTekst(3));

Printer.Print vbCrLf; FULL_LINE

Do While REPORT_TAB(T) <> 0
    Printer.Print Tab(REPORT_TAB(T));
    Printer.Print REPORT_FIELD(T);
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        Printer.Print vbCrLf;
    End If
    T = T + 1
Loop

Printer.Print FULL_LINE; vbCrLf; vbCrLf;
Exit Sub

PrtHandler3:
MsgBox "Kontroleer de printer."
Resume

End Sub

Private Sub PrintTotaal()
Dim T As Integer

On Local Error GoTo PrtHandler4

For T = 0 To 6
    VeldTXT(T) = ""
Next
VeldTXT(3) = "Periodiek totaal :"
VeldTXT(4) = Dec$(Abs(TotaalD), MASK_EURBH)
VeldTXT(5) = Dec$(Abs(TotaalC), MASK_EURBH)

Printer.Print vbCrLf; FULL_LINE

T = 0
Do While T < 8
    Printer.Print ; Tab(REPORT_TAB(T));
    Printer.Print ; VeldTXT(T);
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        Printer.Print vbCrLf;
    End If
    T = T + 1
Loop
Exit Sub

PrtHandler4:
MsgBox "Kontroleer de printer."
Resume

End Sub

Private Sub PrintVelden()
Dim T As Integer

On Local Error GoTo PrtHandler2

T = 0
Do While REPORT_TAB(T) <> 0
    Printer.Print Tab(REPORT_TAB(T));
    Printer.Print VeldTXT(T);
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        Printer.Print vbCrLf;
    End If
    T = T + 1
Loop
If Printer.CurrentY >= Printer.Height - 1200 Then
    Printer.NewPage
    Printer.FontSize = Printer.FontSize
    Printer.Print " "
    Printer.FontSize = Printer.FontSize
    PrintTitel
End If
Exit Sub

PrtHandler2:
MsgBox "Kontroleer de printer."
Printer.KillDoc
Unload Me
Resume

End Sub

Private Sub TekstLijn_GotFocus(Index As Integer)

TekstLijn(Index).SelLength = Len(TekstLijn(Index).Text)

End Sub

Private Sub TekstLijn_LostFocus(Index As Integer)
Dim T As Integer
Dim A As String

Select Case Index
    Case 0
        If DATE_INVALID(Right(TekstLijn(0).Text, 10)) Then
            MsgBox "Respecteer : " + vbCrLf + vbCrLf + "DD/MM/EEJJ - DD/MM/EEJJ a.u.b. !"
            TekstLijn(0).Text = BJPERDAT.PeriodeBoekjaar.Text
            TekstLijn(0).SetFocus
            Exit Sub
        ElseIf Len(TekstLijn(0).Text) <> 23 Then
            MsgBox "Respecteer : " + vbCrLf + vbCrLf + "DD/MM/EEJJ - DD/MM/EEJJ a.u.b. !"
            TekstLijn(0).Text = BJPERDAT.PeriodeBoekjaar.Text
            TekstLijn(0).SetFocus
            Exit Sub
        Else
            PlGrensVan = Mid(TekstLijn(0).Text, 7, 4) + Mid(TekstLijn(0).Text, 4, 2) + Mid(TekstLijn(0).Text, 1, 2)
            PlGrensTot = Mid(TekstLijn(0).Text, 20, 4) + Mid(TekstLijn(0).Text, 17, 2) + Mid(TekstLijn(0).Text, 14, 2)
            KeuzeInfo(0).Clear
            For T = 0 To 9
                bGet TABLE_LEDGERACCOUNTS, 0, RekeningNummer(T)
                If Ktrl Then
                    A = RekeningNummer(T) + Chr$(124) + "Niet aanwezig. Setup Boekjaar!"
                Else
                    RecordToVeld TABLE_LEDGERACCOUNTS
                    A = RekeningNummer(T) + Chr$(124) + RTrim$(vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #"))
                End If
                bFirst TABLE_JOURNAL, 0
                bGetOrGreater TABLE_JOURNAL, 0, RekeningNummer(T) + PlGrensVan
                If Ktrl Then
                Else
                    If KEY_BUF(TABLE_JOURNAL) > RekeningNummer(T) + PlGrensTot Then
                        Else
                        KeuzeInfo(0).AddItem A$
                    End If
                End If
            Next
            If KeuzeInfo(0).ListCount Then KeuzeInfo(0).ListIndex = 0
        End If

    Case 1
        If DATE_INVALID((TekstLijn(1).Text)) Then
            Beep
            TekstLijn(1).Text = MIM_GLOBAL_DATE
            TekstLijn(1).SetFocus
        End If
    Case 3
        TekstLijn(3).Text = Format(Val(TekstLijn(3).Text), "00000")
End Select

End Sub

Private Sub UittrekselsLijst_DblClick()

DetailFinancieelStuk Mid(UittrekselsLijst, 1, 8)

End Sub

Private Sub UittrekselsLijst_GotFocus()

SnelHelpPrint "[Enter] of dubbelklikken voor detail.", BL_LOGGING

End Sub


Private Sub UittrekselsLijst_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then DetailFinancieelStuk Mid(UittrekselsLijst, 1, 8)

End Sub


