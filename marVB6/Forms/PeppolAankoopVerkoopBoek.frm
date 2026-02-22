VERSION 5.00
Begin VB.Form AVBoek 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aankoop / Verkoop"
   ClientHeight    =   2595
   ClientLeft      =   4710
   ClientTop       =   4620
   ClientWidth     =   5670
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2595
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkAfdrukLiggend 
      Caption         =   "Liggende Printerafdruk"
      Height          =   255
      Left            =   60
      TabIndex        =   13
      Top             =   1560
      Width           =   2310
   End
   Begin VB.CheckBox chkAfdrukInVenster 
      Caption         =   "Afdruk in venster"
      Height          =   255
      Left            =   60
      TabIndex        =   12
      Top             =   1080
      Width           =   1830
   End
   Begin VB.CheckBox chkDetailJournaal 
      Caption         =   "Detail Journaal"
      Height          =   225
      Left            =   60
      TabIndex        =   11
      Top             =   1320
      Value           =   1  'Checked
      Width           =   1785
   End
   Begin VB.TextBox TekstLijn0 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   60
      TabIndex        =   2
      Top             =   420
      Width           =   4140
   End
   Begin VB.TextBox TekstLijn1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2880
      TabIndex        =   3
      Top             =   120
      Width           =   1320
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
      Height          =   300
      Left            =   4350
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   420
      Width           =   1200
   End
   Begin VB.OptionButton FaktuurCreditnota 
      Caption         =   "Creditnota's"
      Height          =   240
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   780
      Width           =   1395
   End
   Begin VB.OptionButton FaktuurCreditnota 
      Caption         =   "Facturen"
      Height          =   240
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Value           =   -1  'True
      Width           =   1155
   End
   Begin VB.CommandButton Drukken 
      Caption         =   "Af&drukken"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4350
      TabIndex        =   4
      Top             =   60
      Width           =   1200
   End
   Begin VB.Label TekstTot 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4440
      TabIndex        =   18
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Volgnummer Intern klassement"
      Height          =   615
      Left            =   4440
      TabIndex        =   17
      Top             =   840
      Width           =   1035
   End
   Begin VB.Label LabelBoekEnDocumentDatum 
      Caption         =   "Boekdatum = Documentdatum (Strikt)"
      Height          =   615
      Left            =   2880
      TabIndex        =   16
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label LabelDateTo 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2880
      TabIndex        =   15
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label LabelDateFrom 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2880
      TabIndex        =   14
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label TekstVan 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Sub&Titel"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   6
      Top             =   180
      Width           =   870
   End
   Begin VB.Label Label1 
      Caption         =   "D&atum Drukken"
      Height          =   195
      Index           =   1
      Left            =   1530
      TabIndex        =   7
      Top             =   180
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Van"
      Height          =   240
      Index           =   2
      Left            =   2460
      TabIndex        =   8
      Top             =   1590
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Tot"
      Height          =   240
      Index           =   3
      Left            =   2460
      TabIndex        =   9
      Top             =   2100
      Width           =   480
   End
End
Attribute VB_Name = "AVBoek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit

Dim getal As Long

Dim psTekst(5) As String
Dim RapportVeldNr(23) As Integer
Dim RapportManier(23) As Integer
Dim KolomTotaal(17) As Currency
Dim T               As Integer

Dim LFontSize(20)           As Single
Dim LAantalL(20)            As Integer
Dim FontDefChanged          As Integer

Dim LijstNaam As String

Dim rFlag As String
Dim r(10) As String
Dim FlBtw As Integer
Dim DBVT(23) As Double
Dim DDVT(23) As Double

Dim BedragForfait(100) As Double
Dim PctForfait(100) As Double
Dim BtwForfait(100) As Single
Dim KortingForfait(100) As Single
Dim VakForfait(0 To 3) As Double
Dim BtwTotaalForfait As Double

Dim ForFait As Integer
Dim ForfaitBtw As Double

Dim Ar As Integer

Dim tMaxVeld As Integer
Dim VeldTXT(17) As String
Dim aa As String

Dim MaskHier As String

Dim rsAVBoekHier As ADODB.Recordset
Dim rsJourHier As ADODB.Recordset

Dim typeVATPeriod As String
Dim beginPeriod As String
Dim endPeriod As String

Private Sub Annuleren_Click()

Unload Me

End Sub

Private Sub CumulPrint()
Dim RekeningNaam As String * 30
Dim Tabul As Integer
Dim Tel As Integer

If chkAfdrukInVenster Then Exit Sub

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
    PrintTitel
    PrintTotaal

    Printer.Print
    Printer.Print Tab(2); "** CENTRALISATIE/VIERKANTSCONTROLE **"
    Printer.Print ""

bClose TABLE_DUMMY
bFirst TABLE_DUMMY, 0
RecordToVeld TABLE_DUMMY
If Not ADO_GET(TABLE_LEDGERACCOUNTS, 0, "=", Left(FVT(TABLE_DUMMY, 0), 7)) Then
    RekeningNaam = "Rekening reeds vernietigd !!!"
Else
    RecordToVeld TABLE_LEDGERACCOUNTS
    RekeningNaam = RV(rsMAR(TABLE_LEDGERACCOUNTS), "v020")
End If
Tabul = 0
Printer.Print Tab(Tabul + 2); Dec$(Val(vBibTekst(TABLE_DUMMY, "#v013 #")), "####"); " x "; vSet(vBibTekst(TABLE_DUMMY, "#v089 #"), 7); " "; RekeningNaam; " "; Dec$(Val(vBibTekst(TABLE_DUMMY, "#v068 #")), MaskHier);

Do
    bNext TABLE_DUMMY
    If Ktrl Then
        Exit Do
    End If
    RecordToVeld TABLE_DUMMY
    If Not ADO_GET(TABLE_LEDGERACCOUNTS, 0, "=", Left(FVT(TABLE_DUMMY, 0), 7)) Then
        RekeningNaam = "Rekening reeds vernietigd !!!"
    Else
        RekeningNaam = RV(rsMAR(TABLE_LEDGERACCOUNTS), "v020")
    End If
    If Tabul = 0 Then
        Tabul = 56
        Printer.Print Tab(Tabul + 2); Dec$(Val(vBibTekst(TABLE_DUMMY, "#v013 #")), "####"); " x "; vSet(vBibTekst(TABLE_DUMMY, "#v089 #"), 7); " "; RekeningNaam; " "; Dec$(Val(vBibTekst(TABLE_DUMMY, "#v068 #")), MaskHier); vbCrLf;
    Else
        Tabul = 0
        Printer.Print Tab(Tabul + 2); Dec$(Val(vBibTekst(TABLE_DUMMY, "#v013 #")), "####"); " x "; vSet(vBibTekst(TABLE_DUMMY, "#v089 #"), 7); " "; RekeningNaam; " "; Dec$(Val(vBibTekst(TABLE_DUMMY, "#v068 #")), MaskHier);
    End If
Loop

Dim BedragVK2 As Double
Dim BedragVK As Double

BtwTotaalForfait = 0
VakForfait(0) = 0
VakForfait(1) = 0
VakForfait(2) = 0
VakForfait(3) = 0
If ForFait Then
    Printer.Print vbCrLf; vbCrLf; vbCrLf; vbCrLf;
    Printer.Print Tab(2); "** FORFAITAIR STELSEL **"
    Printer.Print vbCrLf; vbCrLf;
    Printer.Print Tab(2); "FORFAITCODE", "AANKOOPBEDRAG", "FAKTOR", "VERKOOP 1", "KORTING", "VERKOOP NETTO", "BTW %", "BTW BEDRAG"; vbCrLf; vbCrLf;
    For Tel = 0 To 39
        If BedragForfait(Tel) <> 0 Then
            BedragVK2 = CDbl(Format(BedragForfait(Tel) * PctForfait(Tel), "0.00"))
            BedragVK = CDbl(Format(BedragVK2 - (BedragVK2 * KortingForfait(Tel) / 100), "0.00"))
            Printer.Print Tab(2); Dec((Tel), "00"), Dec(BedragForfait(Tel), "########.00"), Dec(PctForfait(Tel), "####.0000"), Dec(BedragVK2, "########.00"), Dec(BedragVK - BedragVK2, "########.00"), Dec(BedragVK, "#######.00"), Dec((BtwForfait(Tel)), "###.0");
            ForfaitBtw = CDbl(Format(BedragVK * BtwForfait(Tel) / 100, "0.00"))
            Printer.Print Dec$(ForfaitBtw, "########.00"); vbCrLf;
            VakForfait(Int(Tel / 10)) = VakForfait(Int(Tel / 10)) + BedragVK
            BtwTotaalForfait = BtwTotaalForfait + ForfaitBtw
        End If
    Next
    For Tel = 40 To 99
        If BedragForfait(Tel) <> 0 Then
            MsgBox "BTW aangifte Vak " + Format(Abs(Tel / 10)) + " bestaat toch niet ?"
        End If
    Next
    Printer.Print vbCrLf; vbCrLf; vbCrLf; vbCrLf;
    Printer.Print Tab(2); "** BTW VAKKEN **"
    For Tel = 0 To 3
        Printer.Print Tab(2); "VAK "; Dec((Tel), "00"); " : "; Dec$(VakForfait(Tel), "########.00"); vbCrLf;
    Next
    Printer.Print vbCrLf; vbCrLf;
    Printer.Print Tab(2); "VAK 54 : "; Dec$(BtwTotaalForfait, "########.00")
End If

End Sub

Private Sub DetailRekeningen()
Dim RekeningNaam As String * 30
Dim Tabul As Integer
Dim ForfaitNr As Integer

If chkDetailJournaal Then
    If Printer.CurrentY >= Printer.Height - 1200 Then
        Printer.NewPage
        Printer.FontSize = Printer.FontSize
        Printer.Print " "
        Printer.FontSize = Printer.FontSize
        PrintTitel
    End If
    Printer.Print vbCrLf;
    Tabul = 56
End If

Do While Not rsJourHier.EOF
    If RV(rsJourHier, "v033") = RV(rsAVBoekHier, "v033") Then
        If Left(RV(rsJourHier, "v019"), 2) = "40" Or Left(RV(rsJourHier, "v019"), 2) = "44" Then
            RekeningNaam = RV(rsJourHier, "v067")
        Else
            If Not ADO_GET(TABLE_LEDGERACCOUNTS, 0, "=", RV(rsJourHier, "v019")) Then
                RekeningNaam = "Rekening reeds vernietigd !!!"
            Else
                RekeningNaam = RV(rsMAR(TABLE_LEDGERACCOUNTS), "v020")
                If ForFait Then
                    GoSub ForFaitBerekening
                End If
            End If
        End If
        If chkDetailJournaal Then
            If Tabul = 0 Then
                Tabul = 56
                Printer.Print Tab(Tabul + 2); RV(rsJourHier, "v019"); " "; RekeningNaam; " "; Dec$((RV(rsJourHier, "dece068")), MaskHier); vbCrLf;
                If Printer.CurrentY >= Printer.Height - 1200 Then
                    Printer.NewPage
                    Printer.FontSize = Printer.FontSize
                    Printer.Print " "
                    Printer.FontSize = Printer.FontSize
                    PrintTitel
                End If
            Else
                Tabul = 0
                Printer.Print Tab(Tabul + 2); RV(rsJourHier, "v019"); " "; RekeningNaam; " "; Dec$((RV(rsJourHier, "dece068")), MaskHier);
            End If
        End If
        GoSub DetailCumul
        rsJourHier.MoveNext
    Else
        Exit Do
    End If
Loop

If chkDetailJournaal Then
    If Tabul = 0 Then
        Printer.Print vbCrLf;
    End If
    If Printer.CurrentY >= Printer.Height - 1200 Then
        Printer.NewPage
        Printer.FontSize = Printer.FontSize
        Printer.Print " "
        Printer.FontSize = Printer.FontSize
        PrintTitel
    End If
End If
If chkAfdrukInVenster = 0 Then
    Printer.Print vbCrLf;
End If
Exit Sub

DetailCumul:
StartPunt:
bGet TABLE_DUMMY, 0, RV(rsJourHier, "v019")
If Ktrl Then
    TLB_RECORD(TABLE_DUMMY) = ""
    vBib TABLE_DUMMY, RV(rsJourHier, "v019"), "v089"
    vBib TABLE_DUMMY, "0", "v013"
    vBib TABLE_DUMMY, "0", "v068"
    bInsert TABLE_DUMMY, 0
    GoTo StartPunt
Else
    RecordToVeld TABLE_DUMMY
    vBib TABLE_DUMMY, Str$(Val(vBibTekst(TABLE_DUMMY, "#v013 #")) + 1), "v013"
    vBib TABLE_DUMMY, Str$(Val(vBibTekst(TABLE_DUMMY, "#v068 #")) + RV(rsJourHier, "dece068")), "v068"
    bUpdate TABLE_DUMMY, 0
End If
Return

ForFaitBerekening:
If RTrim$(RV(rsMAR(TABLE_LEDGERACCOUNTS), "v216")) <> "" Then
    ForfaitNr = Val(RV(rsMAR(TABLE_LEDGERACCOUNTS), "v216"))
    BedragForfait(ForfaitNr) = BedragForfait(ForfaitNr) + RV(rsJourHier, "dece068")
    If PctForfait(ForfaitNr) = 0 Then
        bGet TABLE_VARIOUS, 1, vSet("21" + RV(rsMAR(TABLE_LEDGERACCOUNTS), "v216"), 20)
        If Ktrl Then
            MsgBox "Forfaitaire Kode : " + RV(rsMAR(TABLE_LEDGERACCOUNTS), "v216") + " en waarde nog niet aanwezig.  Eerst inbrengen via diverse gebruikersfiches pér bedrijf a.u.b. !"
        Else
            RecordToVeld TABLE_VARIOUS
            PctForfait(ForfaitNr) = Val(vBibTekst(TABLE_VARIOUS, "#v218 #"))
            BtwForfait(ForfaitNr) = Val(Mid(fmarBoxText("002", "2", vBibTekst(TABLE_VARIOUS, "#v111 #")), 4, 4))
            KortingForfait(ForfaitNr) = Val(vBibTekst(TABLE_VARIOUS, "#v230 #"))
        End If
    End If
End If
Return

End Sub

Private Sub chkAfdrukInVenster_Click()

If chkAfdrukInVenster = 1 Then
    chkDetailJournaal = 0
End If

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
Dim BeginSleutel    As String
Dim EindSleutel     As String

Dim DummySleutel    As String
Dim Tekst           As String

Dim XX As Integer
Dim LengteSleutel   As Integer
Dim sleutelIndex    As Integer
Dim Tel             As Integer

For T = 0 To 17
    KolomTotaal(T) = 0
Next
For T = 0 To 99
    BedragForfait(T) = 0
    PctForfait(T) = 0
Next

Select Case aIndex
    Case TABLE_SUPPLIERS
        BeginSleutel = "A"
        EindSleutel = "A"
    Case TABLE_CUSTOMERS
        BeginSleutel = "V"
        EindSleutel = "V"
    Case Else
        MsgBox "Stop"
End Select
Select Case FaktuurCreditnota(0).Value
    Case -1
        BeginSleutel = BeginSleutel + "0" + Mid(PERIOD_FROMTO, 1, 4) + Format(Val(TekstVan.Caption), "00000")
        EindSleutel = EindSleutel + "0" + Mid(PERIOD_FROMTO, 1, 4) + Format(Val(TekstTot.Caption), "00000")
        Tekst$ = FaktuurCreditnota(0).Caption
    Case Else
        BeginSleutel = BeginSleutel + "1" + Mid(PERIOD_FROMTO, 1, 4) + Format(Val(TekstVan.Caption), "00000")
        EindSleutel = EindSleutel + "1" + Mid(PERIOD_FROMTO, 1, 4) + Format(Val(TekstTot.Caption), "00000")
        Tekst$ = FaktuurCreditnota(1).Caption
End Select

psTekst(2) = LijstNaam + " " + Tekst$ + " " + Mid(Mim.Caption, InStr(Mim.Caption, "["))
psTekst(0) = TekstLijn1.text
psTekst(3) = TekstLijn0.text
InitVelden

Screen.MousePointer = vbHourglass

'recordset maken
Set rsAVBoekHier = New ADODB.Recordset
On Error Resume Next
Err = 0
rsAVBoekHier.CursorLocation = adUseClient

Msg = "SELECT * FROM Dokumenten "
Msg = Msg & "WHERE Mid(v033, 1, 2) = '" & Mid(BeginSleutel, 1, 2) & "' "
Msg = Msg & "AND v035 >= '" & beginPeriod & "' "
Msg = Msg & "AND v035 <= '" & endPeriod & "' "
Msg = Msg & "ORDER BY v033 ASC"

SnelHelpPrint Msg, BL_LOGGING
rsAVBoekHier.Open Msg, adntDB, adOpenForwardOnly, adLockReadOnly
If Err Then
    MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
    Screen.MousePointer = vbNormal
    Exit Sub
ElseIf rsAVBoekHier.RecordCount = 0 Then
    MsgBox "Er zijn géén dokumenten", vbExclamation
    Screen.MousePointer = vbNormal
    Exit Sub
Else
    'TODO update the field TOT in screen
    rsAVBoekHier.MoveLast
    TekstTot.Caption = Mid(rsAVBoekHier("v033"), 7)
    'TODO update the field VAN in screen
    rsAVBoekHier.MoveFirst
    TekstVan.Caption = Mid(rsAVBoekHier("v033"), 7)
End If

'recordset maken
Set rsJourHier = New ADODB.Recordset
On Error Resume Next
Err = 0
rsJourHier.CursorLocation = adUseClient

Msg = "SELECT * FROM Journalen "
Msg = Msg & "WHERE v033 >= '" & rsAVBoekHier("v033") & "' "
Msg = Msg & "AND v035 >= '" & beginPeriod & "' "
Msg = Msg & "AND v035 <= '" & endPeriod & "' "
Msg = Msg & "AND v038 IS NULL "
Msg = Msg & "ORDER BY v033 ASC, v019 ASC"

SnelHelpPrint Msg, BL_LOGGING
rsJourHier.Open Msg, adntDB, adOpenForwardOnly, adLockReadOnly
If Err Then
    MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
    Screen.MousePointer = vbNormal
    Exit Sub
Else
    rsJourHier.MoveFirst
End If
bClose TABLE_DUMMY
ClearFlDummy
Ktrl = bOpen(TABLE_DUMMY)
AVBoek.Refresh

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
PrintTitel


Do While Not rsAVBoekHier.EOF
    PrintVelden
    DetailRekeningen
    rsAVBoekHier.MoveNext
    
    
Loop
PrintTotaal
CumulPrint
    If chkAfdrukInVenster = 0 Then
        Printer.NewPage
        Printer.EndDoc
    End If
    Screen.MousePointer = vbNormal
    Msg = "Totaliseren voor BTW AANGIFTE.  Bent U zeker?"
    If FaktuurCreditnota(0).Value Then
        Msg = Msg + vbCrLf + vbCrLf + "Opgelet !  Creditnota's niet vergeten straks..."
    End If
    Ktrl = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton1, "BTW Aangifte")
    If Ktrl = vbYes Then
        GoSub BTWAangifte
    End If
    AVBoek.Refresh
    AVBoek.SetFocus
    rsAVBoekHier.Close
    rsJourHier.Close
    Set rsAVBoekHier = Nothing
    Set rsJourHier = Nothing
    If FaktuurCreditnota(0).Value = True Then
        FaktuurCreditnota(1).Value = True
    Else
        Unload Me
    End If
Exit Sub

PrtHandler2:
MsgBox "Kontroleer de printer."
Resume

BTWAangifte:
bGet TABLE_VARIOUS, 1, vSet("17" + BJPERDAT.Boekjaar.text + Format(BJPERDAT.PeriodeBoekjaar.ListIndex + 1, "00"), 20)
If Ktrl Then
    MsgBox "Stop"
Else
    RecordToVeld TABLE_VARIOUS
End If

'btwperiode bijvoegen vanaf 07/2008 voor Intervat
If Mid(PERIOD_FROMTO, 5, 2) = Mid(PERIOD_FROMTO, 13, 2) Then
    'ok de periode marIntegraal staan op maandelijks
    vBib TABLE_VARIOUS, Mid(PERIOD_FROMTO, 5, 2), "i001" 'werkelijke maand
    vBib TABLE_VARIOUS, Mid(PERIOD_FROMTO, 1, 4), "i002" 'werkelijk jaar
Else
    MsgBox "marIntegraal boekhoudperiodes staan nog altijd op 3-maandelijks.  Geen Intervat aangifte mogelijk met deze werkwijze die dateert van 1985-1994 en vermoedelijk overgenomen werd uit marIntegraal DOS periode. Contacteer ons 0475/292255 voor manuele tussenkomst!!)", vbCritical
End If
    
Select Case aIndex
    Case TABLE_SUPPLIERS
        Select Case Ar
            Case 1
                'Record Kontroleren, zou MOETEN op nul staan...
                vBib TABLE_VARIOUS, Str$(KolomTotaal(16)), "v045" 'vak 59
                vBib TABLE_VARIOUS, Str$(KolomTotaal(9)), "v052" 'vak 86
                vBib TABLE_VARIOUS, Str$(KolomTotaal(11)), "v053" 'vak 87
                vBib TABLE_VARIOUS, Str$(KolomTotaal(12)), "v054" 'vak 88
                
                vBib TABLE_VARIOUS, Format(Val(TekstVan.Caption), "00000"), "v092"
                vBib TABLE_VARIOUS, Format(Val(TekstTot.Caption), "00000"), "v093"
                
            Case 3
                vBib TABLE_VARIOUS, Str$(KolomTotaal(16)), "v100" 'vak 63
                vBib TABLE_VARIOUS, Str$(KolomTotaal(7)), "v050" 'vak 84
                vBib TABLE_VARIOUS, Str$(KolomTotaal(8)), "v051" 'vak 85
                
                vBib TABLE_VARIOUS, Format(Val(TekstVan.Caption), "00000"), "v094"
                vBib TABLE_VARIOUS, Format(Val(TekstTot.Caption), "00000"), "v095"
                For Tel = 3 To 16
                    KolomTotaal(Tel) = -KolomTotaal(Tel)
                Next
            Case Else
                MsgBox "Stop"
        End Select

        vBib TABLE_VARIOUS, Str$(KolomTotaal(13) + Val(vBibTekst(TABLE_VARIOUS, "#v042 #"))), "v042" 'vak 55
        vBib TABLE_VARIOUS, Str$(KolomTotaal(14) + Val(vBibTekst(TABLE_VARIOUS, "#v043 #"))), "v043" 'vak 56
        vBib TABLE_VARIOUS, Str$(KolomTotaal(15) + Val(vBibTekst(TABLE_VARIOUS, "#v044 #"))), "v044" 'vak 57
        vBib TABLE_VARIOUS, Str$(KolomTotaal(3) + Val(vBibTekst(TABLE_VARIOUS, "#v046 #"))), "v046" 'vak 81
        vBib TABLE_VARIOUS, Str$(KolomTotaal(4) + Val(vBibTekst(TABLE_VARIOUS, "#v047 #"))), "v047" 'vak 82
        vBib TABLE_VARIOUS, Str$(KolomTotaal(5) + Val(vBibTekst(TABLE_VARIOUS, "#v048 #"))), "v048" 'vak 83
        
        If ForFait Then
            vBib TABLE_VARIOUS, Str$(VakForfait(0)), "v055" 'vak 00
            vBib TABLE_VARIOUS, Str$(VakForfait(1)), "v056" 'vak 01
            vBib TABLE_VARIOUS, Str$(VakForfait(2)), "v057" 'vak 02
            vBib TABLE_VARIOUS, Str$(VakForfait(3)), "v058" 'vak 03
            vBib TABLE_VARIOUS, Str$(BtwTotaalForfait), "v064" 'vak 54
        End If
        
    Case TABLE_CUSTOMERS
        Select Case Ar
            Case 12
                vBib TABLE_VARIOUS, Str$(KolomTotaal(12)), "v064" 'vak 54
                vBib TABLE_VARIOUS, Str$(KolomTotaal(2) + Val(vBibTekst(TABLE_VARIOUS, "#v055 #"))), "v055" 'vak 00 bijtellen ?
                
                vBib TABLE_VARIOUS, Str$(KolomTotaal(3)), "v056" 'vak 01
                vBib TABLE_VARIOUS, Str$(KolomTotaal(4)), "v057" 'vak 02
                vBib TABLE_VARIOUS, Str$(KolomTotaal(5)), "v058" 'vak 03
                vBib TABLE_VARIOUS, Str$(KolomTotaal(6)), "v059" 'vak 45
                vBib TABLE_VARIOUS, Str$(KolomTotaal(7)), "v060" 'vak 46
                vBib TABLE_VARIOUS, Str$(KolomTotaal(8)), "v061" 'vak 47
                
                vBib TABLE_VARIOUS, Format(Val(TekstVan.Caption), "00000"), "v096"
                vBib TABLE_VARIOUS, Format(Val(TekstTot.Caption), "00000"), "v097"
                
            Case 14
                vBib TABLE_VARIOUS, Str$(KolomTotaal(12)), "v101" 'vak 64
                vBib TABLE_VARIOUS, Str$(Val(vBibTekst(TABLE_VARIOUS, "#v055 #")) - KolomTotaal(2)), "v055" 'vak 00 aftrekken ?
                
                vBib TABLE_VARIOUS, Str$(KolomTotaal(10)), "v062" 'vak 48
                vBib TABLE_VARIOUS, Str$(KolomTotaal(11)), "v063" 'vak 49
                
                vBib TABLE_VARIOUS, Format(Val(TekstVan.Caption), "00000"), "v098"
                vBib TABLE_VARIOUS, Format(Val(TekstTot.Caption), "00000"), "v099"

            Case Else
                MsgBox "Stop"
        End Select
    Case Else
        MsgBox "Stop"
End Select

If bhEuro Then
    vBib TABLE_VARIOUS, "EUR", "vEUR"
End If

bUpdate TABLE_VARIOUS, 1
Select Case FaktuurCreditnota(0).Value
    Case True 'Faktuur
        If aIndex = TABLE_SUPPLIERS Then
            SS99 TekstTot.Caption, 2
        Else
            SS99 TekstTot.Caption, 12
        End If

    Case False
        If aIndex = TABLE_SUPPLIERS Then
            SS99 TekstTot.Caption, 4
        Else
            SS99 TekstTot.Caption, 14
        End If
End Select
Return

End Sub

Private Sub FaktuurCreditNota_Click(Index As Integer)
Dim T       As Integer


Select Case aIndex
    Case TABLE_SUPPLIERS
        T = 0
    Case TABLE_CUSTOMERS
        T = 10
    Case Else
        MsgBox "Stop"
End Select

getal = 0
Select Case FaktuurCreditnota(0).Value
    Case True 'Faktuur
        FL99_RECORD = String99(READING, 1 + T)
        TekstTot.Caption = Format(Val(FL99_RECORD), "00000")
        FL99_RECORD = String99(READING, 2 + T)
        If Format(Val(FL99_RECORD), "00000") = TekstTot.Caption Then
            Drukken.Enabled = False
            TekstVan.Caption = Format(Val(FL99_RECORD), "00000")
        Else
            Drukken.Enabled = True
            TekstVan.Caption = Format(Val(FL99_RECORD) + 1, "00000")
        End If
        If aIndex = TABLE_SUPPLIERS Then
            Ar = 1
            getal = Val(vBibTekst(TABLE_VARIOUS, "#v092 #")) + Val(vBibTekst(TABLE_VARIOUS, "#v093 #"))
        Else
            Ar = 12
            getal = Val(vBibTekst(TABLE_VARIOUS, "#v096 #")) + Val(vBibTekst(TABLE_VARIOUS, "#v097 #"))
        End If

    Case Else
        FL99_RECORD = String99(READING, 3 + T)
        TekstTot.Caption = Format(Val(FL99_RECORD), "00000")
        FL99_RECORD = String99(READING, 4 + T)
        If Format(Val(FL99_RECORD), "00000") = TekstTot.Caption Then
            Drukken.Enabled = False
            TekstVan.Caption = Format(Val(FL99_RECORD), "00000")
        Else
            Drukken.Enabled = True
            TekstVan.Caption = Format(Val(FL99_RECORD) + 1, "00000")
        End If
        If aIndex = TABLE_SUPPLIERS Then
            Ar = 3
            getal = Val(vBibTekst(TABLE_VARIOUS, "#v094 #")) + Val(vBibTekst(TABLE_VARIOUS, "#v095 #"))
        Else
            Ar = 14
            getal = Val(vBibTekst(TABLE_VARIOUS, "#v098 #")) + Val(vBibTekst(TABLE_VARIOUS, "#v099 #"))
        End If
End Select
If getal Then
    MsgBox "Binnen deze periode, zijn er reeds" + vbCrLf + "dokumenten opgenomen !", 0, "BTW aangifte kontroleren a.u.b. !"
    Drukken.Enabled = False
Else
    'Drukken.Enabled = True
End If

End Sub

Private Sub Form_Load()
Dim PeriodeSleutel  As String * 20
Dim T               As Integer
Dim PeriodeMax      As Integer

'bhEuro
MaskHier = "#####0.00"

Set Printer = Printers(LijstPrinterNr)
On Error Resume Next
Printer.PaperBin = LaadTekst(App.Title, "LIJSTPRINTER")
If Printer.Orientation = vbPRORLandscape Then
    chkAfdrukLiggend.Value = 1
Else
    chkAfdrukLiggend.Value = 0
End If
chkAfdrukLiggend_Click

TekstLijn1.text = MIM_GLOBAL_DATE
Select Case aIndex
    Case TABLE_SUPPLIERS
        LijstNaam = "Aankoopboek"
    Case TABLE_CUSTOMERS
        LijstNaam = "Verkoopboek"
    Case Else
        MsgBox "Stop"
End Select
AVBoek.Caption = LijstNaam

typeVATPeriod = String99(READING, 301)
Select Case typeVATPeriod
    Case "1"
        TekstLijn0 = "Maandelijkse aangifte " & Mid(PERIOD_FROMTO, 5, 2) & "/" & Mid(PERIOD_FROMTO, 1, 4)
        'Monthly periods shoud be (easy)
        beginPeriod = Mid(PERIOD_FROMTO, 1, 8)
        endPeriod = Mid(PERIOD_FROMTO, 9)
        LabelDateFrom.Caption = DATE_TEXT(beginPeriod)
        LabelDateTo.Caption = DATE_TEXT(endPeriod)
    Case "2"
        'Quarter Perdiods should be
        '01/01 - 31/03
        '01/04 - 30/06
        '01/07 - 30/09
        '01/10 - 31/12
        Select Case Mid(PERIOD_FROMTO, 13, 2)
            Case "03"
                beginPeriod = Mid(PERIOD_FROMTO, 1, 4) + "0101"
                endPeriod = Mid(PERIOD_FROMTO, 9)
                TekstLijn0 = "Kwartaal aangifte 03/" & Mid(PERIOD_FROMTO, 1, 4)
                    
            Case "06"
                beginPeriod = Mid(PERIOD_FROMTO, 1, 4) + "0401"
                endPeriod = Mid(PERIOD_FROMTO, 9)
                TekstLijn0 = "Kwartaal aangifte 06/" & Mid(PERIOD_FROMTO, 1, 4)
                    
            Case "09"
                beginPeriod = Mid(PERIOD_FROMTO, 1, 4) + "0701"
                endPeriod = Mid(PERIOD_FROMTO, 9)
                TekstLijn0 = "Kwartaal aangifte 09/" & Mid(PERIOD_FROMTO, 1, 4)
                    
            Case "12"
                beginPeriod = Mid(PERIOD_FROMTO, 1, 4) + "1001"
                endPeriod = Mid(PERIOD_FROMTO, 9)
                TekstLijn0 = "Kwartaal aangifte 12/" & Mid(PERIOD_FROMTO, 1, 4)
                
            Case Else
                MsgBox "Selecteer een geldige maand voor kwartaalaangifte", vbExclamation
                Unload Me
                Exit Sub
        End Select
        LabelDateFrom.Caption = DATE_TEXT(beginPeriod)
        LabelDateTo.Caption = DATE_TEXT(endPeriod)
        
    Case Else
        MsgBox "Volgens Setup BTW geen aangifteplicht. Controleer eventueel.", vbInformation
        Unload Me
        Exit Sub
End Select

Ktrl = 0
SnelHelpPrint "Kontrole hogere periodes...", BL_LOGGING
PeriodeMax = BJPERDAT.PeriodeBoekjaar.ListCount + 1
Do While PeriodeMax > BJPERDAT.PeriodeBoekjaar.ListIndex + 1
    PeriodeSleutel = "17" + BJPERDAT.Boekjaar.text + Format(PeriodeMax, "00")
    bGet TABLE_VARIOUS, 1, PeriodeSleutel
    If Ktrl Then
        TLB_RECORD(TABLE_VARIOUS) = ""
        vBib TABLE_VARIOUS, (BJPERDAT.Boekjaar.text), "v090"
        vBib TABLE_VARIOUS, Format(PeriodeMax, "00"), "v091"
        vBib TABLE_VARIOUS, "17" + vBibTekst(TABLE_VARIOUS, "#v090 #") + vBibTekst(TABLE_VARIOUS, "#v091 #"), "v005"
        bInsert TABLE_VARIOUS, 1
    Else
        RecordToVeld TABLE_VARIOUS
        getal = 0
        For T = 92 To 99
            getal = getal + Val(vBibTekst(TABLE_VARIOUS, "#v" + Format(T, "000") + " #"))
        Next
        If getal Then
            getal = PeriodeMax
            PeriodeMax = 0
            Exit Do
        End If
    End If
    PeriodeMax = PeriodeMax - 1
Loop
jump:
If getal Then
    MsgBox "Periode " + Format(getal, "00") + " reeds afgesloten..."
    Drukken.Visible = False
    Exit Sub
Else
    PeriodeSleutel = "17" + BJPERDAT.Boekjaar.text + Format(BJPERDAT.PeriodeBoekjaar.ListIndex + 1, "00")
    bGet TABLE_VARIOUS, 1, PeriodeSleutel
    If Ktrl Then
        TLB_RECORD(TABLE_VARIOUS) = ""
        vBib TABLE_VARIOUS, (BJPERDAT.Boekjaar.text), "v090"
        vBib TABLE_VARIOUS, Format(BJPERDAT.PeriodeBoekjaar.ListIndex + 1, "00"), "v091"
        vBib TABLE_VARIOUS, "17" + vBibTekst(TABLE_VARIOUS, "#v090 #") + vBibTekst(TABLE_VARIOUS, "#v091 #"), "v005"
        bInsert TABLE_VARIOUS, 1
        GoTo jump
    Else
        RecordToVeld TABLE_VARIOUS
    End If
End If

If Left(String99(READING, 20), 1) = "4" Then
    ForFait = 1
Else
    ForFait = 0
End If

FaktuurCreditNota_Click 0

End Sub

Private Sub InitVelden()
Dim T As Integer
Dim VolgTab As Integer

RapportVeldNr(0) = 33
    RapportManier(0) = 0    'niks omwerken gewoon afdrukken
    REPORT_FIELD(0) = "dokument"
    REPORT_TAB(0) = 2

RapportVeldNr(1) = 35
    RapportManier(1) = 5    'datum omwerken
    REPORT_FIELD(1) = "Datum dok."
    REPORT_TAB(1) = 14

Select Case aIndex
    Case TABLE_SUPPLIERS
        RapportVeldNr(2) = 39
            RapportManier(2) = 0
            REPORT_FIELD(2) = "Referte"
            REPORT_TAB(2) = 25
        
        RapportVeldNr(3) = 46
            RapportManier(3) = 9    'geheel getal geformateerd
            REPORT_FIELD(3) = "   VAK 81"
            REPORT_TAB(3) = 46

        RapportVeldNr(4) = 47
            RapportManier(4) = 9
            REPORT_FIELD(4) = "   VAK 82"
            REPORT_TAB(4) = 56

        RapportVeldNr(5) = 48
            RapportManier(5) = 9
            REPORT_FIELD(5) = "   VAK 83"
            REPORT_TAB(5) = 66

        RapportVeldNr(6) = 49
            RapportManier(6) = 9
            REPORT_FIELD(6) = "   DERDEN"
            REPORT_TAB(6) = 76

        RapportVeldNr(7) = 50
            RapportManier(7) = 9
            REPORT_FIELD(7) = "   VAK 84"
            REPORT_TAB(7) = 86

        RapportVeldNr(8) = 51
            RapportManier(8) = 9
            REPORT_FIELD(8) = "   VAK 85"
            REPORT_TAB(8) = 96

        RapportVeldNr(9) = 52
            RapportManier(9) = 9
            REPORT_FIELD(9) = "   VAK 86"
            REPORT_TAB(9) = 106

        RapportVeldNr(10) = 99
            RapportManier(10) = 1   'zoek flpartij kode+naam1
            REPORT_FIELD(10) = "ID.Kode/Naam"
            REPORT_TAB(10) = 2

        RapportVeldNr(11) = 53
            RapportManier(11) = 9
            REPORT_FIELD(11) = "   VAK 87"
            REPORT_TAB(11) = 56

        RapportVeldNr(12) = 54
            RapportManier(12) = 9
            REPORT_FIELD(12) = "   VAK 88"
            REPORT_TAB(12) = 66

        RapportVeldNr(13) = 42
            RapportManier(13) = 9
            REPORT_FIELD(13) = "   VAK 55"
            REPORT_TAB(13) = 76
        
        RapportVeldNr(14) = 43
            RapportManier(14) = 9
            REPORT_FIELD(14) = "   VAK 56"
            REPORT_TAB(14) = 86

        RapportVeldNr(15) = 44
            RapportManier(15) = 9
            REPORT_FIELD(15) = "   VAK 57"
            REPORT_TAB(15) = 96
        
        RapportVeldNr(16) = 45
            RapportManier(16) = 9
            REPORT_TAB(16) = 106
            
        Select Case Ar
            Case 1
                REPORT_FIELD(16) = "   VAK 59"
            Case Else
                REPORT_FIELD(16) = "   VAK 63"
        End Select
        REPORT_TAB(17) = 0
        tMaxVeld = 16

    Case TABLE_CUSTOMERS
        RapportVeldNr(2) = 55
            RapportManier(2) = 9    'geheel getal geformateerd
            REPORT_FIELD(2) = "VAK 00"
            REPORT_TAB(2) = 44

        RapportVeldNr(3) = 56
            RapportManier(3) = 9
            REPORT_FIELD(3) = "VAK 01"
            REPORT_TAB(3) = 55

        RapportVeldNr(4) = 57
            RapportManier(4) = 9
            REPORT_FIELD(4) = "VAK 02"
            REPORT_TAB(4) = 66

        RapportVeldNr(5) = 58
            RapportManier(5) = 9
            REPORT_FIELD(5) = "VAK 03"
            REPORT_TAB(5) = 77

        RapportVeldNr(6) = 59
            RapportManier(6) = 9
            REPORT_FIELD(6) = "VAK 45"
            REPORT_TAB(6) = 88

        RapportVeldNr(7) = 60
            RapportManier(7) = 9
            REPORT_FIELD(7) = "VAK 46"
            REPORT_TAB(7) = 99

        RapportVeldNr(8) = 61
            RapportManier(8) = 9
            REPORT_FIELD(8) = "VAK 47"
            REPORT_TAB(8) = 110

        RapportVeldNr(9) = 99
            RapportManier(9) = 1
            REPORT_FIELD(9) = "ID.Code/Naam"
            REPORT_TAB(9) = 2

        RapportVeldNr(10) = 62
            RapportManier(10) = 9
            REPORT_FIELD(10) = "VAK 48"
            REPORT_TAB(10) = 77
            
        
        RapportVeldNr(11) = 63
            RapportManier(11) = 9
            REPORT_FIELD(11) = "VAK 49"
            REPORT_TAB(11) = 88

        RapportVeldNr(12) = 64
            RapportManier(12) = 9
            REPORT_TAB(12) = 99
            
        Select Case Ar
            Case 12
                REPORT_FIELD(12) = "VAK 54"
            Case Else
                REPORT_FIELD(12) = "VAK 64"
        End Select
        tMaxVeld = 12
        REPORT_TAB(13) = 0
End Select

For T = 0 To 17
    KolomTotaal(T) = 0
Next

If chkAfdrukInVenster Then
    Me.Hide
    Unload Xlog
    Xlog.Hide
    Xlog.Caption = Me.Caption
    Xlog.X.Cols = tMaxVeld + 1
    Xlog.X.Row = 0
    For T = 0 To tMaxVeld
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
    Printer.Print Tab(1); psTekst(2); Tab(108); "Pagina : " + Dec$((PAGE_COUNTER), "##########");
    Printer.Print Tab(108); "Datum  : "; psTekst(0); vbCrLf$; vbCrLf$;
    Printer.Print Tab(1); UCase$(psTekst(3));

Printer.Print vbCrLf$; FULL_LINE

Do While REPORT_TAB(T) <> 0
    Printer.Print Tab(REPORT_TAB(T));
    Printer.Print REPORT_FIELD(T);
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        Printer.Print vbCrLf;
    End If
    T = T + 1
Loop
Printer.Print FULL_LINE; vbCrLf; vbCrLf;

End Sub

Private Sub PrintTotaal()
Dim T As Integer
Dim VeldTekst As String
Dim TotaalGrens As Long

If chkAfdrukInVenster = 0 Then
    Printer.Print vbCrLf; FULL_LINE; vbCrLf;
End If
    
aa = ""
Do While REPORT_TAB(T) <> 0
    If chkAfdrukInVenster = 0 Then
        Printer.Print Tab(REPORT_TAB(T));
    End If
    Select Case RapportManier(T)
        Case 9
            VeldTekst = Dec((KolomTotaal(T)), MaskHier)
            If chkAfdrukInVenster Then
                aa = aa + VeldTekst + vbTab
            Else
                Printer.Print VeldTekst;
                If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
                    Printer.Print vbCrLf;
                End If
            End If
        Case Else
            If chkAfdrukInVenster Then
                aa = aa + " " + vbTab
            End If
    End Select
    T = T + 1
Loop
If chkAfdrukInVenster Then
    Xlog.X.AddItem aa, Xlog.X.Rows - 1
    Xlog.X.Row = 1
    Xlog.X.Col = 0
    Xlog.X.ColWidth(0) = 1140
    Xlog.WindowState = 2
    'Xlog.X.ColAlignment(2) = flexAlignRightTop
    Screen.MousePointer = vbNormal
    Xlog.WijzigenLijn.Visible = False
    Xlog.Afsluiten.Enabled = False
    Xlog.Afsluiten.TabStop = False
    Xlog.cbAfbeelding.Visible = False
    XLogKey = ""
    Xlog.SSTab1.TabVisible(1) = False
    Xlog.Show 1
    Unload Xlog
End If

End Sub

Private Sub PrintVelden()
Dim T As Integer
Dim VeldTekst As String

Err = 0
On Error Resume Next

aa = ""
Do While REPORT_TAB(T) <> 0
    If chkAfdrukInVenster = 0 Then
        Printer.Print Tab(REPORT_TAB(T));
    End If
    
    Select Case RapportManier(T)
        Case 1
            If Not ADO_GET(aIndex, 0, "=", Mid(RV(rsAVBoekHier, "v034"), 2)) Then
                VeldTekst = Mid(RV(rsAVBoekHier, "v034"), 2) + " is niet meer aanwezig"
                MsgBox VeldTekst
            Else
                VeldTekst = Trim$(RV(rsMAR(aIndex), "A110")) + " " + Trim$(RV(rsMAR(aIndex), "A100"))
            End If
        Case 5
            VeldTekst = DATE_TEXT(RV(rsAVBoekHier, "v" + Format(RapportVeldNr(T), "000")))
        Case 9
            VeldTekst = Dec(Val(RV(rsAVBoekHier, "v" + Format(RapportVeldNr(T), "000"))), MaskHier)
            KolomTotaal(T) = KolomTotaal(T) + Val(VeldTekst)
        Case Else
            VeldTekst = RV(rsAVBoekHier, "v" + Format(RapportVeldNr(T), "000"))
    End Select
        
    If chkAfdrukInVenster = 0 Then
        Printer.Print VeldTekst;
        If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
            Printer.Print vbCrLf;
        End If
    Else
        aa = aa + VeldTekst + vbTab
    End If
    T = T + 1
Loop

If chkAfdrukInVenster Then
    Xlog.X.AddItem aa, Xlog.X.Rows - 1
Else
    If Printer.CurrentY >= Printer.Height - 1200 Then
        Printer.NewPage
        Printer.FontSize = Printer.FontSize
        Printer.Print " "
        Printer.FontSize = Printer.FontSize
        PrintTitel
    End If
End If

End Sub


Private Sub TekstLijn_LostFocus(Index As Integer)

Select Case Index
    Case 1
        If DATE_INVALID((TekstLijn1.text)) Then
            Beep
            TekstLijn1.text = MIM_GLOBAL_DATE
            TekstLijn1.SetFocus
        End If
End Select

End Sub

