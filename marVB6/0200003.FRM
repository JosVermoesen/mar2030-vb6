VERSION 5.00
Begin VB.Form AVBoek 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aankoop / Verkoop"
   ClientHeight    =   1485
   ClientLeft      =   1605
   ClientTop       =   2265
   ClientWidth     =   5610
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
   ScaleHeight     =   1485
   ScaleWidth      =   5610
   Begin VB.CheckBox chkAfdrukInVenster 
      Alignment       =   1  'Right Justify
      Caption         =   "Afdruk in venster"
      Height          =   255
      Left            =   3720
      TabIndex        =   13
      Top             =   1170
      Width           =   1830
   End
   Begin VB.CheckBox chkDetailJournaal 
      Alignment       =   1  'Right Justify
      Caption         =   "Detail Journaal"
      Height          =   225
      Left            =   3870
      TabIndex        =   12
      Top             =   900
      Value           =   1  'Checked
      Width           =   1665
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   300
      Width           =   4140
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   2100
      TabIndex        =   4
      Top             =   930
      Width           =   855
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   2880
      TabIndex        =   3
      Top             =   0
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   420
      Width           =   1200
   End
   Begin VB.OptionButton FaktuurCreditnota 
      Caption         =   "Creditnota's"
      Height          =   240
      Index           =   1
      Left            =   60
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Width           =   1395
   End
   Begin VB.OptionButton FaktuurCreditnota 
      Caption         =   "Fakturen"
      Height          =   240
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   720
      Value           =   -1  'True
      Width           =   1395
   End
   Begin VB.CommandButton Drukken 
      Caption         =   "Af&drukken"
      Height          =   300
      Left            =   4350
      TabIndex        =   5
      Top             =   60
      Width           =   1200
   End
   Begin VB.Label TekstVan 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2100
      TabIndex        =   11
      Top             =   660
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Sub&Titel"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   7
      Top             =   60
      Width           =   870
   End
   Begin VB.Label Label1 
      Caption         =   "D&atum Drukken"
      Height          =   195
      Index           =   1
      Left            =   1530
      TabIndex        =   8
      Top             =   60
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Van"
      Height          =   240
      Index           =   2
      Left            =   1560
      TabIndex        =   9
      Top             =   690
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Tot"
      Height          =   240
      Index           =   3
      Left            =   1560
      TabIndex        =   10
      Top             =   960
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

Dim psTekst(5) As String
Dim RapportVeld(23) As String
Dim RapportVeldNr(23) As Integer
Dim RapportManier(23) As Integer
Dim RapportTab(23) As Integer
Dim KolomTotaal(17) As Long

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
Private Sub Annuleren_Click()

Unload Me

End Sub

Private Sub CumulPrint()
Dim RekeningNaam As String * 30
Dim Tabul As Integer
Dim Tel As Integer

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
Printer.Print Tab(2); "** CENTRALISATIE **"
Printer.Print ""

bFirst FlDummy, 0
RecordToVeld FlDummy
bGet FlRekening, 0, Left$(FVT(FlDummy, 0), 7)
If Ktrl Then
    RekeningNaam = "Rekening reeds vernietigd !!!"
Else
    RecordToVeld FlRekening
    RekeningNaam = vBibTekst(FlRekening, "#v020 #")
End If
Tabul = 0
Printer.Print Tab(Tabul + 2); Dec$(Val(vBibTekst(FlDummy, "#v013 #")), "####"); " x "; vSet(vBibTekst(FlDummy, "#v089 #"), 7); " "; RekeningNaam; " "; Dec$(Val(vBibTekst(FlDummy, "#v068 #")), MaskerSy(0));

Do
    bNext FlDummy
    If Ktrl Then
        Exit Do
    End If
    RecordToVeld FlDummy
    bGet FlRekening, 0, Left$(FVT(FlDummy, 0), 7)
    If Ktrl Then
        RekeningNaam = "Rekening reeds vernietigd !!!"
    Else
        RecordToVeld FlRekening
        RekeningNaam = vBibTekst(FlRekening, "#v020 #")
    End If
    If Tabul = 0 Then
        Tabul = 56
        Printer.Print Tab(Tabul + 2); Dec$(Val(vBibTekst(FlDummy, "#v013 #")), "####"); " x "; vSet(vBibTekst(FlDummy, "#v089 #"), 7); " "; RekeningNaam; " "; Dec$(Val(vBibTekst(FlDummy, "#v068 #")), MaskerSy(0)); vbCrLf;
    Else
        Tabul = 0
        Printer.Print Tab(Tabul + 2); Dec$(Val(vBibTekst(FlDummy, "#v013 #")), "####"); " x "; vSet(vBibTekst(FlDummy, "#v089 #"), 7); " "; RekeningNaam; " "; Dec$(Val(vBibTekst(FlDummy, "#v068 #")), MaskerSy(0));
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
            BedragVK2 = Val(Format(BedragForfait(Tel) * PctForfait(Tel), "0"))
            BedragVK = Val(Format(BedragVK2 - (BedragVK2 * KortingForfait(Tel) / 100), "0"))
            Printer.Print Tab(2); Dec((Tel), "00"), Dec(BedragForfait(Tel), "########.00"), Dec(PctForfait(Tel), "####.0000"), Dec(BedragVK2, "########.00"), Dec(BedragVK - BedragVK2, "########.00"), Dec(BedragVK, "#######.00"), Dec((BtwForfait(Tel)), "###.0");
            ForfaitBtw = Val(Format(BedragVK * BtwForfait(Tel) / 100, "0"))
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
End If

bGetOrGreater FlJournaal, 1, vBibTekst(FlDokument, "#v033 #")
If Ktrl Or KeyBuf(FlJournaal) <> vBibTekst(FlDokument, "#v033 #") Then
    MsgBox "Geen detail journaal aanwezig voor dokument " + vBibTekst(FlDokument, "#v033 #") + " !!!"
    Exit Sub
Else
    RecordToVeld FlJournaal
    If Trim(vBibTekst(FlJournaal, "#v038 #")) <> "" Then
    Else
        If Left$(vBibTekst(FlJournaal, "#v019 #"), 2) = "40" Or Left$(vBibTekst(FlJournaal, "#v019 #"), 2) = "44" Then
            RekeningNaam = vBibTekst(FlJournaal, "#v067 #")
        Else
            bGet FlRekening, 0, vSet(vBibTekst(FlJournaal, "#v019 #"), 7)
            If Ktrl Then
                RekeningNaam = "Rekening reeds vernietigd !!!"
            Else
                RecordToVeld FlRekening
                RekeningNaam = vBibTekst(FlRekening, "#v020 #")
                If ForFait Then
                    GoSub ForFaitBerekening
                End If
            End If
        End If
        If chkDetailJournaal Then
            Printer.Print vbCrLf;
            Tabul = 0
            Printer.Print Tab(Tabul + 2); vBibTekst(FlJournaal, "#v019 #"); " "; RekeningNaam; " "; Dec$(Val(vBibTekst(FlJournaal, "#v068 #")), MaskerSy(0));
        End If
        GoSub DetailCumul
    End If
End If

Do
    bNext FlJournaal
    If Ktrl Or KeyBuf(FlJournaal) <> vBibTekst(FlDokument, "#v033 #") Then
        Exit Do
    End If
    RecordToVeld FlJournaal
    If Trim(vBibTekst(FlJournaal, "#v038 #")) <> "" Then
    Else
        If Left$(vBibTekst(FlJournaal, "#v019 #"), 2) = "40" Or Left$(vBibTekst(FlJournaal, "#v019 #"), 2) = "44" Then
            RekeningNaam = vBibTekst(FlJournaal, "#v067 #")
        Else
            bGet FlRekening, 0, vSet(vBibTekst(FlJournaal, "#v019 #"), 7)
            If Ktrl Then
                RekeningNaam = "Rekening reeds vernietigd !!!"
            Else
                RecordToVeld FlRekening
                RekeningNaam = vBibTekst(FlRekening, "#v020 #")
                If ForFait Then
                    GoSub ForFaitBerekening
                End If
            End If
        End If
        If chkDetailJournaal Then
            If Tabul = 0 Then
                Tabul = 56
                Printer.Print Tab(Tabul + 2); vBibTekst(FlJournaal, "#v019 #"); " "; RekeningNaam; " "; Dec$(Val(vBibTekst(FlJournaal, "#v068 #")), MaskerSy(0)); vbCrLf;
                If Printer.CurrentY >= Printer.Height - 1200 Then
                    Printer.NewPage
                    Printer.FontSize = Printer.FontSize
                    Printer.Print " "
                    Printer.FontSize = Printer.FontSize
                    PrintTitel
                End If
            Else
                Tabul = 0
                Printer.Print Tab(Tabul + 2); vBibTekst(FlJournaal, "#v019 #"); " "; RekeningNaam; " "; Dec$(Val(vBibTekst(FlJournaal, "#v068 #")), MaskerSy(0));
            End If
        End If
        GoSub DetailCumul
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
Printer.Print vbCrLf;
Exit Sub

DetailCumul:
StartPunt:
bGet FlDummy, 0, vSet(vBibTekst(FlJournaal, "#v019 #"), 20)
If Ktrl Then
    TLBRecord(FlDummy) = ""
    vBib FlDummy, vBibTekst(FlJournaal, "#v019 #"), "v089"
    vBib FlDummy, "0", "v013"
    vBib FlDummy, "0", "v068"
    bInsert FlDummy, 0
    GoTo StartPunt
Else
    RecordToVeld FlDummy
    vBib FlDummy, Str$(Val(vBibTekst(FlDummy, "#v013 #")) + 1), "v013"
    vBib FlDummy, Str$(Val(vBibTekst(FlDummy, "#v068 #")) + Val(vBibTekst(FlJournaal, "#v068 #"))), "v068"
    bUpdate FlDummy, 0
End If
Return

ForFaitBerekening:
If RTrim$(vBibTekst(FlRekening, "#v216 #")) <> "" Then
    ForfaitNr = Val(vBibTekst(FlRekening, "#v216 #"))
    BedragForfait(ForfaitNr) = BedragForfait(ForfaitNr) + Val(vBibTekst(FlJournaal, "#v068 #"))
    If PctForfait(ForfaitNr) = 0 Then
        bGet FlAllerlei, 1, vSet("21" + vBibTekst(FlRekening, "#v216 #"), 20)
        If Ktrl Then
            MsgBox "Forfaitaire Kode : " + vBibTekst(FlRekening, "#v216 #") + " en waarde nog niet aanwezig.  Eerst inbrengen via diverse gebruikersfiches pér bedrijf a.u.b. !"
        Else
            RecordToVeld FlAllerlei
            PctForfait(ForfaitNr) = Val(vBibTekst(FlAllerlei, "#v218 #"))
            BtwForfait(ForfaitNr) = Val(Mid$(fmarBoxText("002", "2", vBibTekst(FlAllerlei, "#v111 #")), 4, 4))
            KortingForfait(ForfaitNr) = Val(vBibTekst(FlAllerlei, "#v230 #"))
        End If
    End If
End If
Return

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

Select Case aIndex
    Case FlLeverancier
        BeginSleutel = "A"
        EindSleutel = "A"
    Case FlKlant
        BeginSleutel = "V"
        EindSleutel = "V"
    Case Else
        MsgBox "Stop"
End Select
Select Case FaktuurCreditnota(0).Value
    Case -1
        BeginSleutel = BeginSleutel + "0" + Mid$(PeriodeVanTot, 4, 1) + Format$(Val(TekstVan.Caption), "00000")
        EindSleutel = EindSleutel + "0" + Mid$(PeriodeVanTot, 4, 1) + Format$(Val(TekstLijn(3).Text), "00000")
        Tekst$ = FaktuurCreditnota(0).Caption
    Case Else
        BeginSleutel = BeginSleutel + "1" + Mid$(PeriodeVanTot, 4, 1) + Format$(Val(TekstVan.Caption), "00000")
        EindSleutel = EindSleutel + "1" + Mid$(PeriodeVanTot, 4, 1) + Format$(Val(TekstLijn(3).Text), "00000")
        Tekst$ = FaktuurCreditnota(1).Caption
End Select

psTekst(2) = LijstNaam + " " + Tekst$ + " " + Mid(Mim.Caption, InStr(Mim.Caption, "["))
psTekst(0) = TekstLijn(1).Text
psTekst(3) = TekstLijn(0).Text
InitVelden

bGetOrGreater FlDokument, 0, BeginSleutel
If Ktrl Or UCase$(KeyBuf(FlDokument)) > UCase$(EindSleutel) Then
    Beep
    Exit Sub
ElseIf Left$(KeyBuf(FlDokument), 3) <> Left$(BeginSleutel, 3) Then
    Beep
    Exit Sub
Else
    bClose FlDummy
    ClearFlDummy
    Ktrl = bOpen(FlDummy)
    AVBoek.Refresh
    Screen.MousePointer = 11
    RecordToVeld FlDokument
    
    PaginaTeller = 0
    If chkAfdrukInVenster = 0 Then
        Set Printer = Printers(LijstPrinterNr)
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
    PrintVelden
    DetailRekeningen
End If

Do
    bNext FlDokument
    If Ktrl Or KeyBuf(FlDokument) > EindSleutel Then
        Exit Do
    End If
    RecordToVeld FlDokument
    PrintVelden
    DetailRekeningen
Loop
PrintTotaal
CumulPrint
    Printer.NewPage
    Printer.EndDoc
    Screen.MousePointer = 0
    Msg = "Totaliseren voor BTW AANGIFTE.  Bent U zeker?"
    If FaktuurCreditnota(0).Value Then
        Msg = Msg + vbCrLf + vbCrLf + "Opgelet !  Creditnota's niet vergeten straks..."
    End If
    Ktrl = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton1, "BTW Aangifte")
    If Ktrl = vbYes Then
        GoSub BTWAangifte
    End If
    Annuleren_Click
Exit Sub

PrtHandler2:
MsgBox "Kontroleer de printer."
Resume

BTWAangifte:
bGet FlAllerlei, 1, vSet("17" + BJPERDAT.Boekjaar.Text + Format$(BJPERDAT.PeriodeBoekjaar.ListIndex + 1, "00"), 20)
If Ktrl Then
    MsgBox "Stop"
Else
    RecordToVeld FlAllerlei
End If
Select Case aIndex
    Case FlLeverancier
        Select Case Ar
            Case 1
                'Record Kontroleren, zou MOETEN op nul staan...
                vBib FlAllerlei, Str$(KolomTotaal(16)), "v045" 'vak 59
                vBib FlAllerlei, Str$(KolomTotaal(9)), "v052" 'vak 86
                vBib FlAllerlei, Str$(KolomTotaal(11)), "v053" 'vak 87
                vBib FlAllerlei, Str$(KolomTotaal(12)), "v054" 'vak 88
                
                vBib FlAllerlei, Format$(Val(TekstVan.Caption), "00000"), "v092"
                vBib FlAllerlei, Format$(Val(TekstLijn(3).Text), "00000"), "v093"
                
            Case 3
                vBib FlAllerlei, Str$(KolomTotaal(16)), "v100" 'vak 63
                vBib FlAllerlei, Str$(KolomTotaal(7)), "v050" 'vak 84
                vBib FlAllerlei, Str$(KolomTotaal(8)), "v051" 'vak 85
                
                vBib FlAllerlei, Format$(Val(TekstVan.Caption), "00000"), "v094"
                vBib FlAllerlei, Format$(Val(TekstLijn(3).Text), "00000"), "v095"
                For Tel = 3 To 16
                    KolomTotaal(Tel) = -KolomTotaal(Tel)
                Next
            Case Else
                MsgBox "Stop"
        End Select

        vBib FlAllerlei, Str$(KolomTotaal(13) + Val(vBibTekst(FlAllerlei, "#v042 #"))), "v042" 'vak 55
        vBib FlAllerlei, Str$(KolomTotaal(14) + Val(vBibTekst(FlAllerlei, "#v043 #"))), "v043" 'vak 56
        vBib FlAllerlei, Str$(KolomTotaal(15) + Val(vBibTekst(FlAllerlei, "#v044 #"))), "v044" 'vak 57
        vBib FlAllerlei, Str$(KolomTotaal(3) + Val(vBibTekst(FlAllerlei, "#v046 #"))), "v046" 'vak 81
        vBib FlAllerlei, Str$(KolomTotaal(4) + Val(vBibTekst(FlAllerlei, "#v047 #"))), "v047" 'vak 82
        vBib FlAllerlei, Str$(KolomTotaal(5) + Val(vBibTekst(FlAllerlei, "#v048 #"))), "v048" 'vak 83
        
        If ForFait Then
            vBib FlAllerlei, Str$(VakForfait(0)), "v055" 'vak 00
            vBib FlAllerlei, Str$(VakForfait(1)), "v056" 'vak 01
            vBib FlAllerlei, Str$(VakForfait(2)), "v057" 'vak 02
            vBib FlAllerlei, Str$(VakForfait(3)), "v058" 'vak 03
            vBib FlAllerlei, Str$(BtwTotaalForfait), "v064" 'vak 54
        End If
        
    Case FlKlant
        Select Case Ar
            Case 12
                vBib FlAllerlei, Str$(KolomTotaal(12)), "v064" 'vak 54
                vBib FlAllerlei, Str$(KolomTotaal(2) + Val(vBibTekst(FlAllerlei, "#v055 #"))), "v055" 'vak 00 bijtellen ?
                
                vBib FlAllerlei, Str$(KolomTotaal(3)), "v056" 'vak 01
                vBib FlAllerlei, Str$(KolomTotaal(4)), "v057" 'vak 02
                vBib FlAllerlei, Str$(KolomTotaal(5)), "v058" 'vak 03
                vBib FlAllerlei, Str$(KolomTotaal(6)), "v059" 'vak 45
                vBib FlAllerlei, Str$(KolomTotaal(7)), "v060" 'vak 46
                vBib FlAllerlei, Str$(KolomTotaal(8)), "v061" 'vak 47
                
                vBib FlAllerlei, Format$(Val(TekstVan.Caption), "00000"), "v096"
                vBib FlAllerlei, Format$(Val(TekstLijn(3).Text), "00000"), "v097"
                
            Case 14
                vBib FlAllerlei, Str$(KolomTotaal(12)), "v101" 'vak 64
                vBib FlAllerlei, Str$(Val(vBibTekst(FlAllerlei, "#v055 #")) - KolomTotaal(2)), "v055" 'vak 00 aftrekken ?
                
                vBib FlAllerlei, Str$(KolomTotaal(10)), "v062" 'vak 48
                vBib FlAllerlei, Str$(KolomTotaal(11)), "v063" 'vak 49
                
                vBib FlAllerlei, Format$(Val(TekstVan.Caption), "00000"), "v098"
                vBib FlAllerlei, Format$(Val(TekstLijn(3).Text), "00000"), "v099"

            Case Else
                MsgBox "Stop"
        End Select
    Case Else
        MsgBox "Stop"
End Select

bUpdate FlAllerlei, 1

'Fl99Record = TekstLijn(3).Text
Select Case FaktuurCreditnota(0).Value
    Case True 'Faktuur
        If aIndex = FlLeverancier Then
            SS99 TekstLijn(3).Text, 2
            'DummySleutel = "s002 "
        Else
            SS99 TekstLijn(3).Text, 12
            'DummySleutel = "s012 "
        End If

    Case False
        If aIndex = FlLeverancier Then
            SS99 TekstLijn(3).Text, 4
            'DummySleutel = "s004 "
        Else
            SS99 TekstLijn(3).Text, 14
            'DummySleutel = "s014 "
        End If
End Select
'bGet FlTeller, 0, DummySleutel
'If Ktrl Then
'    MsgBox "TellerStop.  kontakteer vsoft"
'Else
'    RecordToVeld FlTeller
'    If BAModus = 1 Then
'        vBib FlTeller, Fl99Record, "v217 "
'    Else
'        vBib FlTeller, Fl99Record, DummySleutel
'    End If
'    bUpdate FlTeller, 0
'      If Ktrl Then
'        MsgBox "UpdateStop Teller. kontakteer vsoft"
'    End If
'End If
Return

End Sub

Private Sub FaktuurCreditNota_Click(Index As Integer)
Dim T       As Integer
Dim Getal   As Integer

Select Case aIndex
    Case FlLeverancier
        T = 0
    Case FlKlant
        T = 10
    Case Else
        MsgBox "Stop"
End Select

Getal = 0
'GET FlBtw, JaarTal + BJPERDAT.PeriodeBoekjaar.ListIndex + 1
Select Case FaktuurCreditnota(0).Value
    Case True 'Faktuur
        Fl99Record = String99(Lees, 1 + T)
        TekstLijn(3).Text = Format$(Val(Fl99Record), "00000")
        Fl99Record = String99(Lees, 2 + T)
        If Format$(Val(Fl99Record), "00000") = TekstLijn(3).Text Then
            Drukken.Enabled = False
            TekstVan.Caption = Format$(Val(Fl99Record), "00000")
        Else
            Drukken.Enabled = True
            TekstVan.Caption = Format$(Val(Fl99Record) + 1, "00000")
        End If
        If aIndex = FlLeverancier Then
            Ar = 1
            Getal = Val(vBibTekst(FlAllerlei, "#v092 #")) + Val(vBibTekst(FlAllerlei, "#v093 #"))
        Else
            Ar = 12
            Getal = Val(vBibTekst(FlAllerlei, "#v096 #")) + Val(vBibTekst(FlAllerlei, "#v097 #"))
        End If

    Case Else
        Fl99Record = String99(Lees, 3 + T)
        TekstLijn(3).Text = Format$(Val(Fl99Record), "00000")
        Fl99Record = String99(Lees, 4 + T)
        If Format$(Val(Fl99Record), "00000") = TekstLijn(3).Text Then
            Drukken.Enabled = False
            TekstVan.Caption = Format$(Val(Fl99Record), "00000")
        Else
            Drukken.Enabled = True
            TekstVan.Caption = Format$(Val(Fl99Record) + 1, "00000")
        End If
        If aIndex = FlLeverancier Then
            Ar = 3
            Getal = Val(vBibTekst(FlAllerlei, "#v094 #")) + Val(vBibTekst(FlAllerlei, "#v095 #"))
        Else
            Ar = 14
            Getal = Val(vBibTekst(FlAllerlei, "#v098 #")) + Val(vBibTekst(FlAllerlei, "#v099 #"))
        End If
End Select
If Getal Then
    MsgBox "Binnen deze periode, zijn er reeds" + vbCrLf + "dokumenten opgenomen !", 0, "BTW aangifte kontroleren a.u.b. !"
    Drukken.Enabled = False
Else
    'Drukken.Enabled = True
End If

End Sub

Private Sub Form_Load()
Dim Dummy           As String
Dim PeriodeSleutel  As String * 20
Dim T               As Integer
Dim Getal           As Long
Dim PeriodeMax      As Integer

Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2

For T = 0 To 17
    KolomTotaal(T) = 0
Next

For T = 0 To 99
    BedragForfait(T) = 0
    PctForfait(T) = 0
Next

PaginaTeller = 0
TekstLijn(1).Text = Rdt
Select Case aIndex
    Case FlLeverancier
        LijstNaam = "Aankoopboek"
    Case FlKlant
        LijstNaam = "Verkoopboek"
    Case Else
        MsgBox "Stop"
End Select
AVBoek.Caption = LijstNaam

Ktrl = 0
SnelHelpPrint "Kontrole hogere periodes..."
PeriodeMax = BJPERDAT.PeriodeBoekjaar.ListCount + 1
Do While PeriodeMax > BJPERDAT.PeriodeBoekjaar.ListIndex + 1
    PeriodeSleutel = "17" + BJPERDAT.Boekjaar.Text + Format$(PeriodeMax, "00")
    bGet FlAllerlei, 1, PeriodeSleutel
    If Ktrl Then
        TLBRecord(FlAllerlei) = ""
        vBib FlAllerlei, (BJPERDAT.Boekjaar.Text), "v090"
        vBib FlAllerlei, Format$(PeriodeMax, "00"), "v091"
        vBib FlAllerlei, "17" + vBibTekst(FlAllerlei, "#v090 #") + vBibTekst(FlAllerlei, "#v091 #"), "v005"
        bInsert FlAllerlei, 1
    Else
        RecordToVeld FlAllerlei
        Getal = 0
        For T = 92 To 99
            Getal = Getal + Val(vBibTekst(FlAllerlei, "#v" + Format$(T, "000") + " #"))
        Next
        If Getal Then
            Getal = PeriodeMax
            PeriodeMax = 0
            Exit Do
        End If
    End If
    PeriodeMax = PeriodeMax - 1
Loop
Jump:
If Getal Then
    MsgBox "Periode " + Format$(Getal, "00") + " reeds afgesloten..."
    Drukken.Visible = False
    Exit Sub
Else
    PeriodeSleutel = "17" + BJPERDAT.Boekjaar.Text + Format$(BJPERDAT.PeriodeBoekjaar.ListIndex + 1, "00")
    bGet FlAllerlei, 1, PeriodeSleutel
    If Ktrl Then
        TLBRecord(FlAllerlei) = ""
        vBib FlAllerlei, (BJPERDAT.Boekjaar.Text), "v090"
        vBib FlAllerlei, Format$(BJPERDAT.PeriodeBoekjaar.ListIndex + 1, "00"), "v091"
        vBib FlAllerlei, "17" + vBibTekst(FlAllerlei, "#v090 #") + vBibTekst(FlAllerlei, "#v091 #"), "v005"
        bInsert FlAllerlei, 1
        GoTo Jump
    Else
        RecordToVeld FlAllerlei
    End If
End If

If Left$(String99(Lees, 20), 1) = "4" Then
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
    RapportVeld(0) = "Dokument"
    RapportTab(0) = 2

RapportVeldNr(1) = 35
    RapportManier(1) = 5    'datum omwerken
    RapportVeld(1) = "Datum dok."
    RapportTab(1) = 11

Select Case aIndex
    Case FlLeverancier
        RapportVeldNr(2) = 39
            RapportManier(2) = 0
            RapportVeld(2) = "Referte"
            RapportTab(2) = 22
        
        RapportVeldNr(3) = 46
            RapportManier(3) = 9    'geheel getal geformateerd
            RapportVeld(3) = "   VAK 81"
            RapportTab(3) = 43

        RapportVeldNr(4) = 47
            RapportManier(4) = 9
            RapportVeld(4) = "   VAK 82"
            RapportTab(4) = 53

        RapportVeldNr(5) = 48
            RapportManier(5) = 9
            RapportVeld(5) = "   VAK 83"
            RapportTab(5) = 63

        RapportVeldNr(6) = 49
            RapportManier(6) = 9
            RapportVeld(6) = "   DERDEN"
            RapportTab(6) = 73

        RapportVeldNr(7) = 50
            RapportManier(7) = 9
            RapportVeld(7) = "   VAK 84"
            RapportTab(7) = 83

        RapportVeldNr(8) = 51
            RapportManier(8) = 9
            RapportVeld(8) = "   VAK 85"
            RapportTab(8) = 93

        RapportVeldNr(9) = 52
            RapportManier(9) = 9
            RapportVeld(9) = "   VAK 86"
            RapportTab(9) = 103

        RapportVeldNr(10) = 99
            RapportManier(10) = 1   'zoek flpartij kode+naam1
            RapportVeld(10) = "ID.Kode/Naam"
            RapportTab(10) = 2

        RapportVeldNr(11) = 53
            RapportManier(11) = 9
            RapportVeld(11) = "   VAK 87"
            RapportTab(11) = 53

        RapportVeldNr(12) = 54
            RapportManier(12) = 9
            RapportVeld(12) = "   VAK 88"
            RapportTab(12) = 63

        RapportVeldNr(13) = 42
            RapportManier(13) = 9
            RapportVeld(13) = "   VAK 55"
            RapportTab(13) = 73
        
        RapportVeldNr(14) = 43
            RapportManier(14) = 9
            RapportVeld(14) = "   VAK 56"
            RapportTab(14) = 83

        RapportVeldNr(15) = 44
            RapportManier(15) = 9
            RapportVeld(15) = "   VAK 57"
            RapportTab(15) = 93
        
        RapportVeldNr(16) = 45
            RapportManier(16) = 9
            RapportTab(16) = 103
            
        Select Case Ar
            Case 1
                RapportVeld(16) = "   VAK 59"
            Case Else
                RapportVeld(16) = "   VAK 63"
        End Select
        RapportTab(17) = 0

    Case FlKlant
        RapportVeldNr(2) = 55
            RapportManier(2) = 9    'geheel getal geformateerd
            RapportVeld(2) = "VAK 00"
            RapportTab(2) = 44

        RapportVeldNr(3) = 56
            RapportManier(3) = 9
            RapportVeld(3) = "VAK 01"
            RapportTab(3) = 55

        RapportVeldNr(4) = 57
            RapportManier(4) = 9
            RapportVeld(4) = "VAK 02"
            RapportTab(4) = 66

        RapportVeldNr(5) = 58
            RapportManier(5) = 9
            RapportVeld(5) = "VAK 03"
            RapportTab(5) = 77

        RapportVeldNr(6) = 59
            RapportManier(6) = 9
            RapportVeld(6) = "VAK 45"
            RapportTab(6) = 88

        RapportVeldNr(7) = 60
            RapportManier(7) = 9
            RapportVeld(7) = "VAK 46"
            RapportTab(7) = 99

        RapportVeldNr(8) = 61
            RapportManier(8) = 9
            RapportVeld(8) = "VAK 47"
            RapportTab(8) = 110

        RapportVeldNr(9) = 99
            RapportManier(9) = 1
            RapportVeld(9) = "ID.Kode/Naam"
            RapportTab(9) = 2

        RapportVeldNr(10) = 62
            RapportManier(10) = 9
            RapportVeld(10) = "VAK 48"
            RapportTab(10) = 77
            
        
        RapportVeldNr(11) = 63
            RapportManier(11) = 9
            RapportVeld(11) = "VAK 49"
            RapportTab(11) = 88

        RapportVeldNr(12) = 64
            RapportManier(12) = 9
            RapportTab(12) = 99
            
        Select Case Ar
            Case 12
                RapportVeld(12) = "VAK 54"
            Case Else
                RapportVeld(12) = "VAK 64"
        End Select
        RapportTab(13) = 0
End Select

For T = 0 To 17
    KolomTotaal(T) = 0
Next

End Sub

Private Sub PrintTitel()
Dim T As Integer

    PaginaTeller = PaginaTeller + 1
    Printer.CurrentY = 400
    Printer.Print Tab(1); psTekst(2); Tab(108); "Pagina : " + Dec$((PaginaTeller), "##########");
    Printer.Print Tab(108); "Datum  : "; psTekst(0); vbCrLf$; vbCrLf$;
    Printer.Print Tab(1); UCase$(psTekst(3));

Printer.Print vbCrLf$; Lijntje

Do While RapportTab(T) <> 0
    Printer.Print Tab(RapportTab(T));
    Printer.Print RapportVeld(T);
    If RapportTab(T + 1) < RapportTab(T) Then
        Printer.Print vbCrLf;
    End If
    T = T + 1
Loop
Printer.Print Lijntje; vbCrLf; vbCrLf;

End Sub

Private Sub PrintTotaal()
Dim T As Integer
Dim VeldTekst As String
Dim TotaalGrens As Long

Printer.Print vbCrLf; Lijntje
Do While RapportTab(T) <> 0
    Printer.Print Tab(RapportTab(T));
    Select Case RapportManier(T)
        Case 9
            VeldTekst = Dec((KolomTotaal(T)), "########0")
            Printer.Print VeldTekst;
            If RapportTab(T + 1) < RapportTab(T) Then
                Printer.Print vbCrLf;
            End If
    End Select
    T = T + 1
Loop

End Sub

Private Sub PrintVelden()
Dim T As Integer
Dim VeldTekst As String

Do While RapportTab(T) <> 0
    Printer.Print Tab(RapportTab(T));
    Select Case RapportManier(T)
        Case 1
            bGet aIndex, 0, Mid$(vBibTekst(FlDokument, "#v034 #"), 2)
            If Ktrl Then
                VeldTekst = Mid$(vBibTekst(FlDokument, "#v034 #"), 2) + " is niet meer aanwezig"
                MsgBox VeldTekst
            Else
                RecordToVeld aIndex
                VeldTekst = Trim(vBibTekst(aIndex, "#A110 #")) + " " + Trim(vBibTekst(aIndex, "#A100 #"))
            End If
                       
        Case 5
            VeldTekst = fDatumText(vBibTekst(FlDokument, "#v" + Format$(RapportVeldNr(T), "000") + " #"))
        Case 9
            VeldTekst = Dec(Val(vBibTekst(FlDokument, "#v" + Format$(RapportVeldNr(T), "000") + " #")), "########0")
            KolomTotaal(T) = KolomTotaal(T) + Val(VeldTekst)
        Case Else
            VeldTekst = vBibTekst(FlDokument, "#v" + Format$(RapportVeldNr(T), "000") + " #")
    End Select
    Printer.Print VeldTekst;
    If RapportTab(T + 1) < RapportTab(T) Then
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

End Sub

Private Sub TekstLijn_GotFocus(Index As Integer)

TekstLijn(Index).SelStart = 0
TekstLijn(Index).SelLength = Len(TekstLijn(Index).Text)

End Sub

Private Sub TekstLijn_LostFocus(Index As Integer)

Select Case Index
    Case 1
        If DatumFout((TekstLijn(1).Text)) Then
            Beep
            TekstLijn(1).Text = Rdt
            TekstLijn(1).SetFocus
        End If
    Case 3
        TekstLijn(3).Text = Format$(Val(TekstLijn(3).Text), "00000")
End Select

End Sub

