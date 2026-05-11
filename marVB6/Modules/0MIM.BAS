Attribute VB_Name = "Module5"
DefInt A-Z


Sub BalansKontrole(Fl As Integer)
Dim Cumul As Currency
Dim dTotaal As Currency
Dim dBetaald As Currency
Dim A As String
Dim T As Integer
Dim Teller As Integer
Dim VoorLetter As String * 1
Dim EnkelOpenstaand As Integer
Dim KontroleString As String

Unload Xlog
Load Xlog
On Local Error Resume Next
Xlog.Caption = "Balans voor : " & objectValue(rsMAR(Fl)("A100"))
Xlog.Caption = Xlog.Caption & " " & objectValue(rsMAR(Fl)("A101"))
SharedScanFl = TABLE_INVOICES

Xlog.X.Rows = 2
Xlog.X.Cols = 7
Xlog.X.Col = 0
Xlog.X.Row = 0
Xlog.X.text = "Document"
Xlog.X.Col = 1
Xlog.X.text = "Totaal"
Xlog.X.Col = 2
Xlog.X.text = "Datum"
Xlog.X.Col = 3
Xlog.X.text = "Fin.Stuk"
Xlog.X.Col = 4
Xlog.X.text = "Betaald"
Xlog.X.Col = 5
Xlog.X.text = "CumulRest"
Xlog.X.Col = 6
Xlog.X.text = "Referte"

Xlog.X.ColWidth(0) = 1150
Xlog.X.ColWidth(1) = 1110
Xlog.X.ColWidth(2) = 980
Xlog.X.ColWidth(3) = 1005
Xlog.X.ColWidth(4) = 975
Xlog.X.ColWidth(5) = 1185
Xlog.X.ColWidth(6) = 2200

Xlog.X.ColAlignment(0) = flexAlignLeftTop
Xlog.X.ColAlignment(1) = flexAlignRightTop
Xlog.X.ColAlignment(2) = flexAlignLeftTop
Xlog.X.ColAlignment(3) = flexAlignLeftTop
Xlog.X.ColAlignment(4) = flexAlignRightTop
Xlog.X.ColAlignment(5) = flexAlignRightTop
Xlog.X.ColAlignment(6) = flexAlignLeftTop

Opnieuw:
EnkelOpenstaand = False
Select Case Fl
    Case TABLE_CUSTOMERS
        VoorLetter = "K"
    Case TABLE_SUPPLIERS
        VoorLetter = "L"
End Select

On Local Error GoTo ErrorOpvang

bClose TABLE_INVOICES
T = 0
bGetOrGreater TABLE_INVOICES, 1, vSet(VoorLetter + vBibTekst(Fl, "#A110 #"), FLINDEX_LEN(TABLE_INVOICES, 1))
If Ktrl Then
    Beep
    Exit Sub
Else
    RecordToVeld TABLE_INVOICES
End If
If RTrim$(KEY_BUF(TABLE_INVOICES)) <> RTrim$(VoorLetter + vBibTekst(Fl, "#A110 #")) Then
    Exit Sub
Else
    Screen.MousePointer = vbHourglass
    GoSub VolgendeLijn
    Do
        bNext TABLE_INVOICES
        If Ktrl Or RTrim$(KEY_BUF(TABLE_INVOICES)) <> RTrim$(VoorLetter + vBibTekst(Fl, "#A110 #")) Then
            Exit Do
        Else
            RecordToVeld TABLE_INVOICES
            GoSub VolgendeLijn
        End If
    Loop
End If
Screen.MousePointer = vbNormal

OpHetScherm:
Xlog.X.Row = 1
Xlog.X.Col = 0

XLogShow:
Xlog.CmdDetailJournaal.Visible = True
Xlog.WijzigenLijn.Visible = False
Xlog.Afsluiten.TabStop = False
Xlog.cbAfbeelding.Visible = False
Xlog.SSTab1.TabVisible(1) = True
XLogKey = ""
Xlog.Show 1
Unload Xlog
SharedScanFl = 0
If XLogKey <> "" Then
    bGet TABLE_INVOICES, 0, Left(XLogKey, 11)
    RecordToVeld TABLE_INVOICES
    T = TeleBibClick(TABLE_INVOICES)
    If XLogKey <> "" Then
        If VSF_PRO Then
        Else
            'bFirst TABLE_JOURNAL, 1
            bGet TABLE_JOURNAL, 1, vBibTekst(TABLE_INVOICES, "#" + JETTABLEUSE_INDEX(TABLE_INVOICES, 0) + "#")
            If Ktrl Then
                      Msg = "Ja = alle TYPE-dokumenten DAT jaar vernietigen !" + vbCrLf
                Msg = Msg + "Nee = enkel DIT dokument verwijderen."
                Ktrl = MsgBox(Msg$, 19 + 512, vBibTekst(TABLE_INVOICES, "#" + JETTABLEUSE_INDEX(TABLE_INVOICES, 0) + "#") + ": dokument vernietigen !")
                    Select Case Ktrl
                    Case 2
                        'niks
                    Case 6
                        RecordToVeld TABLE_INVOICES
                        KontroleString = Left(vBibTekst(TABLE_INVOICES, "#v035 #"), 4)
                        
                        Msg = "Onvoorwaardelijk meerdere dokumenten in reeks vernietigen van jaar " + KontroleString + vbCrLf + vbCrLf
                        Msg = Msg + "Bent U zeker ?"
                        Ktrl = MsgBox(Msg, vbYesNo + vbDefaultButton2, "Opkuis dokumenten jaar " + KontroleString)
                        If Ktrl = vbYes Then
                            bFirst TABLE_INVOICES, 0
                            If Ktrl Then
                            Else
                                Screen.MousePointer = vbHourglass
                                bBegin
                                Do
                                    If Mid(KEY_BUF(TABLE_INVOICES), 3, 1) = Right(KontroleString, 1) Then
                                        RecordToVeld TABLE_INVOICES
                                        If KontroleString = Left(vBibTekst(TABLE_INVOICES, "#v035 #"), 4) Then
                                            SnelHelpPrint KEY_BUF(TABLE_INVOICES), BL_LOGGING
                                            bDelete TABLE_INVOICES
                                        End If
                                    End If
                                    bNext TABLE_INVOICES
                                    If Ktrl Then
                                        Exit Do
                                    End If
                                Loop
                                bEnd
                                Screen.MousePointer = vbNormal
                            End If
                        End If
                                               
                    Case 7
                        bDelete TABLE_INVOICES
                        If Ktrl Then
                            MsgBox "stop"
                        End If
                End Select
            Else
                MsgBox "Er zijn nog journaallijnen van het boekjaar desbetreffend dokument beschikbaar !  Verwijderen via menuoptie 'opkuis bestanden' a.u.b."
            End If
        End If
    End If
End If
Exit Sub

VolgendeLijn:
T = T + 1
Select Case Fl
    Case TABLE_CUSTOMERS
        A = vBibTekst(TABLE_INVOICES, "#v033 #") & vbTab
        dBetaald = Val(vBibTekst(TABLE_INVOICES, "#v037 #"))
        Select Case Left(vBibTekst(TABLE_INVOICES, "#v033 #"), 1)
            Case "V"
                'dTotaal = 0
                'For Teller = 55 To 64
                '    dTotaal = dTotaal + Val(vBibTekst(TABLE_INVOICES, "#v" + Format(Teller, "000") + " #"))
                'Next
                dTotaal = Val(vBibTekst(TABLE_INVOICES, "#v249 #"))
                If Mid(vBibTekst(TABLE_INVOICES, "#v033 #"), 2, 1) = "1" Then
                    dTotaal = -dTotaal
                    dBetaald = -dBetaald
                End If
            Case "Q"
                'dTotaal = Val(vBibTekst(TABLE_INVOICES, "#B010 #")) - Val(vBibTekst(TABLE_INVOICES, "#B090 #"))
                dTotaal = Val(vBibTekst(TABLE_INVOICES, "#v249 #"))
        End Select
        A = A & Format(dTotaal, "#,##0.00") & vbTab
        Cumul = Cumul + dTotaal - dBetaald
        A = A & DATE_TEXT(vBibTekst(TABLE_INVOICES, "#v035 #")) & vbTab
        A = A & vBibTekst(TABLE_INVOICES, "#v038 #") & vbTab
        A = A & Format(dBetaald, "#,##0.00") & vbTab
        A = A & Format(Cumul, "#,##0.00") & vbTab
        If Trim$(vBibTekst(TABLE_INVOICES, "#A000 #")) <> "" Then
            A = A & vBibTekst(TABLE_INVOICES, "#A000 #")
        Else
            A = A & vBibTekst(TABLE_INVOICES, "#v039 #")
        End If
        
    Case TABLE_SUPPLIERS
        A = vBibTekst(TABLE_INVOICES, "#v033 #") & vbTab
        dBetaald = Val(vBibTekst(TABLE_INVOICES, "#v037 #"))
        'dTotaal = Val(vBibTekst(TABLE_INVOICES, "#v048 #")) + Val(vBibTekst(TABLE_INVOICES, "#v049 #")) + Val(vBibTekst(TABLE_INVOICES, "#v047 #")) + Val(vBibTekst(TABLE_INVOICES, "#v046 #")) - Val(vBibTekst(TABLE_INVOICES, "#v044 #")) - Val(vBibTekst(TABLE_INVOICES, "#v043 #")) + Val(vBibTekst(TABLE_INVOICES, "#v045 #"))
        dTotaal = Val(vBibTekst(TABLE_INVOICES, "#v249 #"))
        If Left(vBibTekst(TABLE_INVOICES, "#v033 #"), 2) = "A1" Then
            dTotaal = -dTotaal
            dBetaald = -dBetaald
        End If

        A = A & Format(dTotaal, "#,##0.00") & vbTab
        Cumul = Cumul + dTotaal - dBetaald
        A = A & DATE_TEXT(vBibTekst(TABLE_INVOICES, "#v035 #")) & vbTab
        A = A & vBibTekst(TABLE_INVOICES, "#v038 #") & vbTab
        A = A & Format(dBetaald, "#,##0.00") & vbTab
        A = A & Format(Cumul, "#,##0.00") & vbTab
        A = A & vBibTekst(TABLE_INVOICES, "#v039 #")

End Select
If EnkelOpenstaand Then
    If dBetaald = dTotaal Then
        Return
    End If
End If
Xlog.X.AddItem A, Xlog.X.Rows - 1
Return

ErrorOpvang:
Screen.MousePointer = vbNormal
Msg = "Er zijn reeds" + Str$(T) + " dokumenten !  Teveel voor het geheugen.  Alleen de openstaande dokumenten weergeven ?"
KtrlBox = MsgBox(Msg, 292)
If KtrlBox = 6 Then
    EnkelOpenstaand = True
    Resume Opnieuw
Else
    Resume OpHetScherm
End If

End Sub


Sub CloseOpenWindows()

On Local Error Resume Next

Unload SqlSearch
XDoEvents = DoEvents
Unload Xlog
XDoEvents = DoEvents
Unload SetupEnParameters
Unload Afbeeldingen
Unload xDokument
Unload VrijBericht
'Unload VoorkeurDefinities
'Unload ProgrammaStruktuur
Unload HistoriekSQL
'Unload Importeren
Unload SQLLijsten
'Unload LayOutDokument
'Unload StandaardKostPrijsKaart
Unload Venster

Unload DirekteAankoop
Unload DirekteVerkoop
Unload DiversePosten
Unload InbrengFinancieel

End Sub

Sub NieuwBoekjaar()
Dim BeginEenheden As Double
Dim BeginBedrag As Currency
Dim Vroeger1 As String
Dim Vroeger2 As String
Dim Nu1 As String
Dim Nu2 As String
Dim bString As String * 16
Dim recNr As Integer
Dim eenSaldo As Currency

'MsgBox "Tijdelijk niet beschikbaar"
'Exit Sub

If ACTIVE_BOOKYEAR Then
    MsgBox "Enkel logisch met hoogste boekjaar actief !  Probeer opnieuw..."
    Exit Sub
End If

Dim TempoBestand As String
Dim TempoBstNaam As String

TempoBestand = TABLEDEF_ONT(TABLE_COUNTERS)
TempoBstNaam = bstNaam(TABLE_COUNTERS)

If Dir$(LOCATION_COMPANYDATA + "DEF01.OCT") = "" Then
Else
    bClose TABLE_COUNTERS
    TABLEDEF_ONT(TABLE_COUNTERS) = "01.ONT"
    bstNaam(TABLE_COUNTERS) = "jr" & Mid(bstNaam(TABLE_COUNTERS), 3) - 1
    If Val(String99(READING, 63)) + Val(String99(READING, 64)) <> 2 Then
        MsgBox "Eerst vorig boekjaar in orde brengen !"
        bClose TABLE_COUNTERS
        TABLEDEF_ONT(TABLE_COUNTERS) = TempoBestand
        Exit Sub
    End If
    
    If Val(String99(READING, 62)) <> 1 Then
        Msg = "Eindinventaris van vorig boekjaar overslaan en stock-roulatie vernietigen !  Bent U zeker ?"
        KtrlBox = MsgBox(Msg, 292)
        If KtrlBox = 6 Then
        Else
            bClose TABLE_COUNTERS
            TABLEDEF_ONT(TABLE_COUNTERS) = TempoBestand
            Exit Sub
        End If
    End If
End If
bClose TABLE_COUNTERS
TABLEDEF_ONT(TABLE_COUNTERS) = TempoBestand

Msg = "Een nieuw boekjaar wordt geïnstalleerd hierna !  Bent U zeker ?"
KtrlBox = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2)
If KtrlBox = vbYes Then
Else
    Exit Sub
End If

'Overdracht rekeningsaldo's
bClose TABLE_LEDGERACCOUNTS
bFirst TABLE_LEDGERACCOUNTS, 0
If Ktrl Then
    Beep
    MsgBox "Voortijdige stop..."
    Exit Sub
Else
    MsgBox "De lijst algemene rekeningen wordt hersamengesteld."
    bBegin
    GoSub InstelSaldos
End If

Do
    bNext TABLE_LEDGERACCOUNTS
    If Ktrl Then
        Exit Do
    Else
        GoSub InstelSaldos
    End If
Loop
bEnd
Close TABLE_LEDGERACCOUNTS

'Initialisatie produktenstock
bClose TABLE_PRODUCTS
bFirst TABLE_PRODUCTS, 0
If Ktrl Then
    Beep
    MsgBox "Er zijn geen produkten..."
    GoTo Tellerbestandvervolg
Else
    MsgBox "De stockaankoop/verkooptellers van het vorig boekjaar worden op nul gezet."
    bBegin
    GoSub InstelVoorraad
End If

Do
    bNext TABLE_PRODUCTS
    If Ktrl Then
        Exit Do
    Else
        GoSub InstelVoorraad
    End If
Loop
bEnd
Close TABLE_PRODUCTS

Tellerbestandvervolg:

Dim cat As New ADOX.Catalog
Dim tbl As New ADOX.Table
Dim Idx As New ADOX.Index

'Open the catalog.
cat.ActiveConnection = jetConnect
bstNaam(TABLE_COUNTERS) = "jr" & Format(Val(BYPERDAT.Boekjaar.text) + 1, "0000")
tbl.Name = bstNaam(TABLE_COUNTERS)
    tbl.Columns.Append "v071", adVarWChar, 5
    tbl.Columns.Append "v217", adVarWChar, 30

' Define an index
With Idx
    .Name = FLINDEX_CAPTION(TABLE_COUNTERS, 0)
    .Columns.Append RTrim$(JETTABLEUSE_INDEX(TABLE_COUNTERS, 0))
    .PrimaryKey = False
    .Unique = True
End With
    ' Append the index to the table
    tbl.Indexes.Append Idx
cat.Tables.Append tbl
cat.Tables.Refresh
DoEvents

'Vervolgens Teller- en definitiebestanden opzij zetten
For COUNT_TO = 9 To 1 Step -1
    Vroeger1 = Format(COUNT_TO - 1, "00") + ".ONT"
    Nu1 = Format(COUNT_TO, "00") + ".ONT"

    Vroeger2 = "DEF" + Format(COUNT_TO - 1, "00") + ".OCT"
    Nu2 = "DEF" + Format(COUNT_TO, "00") + ".OCT"
    If Dir$(LOCATION_COMPANYDATA + Vroeger2) = "" Then
        'niks
    Else
        If Dir$(LOCATION_COMPANYDATA + Nu2) <> "" Then
            Kill LOCATION_COMPANYDATA + Nu2
        End If
        Name LOCATION_COMPANYDATA + Vroeger2 As LOCATION_COMPANYDATA + Nu2
    End If
Next
DoEvents

'Nu zou 00.ONT en DEF00.OCT hierdoor weg moeten zijn !!!
'Kopiëren naar programmalokatie, hernoemen 01 naar 00 en
'Terug kopiëren naar bedrijfslokatie
If Not CopyFile(LOCATION_COMPANYDATA, PROGRAM_LOCATION, "DEF01.OCT") Then
    MsgBox "stop"
End If
Name PROGRAM_LOCATION + "DEF01.OCT" As LOCATION_COMPANYDATA + "DEF00.OCT"

Dim rsTempo As ADODB.Recordset
Set rsTempo = New ADODB.Recordset
Dim PogingTeller As Integer

DoEvents
ProbeerNogEens:
Err = 0
On Error Resume Next
rsTempo.Open "SELECT * FROM " + bstNaam(TABLE_COUNTERS), adntDB, adOpenDynamic, adLockOptimistic, adCmdText
If Err Then
    MsgBox "Even wachten..."
    GoTo ProbeerNogEens
End If

bstNaam(TABLE_COUNTERS) = TempoBstNaam
bFirst TABLE_COUNTERS, 0
If Ktrl Then
    MsgBox "er gaat iets fout bij overdracht van tellers"
End If
Do
    RecordToVeld TABLE_COUNTERS
    rsTempo.AddNew
    rsTempo("v071") = vBibTekst(TABLE_COUNTERS, "#v071 #")
    rsTempo("v217") = vBibTekst(TABLE_COUNTERS, "#v217 #")
    rsTempo.Update
    bNext TABLE_COUNTERS
    If Ktrl Then
        Exit Do
    End If
Loop
bClose TABLE_COUNTERS
rsTempo.Close
Set rsTempo = Nothing

bstNaam(TABLE_COUNTERS) = "jr" & Format(Val(BYPERDAT.Boekjaar.text) + 1, "0000")
FL99_RECORD = "0"

'status boekjaar
SetString99 62
SetString99 63
SetString99 64

If Mid(BOOKYEAR_FROMTO, 13, 2) = "12" Then
Else
    MsgBox "Boekjaar eindigt niet in december.  U dient zelf te beslissen of de tellers van aankoop-, verkoop- en financiële documenten dienen op nul gebracht te worden."
    GoTo TotSlot
End If

'aankoopdokumenten
KtrlBox = MsgBox("Aankooptellers op 0 zetten", vbQuestion + vbYesNo + vbDefaultButton2)
If KtrlBox = vbYes Then
    SetString99 1
    SetString99 2
    SetString99 3
    SetString99 4
End If
SetString99 15
SetString99 205

'verkoopdokumenten
KtrlBox = MsgBox("Verkooptellers op 0 zetten", vbQuestion + vbYesNo + vbDefaultButton2)
If KtrlBox = vbYes Then
    SetString99 11
    SetString99 12
    SetString99 13
    SetString99 14
    SetString99 73
    SetString99 59
    SetString99 188
End If

'fin. dokumenten
SetString99 31
SetString99 32
SetString99 33
SetString99 34
SetString99 35
SetString99 38
SetString99 215
SetString99 216
SetString99 217
SetString99 218

TotSlot:
'PeriodeDefinitiebestand DEF00.OCT bijwerken
If BYPERDAT!PeriodeBoekjaar.ListCount - 1 > 12 Then
    MsgBox "Boekjaar bestaande uit " + Format(BYPERDAT!PeriodeBoekjaar.ListCount) + " periodes wordt hierna tot 12 periodes gebracht.  Kontroleer na installatie a.u.b. !"
    Fl = FreeFile
    recNr = 0
    Open LOCATION_COMPANYDATA + "DEF00.OCT" For Random As Fl Len = 16
    For COUNT_TO = BYPERDAT!PeriodeBoekjaar.ListCount - 12 To BYPERDAT!PeriodeBoekjaar.ListCount - 1
         recNr = recNr + 1
         BYPERDAT!PeriodeBoekjaar.ListIndex = COUNT_TO
         SnelHelpPrint PERIOD_FROMTO, BL_LOGGING
         bString = PERIOD_FROMTO
         Mid(bString, 1, 4) = Format(Val(Mid(bString, 1, 4)) + 1, "0000")
         Mid(bString, 9, 4) = Format(Val(Mid(bString, 9, 4)) + 1, "0000")
         'MsgBox PERIOD_FROMTO + vbCrLf + bString
         Put Fl, recNr, bString
    Next
    bString = Space$(16)
    For COUNT_TO = 13 To 99
        Put Fl, COUNT_TO, bString
    Next
    Close Fl
    GoTo Outchecken
End If

Fl = FreeFile
Open LOCATION_COMPANYDATA + "DEF00.OCT" For Random As Fl Len = 16
For COUNT_TO = 1 To 99
    Get Fl, COUNT_TO, bString
    If bString = Space$(16) Then
    Else
        Mid(bString, 1, 4) = Format(Val(Mid(bString, 1, 4)) + 1, "0000")
        Mid(bString, 9, 4) = Format(Val(Mid(bString, 9, 4)) + 1, "0000")
        Put Fl, COUNT_TO, bString
    End If
Next
Close Fl

'compatibiliteit met .NET versie
If Dir$(LOCATION_COMPANYDATA + "DEF*.OXT") = "" Then
Else
    Kill LOCATION_COMPANYDATA + "DEF*.OXT"
End If

Outchecken:
Ktrl = 100
AutoUnloadBedrijf
Exit Sub

InstelSaldos:
RecordToVeld TABLE_LEDGERACCOUNTS
For COUNT_TO = 30 To 22 Step -1
    vBib TABLE_LEDGERACCOUNTS, vBibTekst(TABLE_LEDGERACCOUNTS, "#v" + Format(COUNT_TO, "000") + " #"), "v" + Format(COUNT_TO + 1, "000")
    vBib TABLE_LEDGERACCOUNTS, vBibTekst(TABLE_LEDGERACCOUNTS, "#e" + Format(COUNT_TO, "000") + " #"), "e" + Format(COUNT_TO + 1, "000")
    Err = 0
    On Error Resume Next
    eenSaldo = rsMAR(TABLE_LEDGERACCOUNTS)("dece" + Format(COUNT_TO, "000")).Value
    If Err Then
        eenSaldo = 0
    End If
    rsMAR(TABLE_LEDGERACCOUNTS)("dece" + Format(COUNT_TO + 1, "000")).Value = eenSaldo
Next
vBib TABLE_LEDGERACCOUNTS, Dec(0, MASK_SY(0)), "v022"
vBib TABLE_LEDGERACCOUNTS, Dec(0, MASK_SY(0)), "e022"
rsMAR(TABLE_LEDGERACCOUNTS)("dece022") = 0
SnelHelpPrint rsMAR(TABLE_LEDGERACCOUNTS)("v019") + ", " + rsMAR(TABLE_LEDGERACCOUNTS)("v020"), BL_LOGGING
bUpdate TABLE_LEDGERACCOUNTS, 0
Return
    
InstelVoorraad:
RecordToVeld TABLE_PRODUCTS

BeginEenheden = Val(vBibTekst(TABLE_PRODUCTS, "#v114 #"))
BeginEenheden = BeginEenheden + Val(vBibTekst(TABLE_PRODUCTS, "#v119 #"))
BeginEenheden = BeginEenheden - Val(vBibTekst(TABLE_PRODUCTS, "#v120 #"))

BeginBedrag = Val(vBibTekst(TABLE_PRODUCTS, "#e123 #"))
BeginBedrag = BeginBedrag + Val(vBibTekst(TABLE_PRODUCTS, "#e121 #"))
BeginBedrag = BeginBedrag - Val(vBibTekst(TABLE_PRODUCTS, "#e122 #"))

vBib TABLE_PRODUCTS, Dec$(0, MASK_SY(2)), "v119"
vBib TABLE_PRODUCTS, Dec$(0, MASK_SY(2)), "v120"
vBib TABLE_PRODUCTS, Dec$(BeginEenheden, MASK_SY(2)), "v114"

vBib TABLE_PRODUCTS, Dec$(0, MASK_EURX), "e121"
vBib TABLE_PRODUCTS, Dec$(0, MASK_EURX), "e122"
vBib TABLE_PRODUCTS, Dec$((BeginBedrag), MASK_EURX), "e123"

'vBib TABLE_PRODUCTS, Dec$(0, MASK_EURX), "e121"
'vBib TABLE_PRODUCTS, Dec$(0, MASK_EURX), "e122"
'vBib TABLE_PRODUCTS, Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#v112 #")), MASK_EURX), "e112"
'vBib TABLE_PRODUCTS, Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#v113 #")), MASK_EURX), "e113"
'vBib TABLE_PRODUCTS, Dec$(BeginBedrag / EURO, MASK_EURX), "e123"
SnelHelpPrint vBibTekst(TABLE_PRODUCTS, "#v102 #") + ", " + vBibTekst(TABLE_PRODUCTS, "#v105 #"), BL_LOGGING
bUpdate TABLE_PRODUCTS, 0
Return

End Sub

Sub SetMenuOptions(MenuType As Integer)

Dim BestandsNummer As Integer
Dim FULL_LINE As String
Dim AantalOptiesMogelijk As Integer

If Dir$(LOCATION_COMPANYDATA + "Menu.Def") = "" Then
    GoSub InitialiseerMenu
End If

BestandsNummer = FreeFile
Open LOCATION_COMPANYDATA + "Menu.def" For Input As BestandsNummer
    On Local Error Resume Next
    Line Input #BestandsNummer, FULL_LINE
    For COUNT_TO = 0 To 7
        Line Input #BestandsNummer, FULL_LINE
        Mim.MenuTitel(COUNT_TO).Enabled = Val(Mid(FULL_LINE, 2))
    Next
    For COUNT_TO = 0 To 10
        Line Input #BestandsNummer, FULL_LINE
        Mim.Basis(COUNT_TO).Enabled = Val(Mid(FULL_LINE, 2))
    Next
    For COUNT_TO = 0 To 9
        Line Input #BestandsNummer, FULL_LINE
        Mim.Bewerken(COUNT_TO).Enabled = Val(Mid(FULL_LINE, 2))
        If Val(Mid(FULL_LINE, 2)) = True Then
            AantalOptiesMogelijk = AantalOptiesMogelijk + 1
        End If
    Next
    For COUNT_TO = 0 To 4
        Line Input #BestandsNummer, FULL_LINE
        Mim.Schikmenu(COUNT_TO).Enabled = Val(Mid(FULL_LINE, 2))
        If Val(Mid(FULL_LINE, 2)) = True Then
            AantalOptiesMogelijk = AantalOptiesMogelijk + 1
        End If
    Next
    For COUNT_TO = 0 To 13
        Line Input #BestandsNummer, FULL_LINE
        Mim.Basisbestanden(COUNT_TO).Enabled = Val(Mid(FULL_LINE, 2))
        If Val(Mid(FULL_LINE, 2)) = True Then
            AantalOptiesMogelijk = AantalOptiesMogelijk + 1
        End If
    Next
    For COUNT_TO = 0 To 12
        Line Input #BestandsNummer, FULL_LINE
        Mim.AV(COUNT_TO).Enabled = Val(Mid(FULL_LINE, 2))
        If Val(Mid(FULL_LINE, 2)) = True Then
            AantalOptiesMogelijk = AantalOptiesMogelijk + 1
        End If
    Next
    For COUNT_TO = 0 To 22
        Line Input #BestandsNummer, FULL_LINE
        Mim.Boekhouding(COUNT_TO).Enabled = Val(Mid(FULL_LINE, 2))
        If Val(Mid(FULL_LINE, 2)) = True Then
            AantalOptiesMogelijk = AantalOptiesMogelijk + 1
        End If
    Next
    For COUNT_TO = 0 To 5
        Line Input #BestandsNummer, FULL_LINE
        Mim.Instellingen(COUNT_TO).Enabled = Val(Mid(FULL_LINE, 2))
        If Val(Mid(FULL_LINE, 2)) = True Then
            AantalOptiesMogelijk = AantalOptiesMogelijk + 1
        End If
    Next
    For COUNT_TO = 0 To 7
        Line Input #BestandsNummer, FULL_LINE
        Mim.Info(COUNT_TO).Enabled = Val(Mid(FULL_LINE, 2))
        If Val(Mid(FULL_LINE, 2)) = True Then
            AantalOptiesMogelijk = AantalOptiesMogelijk + 1
        End If
    Next
If Not EOF(BestandsNummer) Then
    MsgBox "Ongeldige bestandsindeling.  Mogelijk van een vorige versie !"
End If
Close BestandsNummer
If AavntalOptiesMogelijk > 10 Then
    MsgBox "NT Light menudefinitie geeft toegang tot " + Format(AantalOptiesMogelijk) + " menuopties.  Indien U toegang wenst tot meer dan 10 menuopties, schaf U de volledige versie van het pakket aan a.u.b. !"
    End
End If
If BYPERDAT.Boekjaar.ListIndex = 0 Then
    BYPERDAT.Boekjaar.Enabled = False
Else
    MsgBox "Een logbestand aanmaken is enkel voor het hoogste boekjaar toegestaan.  Hoogste boekjaar met PROSERV licentie aanduiden en daarna bedrijf opnieuw openen a.u.b. !"
    VsoftLog = 0
End If
Exit Sub

InitialiseerMenu:
'versie 2.14
MsgBox "Menudefinitie LIGHT niet te vinden !  Schakel menuopties aan '-1' of uit ' 0'.  Gebruik KLADBLOK om het tekstbestand " + LOCATION_COMPANYDATA + "Menu.Def" + " te wijzigen."
On Error GoTo 0
BestandsNummer = FreeFile
Open LOCATION_COMPANYDATA + "Menu.def" For Output As BestandsNummer
    Print #BestandsNummer, ":Wijzig enkel de TWEE eerste tekens van elke lijn !  OFF=' 0' ON='-1'"; vbCrLf;
    For COUNT_TO = 0 To 7
        Print #BestandsNummer, "-1;"; Mim.MenuTitel(COUNT_TO).Caption; vbCrLf;
    Next
    
    For COUNT_TO = 0 To 10
        Print #BestandsNummer, "-1;"; Mim.Basis(COUNT_TO).Caption; vbCrLf;
    Next
    
    For COUNT_TO = 0 To 9
        Print #BestandsNummer, " 0;"; Mim.Bewerken(COUNT_TO).Caption; vbCrLf;
    Next
    For COUNT_TO = 0 To 4
        Print #BestandsNummer, " 0;"; Mim.Schikmenu(COUNT_TO).Caption; vbCrLf;
    Next
    For COUNT_TO = 0 To 13
        Print #BestandsNummer, " 0;"; Mim.Basisbestanden(COUNT_TO).Caption; vbCrLf;
    Next
    For COUNT_TO = 0 To 12
        Print #BestandsNummer, " 0;"; Mim.AV(COUNT_TO).Caption; vbCrLf;
    Next
    For COUNT_TO = 0 To 22
        Print #BestandsNummer, " 0;"; Mim.Boekhouding(COUNT_TO).Caption; vbCrLf;
    Next
    For COUNT_TO = 0 To 5
        Print #BestandsNummer, " 0;"; Mim.Instellingen(COUNT_TO).Caption; vbCrLf;
    Next
    For COUNT_TO = 0 To 7
        Print #BestandsNummer, " 0;"; Mim.Info(COUNT_TO).Caption; vbCrLf;
    Next
Close BestandsNummer
COUNT_TO = Shell("notepad.exe " + LOCATION_COMPANYDATA + "Menu.Def", 1)
Return

End Sub

Function AutoBackupOK() As Boolean
    
    AutoBackupOK = False
    
    Dim TQString    As String
    Dim sTBLName    As String
    Dim SQLstring   As String
    
    Dim cnnSubHier As ADODB.Connection
    Dim jetConnectSubHier As String
    Dim RecordTeller As Long
    
    jetConnectSubHier = ADOJET_PROVIDER & _
            "Data Source=" & LOCATION_COMPANYDATA & _
            "\marnt.mdv;" & _
            "Persist Security Info=False"
    
    Set cnnSubHier = New ADODB.Connection
    cnnSubHier.Open jetConnectSubHier
    If Dir(LOCATION_COMPANYDATA & "\marnt.mdb") <> "" Then
        Kill LOCATION_COMPANYDATA & "\marnt.mdb"
    End If
    If Not CopyFile(PROGRAM_LOCATION, LOCATION_COMPANYDATA, "marnt.mdb") Then
        Exit Function
    End If
    
    TQString = OpenSchemeAsString("TABLE")
    Do While TQString <> ""
        sTBLName = Mid(TQString, 1, InStr(TQString, vbCr) - 1)
        TQString = Mid(TQString, InStr(TQString, vbCr) + 1)
        SQLstring = "SELECT * INTO [" & LOCATION_COMPANYDATA & "marnt.mdb].[" & sTBLName & "] FROM " & sTBLName
        Err = 0
        On Error Resume Next
        SnelHelpPrint "Bezig aan tabel " & sTBLName, BL_LOGGING
        Screen.MousePointer = vbHourglass
        cnnSubHier.Execute SQLstring, RecordTeller
        If Err Then
            Exit Function
        End If
        Screen.MousePointer = vbNormal
    Loop
    cnnSubHier.Close
    Set cnnSubHier = Nothing
    If Not CopyFile(LOCATION_COMPANYDATA, LOCATION_MYDOCUMENTS, "marnt.mdb") Then
        Exit Function
    End If
    Dim OldName As String
    Dim NewName As String
    OldName = LOCATION_MYDOCUMENTS & "\marnt.mdb"
    NewName = LOCATION_MYDOCUMENTS & "\marnt" & Format(Now, "yyyymmddhhmmss") & ".mdb"
    Name OldName As NewName   ' Move and rename file.
    AutoBackupOK = True

End Function


Sub pdfPrintUserDef(TypeEnTaal As String, pdfOVSStrook As Double)
    
    Dim FlFree As Integer
    Dim Teller As Integer
    Dim pdfCmd As String

    Dim Ktrl As Double
    
    Dim adresYpos As Double
    Dim adresXpos As Double
    Dim adresYpos2 As Double
    Dim adresXpos2 As Double
    Dim adresBox As Boolean
    
    Dim YPOS As Double
    Dim xPos As Double
    Dim YPOS2 As Double
    Dim xPos2 As Double
    Dim pdfFontSize As Double
    Dim pdfFontName As String
    Dim pdfFontBold As Double
    Dim pdfFontItalic As Double
    Dim pdfFontUnderLine As Double
    Dim pdfColor As Double
    Dim pdfAlign As Double
    Dim pdfBoxPen As Double
    
    Dim textstring As String
    Dim texttmp As String
      
    Dim COUNT_TO As Integer
                
If Dir(LOCATION_COMPANYDATA + "vpeSjbs\pdfDDEF" + TypeEnTaal + ".Txt") = "" Then
    Exit Sub
Else
    With Mim.Report
        .nTopMargin = 1
        .nLeftMargin = 0.5
        .nRightMargin = 20.8
        .nBottomMargin = 29.8
    End With
    FlFree = FreeFile
    Open LOCATION_COMPANYDATA + "vpeSjbs\pdfDDEF" + TypeEnTaal + ".Txt" For Input As FlFree
    
    Do While Not EOF(FlFree)
        Line Input #FlFree, pdfCmd
        If Left(pdfCmd, 1) = "'" Then
        Else
            Select Case Trim(UCase(pdfCmd))
                Case "CMD-VSOFTSPACE"
                    GoSub CMDVSOFTSPACE
                
                Case "CMD-ADRESSPACE"
                    GoSub CMDADRESSPACE
                
                Case "CMD-WRITE"
                    GoSub CMDWRITE
                
                Case "CMD-WRITEBOX"
                    GoSub CMDWRITEBOX
                
                Case "CMD-PRINT"
                    GoSub CMDPRINT
                
                Case "CMD-PICTURE"
                    GoSub CMDPICTURE
              Case Else
                    MsgBox pfcmd & " nog niet voorzien", vbCritical
            End Select
        End If
    Loop
    Close FlFree
End If
Exit Sub

CMDPICTURE:
Input #FlFree, xPos, YPOS, xPos2, YPOS2, pdfBoxPen, filename$
Mim.Report.PenSize = pdfBoxPen
Mim.Report.PictureBestFit = True
If Left(filename$, 4) = "[BL]" Then
    filename$ = LOCATION_COMPANYDATA & Mid(filename$, 5)
ElseIf Left(filename$, 4) = "[PL]" Then
    filename$ = PROGRAM_LOCATION & Mid(filename$, 5)
End If
'2.5, 0.5, 19, 2.5
Ktrl = Mim.Report.Picture(xPos, YPOS, xPos2, YPOS2, filename$)
'.VpePrint(xPos, YPOS, textstring)
Return

CMDVSOFTSPACE:
Input #FlFree, pdfVsoftVanaf, pdfVsoftTot
Return

CMDADRESSPACE:
Input #FlFree, pdfadresXpos, pdfadresYpos, pdfadresXpos2, pdfadresYpos2 ', adresBox
'Ktrl = Mim.Report.WriteBox(adresXpos, adresYpos, adresXpos2, adresYpos2, "")
Return

CMDWRITE:
Input #FlFree, xPos, YPOS, xPos2, YPOS2, pdfFontSize, pdfFontName, pdfColor, pdfAlign, pdfFontBold, pdfFontItalic, pdfFontUnderLine
Line Input #FlFree, textstring
Do
    Line Input #FlFree, texttmp
    If texttmp = "CMD-ENDWRITE" Then
        Exit Do
    Else
        textstring = textstring & vbCrLf & texttmp
    End If
Loop
Mim.Report.FontName = pdfFontName
Mim.Report.FontSize = pdfFontSize
Mim.Report.TextColor = pdfColor
Mim.Report.TextBold = pdfFontBold
Mim.Report.TextItalic = pdfFontItalic
Mim.Report.TextUnderline = pdfFontUnderLine

Mim.Report.TextAlignment = pdfAlign
Ktrl = Mim.Report.Write(xPos, YPOS, xPos2, YPOS2, textstring)
Return
           
CMDWRITEBOX:
Input #FlFree, xPos, YPOS, xPos2, YPOS2, pdfFontSize, pdfFontName, pdfColor, pdfAlign, pdfFontBold, pdfFontItalic, pdfFontUnderLine
Line Input #FlFree, textstring
Do
    Line Input #FlFree, texttmp
    If texttmp = "CMD-ENDWRITE" Then
        Exit Do
    Else
        textstring = textstring & vbCrLf & texttmp
    End If
Loop
Mim.Report.FontName = pdfFontName
Mim.Report.FontSize = pdfFontSize
Mim.Report.TextColor = pdfColor
Mim.Report.TextBold = pdfFontBold
Mim.Report.TextItalic = pdfFontItalic
Mim.Report.TextUnderline = pdfFontUnderLine

Mim.Report.TextAlignment = pdfAlign
Ktrl = Mim.Report.WriteBox(xPos, YPOS, xPos2, YPOS2, textstring)
Return
           
CMDPRINT:
Input #FlFree, xPos, YPOS, pdfFontSize, pdfFontName, pdfColor, pdfFontBold, pdfFontItalic, pdfFontUnderLine
Line Input #FlFree, textstring
Mim.Report.FontName = pdfFontName
Mim.Report.FontSize = pdfFontSize
Mim.Report.TextColor = pdfColor
Mim.Report.TextBold = pdfFontBold
Mim.Report.TextItalic = pdfFontItalic
Mim.Report.TextUnderline = pdfFontUnderLine
If YPOS > pdfVsoftTot Then
    Ktrl = Mim.Report.VPEPRINT(xPos, YPOS - pdfOVSStrook, textstring)
Else
    Ktrl = Mim.Report.VPEPRINT(xPos, YPOS, textstring)
End If
Return

End Sub

