Attribute VB_Name = "Module5"

DefInt A-Z

Function InitBestanden() As Integer
Dim T           As Integer
Dim TT          As Integer
Dim TTT         As Integer
Dim FlTemp      As Integer
Dim aa           As String
Dim PlTT        As Integer
Dim DeNaam      As String

InitBestanden = True
FL_NUMBEROFINDEXEN(TABLE_VARIOUS) = 1
    JETTABLEUSE_INDEX(TABLE_VARIOUS, 0) = "v004 ": FLINDEX_LEN(TABLE_VARIOUS, 0) = 13: FLINDEX_CAPTION(TABLE_VARIOUS, 0) = "Partij"
    JETTABLEUSE_INDEX(TABLE_VARIOUS, 1) = "v005 ": FLINDEX_LEN(TABLE_VARIOUS, 1) = 20: FLINDEX_CAPTION(TABLE_VARIOUS, 1) = "SPtype"

FL_NUMBEROFINDEXEN(TABLE_CUSTOMERS) = 1
    JETTABLEUSE_INDEX(TABLE_CUSTOMERS, 0) = "A110 ": FLINDEX_LEN(TABLE_CUSTOMERS, 0) = 12: FLINDEX_CAPTION(TABLE_CUSTOMERS, 0) = "Nummer"
    JETTABLEUSE_INDEX(TABLE_CUSTOMERS, 1) = "A100 ": FLINDEX_LEN(TABLE_CUSTOMERS, 1) = 10: FLINDEX_CAPTION(TABLE_CUSTOMERS, 1) = "Bedrijfsnaam"

FL_NUMBEROFINDEXEN(TABLE_SUPPLIERS) = 1
    JETTABLEUSE_INDEX(TABLE_SUPPLIERS, 0) = "A110 ": FLINDEX_LEN(TABLE_SUPPLIERS, 0) = 12: FLINDEX_CAPTION(TABLE_SUPPLIERS, 0) = "Nummer"
    JETTABLEUSE_INDEX(TABLE_SUPPLIERS, 1) = "A100 ": FLINDEX_LEN(TABLE_SUPPLIERS, 1) = 10: FLINDEX_CAPTION(TABLE_SUPPLIERS, 1) = "Bedrijfsnaam"

FL_NUMBEROFINDEXEN(TABLE_LEDGERACCOUNTS) = 1
    JETTABLEUSE_INDEX(TABLE_LEDGERACCOUNTS, 0) = "v019 ": FLINDEX_LEN(TABLE_LEDGERACCOUNTS, 0) = 7: FLINDEX_CAPTION(TABLE_LEDGERACCOUNTS, 0) = "RekeningNummer"
    JETTABLEUSE_INDEX(TABLE_LEDGERACCOUNTS, 1) = "v020 ": FLINDEX_LEN(TABLE_LEDGERACCOUNTS, 1) = 10: FLINDEX_CAPTION(TABLE_LEDGERACCOUNTS, 1) = "Omschrijving"
    
FL_NUMBEROFINDEXEN(TABLE_PRODUCTS) = 1
    JETTABLEUSE_INDEX(TABLE_PRODUCTS, 0) = "v102 ": FLINDEX_LEN(TABLE_PRODUCTS, 0) = 13: FLINDEX_CAPTION(TABLE_PRODUCTS, 0) = "Artikelkode EAN"
    JETTABLEUSE_INDEX(TABLE_PRODUCTS, 1) = "v105 ": FLINDEX_LEN(TABLE_PRODUCTS, 1) = 10: FLINDEX_CAPTION(TABLE_PRODUCTS, 1) = "Omschrijving"

FL_NUMBEROFINDEXEN(TABLE_JOURNAL) = 4
    JETTABLEUSE_INDEX(TABLE_JOURNAL, 0) = "v070 ": FLINDEX_LEN(TABLE_JOURNAL, 0) = 15: FLINDEX_CAPTION(TABLE_JOURNAL, 0) = "Rekening Boekdatum"
    JETTABLEUSE_INDEX(TABLE_JOURNAL, 1) = "v033 ": FLINDEX_LEN(TABLE_JOURNAL, 1) = 11: FLINDEX_CAPTION(TABLE_JOURNAL, 1) = "Dokumentnummer"
    JETTABLEUSE_INDEX(TABLE_JOURNAL, 2) = "v038 ": FLINDEX_LEN(TABLE_JOURNAL, 2) = 8: FLINDEX_CAPTION(TABLE_JOURNAL, 2) = "Betalingsstuk"
    JETTABLEUSE_INDEX(TABLE_JOURNAL, 3) = "v041 ": FLINDEX_LEN(TABLE_JOURNAL, 3) = 1: FLINDEX_CAPTION(TABLE_JOURNAL, 3) = "Bewerkingsvlag"
    JETTABLEUSE_INDEX(TABLE_JOURNAL, 4) = "v066 ": FLINDEX_LEN(TABLE_JOURNAL, 4) = 7: FLINDEX_CAPTION(TABLE_JOURNAL, 4) = "Boekdatum"

FL_NUMBEROFINDEXEN(TABLE_INVOICES) = 2
    JETTABLEUSE_INDEX(TABLE_INVOICES, 0) = "v033 ": FLINDEX_LEN(TABLE_INVOICES, 0) = 11: FLINDEX_CAPTION(TABLE_INVOICES, 0) = "DokumentNummer"
    JETTABLEUSE_INDEX(TABLE_INVOICES, 1) = "v034 ": FLINDEX_LEN(TABLE_INVOICES, 1) = 13: FLINDEX_CAPTION(TABLE_INVOICES, 1) = "Partij"
    JETTABLEUSE_INDEX(TABLE_INVOICES, 2) = "A000 ": FLINDEX_LEN(TABLE_INVOICES, 2) = 12: FLINDEX_CAPTION(TABLE_INVOICES, 2) = "KontraktNummer"

FL_NUMBEROFINDEXEN(TABLE_CONTRACTS) = 3
    JETTABLEUSE_INDEX(TABLE_CONTRACTS, 0) = "A000 ": FLINDEX_LEN(TABLE_CONTRACTS, 0) = 12: FLINDEX_CAPTION(TABLE_CONTRACTS, 0) = "Polisnummer"
    JETTABLEUSE_INDEX(TABLE_CONTRACTS, 1) = "A110 ": FLINDEX_LEN(TABLE_CONTRACTS, 1) = 12: FLINDEX_CAPTION(TABLE_CONTRACTS, 1) = "Klantkode"
    JETTABLEUSE_INDEX(TABLE_CONTRACTS, 2) = "A010 ": FLINDEX_LEN(TABLE_CONTRACTS, 2) = 4: FLINDEX_CAPTION(TABLE_CONTRACTS, 2) = "Maatschappij"
    JETTABLEUSE_INDEX(TABLE_CONTRACTS, 3) = "v167 ": FLINDEX_LEN(TABLE_CONTRACTS, 3) = 30: FLINDEX_CAPTION(TABLE_CONTRACTS, 3) = "MaandKlantMijPolis"
    
FL_NUMBEROFINDEXEN(TABLE_COUNTERS) = 0
    JETTABLEUSE_INDEX(TABLE_COUNTERS, 0) = "v071 ": FLINDEX_LEN(TABLE_COUNTERS, 0) = 5: FLINDEX_CAPTION(TABLE_COUNTERS, 0) = "Setup Parameter"
    
FL_NUMBEROFINDEXEN(TABLE_DUMMY) = 0
    JETTABLEUSE_INDEX(TABLE_DUMMY, 0) = "v089 ": FLINDEX_LEN(TABLE_DUMMY, 0) = 20: FLINDEX_CAPTION(TABLE_DUMMY, 0) = "Plaatselijk sorteren"


For T = TABLE_VARIOUS To TABLE_COUNTERS
    bClose T
    bOpen T
    If TeleBibPagina(T) Then
    Else
        MsgBox "Fout tijdens inladen bestandsdefinities.  Herinstalleer het programma en/of contacteer Vsoft"
        InitBestanden = False
    End If
    'On Local Error Resume Next
    If T = TABLE_VARIOUS Or T = TABLE_COUNTERS Then
    Else
        'Fase 1 : Huidige databaseindexen (her)samenstellen
        aa = ""
        
        For TT = 0 To ntDB.TableDefs(bstNaam(T)).Indexes.Count - 1
            aa = aa + ntDB.TableDefs(bstNaam(T)).Indexes(TT).Name + ";"
        Next
        
        'Fase 2 : Standaard definitie aanwezigheid controleren
        For TT = 0 To FL_NUMBEROFINDEXEN(T)
            PlTT = InStr(aa, FLINDEX_CAPTION(T, TT))
            If PlTT Then
                If PlTT = 1 Then
                    aa = Mid(aa, Len(FLINDEX_CAPTION(T, TT)) + 2)
                Else
                    aa = Left(aa, PlTT - 1) + Mid(aa, PlTT + Len(FLINDEX_CAPTION(T, TT)) + 1)
                End If
            ElseIf FLINDEX_CAPTION(T, TT) = "Boekdatum" Then
            Else
                MsgBox "Index '" + FLINDEX_CAPTION(T, TT) + "' van tabel '" + bstNaam(T) + "' bestaat niet meer !!!"
                InitBestanden = False
            End If
        Next
        If aa = "" Then
        Else
          'Deze indexen zijn door gebruiker zelf bijgevoegd
            Do While aa <> ""
                FL_NUMBEROFINDEXEN(T) = FL_NUMBEROFINDEXEN(T) + 1
                FLINDEX_CAPTION(T, FL_NUMBEROFINDEXEN(T)) = Left(aa, InStr(aa, ";") - 1)
                If ntDB.TableDefs(bstNaam(T)).Indexes(FLINDEX_CAPTION(T, FL_NUMBEROFINDEXEN(T))).fields.Count - 1 <> 0 Then
                    MsgBox "Index " + FLINDEX_CAPTION(T, FL_NUMBEROFINDEXEN(T)) + " van tabel " + bstNaam(T) + " is samengesteld uit meerdere velden..." + vbCrLf + "Deze index enkel te gebruiken voor lijsten van " + bstNaam(T) + ".  Bij geïndexeerd zoeken wordt enkel het eerste veld opgenomen in het rooster."
                    JETTABLEUSE_INDEX(T, FL_NUMBEROFINDEXEN(T)) = ntDB.TableDefs(bstNaam(T)).Indexes(FLINDEX_CAPTION(T, FL_NUMBEROFINDEXEN(T))).fields(0).Name
                    For TTT = 1 To ntDB.TableDefs(bstNaam(T)).Indexes(FLINDEX_CAPTION(T, FL_NUMBEROFINDEXEN(T))).fields.Count - 1
                        JETTABLEUSE_INDEX(T, FL_NUMBEROFINDEXEN(T)) = JETTABLEUSE_INDEX(T, FL_NUMBEROFINDEXEN(T)) + "+" + ntDB.TableDefs(bstNaam(T)).Indexes(FLINDEX_CAPTION(T, FL_NUMBEROFINDEXEN(T))).fields(TTT).Name
                    Next
                    FLINDEX_LEN(T, FL_NUMBEROFINDEXEN(T)) = 0
                Else
                    JETTABLEUSE_INDEX(T, FL_NUMBEROFINDEXEN(T)) = ntDB.TableDefs(bstNaam(T)).Indexes(FLINDEX_CAPTION(T, FL_NUMBEROFINDEXEN(T))).fields(0).Name
                    FLINDEX_LEN(T, FL_NUMBEROFINDEXEN(T)) = ntDB.TableDefs(bstNaam(T)).fields(RTrim$(JETTABLEUSE_INDEX(T, FL_NUMBEROFINDEXEN(T)))).Size
                End If
                PlTT = InStr(aa, FLINDEX_CAPTION(T, FL_NUMBEROFINDEXEN(T)))
                If PlTT = 1 Then
                    aa = Mid(aa, Len(FLINDEX_CAPTION(T, FL_NUMBEROFINDEXEN(T))) + 2)
                Else
                    aa = Left(aa, PlTT - 1) + Mid(aa, PlTT + Len(FLINDEX_CAPTION(T, FL_NUMBEROFINDEXEN(T))) + 1)
                End If
            Loop
        End If
    End If
Next
        
End Function

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
Xlog.Caption = "Balans voor : " & oWaarde(rsMAR(Fl)("A100"))
Xlog.Caption = Xlog.Caption & " " & oWaarde(rsMAR(Fl)("A101"))
SharedScanFl = TABLE_INVOICES

Xlog.X.Rows = 2
Xlog.X.Cols = 7
Xlog.X.Col = 0
Xlog.X.Row = 0
Xlog.X.Text = "dokument"
Xlog.X.Col = 1
Xlog.X.Text = "Totaal"
Xlog.X.Col = 2
Xlog.X.Text = "Datum"
Xlog.X.Col = 3
Xlog.X.Text = "Fin.Stuk"
Xlog.X.Col = 4
Xlog.X.Text = "Betaald"
Xlog.X.Col = 5
Xlog.X.Text = "CumulRest"
Xlog.X.Col = 6
Xlog.X.Text = "Referte"

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


Sub AutoLoadBedrijf()
Dim T                   As Integer
Dim FlTemp              As Integer
Dim FlTemp2             As Integer
Dim AktievePeriode      As Integer

Dim XX                  As String
Dim YY                  As String
Dim Titel               As String
Dim A                   As String * 16
Dim aa                  As String * 4
Dim AAA                 As String * 30
Dim AT                  As String

Dim updateLengths20250809 As Boolean
Dim updateLengthsBis20250809 As Boolean
Dim updateLengths20251025 As Boolean

Mim.SnelHelp.Panels(2).Text = "---"
XisEuroWisBEF = False

Static ProducentKopij As String * 8

Err = 0
On Error Resume Next

'3. Menuopties toegang geven
For COUNT_TO = 1 To 5
    Mim.MenuTitel(COUNT_TO).Enabled = True
Next
For T = 1 To 4
    BasisB(T).Enabled = True
Next
Mim.Basis(11).Enabled = True

BJPERDAT.Hide
BJPERDAT.DatumVerwerking.Value = MIM_GLOBAL_DATE
Mim.DatumVerwerking.Value = MIM_GLOBAL_DATE
BJPERDAT.PeriodeBoekjaar.Clear

Ktrl = netVoorbereiden
If Ktrl = True Then
    'MsgBox LOCATION_COMPANYDATA & "is voorbereid voor marintegraal.net"
End If

FlTemp = FreeFile
Open PROGRAM_LOCATION + "9999.OCT" For Random As FlTemp Len = 4
    Get FlTemp, 1, aa
    ACTIVE_BOOKYEAR = 0
    Get FlTemp, 2, aa
    For COUNT_TO = 9 To 0 Step -1
        If Dir$(LOCATION_COMPANYDATA + "DEF" + Format(COUNT_TO, "00") + ".OCT") <> "" Then
            FlTemp2 = FreeFile
            Open LOCATION_COMPANYDATA + "DEF" + Format(COUNT_TO, "00") + ".OCT" For Random As FlTemp2 Len = 16
                Get FlTemp2, 1, A
                XX = Left(A, 4)
            Close FlTemp2
            BJPERDAT.Boekjaar.AddItem XX, 0
        End If
    Next
    Get FlTemp, 3, aa
    AktievePeriode = Val(aa)
Close FlTemp
BJPERDAT.Boekjaar.ListIndex = ACTIVE_BOOKYEAR
bstNaam(TABLE_COUNTERS) = "jr" & BJPERDAT.Boekjaar.Text

FlTemp = FreeFile
Open LOCATION_COMPANYDATA + "DEF" + Format(ACTIVE_BOOKYEAR, "00") + ".OCT" For Random As FlTemp Len = 16
    For T = 1 To 99
        Get FlTemp, T, A
        If A = Space$(16) Then
            BJPERDAT.PeriodeBoekjaar.ListIndex = 0
            YY = BJPERDAT.PeriodeBoekjaar.Text
            BOOKYEAR_FROMTO = Mid(YY, 7, 4) + Mid(YY, 4, 2) + Mid(YY, 1, 2) + Mid(XX, 20, 4) + Mid(XX, 17, 2) + Mid(XX, 14, 2)
            Exit For
        Else
            XX = Mid(A, 7, 2) + "/" + Mid(A, 5, 2) + "/" + Left(A, 4) + " - " + Right(A, 2) + "/" + Mid(A, 13, 2) + "/" + Mid(A, 9, 4)
            BJPERDAT.PeriodeBoekjaar.AddItem XX
        End If
    Next T
    If AktievePeriode - 1 > BJPERDAT.PeriodeBoekjaar.ListCount Then
        MsgBox "Het hoogste boekjaar wordt automatisch ingeladen.  Laatste bewerking gebeurde in een boekjaar met meer periodes dan nu mogelijk.  De eerste periode van het hoogste boekjaar wordt hierna automatisch geaktiveerd"
        AktievePeriode = 1
    End If
    BJPERDAT.PeriodeBoekjaar.ListIndex = AktievePeriode - 1
    AT = BJPERDAT.PeriodeBoekjaar.Text
    PERIOD_FROMTO = Mid(AT, 7, 4) + Mid(AT, 4, 2) + Left(AT, 2) + Right(AT, 4) + Mid(AT, 17, 2) + Mid(AT, 14, 2)
Close FlTemp
BJPERDAT.Show

Mim.Basis(3).Enabled = True
'Mim.Basis(3).Caption = "Bedrijf Sluiten en MarSync Openen"

Set adntDB = New ADODB.Connection
Set adntDBSQLS = New ADODB.Connection

On Error Resume Next
If Mim.Instellingen(4).Checked Then
    Err = 0
    ' voor SQL SERVER
    Set ntDB = NTRuimte.OpenDatabase("", dbDriverPrompt)
    If Err Then
        AutoUnloadBedrijf
        Exit Sub
    ElseIf InStr(ntDB.Connect, "DATABASE=") Then
        'Dim dbNaam As String
        'dbNaam = Mid(ntDB.Connect, InStr$(ntDB.Connect, "DATABASE=") + 9)
        'dbNaam = Left(dbNaam, InStr$(dbNaam, ";") - 1)
    Else
        MsgBox "Er is geen database opgegeven", vbCritical
        AutoUnloadBedrijf
        Exit Sub
    End If
    If InitBestanden Then
    Else
        MsgBox "(Laat de) DATABASE KONTROLEREN !!!"
        AutoUnloadBedrijf
        Exit Sub
    End If
Else
    If Dir$(LOCATION_COMPANYDATA + "Marnt.MDV") = "" Then
        'zoeken is subfolder marNET..
        LOCATION_NETDATA = "marNET\App_Data\"
        If Dir$(LOCATION_COMPANYDATA + LOCATION_NETDATA + "Marnt.MDV") = "" Then
            'problems
            MsgBox "stop"
            AutoUnloadBedrijf
            Screen.MousePointer = vbNormal
            Exit Sub
        Else
            'ok
        End If
    Else
        LOCATION_NETDATA = ""
    End If
        
    If Dir$(LOCATION_COMPANYDATA + LOCATION_NETDATA + "Marnt.MDV") <> "" Then
        'naam jaartabellen aanpassen voor SQL SERVER
        '2001 wordt jr2001 enz...
        Ktrl = TabelKontrole
    
        Err = 0
        ' voor JET
        Set ntDB = NTRuimte.OpenDatabase(LOCATION_COMPANYDATA + LOCATION_NETDATA + "marnt.MDV", False, False)
        Mim.SnelHelp.Panels(3).ToolTipText = LOCATION_COMPANYDATA & LOCATION_NETDATA & "marnt.mdv"
        BAModus = 1
        Screen.MousePointer = vbHourglass
        jetConnect = ADOJET_PROVIDER & _
        "Data Source=" & LOCATION_COMPANYDATA & LOCATION_NETDATA & _
        "\marnt.mdv;" & _
        "Persist Security Info=False"
        If ntDB.Connect <> "" Then
            SnelHelpPrint "Verbinding maken via SQL Server", BL_LOGGING
            adntDB.Open SQLConnect
        Else
            SnelHelpPrint "Database verbinding maken via MS-Jet", BL_LOGGING
            adntDB.Open jetConnect
        End If
        
        Dim tdf  As DAO.TableDef
        Dim fld  As DAO.Field
    
        ' Point to the table and field you care about
        Set tdf = ntDB.TableDefs("Klanten")
        Set fld = tdf.fields("A100")
    
        ' Size holds the defined max length for text fields
        If fld.Size < 50 Then
            updateLengths20250809 = True
            MsgBox "Een reeks velden rond naam en adres zullen worden uitgebreid in het klanten- en leveranciersbestand", vbInformation
        Else
            updateLengths20250809 = False
        End If
                
        Set fld = tdf.fields("A125")
        ' Size holds the defined max length for text fields
        If fld.Size < 50 Then
            updateLengthsBis20250809 = True
        Else
            updateLengthsBis20250809 = False
        End If
        
        ' Point to the table and field you care about
        Set tdf = ntDB.TableDefs("Dokumenten")
        Set fld = tdf.fields("v039")
        
        ' Size holds the defined max length for text fields
        If fld.Size < 35 Then
            updateLengths20251025 = True
            MsgBox "Een reeks velden rond betaalreferte zullen worden uitgebreid in het dokumentenbestand", vbInformation
        Else
            updateLengths20251025 = False
        End If
        
        Dim TQString As String
        Dim sTBLName As String
        
        TQString = OpenSchemaString("TABLE")
        
        'add the Tables
        Do While TQString <> ""
            sTBLName = Mid(TQString, 1, InStr(TQString, vbCr) - 1)
            TQString = Mid(TQString, InStr(TQString, vbCr) + 1)
            If Left(sTBLName, 2) = "jr" Then
                ' Point to the table and field you care about
                Set tdf = ntDB.TableDefs(sTBLName)
                Set fld = tdf.fields("v217")
                If fld.Size < 50 Then
                    bClose TABLE_COUNTERS
                    Err = 0
                    On Error Resume Next
                    Msg = "ALTER TABLE " & sTBLName & " ALTER COLUMN v217 TEXT(50)"
                    adntDB.Execute Msg
                    If Err Then
                        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
                        End
                    End If
                Else
                    Exit Do
                End If
            End If
        Loop
        
        'db.Close
        Set fld = Nothing
        Set tdf = Nothing
        'Set db = Nothing
    
        'SQLConnect = "Provider=SQLNCLI.1;Integrated Security=SSPI;" & _
        '            "Persist Security Info=False;" & _
        '            "AttachDBFileName=" & LOCATION_COMPANYDATA & "\marnet001_Data.mdf;"

        'SQLConnect = "Data Source=breffql891.database.windows.net;Initial Catalog=marCLOUD;Persist Security Info=True;User ID=josvermoesen@breffql891;Password=rvb570428!"
        'adntDBSQLS.Open SQLConnect
        'Stop
        
        If InitBestanden Then
        Else
            MsgBox "(Laat de) DATABASE KONTROLEREN !!!"
            AutoUnloadBedrijf
            Exit Sub
        End If
    End If
End If
             
If ntDB.Connect <> "" Then
    SnelHelpPrint "Verbinding maken via SQL Server CLOUD (SQL Azure)", BL_LOGGING
    adntDB.Open SQLConnect
End If

EerstEnVooral:
'Dim Dummietje As String
Dim Dummietje2 As String

For ktrlfl = TABLE_CUSTOMERS To TABLE_JOURNAL
    ptel = 0
    bClose ktrlfl
    bOpen ktrlfl
    Do While Asc(vBC(ktrlfl, ptel)) <> 0
        'Dummietje = ""
        On Error Resume Next
        Dummietje2 = ""
        Err = 0
        'Dummietje = RTrim$(ntDB.TableDefs(BstNaam(ktrlfl)).Fields(vBC(ktrlfl, ptel)).Name)
        Dummietje2 = rsMAR(ktrlfl)(vBC(ktrlfl, ptel)).Name
        If Err Then
            bClose ktrlfl
            If adxKolom(bstNaam(ktrlfl), vBC(ktrlfl, ptel), adVarWChar, 50) Then
                MsgBox "Extra veld : " + vBC(ktrlfl, ptel) + " met succes bijgevoegd in verzameling : " + bstNaam(ktrlfl)
            Else
                MsgBox "Extra veld : " + vBC(ktrlfl, ptel) + " NIET MET SUCCES bijgevoegd in verzameling : " + bstNaam(ktrlfl)
            End If
            bOpen ktrlfl
        End If
        ptel = ptel + 1
    Loop
Next

Dim TempoVar As String
Dim tempovar2 As String
TempoVar = Trim(String99(READING, 292))
tempovar2 = Trim(String99(READING, 51))
If TempoVar = tempovar2 And Len(TempoVar + tempovar2) = 20 Then
Else
    MsgBox "Breng SPOEDIG Setup Bedrijfsinformatie van de 2 hoogste boekjaren in orde a.u.b.", vbExclamation, "Ontbrekende bedrijfsinfo voor Peppol"
End If

'supplierElectronicMail
If IsValidEmail(Trim(String99(READING, 295))) Then
Else
    MsgBox "Mailadres van uw onderneming is ongeldig.", vbExclamation, "Ontbrekende bedrijfsinfo voor Peppol"
End If

'ALTER TABLE Dokumenten ADD COLUMN v405 MEMO
'ALTER TABLE Dokumenten ADD COLUMN v406 MEMO
bClose TABLE_INVOICES
Err = 0
On Error Resume Next
Dim recDokument As ADODB.Recordset
Set recDokument = New ADODB.Recordset
Msg = "SELECT TOP 1 v405 FROM Dokumenten"
recDokument.CursorLocation = adUseClient
recDokument.Open Msg, adntDB, adOpenForwardOnly, adLockOptimistic
If Err Then
    bClose TABLE_INVOICES
    Err = 0
    On Error Resume Next
    Msg = "ALTER TABLE Dokumenten ADD COLUMN v405 MEMO"
    adntDB.Execute Msg
    If Err = -2147217887 Then
        'Stop
    ElseIf Err = -2147217900 Then
        'Stop
        MsgBox "De database kan niet bijgewerkt worden wanneer in gebruik (netwerk?)" & vbCrLf & "Veiligheidshalve wordt het programma beëindigd.", vbInformation
        End
    ElseIf Err = 0 Then
        Msg = "ALTER TABLE Dokumenten ADD COLUMN v406 MEMO"
        adntDB.Execute Msg
        If Err = 0 Then
        Else
            MsgBox Err.Description
        End If
    Else
        MsgBox "Een foutmelding buiten controle:" & vbCrLf & Err.Description & vbCrLf & "Veiligheidshalve wordt het programma beëindigd.", vbInformation
        End
    End If
End If

'ALTER TABLE Klanten ADD COLUMN v407 MEMO
bClose TABLE_CUSTOMERS
Err = 0
On Error Resume Next
Dim recCustomer As ADODB.Recordset
Set recCustomer = New ADODB.Recordset
Msg = "SELECT TOP 1 v407 FROM Klanten"
recCustomer.CursorLocation = adUseClient
recCustomer.Open Msg, adntDB, adOpenForwardOnly, adLockOptimistic
If Err Then
    bClose TABLE_CUSTOMERS
    Err = 0
    On Error Resume Next
    Msg = "ALTER TABLE Klanten ADD COLUMN v407 MEMO"
    adntDB.Execute Msg
    If Err = -2147217887 Then
        'Stop
    ElseIf Err = -2147217900 Then
        'Stop
        MsgBox "De database kan niet bijgewerkt worden wanneer in gebruik (netwerk?)" & vbCrLf & "Veiligheidshalve wordt het programma beëindigd.", vbInformation
        End
    ElseIf Err = 0 Then
        'ok
    End If
Else
    recCustomer.Close
End If

'ALTER TABLE Leveranciers ADD COLUMN v410 MEMO
bClose TABLE_SUPPLIERS
Err = 0
On Error Resume Next
Dim recSupplier As ADODB.Recordset
Set recSupplier = New ADODB.Recordset
Msg = "SELECT TOP 1 v410 FROM Leveranciers"
recSupplier.CursorLocation = adUseClient
recSupplier.Open Msg, adntDB, adOpenForwardOnly, adLockOptimistic
If Err Then
    bClose TABLE_SUPPLIERS
    Err = 0
    On Error Resume Next
    Msg = "ALTER TABLE Leveranciers ADD COLUMN v410 MEMO"
    adntDB.Execute Msg
    If Err = -2147217887 Then
        'Stop
    ElseIf Err = -2147217900 Then
        'Stop
        MsgBox "De database kan niet bijgewerkt worden wanneer in gebruik (netwerk?)" & vbCrLf & "Veiligheidshalve wordt het programma beëindigd.", vbInformation
        End
    ElseIf Err = 0 Then
        'ok
    End If
Else
    recSupplier.Close
End If

'ALTER TABLE Dokumenten ADD COLUMN v408 MEMO
bClose TABLE_CUSTOMERS
Err = 0
On Error Resume Next
Set recDokument = New ADODB.Recordset
Msg = "SELECT TOP 1 v408 FROM Dokumenten"
recDokument.CursorLocation = adUseClient
recDokument.Open Msg, adntDB, adOpenForwardOnly, adLockOptimistic
If Err Then
    bClose TABLE_INVOICES
    Err = 0
    On Error Resume Next
    Msg = "ALTER TABLE Dokumenten ADD COLUMN v408 MEMO"
    adntDB.Execute Msg
    If Err = -2147217887 Then
        'Stop
    ElseIf Err = -2147217900 Then
        'Stop
        MsgBox "De database kan niet bijgewerkt worden wanneer in gebruik (netwerk?)" & vbCrLf & "Veiligheidshalve wordt het programma beëindigd." & vbCrLf & vbCrLf & "Sluit het bedrijf op alle toestellen en probeer opnieuw", vbInformation
        End
    ElseIf Err = 0 Then
        'ok
    End If
Else
    recDokument.Close
End If


'hier de velden 20250809 bijwerken klanten- en leveranciersbestand
If updateLengths20251025 = True Then
    bClose TABLE_INVOICES
    Err = 0
    On Error Resume Next
    Msg = "ALTER TABLE Dokumenten ALTER COLUMN v039 TEXT(35)"
    adntDB.Execute Msg
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        End
    End If
End If

'hier de velden 20250809 bijwerken klanten- en leveranciersbestand
If updateLengths20250809 = True Then
    bClose TABLE_CUSTOMERS
    Err = 0
    On Error Resume Next
    Msg = "ALTER TABLE Klanten ALTER COLUMN A100 TEXT(50)"
    adntDB.Execute Msg
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        End
    End If
    Msg = "ALTER TABLE Klanten ALTER COLUMN A101 TEXT(25)"
    adntDB.Execute Msg
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        End
    End If
    Msg = "ALTER TABLE Klanten ALTER COLUMN A104 TEXT(50)"
    adntDB.Execute Msg
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        End
    End If
    Msg = "ALTER TABLE Klanten ALTER COLUMN A108 TEXT(50)"
    adntDB.Execute Msg
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        End
    End If
    Msg = "ALTER TABLE Klanten ALTER COLUMN A120 TEXT(50)"
    adntDB.Execute Msg
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        End
    End If
    Msg = "ALTER TABLE Klanten ALTER COLUMN A127 TEXT(25)"
    adntDB.Execute Msg
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        End
    End If
    Msg = "ALTER TABLE Klanten ALTER COLUMN v254 TEXT(100)"
    adntDB.Execute Msg
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        End
    End If
    Msg = "ALTER TABLE Klanten ALTER COLUMN v247 TEXT(100)"
    adntDB.Execute Msg
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        End
    End If
    
    'Leveranciers
    bClose TABLE_SUPPLIERS
    Err = 0
    On Error Resume Next
    Msg = "ALTER TABLE Leveranciers ALTER COLUMN A100 TEXT(50)"
    adntDB.Execute Msg
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        End
    End If
    Msg = "ALTER TABLE Leveranciers ALTER COLUMN A104 TEXT(50)"
    adntDB.Execute Msg
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        End
    End If
    Msg = "ALTER TABLE Leveranciers ALTER COLUMN A108 TEXT(50)"
    adntDB.Execute Msg
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        End
    End If
    Msg = "ALTER TABLE Leveranciers ALTER COLUMN A125 TEXT(50)"
    adntDB.Execute Msg
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        End
    End If
End If

'updateLengthsBis20250809
If updateLengthsBis20250809 = True Then
    'KLANTENBESTAND VERBETERING!:
    'A125: naam2 (vanaf 09/2025 50 tekens) i.p.v. A120!!
    bClose TABLE_CUSTOMERS
    Err = 0
    On Error Resume Next
    Msg = "ALTER TABLE Klanten ALTER COLUMN A125 TEXT(50)"
    adntDB.Execute Msg
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        End
    End If
End If

Volgende:
Mim.SnelHelp.Panels(1).Text = Mid(fmarBoxText("899", "2", Left(String99(READING, 20), 1)), 4)
If Mim.SnelHelp.Panels(1).Text = "marIntegraal NT Light" Then
    MsgBox PROGRAM_LOCATION & "NTIMPORT.LOG is geactiveerd.", vbInformation
End If
Mim.AV(1).Enabled = True
Mim.Boekhouding(3).Enabled = True
Dim txtTest As String
txtTest = LaadTekst("Algemeen", "BoekInfoModus")
If Len(txtTest) = 0 Then
    If TestEuroModus Then
        Mim.cmdWegBoekModus.Visible = True
        Mim.cmdWegBoekModus.Enabled = True
        Mim.cmdWegBoekModus.ListIndex = 2
    Else
        Mim.cmdWegBoekModus.ListIndex = 0
        Mim.cmdWegBoekModus.Visible = True
        Mim.cmdWegBoekModus.Enabled = False
    End If
Else
    Mim.cmdWegBoekModus.ListIndex = Val(Left(txtTest, 1))
End If

ProducentNummer = ""
Select Case Left(String99(READING, 20), 1)
    Case "0", "M"
        VsoftLog = 0
    Case ""
        If BJPERDAT.Boekjaar.ListIndex <> 0 Then MsgBox "Draai de Light versie enkel voor het hoogste boekjaar a.u.b. !"
        VsoftLog = 0
    Case "1"
        If BJPERDAT.Boekjaar.ListIndex <> 0 Then MsgBox "Draai de Light versie enkel voor het hoogste boekjaar a.u.b. !"
        VsoftLog = 1
        SetMenuOptions 1
    Case "2"
        VsoftLog = 0
    Case "4"
        'Mim.AV(1).Enabled = False
        Mim.Boekhouding(3).Enabled = False
    Case "5"
        Mim.AV(11).Enabled = True
        ProducentNummer = ProducentKopij
    Case Else
        'MsgBox "Stop"
End Select
If ProducentNummer <> Space$(8) Then
    ProducentKopij = ProducentNummer
    Mim.MenuTitel(6).Enabled = True
Else
    Mim.MenuTitel(6).Enabled = False
End If


If Len(Trim(String99(READING, 296))) = 0 Then
    MsgBox "Gelieve Setup Boekingen en algemene instellingen : munt van de Boekhouding in te stellen a.u.b.  Pér bedrijf, pér boekjaar.  Hierna wordt voorlopig verder gewerkt in EUR."
    bhEuro = True
    SS99 "EUR", 296
ElseIf String99(READING, 296) = "BEF" Then
    bhEuro = False
ElseIf String99(READING, 296) = "EUR" Then
    bhEuro = True
Else
    SnelHelpPrint "Onlogische situatie", BL_LOGGING
    bhEuro = False
End If
If bhEuro = False Then
    MsgBox "Hoogste boekjaar enkel nog in EUR vanaf versie 6.5.301 of hoger.  Indien U nog BEF verrichtingen wenst uit te voeren, gelieve een eerdere versie opnieuw te installeren a.u.b.", vbInformation
    End
Else
    Mim.SnelHelp.Panels(2).Text = "EUR"
End If

PeppolDirCheck:
Dim peppolDir As String
peppolDir = LOCATION_COMPANYDATA + "peppol"
On Local Error Resume Next
ChDir peppolDir

If Err = 76 Then
    'Msg = "Vanaf versie 11.3.026 dient de datainhoudsopgave over submap 'peppol' te beschikken." & vbCrLf & vbCrLf
    'Msg = Msg & "Deze map wordt hierna aangemaakt"
    'MsgBox Msg, vbInformation
    MkDir (LOCATION_COMPANYDATA + "peppol")
    ChDir peppolDir
    MkDir (LOCATION_COMPANYDATA + "peppol\in")
    ChDir peppolDir
    MkDir (LOCATION_COMPANYDATA + "peppol\out")
    ChDir LOCATION_COMPANYDATA
Else
    ChDir LOCATION_COMPANYDATA
End If
On Local Error GoTo 0


QRCodeDIRCheck:
Dim qrDir As String
qrDir = LOCATION_COMPANYDATA + "BMP-qr"
On Local Error Resume Next
ChDir qrDir

If Err = 76 Then
    'Msg = "Vanaf versie 11.2.007 dient de datainhoudsopgave over submap 'BMP-qr' te beschikken." & vbCrLf & vbCrLf
    'Msg = Msg & "Deze map wordt hierna aangemaakt"
    'MsgBox Msg, vbInformation
    MkDir (LOCATION_COMPANYDATA + "BMP-qr")
    ChDir LOCATION_COMPANYDATA
Else
    ChDir LOCATION_COMPANYDATA
End If
On Local Error GoTo 0

VPESJBSCheck:
Dim vpeSjbsDir As String
vpeSjbsDir = LOCATION_COMPANYDATA + "vpeSjbs"
On Local Error Resume Next
ChDir vpeSjbsDir

If Err = 76 Then
    'Msg = "Vanaf versie 11.3.000 dient de datainhoudsopgave van een bedrijf over submap 'vpeSjbs' te beschikken." & vbCrLf & vbCrLf
    'Msg = Msg & "Deze map wordt hierna aangemaakt"
    'MsgBox Msg, vbInformation
    MkDir (LOCATION_COMPANYDATA + "vpeSjbs")
    ChDir LOCATION_COMPANYDATA
Else
    ChDir LOCATION_COMPANYDATA
End If
On Local Error GoTo 0

XLSXCheck:
Dim xlsxDir As String
xlsxDir = LOCATION_COMPANYDATA + "xlsx-templates"
On Local Error Resume Next
ChDir xlsxDir

If Err = 76 Then
    MkDir (LOCATION_COMPANYDATA + "xlsx-templates")

    'copy the sample templates
    If Not CopyFile(App.path & "\xlsx-templates", LOCATION_COMPANYDATA + "xlsx-templates\", "budget.xml") Then
        'Stop
    End If
    If Not CopyFile(App.path & "\xlsx-templates", LOCATION_COMPANYDATA + "xlsx-templates\", "intrest.xml") Then
        'Stop
    End If
    ChDir LOCATION_COMPANYDATA
    
Else
    ChDir LOCATION_COMPANYDATA
End If
On Local Error GoTo 0

VATCheck:
Dim vatDir As String
vatDir = LOCATION_COMPANYDATA + "vat"
On Local Error Resume Next
ChDir vatDir

If Err = 76 Then
    'Msg = "Vanaf versie 11.3.007 dient de datainhoudsopgave van een bedrijf over submap 'vat' te beschikken." & vbCrLf & vbCrLf
    'Msg = Msg & "Deze map wordt hierna aangemaakt"
    'MsgBox Msg, vbInformation
    MkDir (LOCATION_COMPANYDATA + "vat")
    ChDir LOCATION_COMPANYDATA
Else
    ChDir LOCATION_COMPANYDATA
End If
On Local Error GoTo 0

VPEDOCCheck:
Dim vpeSjbsInRoot As String
vpeSjbsInRoot = LOCATION_COMPANYDATA + "pdfDDEF*.txt"
Dim docName As String

On Local Error Resume Next

docName = Dir(vpeSjbsInRoot)
If docName = "" Then
Else
    Msg = "Vanaf versie 11.3.001 dienen pdfDDEF definities zich te bevinden in de vpeSjbs inhoudsopgave van een bedrijf." & vbCrLf & vbCrLf
    Msg = Msg & "De transfert van bestaande definities wordt hierna uitgevoerd."
    MsgBox Msg, vbInformation
    Do While docName <> ""
        Ktrl = CopyFile(LOCATION_COMPANYDATA, LOCATION_COMPANYDATA & "vpeSjbs\", docName)
        If Ktrl = 0 Then
            MsgBox "fout tijdens kopieren"
        Else
            Kill LOCATION_COMPANYDATA & docName
            If Err Then
                MsgBox Err.Description
            End If
        End If
        docName = Dir(vpeSjbsInRoot)
    Loop
End If
On Local Error GoTo 0

BEBtwNummerCheck:
Dim beBtwCheck As String
beBtwCheck = LOCATION_COMPANYDATA + "btwBEChecked-donotremove.txt"
Dim checkFileAvailable As String
Dim recClient As ADODB.Recordset
Dim recLeverancier As ADODB.Recordset
Dim AantalRC As Long

On Local Error Resume Next
checkFileAvailable = Dir(beBtwCheck)
If Len(checkFileAvailable) <> 0 Then
Else
    Msg = "Vanaf versie 11.3.005 wordt het Belgisch BTW nummer van uw klanten bewaard als 10 cijfers (zonder leestekens!)" & vbCrLf & vbCrLf & "bvb. 0234234234 of 1234234234 enz." & vbCrLf & vbCrLf
    Msg = Msg & "De omwerking wordt hierna uitgevoerd."
    MsgBox Msg, vbOKOnly + vbInformation, "Eénmalig controle btw nummers"
    
    Screen.MousePointer = vbHourglass
    
    On Error Resume Next
    Set recClient = New ADODB.Recordset
    Err = 0
    Msg = "SELECT * FROM Klanten"
    recClient.CursorLocation = adUseClient
    recClient.Open Msg, adntDB, adOpenForwardOnly, adLockOptimistic
    If Err Then MsgBox "stop"
    recClient.MoveFirst
    
    Dim btwKlantVeldje As String
    Dim tmpKlantVeldje As String
    Dim manueelKlantVeldje As String
    Dim manueelList As String
    Dim okKlantVeldje As String
    Dim okList As String
    Dim warningMessage As String
    Dim warningMessage2 As String
    Dim warningMessage3 As String
    
    warningMessage = "Sinds 1985 tot heden werd notatie verwacht 000.000.000 of 000-000-000" & vbCrLf & "marIntegraal plaatste daar later automatsch een '0' en 'BE' voor." & vbCrLf & vbCrLf & "Na update manueel zelf verbeteren a.u.b. voor deze klant"
    warningMessage2 = "Sinds 1985 tot heden werd notatie verwacht 000.000.000 of 000-000-000" & vbCrLf & "Btw nummer 9 tekens lang? Een '0' wordt toegevoegd." & vbCrLf & vbCrLf & "Na update manueel controleren en verbeteren a.u.b. voor deze klant"
    warningMessage3 = "Sinds 1985 tot heden werd notatie verwacht 000.000.000 of 000-000-000" & vbCrLf & "Btw nummer reeds 10 tekens lang?" & vbCrLf & vbCrLf & "Na update manueel controleren en verbeteren a.u.b. voor deze klant"
        
    On Error GoTo 0
    Do While Not recClient.EOF
        If recClient("v150") = "BE" Then
            If Trim(recClient("A161")) = "" Or IsNull(recClient("A161")) Then
            Else
                btwKlantVeldje = Trim(recClient("A161"))
                If Len(btwKlantVeldje) = 11 Then 'zou zo moeten zijn, notering sinds 1985!
                    tmpKlantVeldje = "0" + Mid(btwKlantVeldje, 1, 3) + Mid(btwKlantVeldje, 5, 3) + Mid(btwKlantVeldje, 9, 3)
                    okKlantVeldje = Trim(recClient(0)) & " " & Trim(recClient(2)) & " " & btwKlantVeldje
                    recClient("A161").Value = tmpKlantVeldje
                    recClient("dnnSync").Value = False
                    okList = okList + okKlantVeldje + vbCrLf
                ElseIf Len(btwKlantVeldje) = 9 Then
                    manueelKlantVeldje = Trim(recClient(0)) & " " & Trim(recClient(2)) & " " & btwKlantVeldje
                    MsgBox manueelKlantVeldje & vbCrLf & vbCrLf & warningMessage2
                    manueelList = manueelList + manueelKlantVeldje + vbCrLf
                    recClient("A161").Value = "0" + btwKlantVeldje
                    recClient("dnnSync").Value = False
                ElseIf Len(btwKlantVeldje) = 10 Then
                    manueelKlantVeldje = Trim(recClient(0)) & " " & Trim(recClient(2)) & " " & btwKlantVeldje
                    MsgBox manueelKlantVeldje & vbCrLf & vbCrLf & warningMessage3
                    manueelList = manueelList + manueelKlantVeldje + vbCrLf
                Else
                    manueelKlantVeldje = Trim(recClient(0)) & " " & Trim(recClient(2)) & " " & btwKlantVeldje
                    MsgBox manueelKlantVeldje & vbCrLf & vbCrLf & warningMessage
                    manueelList = manueelList + manueelKlantVeldje + vbCrLf
                End If
            End If
        End If
        recClient.MoveNext
    Loop
    Screen.MousePointer = vbNormal
    MsgBox "Einde update BE BTW nummers"
    recClient.Close
    Set recClient = Nothing
    
    flHier = FreeFile
    Open beBtwCheck For Output As flHier
        Print #flHier, "done"
    Close flHier
    
    Dim btwOudNummerBedrijf As String
    Dim btwNieuwNummerBedrijf As String
    
    btwOudNummerBedrijf = Trim(String99(READING, 51))
    If Len(btwOudNummerBedrijf) = 11 Then 'zou zo moeten zijn, notering sinds 1985 tot heden!
        btwNieuwNummerBedrijf = "0" + Mid(btwOudNummerBedrijf, 1, 3) + Mid(btwOudNummerBedrijf, 5, 3) + Mid(btwOudNummerBedrijf, 9, 3)
        SS99 btwNieuwNummerBedrijf, 51
    Else
        'MsgBox btwOudNummerBedrijf & " ???" & vbCrLf & vbCrLf & "Er klopt iets niet! Gelieve zelf uw btw nummer in setup boekjaar en parameters correct te vermelden."
    End If
    
    flHier = FreeFile
    Open LOCATION_COMPANYDATA + "ManueelControleren.txt" For Output As flHier
        Print #flHier, manueelList
    Close flHier
    DoEvents
    Shell "notepad.exe " + LOCATION_COMPANYDATA + "ManueelControleren.txt", vbNormalFocus
            
    flHier = FreeFile
    Open LOCATION_COMPANYDATA + "BtwOk.txt" For Output As flHier
        Print #flHier, okList
    Close flHier
    DoEvents
    Shell "notepad.exe " + LOCATION_COMPANYDATA + "BtwOk.txt", vbNormalFocus
End If
    
On Local Error GoTo 0

'versie 11.3.015
Dim aantalON As Integer

'CompanyNumberChecked-doNotRemove
ONNummerCheck:
Dim ONNRCheck As String
ONNRCheck = LOCATION_COMPANYDATA + "CompanyNumberChecked-doNotRemove.txt"
Dim checkONFileAvailable As String

On Local Error Resume Next
checkONFileAvailable = Dir(ONNRCheck)

If Len(checkONFileAvailable) <> 0 Then
Else
    'MsgBox "Tabellen Klanten en Leveranciers." & vbCrLf & vbCrLf & "Hierna wordt het btwnummer eveneens aangeduid als ondernemingsnummer. Gelieve zelf manueel te controleren.", vbInformation
    
    Screen.MousePointer = vbHourglass
    bClose TABLE_CUSTOMERS
    
    On Error Resume Next
    Set recClient = New ADODB.Recordset
    Err = 0
    Msg = "SELECT * FROM Klanten"
    recClient.CursorLocation = adUseClient
    recClient.Open Msg, adntDB, adOpenForwardOnly, adLockOptimistic
    If Err Then MsgBox "stop"
    recClient.MoveFirst
                    
    aantalON = 0
    On Error GoTo 0
    Do While Not recClient.EOF
        If Trim(recClient("A161")) = "" Then
        Else
            aantalON = aantalON + 1
            recClient("v404").Value = recClient("A161").Value
        End If
        recClient.MoveNext
    Loop
    Screen.MousePointer = vbNormal
    'MsgBox "Einde update Ondernemingsnummers voor " & aantalON & " ondernemingen klanten"
    recClient.Close
    Set recClient = Nothing
    
        Screen.MousePointer = vbHourglass
    
    On Error Resume Next
    Set recLeverancier = New ADODB.Recordset
    Err = 0
    Msg = "SELECT * FROM Leveranciers"
    recLeverancier.CursorLocation = adUseClient
    recLeverancier.Open Msg, adntDB, adOpenForwardOnly, adLockOptimistic
    If Err Then MsgBox "stop"
    recLeverancier.MoveFirst
                    
    aantalON = 0
    On Error GoTo 0
        
    Do While Not recLeverancier.EOF
        If Trim(recLeverancier("A161")) = "" Then
        Else
            aantalON = aantalON + 1
            recLeverancier("v404").Value = recLeverancier("A161").Value
        End If
        recLeverancier.MoveNext
    Loop
    Screen.MousePointer = vbNormal
    'MsgBox "Einde update Ondernemingsnummers voor " & aantalON & " ondernemingen leveranciers"
    recLeverancier.Close
    Set recLeverancier = Nothing
    flHier = FreeFile
    Open ONNRCheck For Output As flHier
    Print #flHier, "done"
    Close flHier
End If


X = bOpen(TABLE_JOURNAL)
If Left(usrLicentieInfo, 9) = "DemoModus" Then
    If rsMAR(TABLE_JOURNAL).RecordCount > 50 Then
        JournaalLocked = True
    Else
        JournaalLocked = False
    End If
End If

For COUNT_TO = 22 To 31
    KtrlLong = VeldOK(TABLE_LEDGERACCOUNTS, "dece" & Format(COUNT_TO, "000"), "CURRENCY")
Next

If InStr(UCase(Command), "/J=TRUE") Then
    MsgBox Command
    GoTo JumpToJournaal
End If

'If VeldOK(TABLE_INVOICES, "e069") <> 0 Then
'    KtrlLong = VeldOK(TABLE_INVOICES, "e069", "TEXT[12]")
'End If
'If VeldOK(TABLE_CONTRACTS, "dece069") <> 0 Then
'    KtrlLong = VeldOK(TABLE_CONTRACTS, "e069", "TEXT[12]")
'End If

Dim resultDir As Boolean
destpath$ = LOCATION_COMPANYDATA & "sjb"
Err = 0
On Local Error Resume Next
If Dir(destpath & "\sBrief.doc") = "" Then
    resultDir = CreatePath(destpath$)
    If Not CopyFile(App.path, destpath$, "sbrief.doc") Then
       MsgBox "Fout tijdens installatie briefsjabloon: sbrief.doc", vbExclamation
    End If
End If

'versie 7.1.2xx
If VeldOK(TABLE_CUSTOMERS, "v263") = 0 Then
    'het bestaat al dus ok
Else
    MsgBox "Tabel Klanten dient vanaf versie 7.1.200 te beschikken over veld v263 (Outlook EntryID).  Hierna wordt automatisch de herwerking uitgevoerd.", vbInformation
    bClose TABLE_CUSTOMERS
    Msg = "ALTER TABLE Klanten ADD COLUMN v263 TEXT[64];"
    Err = 0

    adntDB.Execute Msg, AantalRC
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        End
    Else
        MsgBox Msg & " met succes uitgevoerd", vbInformation
    End If

    MsgBox "Tabel Klanten dient vanaf versie 7.1.200 te beschikken over veld v262 (Outlook EntryID).  Hierna wordt automatisch de herwerking uitgevoerd.", vbInformation
    bClose TABLE_CUSTOMERS
    Msg = "ALTER TABLE Klanten ADD COLUMN v262 TEXT[64];"
    Err = 0

    adntDB.Execute Msg, AantalRC
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        End
    Else
        MsgBox Msg & " met succes uitgevoerd", vbInformation
    End If
    MsgBox "Tabel Leveranciers dient vanaf versie 7.1.200 te beschikken over veld v262 (Outlook EntryID).  Hierna wordt automatisch de herwerking uitgevoerd.", vbInformation
    bClose TABLE_SUPPLIERS
    Msg = "ALTER TABLE Leveranciers ADD COLUMN v262 TEXT[64];"
    Err = 0

    adntDB.Execute Msg, AantalRC
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        End
    Else
        MsgBox Msg & " met succes uitgevoerd", vbInformation
    End If
End If

'versie 6.6.8xx
If VeldOK(TABLE_JOURNAL, "v102") = 0 Then
    'het bestaat al
Else
    MsgBox "Journaaltabel dient vanaf versie 6.6.800 te beschikken over veld v102 (productnummer).  Hierna wordt automatisch de herwerking uitgevoerd.", vbInformation
    bClose TABLE_JOURNAL
    Msg = "ALTER TABLE Journalen ADD COLUMN v102 TEXT[13];"
    Err = 0
    adntDB.Execute Msg, AantalRC
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        End
    Else
        MsgBox Msg & " met succes uitgevoerd", vbInformation
    End If
End If

'versie 6.5.3xx
If VeldOK(TABLE_JOURNAL, "dece068") = 0 Then
    'het bestaat al
Else
    MsgBox "Alle rekening- en journaalvelden voor cijfers/bedragen dienen vanaf versie 6.5.301 formaat DECIMAL te zijn (of ook CURRENCY).  Hierna worden automatisch de herwerkingen uitgevoerd.  STRAKS VOORAL AANDACHT voor juiste datum overstap naar EURO a.u.b.  GEEN NOOD ECHTER indien U zich vergist, U kan steeds opnieuw hersamenstellen tot uiterlijk 6.5.400" & vbCr & vbCr & "Aarzel niet ons te contacteren voor bijkomende inlichtingen.", vbInformation
    bClose TABLE_JOURNAL
    Msg = "ALTER TABLE Journalen ADD COLUMN dece068 CURRENCY;"
    Err = 0
    adntDB.Execute Msg, AantalRC
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        End
    Else
        MsgBox Msg & " met succes uitgevoerd", vbInformation
    End If
    Msg = "UPDATE Journalen SET dece068=0 WHERE dece068=NULL"
    adntDB.Execute Msg, AantalRC
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        End
    Else
        MsgBox Msg & " met succes uitgevoerd", vbInformation
    End If
    txtTest = InputBox("ALLEREerste dag invoering EURO in uw boekhouding (gebruik formaat JJJJMMDD a.u.b. bvb. 20010701 voor één juli 2002) ?", , "20020101")
    If Len(txtTest) <> 8 Then
        MsgBox "Begin opnieuw", vbCritical
        Msg = "ALTER TABLE Journalen DROP COLUMN dece068"
        adntDB.Execute Msg, AantalRC
        If Err Then
            MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
            End
        Else
            MsgBox Msg & " met succes uitgevoerd", vbInformation
        End If
    End If
    Msg = "UPDATE Journalen SET dece068=(Val(v068)/40.3399) WHERE v066 <= '" & txtTest & "'"
    adntDB.Execute Msg, AantalRC
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        MsgBox "Begin opnieuw", vbCritical
        Msg = "ALTER TABLE Journalen DROP COLUMN dece068"
        adntDB.Execute Msg, AantalRC
        If Err Then
            MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
            End
        Else
            MsgBox Msg & " met succes uitgevoerd", vbInformation
        End If
        End
    Else
        MsgBox Msg & " met succes uitgevoerd voor " & AantalRC & " lijnen.", vbInformation
    End If
    Msg = "UPDATE Journalen SET dece068=(Val(v068)) WHERE v066 >= '" & txtTest & "'"
    adntDB.Execute Msg, AantalRC
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        End
    Else
        MsgBox Msg & " met succes uitgevoerd voor " & AantalRC & " lijnen.", vbInformation
    End If
    Screen.MousePointer = vbHourglass
    Dim rcDummy As ADODB.Recordset
    
    On Error Resume Next
    Set rcDummy = New ADODB.Recordset
    Err = 0
    Msg = "SELECT * FROM Rekeningen"
    rcDummy.CursorLocation = adUseClient
    rcDummy.Open Msg, adntDB, adOpenForwardOnly, adLockOptimistic
    If Err Then MsgBox "stop"
    rcDummy.MoveFirst
    
    Dim bVeldje As String
    Dim eVeldje As String
        
    On Error GoTo 0
    Do While Not rcDummy.EOF
        For COUNT_TO = 22 To 31
            bVeldje = "e0" + Dec((COUNT_TO), "00")
            eVeldje = "dece0" + Dec((COUNT_TO), "00")
            Err = 0
            On Error Resume Next
            rcDummy(eVeldje) = Val(rcDummy(bVeldje))
            If Err Then rcDummy(eVeldje) = 0
            SnelHelpPrint rcDummy("v019") & rcDummy(eVeldje), BL_LOGGING
        Next
        rcDummy.MoveNext
    Loop
    Screen.MousePointer = vbNormal
    MsgBox "Einde update DECIMAAL veld(en) rekeningen"
    rcDummy.Close
    Set rcDummy = Nothing
End If

Dim DitTellertje As Integer

If VeldOK(TABLE_VARIOUS, "rvID") = 0 Then
Else
    For DitTellertje = TABLE_VARIOUS To TABLE_JOURNAL
        bClose DitTellertje
        'sysID autonummering toevoegen 22/04/2004
        'eerst contoleren of nog niet gewijzigd !!
        Msg = "ALTER TABLE " & _
            bstNaam(DitTellertje) & _
            " ADD COLUMN rvID int IDENTITY (1,1)"
        Err = 0
        adntDB.Execute Msg, AantalRC
        If Err Then
            'MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        Else
            MsgBox Msg & " met succes uitgevoerd", vbInformation
        End If
    Next
End If

If VeldOK(TABLE_VARIOUS, "dnnSync") = 0 Then
Else
    For DitTellertje = TABLE_VARIOUS To TABLE_JOURNAL
        'Debug.Print bstNaam(DitTellertje)
        bClose DitTellertje
        'sysID autonummering toevoegen 22/04/2004
        'eerst contoleren of nog niet gewijzigd !!
        Msg = "ALTER TABLE " & _
            bstNaam(DitTellertje) & _
            " ADD COLUMN [dnnSync] Bit"
        Err = 0
        adntDB.Execute Msg, AantalRC
        If Err Then
            MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        Else
            MsgBox Msg & " met succes uitgevoerd", vbInformation
        End If
        Msg = "ALTER TABLE " & _
            bstNaam(DitTellertje) & _
            " ADD COLUMN [dnnSyncAlready] Bit"
        Err = 0
        adntDB.Execute Msg, AantalRC
        If Err Then
            MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        Else
            MsgBox Msg & " met succes uitgevoerd", vbInformation
        End If
    Next
    'vervolgens nog voor TB2 en Briefwisseling maar TB2 enkel voor makelaars !!!!
    'Msg = "ALTER TABLE TB2 ADD COLUMN [dnnSync] Bit"
    'Err = 0
    'adntDB.Execute Msg, AantalRC
    'If Err Then
    '    MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
    'Else
    '    MsgBox Msg & " met succes uitgevoerd", vbInformation
    'End If
    Msg = "ALTER TABLE Briefwisseling ADD COLUMN [dnnSync] Bit"
    Err = 0
    adntDB.Execute Msg, AantalRC
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
    Else
        MsgBox Msg & " met succes uitgevoerd", vbInformation
    End If
    Msg = "ALTER TABLE Briefwisseling ADD COLUMN [dnnSyncAlready] Bit"
    Err = 0
    adntDB.Execute Msg, AantalRC
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
    Else
        MsgBox Msg & " met succes uitgevoerd", vbInformation
    End If
End If

'If VeldOK(TABLE_SUPPLIERS, "sqlID") = 0 Then
'Else
'    Stop
'    bClose TABLE_SUPPLIERS
'    Msg = "ALTER TABLE " & bstNaam(TABLE_SUPPLIERS) & " ADD COLUMN sqlID LONG"
'    Err = 0
'    adntDB.Execute Msg, AantalRC
'    If Err Then
'        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
'    Else
'        MsgBox Msg & " met succes uitgevoerd", vbInformation
'    End If
'    Msg = "UPDATE " & bstNaam(TABLE_SUPPLIERS) & " SET sqlID = 0"
'    Err = 0
'    adntDB.Execute Msg, AantalRC
'    If Err Then
'        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
'    Else
'        MsgBox Msg & " met succes uitgevoerd", vbInformation
'    End If
'End If


If VeldOK(TABLE_VARIOUS, "sqlID") = 0 Then
Else
    For DitTellertje = TABLE_VARIOUS To TABLE_INVOICES
        bClose DitTellertje
        Msg = "ALTER TABLE " & bstNaam(DitTellertje) & " ADD COLUMN sqlID INT"
        Err = 0
        adntDB.Execute Msg, AantalRC
        If Err Then
            MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        Else
            Msg = "UPDATE " & bstNaam(DitTellertje) & " SET sqlID = 0"
            Err = 0
            adntDB.Execute Msg, AantalRC
            If Err Then
                MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
            Else
                MsgBox Msg & " met succes uitgevoerd", vbInformation
            End If
        End If
    Next
End If

If VeldOK(TABLE_JOURNAL, "sqlID") = 0 Then
    bClose TABLE_JOURNAL
    Msg = "ALTER TABLE " & bstNaam(TABLE_JOURNAL) & " DROP COLUMN sqlID"
    Err = 0
    adntDB.Execute Msg, AantalRC
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
    Else
        MsgBox Msg & " met succes uitgevoerd", vbInformation
    End If
End If
    
If VeldOK(TABLE_INVOICES, "vsoftKLID") = 0 Then
Else
    'TABLE_CUSTOMERS = 1 : TABLE_INVOICES = 6 'K' vsoftKLID
    'TABLE_SUPPLIERS = 2 : 'TABLE_INVOICES = 6 'L' vsoftKLID
    bClose TABLE_INVOICES
    Msg = "ALTER TABLE " & bstNaam(TABLE_INVOICES) & " ADD COLUMN vsoftKLID INT"
    Err = 0
    adntDB.Execute Msg, AantalRC
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
    Else
        Msg = "UPDATE " & bstNaam(TABLE_INVOICES) & " SET vsoftKLID = 0"
        Err = 0
        adntDB.Execute Msg, AantalRC
        If Err Then
            MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        Else
            MsgBox Msg & " met succes uitgevoerd", vbInformation
        End If
    End If
End If

If VeldOK(TABLE_JOURNAL, "vsoftRID") = 0 Then
Else
    'TABLE_JOURNAL = 7 : vsoftRID
    bClose TABLE_JOURNAL
    Msg = "ALTER TABLE " & bstNaam(TABLE_JOURNAL) & " ADD COLUMN vsoftRID INT"
    Err = 0
    adntDB.Execute Msg, AantalRC
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
    Else
        Msg = "UPDATE " & bstNaam(TABLE_JOURNAL) & " SET vsoftRID = 0"
        Err = 0
        adntDB.Execute Msg, AantalRC
        If Err Then
            MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        Else
            MsgBox Msg & " met succes uitgevoerd", vbInformation
        End If
    End If
End If

If VeldOK(TABLE_CONTRACTS, "vsoftKID") = 0 Then
Else
    'TABLE_CONTRACTS = 5 vsoftKID
    bClose TABLE_CONTRACTS
    Msg = "ALTER TABLE " & bstNaam(TABLE_CONTRACTS) & " ADD COLUMN vsoftKID INT"
    Err = 0
    adntDB.Execute Msg, AantalRC
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
    Else
        Msg = "UPDATE " & bstNaam(TABLE_CONTRACTS) & " SET vsoftKID = 0"
        Err = 0
        adntDB.Execute Msg, AantalRC
        If Err Then
            MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        Else
            MsgBox Msg & " met succes uitgevoerd", vbInformation
        End If
    End If
End If

JumpToJournaal:
Set rsJournaal = New ADODB.Recordset
If ntDB.Connect <> "" Then
    rsJournaal.CursorLocation = adUseClient
Else
    rsJournaal.CursorLocation = adUseServer
End If
rsJournaal.Open "SELECT TOP 1 * FROM Journalen", adntDB, adOpenForwardOnly, adLockOptimistic, adCmdText

Cijfermaskers
Dim VarString As String
Dim VanString As String
Dim NaarString As String
If Dir(LOCATION_COMPANYDATA + "DDEF*.txt") = "" Then
    VarString = Dir(LOCATION_COMPANYDATA + "Dok*.txt")
    Do While VarString <> ""
        If Mid(VarString, 4, 1) = "2" Then
        Else
            VanString = LOCATION_COMPANYDATA + VarString
            For COUNT_TO = 0 To 5
                NaarString = LOCATION_COMPANYDATA + "DDEF" + Mid(VarString, 4, 2) + Format(COUNT_TO) + ".TXT"
                FileCopy VanString, NaarString
            Next
        End If
        VarString = Dir
    Loop
End If
If Dir(LOCATION_COMPANYDATA + "DDEF125.TXT") = "" Then
    For COUNT_TO = 1 To 4
        VanString = LOCATION_COMPANYDATA + "DDEF1" + Format(COUNT_TO) + "0.TXT"
        NaarString = LOCATION_COMPANYDATA + "DDEF1" + Format(COUNT_TO) + "5.TXT"
        FileCopy VanString, NaarString
    Next
End If

Screen.MousePointer = vbNormal
Exit Sub

ErrorOpvang:
MsgBox vBC(ktrlfl, ptel) + " " + Error
Resume Next

End Sub

Sub BeWaarTekst(Onderdeel As String, SubDeel As String, Element As String)

SaveSetting App.ProductName, Onderdeel, SubDeel, Element

End Sub

Sub Cijfermaskers()

If bhEuro Then
    MASK_2002 = MASK_EUR
Else
    MASK_2002 = MASK_BEF
End If

MASK_SY(0) = "#########"
MASK_SY(1) = "###0"
MASK_SY(2) = "######0.00"
MASK_SY(3) = "##0.00000000"
MASK_SY(4) = "#######0.00"
MASK_SY(5) = "##0"
MASK_SY(6) = "#0"
MASK_SY(7) = "#####0.0"
MASK_SY(8) = "#######0"

End Sub

Sub AutoUnloadBedrijf()
Dim T           As Integer
Dim LastUsed As String

Screen.MousePointer = vbHourglass
Ktrl = 100
Unload BJPERDAT
CloseOpenWindows

'2. Tabellen sluiten
bClose 99

'adoverbindingen vrijmaken
Screen.MousePointer = vbHourglass
On Error Resume Next
rsJournaal.Close
Set rsJournaal = Nothing
For COUNT_TO = TABLE_VARIOUS To TABLE_CONTRACTS
    rsMAR(COUNT_TO).Close
Next

'3. Menuopties beperken
For COUNT_TO = 1 To 6
    Mim.MenuTitel(COUNT_TO).Enabled = False
Next

For T = 1 To 3
    BasisB(T).WindowState = 1
    BasisB(T).Enabled = False
Next
Mim.AV(11).Enabled = False
Mim.Basis(11).Enabled = False

LOCATION_COMPANYDATA = LOCATION_
Mim.Caption = App.Title & "-{Local First} v." & App.Major & "." & App.Minor & "." & App.Revision
Screen.MousePointer = vbNormal

On Error Resume Next
adntDB.Close
Set adntDB = Nothing
Mim.Basis(3).Enabled = False
Close

LOCATION_COMPANYDATA = ""

End Sub

Function InstellingenBewaard(frmVenster As Form) As Boolean

Err = 0: InstellingenBewaard = True
On Error Resume Next
SaveSetting App.ProductName, frmVenster.Name, "Top", frmVenster.Top
SaveSetting App.ProductName, frmVenster.Name, "Links", frmVenster.Left
SaveSetting App.ProductName, frmVenster.Name, "Breedte", frmVenster.Width
SaveSetting App.ProductName, frmVenster.Name, "Hoogte", frmVenster.Height
If Err Then InstellingenBewaard = False

End Function


Sub LaadInstellingen(frmVenster As Form)

On Local Error GoTo EenFoutBijINLaden

frmVenster.Top = GetSetting(App.ProductName, frmVenster.Name, "Top")
frmVenster.Left = GetSetting(App.ProductName, frmVenster.Name, "Links")
frmVenster.Width = GetSetting(App.ProductName, frmVenster.Name, "Breedte")
frmVenster.Height = GetSetting(App.ProductName, frmVenster.Name, "Hoogte")
Exit Sub

EenFoutBijINLaden:
Exit Sub

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

Function LaadTekst(Onderdeel As String, SubDeel As String) As Variant

On Local Error Resume Next
If InStr(Onderdeel, ";") Then
    LaadTekst = GetSetting(Left(Onderdeel, InStr(Onderdeel, ";") - 1), Mid(Onderdeel, InStr(Onderdeel, ";") + 1), SubDeel)
Else
    LaadTekst = GetSetting(App.ProductName, Onderdeel, SubDeel)
End If

EenFoutBijINLaden:
Exit Function

End Function

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
bstNaam(TABLE_COUNTERS) = "jr" & Format(Val(BJPERDAT.Boekjaar.Text) + 1, "0000")
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

bstNaam(TABLE_COUNTERS) = "jr" & Format(Val(BJPERDAT.Boekjaar.Text) + 1, "0000")
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
If BJPERDAT!PeriodeBoekjaar.ListCount - 1 > 12 Then
    MsgBox "Boekjaar bestaande uit " + Format(BJPERDAT!PeriodeBoekjaar.ListCount) + " periodes wordt hierna tot 12 periodes gebracht.  Kontroleer na installatie a.u.b. !"
    Fl = FreeFile
    recNr = 0
    Open LOCATION_COMPANYDATA + "DEF00.OCT" For Random As Fl Len = 16
    For COUNT_TO = BJPERDAT!PeriodeBoekjaar.ListCount - 12 To BJPERDAT!PeriodeBoekjaar.ListCount - 1
         recNr = recNr + 1
         BJPERDAT!PeriodeBoekjaar.ListIndex = COUNT_TO
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
If BJPERDAT.Boekjaar.ListIndex = 0 Then
    BJPERDAT.Boekjaar.Enabled = False
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

Sub SetString99(NummerSleutel As Integer)
Dim DummySleutel As String

DummySleutel = "s" + Format(NummerSleutel, "000") + " "
bGet TABLE_COUNTERS, 0, DummySleutel
If Ktrl Then
    If BAModus = 1 Then
        SnelHelpPrint "Teller " + DummySleutel + "niet te vinden.  Wordt hierna ingevoegd.", BL_LOGGING
        TLB_RECORD(TABLE_COUNTERS) = ""
        vBib TABLE_COUNTERS, DummySleutel, "v071"
        vBib TABLE_COUNTERS, FL99_RECORD, "v217"
        bInsert TABLE_COUNTERS, 0
    Else
        MsgBox "Onlogika btrieve versie !"
    End If
Else
    RecordToVeld TABLE_COUNTERS
    If BAModus = 1 Then
        vBib TABLE_COUNTERS, FL99_RECORD, "v217 "
        bUpdate TABLE_COUNTERS, 0
        If Ktrl Then
            MsgBox "UpdateStop Teller. kontakteer vsoft"
        End If
    Else
        MsgBox "Onlogika btrieve versie !"
    End If
End If
bClose TABLE_COUNTERS

End Sub



Sub SS99(StringInhoud As String, NummerRec As Integer)

FL99_RECORD = StringInhoud
SetString99 NummerRec

End Sub



Public Function VeldOK(flHier As Integer, VeldNaam As String, Optional VeldDef As String) As Long

    If rsMAR(flHier).State = adStateClosed Then
        Ktrl = bOpen(flHier)
    End If
    On Error Resume Next
    Err = 0
    Msg = rsMAR(flHier)(VeldNaam).Name
    If Err = 0 Then
        VeldOK = Err
        Exit Function
    ElseIf VeldDef = "" Then
        VeldOK = Err
        Exit Function
    Else
        bClose flHier
        Msg = "ALTER TABLE " & bstNaam(flHier) & " ADD COLUMN " & VeldNaam & " " & VeldDef & ";"
        If MsgBox(Msg & vbCr & vbCr & "SQL-instructie uitvoeren", vbYesNo + vbDefaultButton1 + vbQuestion) = vbYes Then
            Err = 0
            adntDB.Execute Msg, AantalRC
            If Err Then
                MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
            Else
                MsgBox Msg & " met succes uitgevoerd", vbInformation
                
            End If
            VeldOK = Err
        Else
            VeldOK = 99
        End If
    End If

End Function

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
    
    TQString = OpenSchemaString("TABLE")
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

Function netVoorbereiden() As Boolean
netVoorbereiden = False

'klaarmaken voor .NET versie
Dim netFlTemp As Integer
Dim netFlFree3 As Integer
Dim netdummyLinE As String
Dim netCOUNT_TO As Integer
Dim netT As Integer
Dim netA As String * 16

For netCOUNT_TO = 9 To 0 Step -1
    If Dir$(LOCATION_COMPANYDATA + "DEF" + Format(netCOUNT_TO, "00") + ".OXT") = "" Then
        If Dir$(LOCATION_COMPANYDATA + "DEF" + Format(netCOUNT_TO, "00") + ".OCT") <> "" Then
            netFlTemp = FreeFile
            Open LOCATION_COMPANYDATA + "DEF" + Format(netCOUNT_TO, "00") + ".OCT" For Random As netFlTemp Len = 16
                netdummyLinE = ""
                For netT = 1 To 99
                    Get netFlTemp, netT, netA
                    If netA = Space$(16) Then
                        Exit For
                    Else
                        netdummyLinE = netdummyLinE & "," & netA
                    End If
                Next
                netdummyLinE = Mid(netdummyLinE, 2)
                'wegschrijven
                netFlFree3 = FreeFile
                Open LOCATION_COMPANYDATA + "DEF" + Format(netCOUNT_TO, "00") + ".OXT" For Output As netFlFree3
                    Print #netFlFree3, netdummyLinE
                Close netFlFree3
                netVoorbereiden = True
            Close netFlTemp
        End If
    End If
Next

'If Dir$(PROGRAM_LOCATION + "9999.OXT") = "" Then
'    netFlTemp = FreeFile
'    Open PROGRAM_LOCATION + "9999.OCT" For Random As netFlTemp Len = 4
'    Get netFlTemp, 1, aa
'    dummyLinE = Trim(aa)
'    Get netFlTemp, 2, aa
'    dummyLinE = dummyLinE & "," & Trim(aa)
'    Get netFlTemp, 3, aa
'    dummyLinE = dummyLinE & "," & Trim(aa) '''
'
'   netFlFree3 = FreeFile
'    Open PROGRAM_LOCATION + "9999.OXT" For Output As netFlFree3
'    Print #netFlFree3, dummyLinE
'    Close netFlFree3
'    Close netFlTemp
'End If


End Function

