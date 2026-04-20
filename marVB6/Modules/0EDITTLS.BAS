Attribute VB_Name = "Module3"

DefInt A-Z

Option Explicit

Dim DataPos                As Integer
Dim DataLen                As Integer
Dim OptieTxt               As String
Dim TempoTxt               As String
Dim CrText                 As String
Dim CrText2                As String

Dim T                       As Integer


Sub DbKontrole(StringTeZoeken As String, FlNr As Integer)
Dim aa As Variant

Dim Teller As Integer

Select Case FlNr
    Case TABLE_LEDGERACCOUNTS
        Dim rsKBRecord As ADODB.Recordset
        Set rsKBRecord = New ADODB.Recordset
        rsKBRecord.CursorLocation = adUseServer
        rsKBRecord.Open "MinimumIndeling", adKBDB, adOpenKeyset, adLockOptimistic, adCmdTableDirect
        rsKBRecord.Index = "RekeningNummer"
        aa = ""
       
        rsKBRecord.Seek StringTeZoeken, adSeekBeforeEQ
        If rsKBRecord.BOF Then
            If rsKBRecord.EOF Then
                rsKBRecord.MoveFirst
            End If
            aa = rsKBRecord.fields(0) & " " + rsKBRecord.fields(1) & vbCrLf & aa
            Teller = 0
            rsKBRecord.MovePrevious
            Do While Not rsKBRecord.BOF
                Teller = Teller + 1
                rsKBRecord.MovePrevious
                aa = rsKBRecord.fields(0) & " " & rsKBRecord.fields(1) & vbCrLf & aa
                If Teller >= 10 Then Exit Do
            Loop
        End If
        
        Teller = 0
        rsKBRecord.Seek StringTeZoeken, adSeekAfterEQ
        If rsKBRecord.EOF Then
        Else
            rsKBRecord.MoveNext
            Do While Not rsKBRecord.EOF
                Teller = Teller + 1
                aa = aa & rsKBRecord.fields(0) & " " & rsKBRecord.fields(1) & vbCrLf
                If Teller >= 15 Then Exit Do
                rsKBRecord.MoveNext
            Loop
        End If
        Mim.InfoData.Visible = True
        Mim.InfoData.Cls
        Mim.InfoData.Print aa
        rsKBRecord.Close
        Set rsKBRecord = Nothing
                
    Case Else
        MsgBox "stop"
End Select

End Sub

Function TeleBibClick(Fl As Integer) As Integer
Dim PositieX       As Integer
Dim aa             As String
Dim LogTekst       As String
Dim BoxType         As Integer
Dim BoxMask         As String

TeleBibClick = False
Select Case Fl
    Case TABLE_CUSTOMERS To TABLE_CONTRACTS 'Hoofdfiches
        If TeleBibPagina(Fl) = False Then
            Beep
            Exit Function
        End If
    Case TABLE_INVOICES 'Aankoop verkoopdokumenten
        If TLBPag2("020" + Left(FVT(TABLE_INVOICES, 0), 1)) = False Then
            Beep
            Exit Function
        End If
    Case 10, 12, 18, 21, 28 'Diverse gebruikersfiches
        If TLBPag2(Format(Fl, "000")) = False Then
            Beep
            Exit Function
        End If
    Case 1000 To 1999   'As1 verzoekdokument
        If TLBPag3("AS1" + Format(Fl - 1000, "000")) = False Then
            Beep
            Exit Function
        End If
     Case 2000 To 2099  'Diverse groepen eigen aan ROELANDT systeem
        If TLBPag2("GROEP" + Format(Fl - 2000, "00")) = False Then
            Beep
            Exit Function
        End If
    Case 3000 To 3099   'Schadedossiers e.a.
        If TLBPag2("SCHADE" + Format(Fl - 3000, "00")) = False Then
            Beep
            Exit Function
        End If
    Case 4000 To 4999   'takkodes
        If TLBPag3("TAK" + Format(Fl - 4000)) = False Then
            Beep
            Exit Function
        End If
    Case Else
        MsgBox "stop in telebibclick, fl=" + Format(Fl)
End Select
Select Case Fl
    Case TABLE_CUSTOMERS
        LogTekst = "BIB voor Klanten"
    Case TABLE_SUPPLIERS
        LogTekst = "BIB voor Leveranciers"
    Case TABLE_LEDGERACCOUNTS
        LogTekst = "BIB voor Algemene Rekeningen"
    Case TABLE_PRODUCTS
        LogTekst = "BIB voor Artikels/Diensten"
    Case TABLE_CONTRACTS
        LogTekst = "BIB voor contracten"
    Case TABLE_INVOICES
        LogTekst = "BIB voor Aan- en Verkoopdokumenten"
    Case 1000 To 1999
        Fl = TABLE_VARIOUS
        LogTekst = "BIB AS1/verzoeken"
    Case 2000 To 2099
        Fl = TABLE_VARIOUS
        LogTekst = "BIB Polis " + vBibTekst(TABLE_CONTRACTS, "#A000 #")
    Case 3000 To 3099
        Fl = TABLE_VARIOUS
        LogTekst = "Bib Schade " + vBibTekst(TABLE_VARIOUS, "#C000 #")
    Case 4000 To 4099
        Fl = TABLE_VARIOUS
        LogTekst = "BIB DetailPolis " + vBibTekst(TABLE_CONTRACTS, "#A000 #")
    Case Else
        Fl = TABLE_VARIOUS
        LogTekst = " BIB Allerlei"
End Select

Unload Xlog
Load Xlog
Xlog.Caption = Xlog.Caption + LogTekst
Xlog.Height = Xlog.Height + 3
Xlog.X.Rows = 1
Xlog.X.Cols = 3

Xlog.X.Col = 0
Xlog.X.Row = 0
Xlog.X.text = "vsfKode"
Xlog.X.Col = 1
Xlog.X.text = "Veldomschrijving"
Xlog.X.Col = 2
Xlog.X.text = "Veldgegevens"

If Fl = TABLE_INVOICES Then
    If VSF_PRO Then
        Xlog.WijzigenLijn.Enabled = True
        Xlog.Afsluiten.Caption = "Speciaal"
        Xlog.Afsluiten.Visible = True
    Else
        Xlog.WijzigenLijn.Enabled = False
        Xlog.Afsluiten.Caption = "Vernietig!"
        Xlog.Afsluiten.Visible = False
    End If
End If

aa = ""
T = 0
Do While TELEBIB_CODE(T) <> Space$(10)
    CrText = vBibTekst(Fl, "#" + Mid(TELEBIB_CODE(T), 5, 5) + "#")
    Select Case Mid(TELEBIB_CODE(T), 2, 2)
        Case "  ", "K ", "L ", "LC", "R ", "R3", "R4", "R6", "R7"
            'niks
        Case Else
            Select Case Mid(TELEBIB_CODE(T), 1, 1)
                Case " "
                    BoxMask = "00"
                    BoxType = 0
                Case "0" To "9"
                    BoxMask = "000"
                    BoxType = 1
            End Select
            If Left(TELEBIB_CODE(T), 1) = "@" Or CrText = "" Then
            Else
                CrText = fmarBoxText(Format(Val(Mid(TELEBIB_CODE(T), 1, 3)), BoxMask), "2", CrText)    'hier eventueel taaloptie
            End If
    End Select
    If Mid(TELEBIB_CODE(T), 10, 1) = "x" Then
        CrText = Str(Fl) + "{...}" 'rsMAR(Fl)(Mid(TELEBIB_CODE(T), 5, 4))
    End If
    aa = TELEBIB_CODE(T) + vbTab + TELEBIB_TEXT(T) + vbTab + CrText
    Xlog.X.AddItem aa
    T = T + 1
Loop

Xlog.X.ColWidth(0) = 45
Xlog.X.ColWidth(1) = 2805
Xlog.X.ColWidth(2) = 6165
Xlog.X.ColAlignment(0) = flexAlignLeftTop
Xlog.X.ColAlignment(1) = flexAlignLeftTop
Xlog.X.ColAlignment(2) = flexAlignLeftTop
Xlog.X.Row = 1
Xlog.X.Col = 2

XLogShow:
Xlog.WijzigenLijn.Default = True
Xlog.Afsluiten.TabStop = True
XLogKey = ""
Xlog.SSTab1.TabVisible(1) = False
Xlog.Show 1
PeppolFlag = False
If XLogKey <> "" Then
    T = 0
    Xlog.X.Col = 2
    Do While TELEBIB_CODE(T) <> Space$(10)
        Xlog.X.Row = T + 1
        CrText2 = Xlog.X.text
        If Mid(TELEBIB_CODE(Xlog.X.Row - 1), 10, 1) = "*" And CrText2 = "" Then
            MsgBox "Invoer voor '" + RTrim$(TELEBIB_TEXT(T)) + "'" + vbCrLf + vbCrLf + "is verplicht !", 0, "Vervolledig a.u.b."
            GoTo XLogShow
        ElseIf CrText2 <> "" Then
            If Mid(TELEBIB_CODE(T), 10, 1) = "x" Then
            Else
                Select Case Mid(TELEBIB_CODE(T), 2, 2)
                    Case "  "
                    Case Else
                        If Left(TELEBIB_CODE(T), 1) = "@" Then
                        Else
                            On Local Error Resume Next
                            CrText2 = Left(CrText2, InStr(CrText2, ":") - 1)
                            On Local Error GoTo 0
                        End If
                End Select
                vBib Fl, CrText2, Mid(TELEBIB_CODE(T), 5, 5)
            End If
        End If
        T = T + 1
    Loop
    If Xlog.Afsluiten.Caption = "Speciaal" Then
        Msg = "Gegevens bestaande fiche wijzigen.  Bent U zeker ?"
        Ktrl = MsgBox(Msg, 292)
        If Ktrl = 6 Then
            bUpdate Fl, 0
        End If
    End If
    TeleBibClick = True
End If

End Function

Function TLBPag3(BsDef As String) As Integer

TLBPag3 = False

Dim FlTemp As Integer
'code= 1 - 1 : Poliskontrole 1 = ON
'      2 - 3 : Selektiekeuze waarde (ListIDX)
'      4 - 4 :
'      5 - 8 : TeleBib
'      9 - 9 : * voor verplichte invulling

If Dir$(PROGRAM_LOCATION + "Def\" + BsDef + ".Def") = "" Then
    MsgBox "Geen VsoftBib definitie " + BsDef + ".Def"
    Exit Function
End If

FlTemp = FreeFile
Open PROGRAM_LOCATION + "Def\" + BsDef + ".DEF" For Input As FlTemp
T = 0
Do While Not EOF(FlTemp)
    Input #FlTemp, TELEBIB_CODE(T), TELEBIB_TEXT(T), TELEBIB_TYPE(T), TELEBIB_LENGTH(T), TELEBIB_POS(T)
    T = T + 1
Loop
Close FlTemp
TELEBIB_CODE(T) = ""
TLBPag3 = True

End Function




Function TLBPag2(BsDef As String) As Integer
Dim FlInput As Integer
Dim FFDefinitie As String

'On Local Error GoTo 0

'code= 1 - 1 : Poliskontrole 1 = ON
'      2 - 3 : Selektiekeuze waarde (ListIDX)
'      4 - 4 :
'      5 - 8 : TeleBib
'      9 - 9 : Volgnummer eventueel 1 tot 9
'      10    : * voor verplichte invulling
'Inlaadvolgorde 1 : 000     'hoofddefinitie (VSOFT)
'               2 : xxxG    'uitbreiding gebruiker
'               2 : xxxM    'uitbreiding makelaar

If Dir$(PROGRAM_LOCATION + "Def\" + BsDef + ".Def") = "" Then
    MsgBox "Geen VsoftBib definitie " + BsDef + ".Def"
    Exit Function
End If

If Dir$(PROGRAM_LOCATION + "Def\" + BsDef + "U.Def") = "" Then
    GoTo GeenUserDef
Else
    FlInput = FreeFile
    T = 0
    Open PROGRAM_LOCATION + "Def\" + BsDef + "U.Def" For Input As FlInput
    Do While Not EOF(FlInput)
        Input #FlInput, TELEBIB_CODE(T), TELEBIB_TEXT(T), TELEBIB_TYPE(T), TELEBIB_LENGTH(T)
        T = T + 1
    Loop
    Close FlInput
    TELEBIB_CODE(T) = ""
    TELEBIB_LAST = T - 1
    TLBPag2 = True
    Exit Function
End If

GeenUserDef:
FlInput = FreeFile
Open PROGRAM_LOCATION + "Def\" + BsDef + ".Def" For Input As FlInput
T = 0
Do While Not EOF(FlInput)
    Input #FlInput, TELEBIB_CODE(T), TELEBIB_TEXT(T), TELEBIB_TYPE(T), TELEBIB_LENGTH(T)
    T = T + 1
Loop
Close FlInput
TELEBIB_CODE(T) = ""

If ProducentNummer = Space$(8) Then
ElseIf Dir$(PROGRAM_LOCATION + "Def\" + BsDef + "M.Def") = "" Then
Else
    FlInput = FreeFile
    Open PROGRAM_LOCATION + "Def\" + BsDef + "M.Def" For Input As FlInput
    Do While Not EOF(FlInput)
        Input #FlInput, TELEBIB_CODE(T), TELEBIB_TEXT(T), TELEBIB_TYPE(T), TELEBIB_LENGTH(T)
        T = T + 1
    Loop
    Close FlInput
    TELEBIB_CODE(T) = ""
End If
TLBPag2 = True
Exit Function

TLBError:
MsgBox "Telebibinlaadfout" + Str$(T) + " error:" + Error$
Close FlInput
TLBPag2 = False
Exit Function
Resume

End Function

Function vsfInputBox(InfoTekst As String, Titel As String, TekstZelf As String, Paswoord As String) As String
Dim T       As Integer
Dim TT      As Integer

'On Local Error GoTo ErrInput

ReDim ToolDef(3) As String
Load ntInputbox
ntInputbox.Caption = Titel
If Mid(InfoTekst, 1, 1) = "@" Then
    ntInputbox.Hernieuw.Visible = True
    ntInputbox.cmdVooruitAchteruit(0).Visible = True
    ntInputbox.cmdVooruitAchteruit(1).Visible = True
    ntInputbox.lblInfo.Visible = True
    Select Case Mid(InfoTekst, 2, 2)
        Case "00"
            SQLBevel = "SELECT * FROM ISOLandKodes WHERE ISOLandNummer LIKE '"
            ntInputbox.DefaultData.RecordSource = SQLBevel + RTrim$(TekstZelf) + "%';"
            ReDim ToolDef(3) As String
            ToolDef(0) = "00=v149 "     'Landnummer  ISO kode
            ToolDef(1) = "01=vs03 "     'Munteenheid ISO kode
            ToolDef(2) = "02=v150 "     'Landkode    ISO kode
        Case "01"
            SQLBevel = "SELECT * FROM PostKodesWoonplaatsen WHERE PostKode LIKE '"
            ntInputbox.DefaultData.RecordSource = SQLBevel + RTrim$(TekstZelf) + "%';"
            ReDim ToolDef(3) As String
            ToolDef(0) = "01=A107 "     'PostKode volgens Postkantoor
            ToolDef(1) = "02=A108 "     'Plaatsnaam
            
        Case "02"
            SQLBevel = "SELECT * FROM PostKodesWoonplaatsen WHERE PlaatsNaam LIKE '"
            ntInputbox.DefaultData.RecordSource = SQLBevel + RTrim$(TekstZelf) + "%';"
            ReDim ToolDef(3) As String
            ToolDef(0) = "02=A108 "     'Plaatsnaam
            ToolDef(1) = "01=A107 "     'PostKode volgens Postkantoor
    End Select
Else
    ntInputbox.Hernieuw.Visible = False
    ntInputbox.cmdVooruitAchteruit(0).Visible = False
    ntInputbox.cmdVooruitAchteruit(1).Visible = False
    ntInputbox.lblInfo.Visible = False
End If
ntInputbox.Tag = GridText
If GridText = "Edit No" Then
    ntInputbox!Ok.Visible = False
    ntInputbox!Sluiten.Default = True
Else
    ntInputbox!Ok.Visible = True
    ntInputbox!Ok.Default = True
End If
GridText = InfoTekst
If Mid(InfoTekst, 1, 1) = "@" Then
    ntInputbox.MedeDeling.SimpleText = GridText + ntInputbox.DefaultData.RecordSource
Else
    ntInputbox.MedeDeling.SimpleText = GridText
End If
ntInputbox.TekstInfo.text = TekstZelf
'nTInputBox.TekstInfo.PasswordChar = Paswoord
ntInputbox.Show 1
    
    If ntInputbox.TekstInfo.text = Chr$(255) Then
    vsfInputBox = TekstZelf
Else
    If ntInputbox.DefaultData.RecordSource = "" Then
        vsfInputBox = ntInputbox.TekstInfo.text
    ElseIf Mid(InfoTekst, 1, 1) = "@" And ntInputbox.TekstInfo.text = ntInputbox.DefaultData.Recordset(Val(Mid(ToolDef(0), 1, 2))) Then
        Dim AantalRijen As Integer
        Select Case Mid(InfoTekst, 2, 2)
            Case "00"
                AantalRijen = 2
            Case "01", "02"
                AantalRijen = 1
            Case Else
                MsgBox "Stop"
        End Select
        For T = AantalRijen To 0 Step -1
            Xlog.X.Col = 0
            For TT = 1 To Xlog.X.Rows
                Xlog.X.Row = TT
                If Mid(Xlog.X.text, 5, 5) = Mid(ToolDef(T), 4) Then
                    Xlog.X.Col = 2
                    Xlog.X.text = ntInputbox.DefaultData.Recordset(Val(Mid(ToolDef(T), 1, 2)))
                    Exit For
                End If
            Next
        Next
        vsfInputBox = ntInputbox.TekstInfo.text
    Else
        vsfInputBox = ntInputbox.TekstInfo.text
    End If
End If
ntInputbox.DefaultData.RecordSource = ""
Exit Function

ErrInput:
Beep
vsfInputBox = TekstZelf
Exit Function

End Function

Function ZoekEnPlaats(DeKontrol As Control, ZoekTekst As String, ALijnen As Integer, OptieNr As Integer, OptieTxt As String) As String
Dim OptieLen As Integer
'Dim CRLFLokatie As Integer
Dim PuntKommaLokatie As Integer
Dim TeDoorZoeken As String
Dim DeString As String
Dim joinStringHier As String
Dim reeksHier() As String

ZoekEnPlaats = ""
adKBTable.Seek ZoekTekst, adSeekFirstEQ
If adKBTable.EOF Then
    MsgBox "Stop !  Keuzebox " + ZoekTekst + " niet te vinden..."
    Exit Function
Else
    'adKBTable.Fields("BestandsNaam")
    joinStringHier = adKBTable.fields("splitDefinitie")
End If

If Right(joinStringHier, 1) = ";" Then
Else
    joinStringHier = joinStringHier + ";"
End If

PuntKommaLokatie = 1
ALijnen = 0
OptieNr = 0
OptieLen = Len(OptieTxt)
DeKontrol.Clear

'Dim stringArray() As String
'Dim counter As Integer
'stringArray = Split(joinStringHier, ";")
'For counter = 0 To UBound(stringArray)
'    DeKontrol.AddItem stringArray(counter)
'Next

Do While InStr(PuntKommaLokatie, joinStringHier, ";")
    ALijnen = ALijnen + 1
    DeString = Mid(joinStringHier, PuntKommaLokatie, InStr(PuntKommaLokatie, joinStringHier, ";") - PuntKommaLokatie)
    'KeuzeVSF.NTBoxLijst.AddItem DeString
    DeKontrol.AddItem DeString
    If Mid(joinStringHier, PuntKommaLokatie, OptieLen) = OptieTxt Then
        OptieNr = ALijnen - 1
        ZoekEnPlaats = DeString
    End If
    PuntKommaLokatie = InStr(PuntKommaLokatie, joinStringHier, ";") + 1
Loop
If Len(OptieTxt) = 0 Then OptieNr = 0

End Function
Function Toegankelijk(frmHier As Form) As Boolean

'If objUsers.CheckRights(Left(frmHier.Name, 20), 1) Then
    Toegankelijk = True
'    objAudit.WriteLog frmHier.Name, "Toegang met Succes"
'Else
'    Toegankelijk = False
'    MsgBox "Deze taak is (nog) niet toegekend aan huidige gebruiker. Neem contact op met systeembeheerder.", vbExclamation, frmHier.Name
'    objAudit.WriteLog "LogIn", frmHier.Name + ":Toegang voorbehouden"
'End If

End Function

Sub GridSchoon(GelijkWelkRooster As MSFlexGrid)
Dim T As Integer

GelijkWelkRooster.Rows = 2
GelijkWelkRooster.Row = 1
For T = 0 To GelijkWelkRooster.Cols - 1
    GelijkWelkRooster.Col = T
    GelijkWelkRooster.text = ""
Next

End Sub

