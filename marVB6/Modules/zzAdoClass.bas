Attribute VB_Name = "thisAdoToolsClass"
DefInt A-Z
Option Explicit

Dim X As Integer
Public Const BATCH_SIZE As Long = 5000

Dim DataPos                As Integer
Dim DataLen                As Integer
Dim OptieTxt               As String
Dim TempoTxt               As String
Dim CrText                 As String
Dim CrText2                As String
Dim T                       As Integer


Function NieuwBoekjaar() As Boolean
Dim BeginEenheden As Double
Dim BeginBedrag As Currency
Dim Vroeger1 As String
Dim Vroeger2 As String
Dim Nu1 As String
Dim Nu2 As String
Dim bString As String * 16
Dim recNr As Integer
Dim eenSaldo As Currency

NieuwBoekjaar = False
'MsgBox "Tijdelijk niet beschikbaar"
'Exit Sub

If ACTIVE_BOOKYEAR Then
    MsgBox "Enkel logisch met hoogste boekjaar actief !  Probeer opnieuw..."
    Exit Function
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
        Exit Function
    End If
    
    If Val(String99(READING, 62)) <> 1 Then
        Msg = "Eindinventaris van vorig boekjaar overslaan en stock-roulatie vernietigen !  Bent U zeker ?"
        KtrlBox = MsgBox(Msg, 292)
        If KtrlBox = 6 Then
        Else
            bClose TABLE_COUNTERS
            TABLEDEF_ONT(TABLE_COUNTERS) = TempoBestand
            Exit Function
        End If
    End If
End If
bClose TABLE_COUNTERS
TABLEDEF_ONT(TABLE_COUNTERS) = TempoBestand

Msg = "Een nieuw boekjaar wordt geïnstalleerd hierna !  Bent U zeker ?"
KtrlBox = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2)
If KtrlBox = vbYes Then
Else
    Exit Function
End If

'Overdracht rekeningsaldo's
bClose TABLE_LEDGERACCOUNTS
bFirst TABLE_LEDGERACCOUNTS, 0
If Ktrl Then
    Beep
    MsgBox "Voortijdige stop..."
    Exit Function
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
NieuwBoekjaar = True

Outchecken:
Ktrl = 100
AutoUnloadBedrijf
Exit Function

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

End Function


Function adxMaakDatabase(dbNaam As String, dbPath As String) As Boolean
    
    adxMaakDatabase = False
    Dim cat As New ADOX.Catalog
    On Error Resume Next
    cat.Create ADOJET_PROVIDER & _
        "Data Source=" + dbPath + "\" + dbNaam + ".mdb"
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
    Else
        adxMaakDatabase = True
    End If
    Set cat.ActiveConnection = Nothing
        
End Function


Function adxMaakTabel(tbNaam As String) As Boolean

    Dim cat     As New ADOX.Catalog
    Dim tbl     As New ADOX.Table
                           
    adxMaakTabel = False
    Err = 0
    On Error Resume Next
    If InStr(adntDB.Properties("DBMS Name"), "SQL Server") Then
        ' De catalogus openen voor SQL-SERVER
        cat.ActiveConnection = SQLConnect
    Else
        ' De catalogus openen voor JET4
        cat.ActiveConnection = jetConnect
    End If
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        Set cat.ActiveConnection = Nothing
        Exit Function
    Else
        With tbl
            .Name = tbNaam
            Set .ParentCatalog = cat
            .Columns.Append "ID", adInteger, adBigInt
                                                          
            ' De kolom volgnummer moet automatisch groter worden
            .Columns("ID").Properties("AutoIncrement") = True
        End With
        cat.Tables.Append tbl
        Set cat = Nothing
        If Err Then
            MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        Else
            MsgBox "Aanmaak tabel " + tbNaam + " met succes.", vbInformation
            adxMaakTabel = True
        End If
    End If
    Set cat.ActiveConnection = Nothing
        
End Function


Public Function SQLPopUp(OpzoekReeks As String, tbNaam As String, vldNaam As String, idKode As String) As Boolean

Dim Zoekstring As String
Dim SQLstring As String
Dim SQLPopString As String
Dim strSQL As String

Dim FormHier As Object

Dim rs  As ADODB.Recordset

SQLPopUp = False
Do While OpzoekReeks <> ""
    If InStr(OpzoekReeks, ";") = 0 Then
        Zoekstring = OpzoekReeks
        OpzoekReeks = ""
    Else
        Zoekstring = Left(OpzoekReeks, InStr(OpzoekReeks, ";") - 1)
        OpzoekReeks = Mid(OpzoekReeks, InStr(OpzoekReeks, ";") + 1)
    End If
    bGet TABLE_VARIOUS, 1, "29" + Zoekstring
    If Ktrl Then
    Else
        RecordToVeld TABLE_VARIOUS
        If InStr(vBibTekst(TABLE_VARIOUS, "#v132 #"), "[Colwidth]") Then
            SQLstring = Left(vBibTekst(TABLE_VARIOUS, "#v132 #"), InStr(vBibTekst(TABLE_VARIOUS, "#v132 #"), "[Colwidth]") - 1)
        Else
            SQLstring = vBibTekst(TABLE_VARIOUS, "#v132 #")
        End If
        SQLPopString = Mid(SQLstring, InStr(SQLstring, "FROM ") + 5)
        If SQLPopString = "" Then
            MsgBox "Ongeldige syntax in : " + SQLstring, vbCritical
            Exit Function
        End If
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
         
        Err = 0
        On Error Resume Next
        rs.Open "SELECT * FROM " + SQLPopString + " WHERE " + Left(tbNaam, 1) + vldNaam + " Like '" + idKode + "%'", adntDB, adOpenStatic, adLockOptimistic
        If Err Then
            'Dus eerst verbindingsveld nog invoegen
            MsgBox "Verbindingsveld bestaat nog niet.  Voeg eerst een veld toe: " + Left(tbNaam, 1) + vldNaam
        ElseIf rs.RecordCount = 0 Then
            'deze partij bestaat nog niet, dus invoegen
            strSQL = "INSERT INTO " & SQLPopString & " (" & Left(tbNaam, 1) & vldNaam & ") VALUES ('" & idKode & "');"
            adntDB.Execute strSQL, , adCmdText
        End If
    
        'Msg = Mid(vBibTekst(TABLE_VARIOUS, "#v132 #"), InStr$(vBibTekst(TABLE_VARIOUS, "#v132 #"), "[Colwidth]") + 10)
        'If Msg = "" Then
        '    grdColWidth(0) = 0
        'Else
        'COUNT_TO = 0
        'Do While Msg <> ""
        '    If InStr$(Msg, vbTab) <> 0 Then
        '        grdColWidth(COUNT_TO) = Val(Left(Msg, InStr$(Msg, vbTab) - 1))
        '        Msg = Mid(Msg, InStr$(Msg, vbTab) + 1)
        '        COUNT_TO = COUNT_TO + 1
        '    Else
        '        Exit Do
        '    End If
        'Loop
        'grdColWidth(COUNT_TO) = 0
        Set FormHier = New mijnSQLEditor
        'If optBewerkbaar.Value = True Then
            FormHier!msfSQL.AllowUpdate = True
            FormHier!msfSQL.AllowAddNew = False
            FormHier!msfSQL.AllowDelete = False
        'Else
        '    FormHier!msfSQL.AllowUpdate = False
        '    FormHier!msfSQL.AllowAddNew = False
        '    FormHier!msfSQL.AllowDelete = False
        'End If
        FormHier.txtSQL = SQLstring + " WHERE " + Left(tbNaam, 1) + vldNaam + " Like '" + idKode + "%'"
        FormHier.Caption = SQLstring
        FormHier.Show

        FormHier.SetFocus
        SendKeys "{ENTER}"
    End If
Loop
'DirekteVerkoop.SetFocus
SQLPopUp = True

End Function

Public Function OpenSchemeAsString(tbType As String) As Variant

   Dim rstSchema As ADODB.Recordset
   Dim strCnn As String
   Dim retString As Variant
   
   retString = ""
   
   Err = 0
   On Error Resume Next
   Set rstSchema = adntDB.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))
   If Err Then MsgBox Error: Exit Function
   Do Until rstSchema.EOF
        retString = retString & rstSchema!TABLE_NAME & vbCr
        rstSchema.MoveNext
   Loop
   rstSchema.Close
   Set rstSchema = Nothing
   OpenSchemeAsString = retString
        
End Function

Function vBT(TLBR As Variant, TBS As String) As String

    Dim tbsHier As String * 7
        
    tbsHier = "#     #"
    Mid(tbsHier, 2) = TBS
    Err = 0
    On Error Resume Next
    If TLBR = "" Then
        vBT = ""
    Else
        vBT = Mid(TLBR, InStr(TLBR, tbsHier) + 7, InStr(InStr(TLBR, tbsHier) + 7, TLBR, "#") - (InStr(TLBR, tbsHier) + 7))
    End If
    If Err Then vBT = ""

End Function


Public Sub GetAllIndexes(tbNaam As String, obObject As Object)

   Dim rstSchema As ADODB.Recordset
       
   obObject.Clear
  
   Set rstSchema = adntDB.OpenSchema(adSchemaIndexes)
  
   Do Until rstSchema.EOF
    If UCase(tbNaam) = UCase(rstSchema!TABLE_NAME) Then
        obObject.AddItem "+" & rstSchema!COLUMN_NAME & "; " & rstSchema!INDEX_NAME
        'for t=0 to rstschema.Fields.Count-1:print t;" ";rstschema.Fields(t).Name;" = ";rstschema.Fields(t).Value:next
    End If
      rstSchema.MoveNext
   Loop
   rstSchema.Close
   
End Sub

' Copy a file's contents into a BLOB field.
Function FileToBlob(fld As ADODB.Field, fileName As String, _
    Optional ChunkSize As Long = 8192)
    Dim fnum As Integer, bytesLeft As Long, bytes As Long
    Dim tmp() As Byte
    
    ' Raise an error if the field doesn't support GetChunk.
    If (fld.Attributes And adFldLong) = 0 Then
        Err.Raise 1001, , "Field doesn't support the GetChunk method."
    End If
    ' Open the file; raise an error if the file doesn't exist.
    If Dir$(fileName) = "" Then Err.Raise 53, , "File not found"
    
    fnum = FreeFile
    Open fileName For Binary As fnum
    ' Read the file in chunks, and append data to the field.
    bytesLeft = LOF(fnum)
    Do While bytesLeft
        bytes = bytesLeft
        If bytes > ChunkSize Then bytes = ChunkSize
        ReDim tmp(1 To bytes) As Byte
        Get #1, , tmp
        fld.AppendChunk tmp
        bytesLeft = bytesLeft - bytes
    Loop
    Close #fnum
    
End Function


Function ktrlBLOBRecord() As Boolean

Dim ktrlString As String

Err = 0
On Error Resume Next

'nieuwe poging 03/2007

ktrlBLOBRecord = True
ktrlString = rsMAR(TABLE_INVOICES)("bstndNaam37").Name

'bstndNaam
'typeZending 'TIF, PDF, enz...
'bstBLOB inhoud bestand
If Err = 3265 Then
    rsMAR(TABLE_INVOICES).Close
    Msg = "ALTER TABLE Dokumenten ADD COLUMN bstndNaam37 varchar;"
    Err = 0
    adntDB.Execute Msg
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
    Else
        MsgBox Msg, vbInformation, "Met succes"
    End If
    Msg = "ALTER TABLE Dokumenten ADD COLUMN typeZending37 TEXT(5);"
    Err = 0
    adntDB.Execute Msg
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
    Else
        MsgBox Msg, vbInformation, "Met succes"
    End If
    Msg = "ALTER TABLE Dokumenten ADD COLUMN bstBLOB37 OLEobject;"
    Err = 0
    adntDB.Execute Msg
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
    Else
        MsgBox Msg, vbInformation, "Met succes"
    End If
    MsgBox "Belangrijke velden werden toegevoegd. Gelieve het bedrijf opnieuw te openen a.u.b."
    ktrlBLOBRecord = False
End If

End Function


Function xrsMar(Fl As Integer, TBS As String) As Variant

    If IsNull(rsMAR(Fl)(TBS)) Or IsEmpty(rsMAR(Fl)(TBS)) Then
        xrsMar = ""
    Else
        xrsMar = rsMAR(Fl)(TBS)
    End If
    
End Function

Function objectValue(dbValue As Variant) As Variant

    'On Error Resume Next
    If IsNull(dbValue) Or IsEmpty(dbValue) Then
        objectValue = ""
    Else
        objectValue = dbValue
    End If
    
End Function


Function RTV(Fl As Integer) As Integer
Dim T As Integer

RTV = False
TLB_RECORD(Fl) = ""
On Local Error Resume Next
T = 0
Do While Asc(vBC(Fl, T)) <> 0
    'vBib Fl, RTrim$(ntRS(Fl).Fields(vBC(Fl, T))), vBC(Fl, T)
    vBib Fl, RTrim$(rsMAR(Fl).fields(vBC(Fl, T))), vBC(Fl, T)
    T = T + 1
Loop
For T = 0 To FL_NUMBEROFINDEXEN(Fl)
    FVT(Fl, T) = vBibTekst(Fl, "#" + JETTABLEUSE_INDEX(Fl, T) + "#")
Next
RTV = True

End Function


Sub SetFields(Fl As Integer, vBibCode As String, StringData As String)
Dim vBCode As String

vBCode = RTrim(vBibCode)

Kop:
On Local Error GoTo 0  'WatGaatErFout

'Debug.Print vBCode
If RTrim$(vBCode) = "MEMO" Then
    'ntRS(Fl).Fields(vBCode) = TLB_RECORD(Fl)
    rsMAR(Fl).fields(vBCode) = TLB_RECORD(Fl)
Else
    'ntRS(Fl).Fields(vBCode) = StringData
    If rsMAR(Fl).fields(vBCode) = Left(StringData, rsMAR(Fl).fields(vBCode).DefinedSize) Then 'StringData Then
    ElseIf StringData = "" Then
    Else
    On Local Error GoTo 0  'WatGaatErFout
        rsMAR(Fl).fields(vBCode) = Left(StringData, rsMAR(Fl).fields(vBCode).DefinedSize)
    End If
End If
Exit Sub

WatGaatErFout:
Select Case Err

Case Else
    MsgBox vBCode & vbCr & Error & vbCr & Err & vbCr & vbCr & "Geef deze foutcode melding door a.u.b. 053/21.59.25.  Dank U", vbExclamation, "Velden in rijen plaatsen (SetFields)"
    Exit Sub
End Select
Resume

WatGaatErDaarnaFout:
MsgBox "watgaaterdaarnafoutStop"

End Sub


Function daoBlankoRecord(Fl As Integer) As Boolean

TLB_RECORD(Fl) = ""
Select Case Fl
    Case TABLE_CUSTOMERS, TABLE_SUPPLIERS
        vBib Fl, "2", "A10C"    'Taalkode
        vBib Fl, "002", "v149"  'Landnummer  ISO kode
        vBib Fl, "B  ", "A109"  'Landkode Postkantoor
        vBib Fl, "BE", "v150"   'Landkode    ISO kode
        If bhEuro Then
            vBib Fl, "EUR", "vs03"  'Munteenheid ISO kode
        Else
            vBib Fl, "BEF", "vs03"  'Munteenheid ISO kode
        End If
        vBib Fl, "1", "vs07"    'exemplaren dokumenten
    Case TABLE_LEDGERACCOUNTS
        vBib Fl, "O", "v032"    'Budgetcode
    Case TABLE_PRODUCTS
        vBib Fl, fmarBoxText("004", "2", "0"), "v106"
        vBib Fl, Dec$(1, "#####.00"), "v107"
        vBib Fl, fmarBoxText("022", "2", "N"), "v108"
        vBib Fl, fmarBoxText("002", "2", String99(READING, 183)), "v111"
        vBib Fl, String99(READING, 77), "v116"
        vBib Fl, String99(READING, 78), "v117"
        vBib Fl, String99(READING, 79), "v118"
End Select

End Function


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

Sub ClearFlDummy()

bClose TABLE_DUMMY
Ktrl = bOpen(TABLE_DUMMY)
bFirst TABLE_DUMMY, 0
If Ktrl Then
    bClose TABLE_DUMMY
    Exit Sub
Else
    BeginTrans
    'Do While Not ntRS(TABLE_DUMMY).EOF
    '    ntRS(TABLE_DUMMY).Delete
    '    ntRS(TABLE_DUMMY).MoveNext
    'Loop
    Do While Not rsMAR(TABLE_DUMMY).EOF
        rsMAR(TABLE_DUMMY).Delete
        rsMAR(TABLE_DUMMY).MoveNext
    Loop
End If
CommitTrans
bClose TABLE_DUMMY

End Sub


Public Function Editmogelijk(Fl As Integer)
Dim TTT As Integer

Editmogelijk = True
Exit Function
    
ProbeerEdit:
'On Local Error GoTo OpnieuwEdit
Err = 0
'If ntRS(Fl).EditMode = dbEditNone Then
If rsMAR(Fl).EditMode = dbEditNone Then
    MsgBox "stop"
    Editmogelijk = False
Else
    Editmogelijk = True
End If
'If IsNull(ntRS(Fl).Fields(0)) Then
'    ntRS(Fl).AddNew
'Else
'    ntRS(Fl).Edit
'End If
Exit Function

OpnieuwEdit:
If Err = 91 Then
    bOpen Fl
    GoTo ProbeerEdit
ElseIf Err = 3021 Then
    'MsgBox "addnew niet uitgevoerd voor bestand " + ntRS(Fl).Name + " !  Verwittig ons."
    MsgBox "addnew niet uitgevoerd voor tabel " + bstNaam(Fl) + " !  Verwittig ons."
    Err = 0
    'ntRS(Fl).AddNew
    rsMAR(Fl).AddNew
    'msgBox "Addnew 2e poging met succes " + ntRS(Fl).Name + " !  Vergeet ons niet te verwittigen a.u.b. !"
    MsgBox "Addnew 2e poging met succes " + bstNaam(Fl) + " !  Vergeet ons niet te verwittigen a.u.b. !"
    Editmogelijk = True
    Exit Function
End If
    MsgBox Error
      Msg = "De recordgegevens zelf zijn op dit ogenblik in bewerking bij een andere gebruiker !" + vbCrLf
Msg = Msg + "Steeds opnieuw proberen ?" + vbCrLf + vbCrLf
Ktrl = MsgBox(Msg$, 5 + 16, "Schrijfbeveiliging andere gebruiker")
Select Case Ktrl
    Case 4
        Resume
    Case Else
        Exit Function
End Select

End Function

Sub bAbort()

On Error Resume Next
Err = 0
'NTRuimte.Rollback
If InStr(SQLConnect, "Provider=SQL") Then
    adntDBSQLS.RollbackTrans
Else
    adntDB.RollbackTrans
End If
    Ktrl = Err
    If Err Then
        MsgBox Error
    End If

End Sub

Sub bBegin()

On Error Resume Next
Err = 0
'NTRuimte.BeginTrans
If InStr(SQLConnect, "Provider=SQL") Then
    adntDBSQLS.BeginTrans
Else
    adntDB.BeginTrans
End If
Ktrl = Err
If Err Then
    MsgBox Error
End If

End Sub


Sub bEnd()

On Error Resume Next
'NTRuimte.CommitTrans
If InStr(SQLConnect, "Provider=SQL") Then
    adntDBSQLS.CommitTrans
Else
    adntDB.CommitTrans
End If
Ktrl = Err
If Err Then
    MsgBox Error
End If

End Sub

Sub bFirst(Fl As Integer, fIndex As Integer)

MoveFirstNogEens:
If rsMAR(Fl).State = adStateClosed Then
    Ktrl = bOpen(Fl)
End If
If VsoftLog Then
    SchrijfLog "FIRST ", Fl, fIndex, ""
End If

'ntRS(Fl).Index = FLINDEX_CAPTION(Fl, fIndex)
Err = 0
On Error Resume Next
rsMAR(Fl).Index = FLINDEX_CAPTION(Fl, fIndex)
If Err Then
    If Err = -2147217883 Then
        MsgBox Error, vbInformation, "Gecontroleerde foutopvang..."
        bClose Fl: GoTo MoveFirstNogEens
    End If
    'Err = 0
    'On Error Resume Next
    'rsMAR(Fl).MoveFirst
    'If Err Then
    '    MsgBox "Stop"
    '    Exit Sub
    'Else
    '    bClose Fl
    '    GoTo MoveFirstNogEens
    'End If
    Ktrl = 9
    Exit Sub
End If

'If ntRS(Fl).RecordCount = 0 Then
If rsMAR(Fl).RecordCount = 0 Then
    Ktrl = 9
    Exit Sub
Else
    rsMAR(Fl).MoveFirst
'    ntRS(Fl).MoveFirst
    Ktrl = 0
End If
KEY_INDEX(Fl) = fIndex

On Local Error Resume Next
KEY_BUF(Fl) = ""
If IsNull(rsMAR(Fl).fields(Trim$(JETTABLEUSE_INDEX(Fl, fIndex)))) Then
ElseIf FLINDEX_LEN(Fl, fIndex) = 0 Then
    'KEY_BUF(Fl) = Trim$(ntRS(Fl).Fields(Left(JETTABLEUSE_INDEX(Fl, KEY_INDEX(Fl)), 4)))
    KEY_BUF(Fl) = Trim$(rsMAR(Fl).fields(Left(JETTABLEUSE_INDEX(Fl, KEY_INDEX(Fl)), 4)))
Else
    'KEY_BUF(Fl) = vSet(ntRS(Fl).Fields(RTrim$(JETTABLEUSE_INDEX(Fl, fIndex))), FLINDEX_LEN(Fl, fIndex))
    KEY_BUF(Fl) = vSet(rsMAR(Fl).fields(RTrim$(JETTABLEUSE_INDEX(Fl, fIndex))), FLINDEX_LEN(Fl, fIndex))
End If
If Err Then KEY_BUF(Fl) = ""
Exit Sub

AccesErrorFirst:
Select Case Err
    Case 94
        MsgBox Error
        Ktrl = 9
    Case Else
        MsgBox Error + vbCrLf + vbCrLf + "Bestand : " + bstNaam(Fl) + vbCrLf + vbCrLf + "Mogelijke sleutel : " + KEY_BUF(Fl)
End Select
Resume Next

End Sub

Sub bLast(Fl As Integer, fIndex As Integer)

MoveLastNogEens:
If rsMAR(Fl).State = adStateClosed Then
    Ktrl = bOpen(Fl)
End If
If VsoftLog Then
    SchrijfLog "LAST  ", Fl, fIndex, ""
End If

'ntRS(Fl).Index = FLINDEX_CAPTION(Fl, fIndex)
Err = 0
On Error Resume Next
rsMAR(Fl).Index = FLINDEX_CAPTION(Fl, fIndex)
If Err Then
    Err = 0
    On Error Resume Next
    rsMAR(Fl).MoveLast
    If Err Then
        MsgBox "Stop"
        Exit Sub
    Else
        bClose Fl
        GoTo MoveLastNogEens
    End If
End If
        
'If ntRS(Fl).RecordCount = 0 Then
If rsMAR(Fl).RecordCount = 0 Then
    Ktrl = 9
    Exit Sub
Else
    On Error Resume Next
    Err = 0
    'ntRS(Fl).MoveLast
    rsMAR(Fl).MoveLast
    Ktrl = Err
    If Err Then MsgBox Error
End If
KEY_INDEX(Fl) = fIndex
KEY_BUF(Fl) = ""
If IsNull(ntRS(Fl).fields(RTrim$(JETTABLEUSE_INDEX(Fl, fIndex)))) Then
    bPrev Fl
ElseIf FLINDEX_LEN(Fl, fIndex) = 0 Then
    KEY_BUF(Fl) = "..."
Else
    KEY_BUF(Fl) = vSet(ntRS(Fl).fields(RTrim$(JETTABLEUSE_INDEX(Fl, fIndex))), FLINDEX_LEN(Fl, fIndex))
End If

End Sub


Sub bPrev(Fl As Integer)

If rsMAR(Fl).State = adStateClosed Then
    If Fl = TABLE_VARIOUS Then
        bFirst Fl, 1
    Else
        bFirst Fl, 0
    End If
    Exit Sub
End If
If VsoftLog Then
    SchrijfLog "PREV  ", Fl, 0, ""
End If

'If ntRS(Fl).BOF Or ntRS(Fl).EOF Then
If rsMAR(Fl).BOF Or rsMAR(Fl).EOF Then
    bFirst Fl, 0
    Exit Sub
Else
    'ntRS(Fl).MovePrevious
    rsMAR(Fl).MovePrevious
    'If ntRS(Fl).BOF Then
    If rsMAR(Fl).BOF Then
        Ktrl = 9
    Else
        Ktrl = 0
        On Error Resume Next
        KEY_BUF(Fl) = rsMAR(Fl).fields(RTrim$(JETTABLEUSE_INDEX(Fl, KEY_INDEX(Fl))))
    End If
End If

End Sub



Sub bNext(Fl As Integer)
Dim X As Integer

If rsMAR(Fl).State = adStateClosed Then
    If Fl = TABLE_VARIOUS Then
        bLast Fl, 1
    Else
        bLast Fl, 0
    End If
End If
If VsoftLog Then
    SchrijfLog "NEXT  ", Fl, 0, ""
End If

AccessNext:
'KEY_BUF(Fl) = ""
'If ntRS(Fl).BOF Or ntRS(Fl).EOF Then
If rsMAR(Fl).BOF Or rsMAR(Fl).EOF Then
    Ktrl = 9
    MsgBox "Er is geen record (meer)."
    Exit Sub
Else
    On Local Error GoTo AccessNextError
    'ntRS(Fl).MoveNext
    rsMAR(Fl).MoveNext
    'If ntRS(Fl).EOF Then
    If rsMAR(Fl).EOF Then
        Ktrl = 9
    Else
        Ktrl = 0
        'KEY_BUF(Fl) = ntRS(Fl).Fields(Left(JETTABLEUSE_INDEX(Fl, KEY_INDEX(Fl)), 4))
        KEY_BUF(Fl) = rsMAR(Fl).fields(Left(JETTABLEUSE_INDEX(Fl, KEY_INDEX(Fl)), 4))
    End If
End If
Exit Sub

AccessNextError:
Msg = "Database stopkode " + Format(Err) + vbCrLf + vbCrLf
Msg = Msg + "Mededeling :" + vbCrLf
Msg = Msg + Error + vbCrLf + vbCrLf
Msg = Msg + "Steeds opnieuw proberen ?"
KtrlBox = MsgBox(Msg, vbYesNo)
If KtrlBox = vbYes Then
    Resume AccessNext
Else
    Exit Sub
End If

End Sub

Sub bInsert(Fl As Integer, fIndex As Integer)       ', fSleutel As String)
Dim XXXXX As Integer

If Fl = TABLE_INVOICES Then
Else
    If rsMAR(Fl).State = adStateClosed Then
        Ktrl = bOpen(Fl)
    End If
    'ntRS(Fl).AddNew
    rsMAR(Fl).AddNew
End If
XXXXX = VeldToRecord(Fl)
If Ktrl = 32000 Then Exit Sub
KEY_INDEX(Fl) = fIndex
KEY_BUF(Fl) = FVT(Fl, fIndex)
If VsoftLog Then
    SchrijfLog "INSERT", Fl, fIndex, ""
End If

If Fl = TABLE_JOURNAL Then
    DKTRL_CUMUL = DKTRL_CUMUL + Val(rsMAR(TABLE_JOURNAL)("v068").Value)
    If bhEuro Then
        DKTRL_BEF = DKTRL_BEF + Round(Val(rsMAR(TABLE_JOURNAL)("v068").Value) * EURO, 0)
        DKTRL_EUR = DKTRL_EUR + Round(Val(rsMAR(TABLE_JOURNAL)("v068").Value), 2)
        On Error Resume Next
        rsMAR(TABLE_JOURNAL)("dece068") = Val(rsMAR(TABLE_JOURNAL)("v068").Value)
        Err = 0
        On Error Resume Next
    Else
        DKTRL_BEF = DKTRL_BEF + Round(Val(rsMAR(TABLE_JOURNAL)("v068").Value), 0)
        DKTRL_EUR = DKTRL_EUR + Round(Val(rsMAR(TABLE_JOURNAL)("v068").Value) / EURO, 2)
    End If
                       
    Dim Pipo As Variant
    Pipo = rsMAR(Fl)("v019") & vbTab & rsMAR(Fl)("v067") & vbTab
    If bhEuro Then
        If Val(rsMAR(Fl)("v068")) < 0 Then
            Pipo = Pipo & "" & vbTab & Format(-Val(rsMAR(Fl)("v068")), "#,##0.00")
            Pipo = Pipo & vbTab & "" & vbTab & Format(Round(-Val(rsMAR(Fl)("v068")) * EURO, 0), "#,##0.00")
        Else
            Pipo = Pipo & Format(Val(rsMAR(Fl)("v068")), "#,##0.00") & vbTab & ""
            Pipo = Pipo & vbTab & Format(Round(Val(rsMAR(Fl)("v068")) * EURO, 0), "#,##0.00") & vbTab & ""
        End If
    Else
        If Val(rsMAR(Fl)("v068")) < 0 Then
            Pipo = Pipo & "" & vbTab & Format(-Val(rsMAR(Fl)("v068")) / EURO, "#,##0.00")
            Pipo = Pipo & vbTab & "" & vbTab & Format(-Val(rsMAR(Fl)("v068")), "#,##0.00")
        Else
            Pipo = Pipo & Format(Val(rsMAR(Fl)("v068")) / EURO, "#,##0.00") & vbTab & ""
            Pipo = Pipo & vbTab & Format(Val(rsMAR(Fl)("v068")), "#,##0.00") & vbTab & ""
        End If
    End If
    frmBoeking.mshfBoekLijst.AddItem Pipo, frmBoeking.mshfBoekLijst.Rows - 1
End If
    
On Error GoTo AccesErrorInsert
Ktrl = 0
'ntRS(Fl).Update
rsMAR(Fl).Update
If Fl = TABLE_JOURNAL Then
    If Ktrl Then
        MsgBox "bInsert journaal stopkode " + Str$(Ktrl)
    Else
        bGet TABLE_LEDGERACCOUNTS, 0, Left(FVT(TABLE_JOURNAL, 0), 7)
        If Ktrl Then
            MsgBox "Rekening " + Left(FVT(TABLE_JOURNAL, 0), 7) + " niet te vinden." + vbCrLf + "Eerst SETUPrekening inbrengen a.u.b. !"
            DKTRL_CUMUL = DKTRL_CUMUL + 99
            Exit Sub
        ElseIf ACTIVE_BOOKYEAR Then
            RecordToVeld TABLE_LEDGERACCOUNTS
            vBib TABLE_LEDGERACCOUNTS, Str$(Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#e023 #")) + Val(vBibTekst(TABLE_JOURNAL, "#v068 #"))), "e023"
            rsMAR(TABLE_LEDGERACCOUNTS)("dece023") = rsMAR(TABLE_LEDGERACCOUNTS)("dece023") + rsMAR(TABLE_JOURNAL)("dece068")
        Else
            RecordToVeld TABLE_LEDGERACCOUNTS
            vBib TABLE_LEDGERACCOUNTS, Str$(Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#e022 #")) + Val(vBibTekst(TABLE_JOURNAL, "#v068 #"))), "e022"
            rsMAR(TABLE_LEDGERACCOUNTS)("dece022") = rsMAR(TABLE_LEDGERACCOUNTS)("dece022") + rsMAR(TABLE_JOURNAL)("dece068")
        End If
        bUpdate TABLE_LEDGERACCOUNTS, 0
    End If
End If
Select Case Ktrl
    Case 0
    Case 5
        Msg = "Dergelijke ID.Kode Bestaat reeds : " + KEY_BUF(Fl) + " : " + Str$(Fl)
        MsgBox Msg
    Case 46
        Msg$ = "Bestand werd geopend in LEES-modus." + vbCrLf + "Schrijven is niet mogelijk..."
        MsgBox Msg$, 0, "Database beveiliging"
    
    Case Else
        Msg$ = "Stopkode " + Str$(Ktrl) + " tijdens invoegen nieuwe record."
        MsgBox Msg$
End Select
Exit Sub

AccesErrorInsert:
Select Case Err
    Case 3022
        MsgBox "Unieke sleutel reeds aanwezig in bestand : " + bstNaam(Fl) + vbCrLf + vbCrLf + "Mogelijke sleutel : " + FVT(Fl, fIndex)
        Ktrl = Err
    Case Else
        If IsNull(rsMAR(Fl)(fIndex)) Then
            MsgBox Error + vbCrLf + vbCrLf + "Bestand : " + bstNaam(Fl) + vbCrLf + vbCrLf + "De sleutel heeft 'null' waarde", vbExclamation
        Else
            MsgBox Error + vbCrLf + vbCrLf + "Bestand : " + bstNaam(Fl) + vbCrLf + vbCrLf + "Mogelijke sleutel : " + FVT(Fl, fIndex), vbExclamation
        End If
        Ktrl = Err
End Select
Resume Next

End Sub


Sub bGetOrGreater(Fl As Integer, fIndex As Integer, fSleutel As String)

On Local Error Resume Next

opnieuwGOG:
If rsMAR(Fl).State = adStateClosed Then
    Ktrl = bOpen(Fl)
End If
If VsoftLog Then
    SchrijfLog "GETOG ", Fl, fIndex, fSleutel
End If

Err = 0
If InStr(SQLConnect, "SQLOLEDB") Then
    If ADO_GET(Fl, fIndex, "Like", fSleutel & "%") Then
        Ktrl = 0
    Else
        Ktrl = 4
    End If
Else
    fSleutel = vSet(fSleutel, FLINDEX_LEN(Fl, fIndex))
    'If ntRS(Fl).Index = FLINDEX_CAPTION(Fl, fIndex) Then
    If rsMAR(Fl).Index = FLINDEX_CAPTION(Fl, fIndex) Then
    Else
        'ntRS(Fl).Index = FLINDEX_CAPTION(Fl, fIndex)
        rsMAR(Fl).Index = FLINDEX_CAPTION(Fl, fIndex)
        If Err = -2147217883 Then
            'MsgBox "Stop"
            bClose Fl
            DoEvents
            GoTo opnieuwGOG
        End If
    End If
    'ntRS(Fl).Seek "=", fSleutel
    rsMAR(Fl).Seek fSleutel, adSeekAfterEQ
    
    'If ntRS(Fl).NoMatch Then
    If rsMAR(Fl).EOF Then
        Ktrl = 4
    Else
        Ktrl = 0
    End If
    KEY_BUF(Fl) = rsMAR(Fl).fields(RTrim(JETTABLEUSE_INDEX(Fl, fIndex)))
    KEY_INDEX(Fl) = fIndex
    'KEY_BUF(Fl) = vSet(fSleutel, FLINDEX_LEN(Fl, fIndex))
End If

End Sub




Sub bDelete(Fl As Integer)

Ktrl = 0
If VsoftLog Then
    SchrijfLog "DELETE", Fl, 0, ""
End If

On Error Resume Next
Err = 0
'ntRS(Fl).Delete
rsMAR(Fl).Delete
Ktrl = Err
If Err Then
    MsgBox Error
End If

End Sub



Public Function VeldOK(flHier As Integer, VeldNaam As String, Optional VeldDef As String) As Long

    Dim AantalRC As Integer
    
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


Function TabelKontrole() As Boolean

   Dim cnnHier As New ADODB.Connection
   Dim catHier As New ADOX.Catalog

   TabelKontrole = False
   cnnHier.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source= " & LOCATION_COMPANYDATA & _
      "marnt.mdv;"
   Set catHier.ActiveConnection = cnnHier
   
   For COUNT_TO = 0 To catHier.Tables.Count - 1
        Select Case catHier.Tables(COUNT_TO).Name
            Case "1985" To "2061"
                catHier.Tables(COUNT_TO).Name = "jr" & catHier.Tables(COUNT_TO).Name
        End Select
   Next
   catHier.Tables.Refresh
   cnnHier.Close
   Set catHier = Nothing
   Set cnnHier = Nothing
   
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


Function fmarBoxText(marBoxNumber As String, Taal As String, marBoxOption As String) As String
Dim FlTemp As Integer
Dim OptieLen As Integer
Dim ListOptie As String
Dim ZoekTekst As String

If Len(marBoxNumber) = 2 Then
    ZoekTekst = "NTKB" + Taal + "9"
ElseIf Len(marBoxNumber) = 3 Then
    ZoekTekst = "NTKB" + Taal
ElseIf Len(marBoxNumber) = 4 Then
    MsgBox "Stop"
    ZoekTekst = "NT"
Else
    MsgBox "fmarBoxText fout"
    Exit Function
End If
ZoekTekst = ZoekTekst + marBoxNumber
fmarBoxText = ZoekEnPlaats(KeuzeVSF.NTBoxLijst, ZoekTekst, 0, 0, marBoxOption)

End Function



Function adxKolom(tbNaam As String, clNaam As String, clType As Long, clLengte As Long) As Boolean

    Dim cat     As New ADOX.Catalog
    Dim clCol   As New ADOX.Column
                  
    ' Waarde van clType:
    ' adDate voor DATUM/TIJD
    ' adInteger voor INTEGER
    ' adCurrency  'voor geldwaarden (4 cijfers na de comma)
    ' adSingle voor SINGLE
    ' adDouble voor DUBBELE NAUWKEURIGHEID
    ' adBoolean voor BOOLEAN
    ' adVarWChar voor TEKST. Parameters Size indien vaste lengte.
    ' adLongVarWChar voor MEMO
    
    adxKolom = False
    Err = 0
    On Error Resume Next
    If InStr(adntDB.Properties("DBMS Name"), "SQL Server") Then
        ' De catalogus openen voor SQL-SERVER
        cat.ActiveConnection = SQLConnect
    Else
        ' De catalogus openen voor JET4
        cat.ActiveConnection = jetConnect
    End If
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        Set cat.ActiveConnection = Nothing
        Exit Function
    Else
        With clCol
            .Name = clNaam
            .ParentCatalog = cat
            .Type = clType
            If clLengte <> 0 Then
                .DefinedSize = clLengte
            End If
            .Properties("Nullable") = True
        End With
        cat.Tables(tbNaam).Columns.Append clCol
        Set cat = Nothing
        If Err Then
            MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
            MsgBox "Aanmaak Kolom " + clNaam + " zonder succes.", vbInformation
        Else
            adxKolom = True
        End If
    End If
    Set cat.ActiveConnection = Nothing

End Function


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


Function String99(LockModus As Integer, SZNummer As Integer) As String
    Dim TlString As String * 5

    TlString = "s" + Format(SZNummer, "000")
    If LockModus = READING Then
        LockHold = False
    Else
        LockHold = True
    End If
    bGet TABLE_COUNTERS, 0, TlString

    If Ktrl = 99 Then
    ElseIf Ktrl Then
        MsgBox "Tellers Stopkode " + Format(Ktrl) + ", voor setup-tellersleutel " + TlString + vbCrLf + vbCrLf + "Controleer setup instellingen vooraleer op te starten of verder te werken!" + vbCrLf + "Wij staan tot uw beschikking om U hierbij te helpen."
    Else
        RecordToVeld TABLE_COUNTERS
        'String99 = ntRS(TABLE_COUNTERS).Fields("v217")
        On Error Resume Next
        Err = 0
        String99 = rsMAR(TABLE_COUNTERS).fields("v217")
        If Err Then MsgBox Error
    End If
    bClose TABLE_COUNTERS

End Function


Function VeldToRecord(Fl As Integer) As Integer
Dim T As Integer

On Error GoTo 0
If rsMAR(Fl).State = adStateClosed Then
    Ktrl = bOpen(Fl)
'ElseIf Editmogelijk(Fl) = False Then
'    Ktrl = 32000
'    Exit Function
End If

If Fl = TABLE_CONTRACTS Then
    vBib Fl, vSet(vBibTekst(Fl, "#v164 #"), 2) + vSet(vBibTekst(Fl, "#A110 #"), 12) + vSet(vBibTekst(Fl, "#A010 #"), 4) + vSet(vBibTekst(Fl, "#A000 #"), 12), "v167" 'MaandKlantMaatschappijPolis
ElseIf Fl = TABLE_JOURNAL Then
    vBib Fl, vSet(vBibTekst(Fl, "#v019 #"), 7) + vBibTekst(Fl, "#v066 #"), "v070"
End If
For T = 0 To FL_NUMBEROFINDEXEN(Fl)
    FVT(Fl, T) = vSet(vBibTekst(Fl, "#" + JETTABLEUSE_INDEX(Fl, T) + "#"), FLINDEX_LEN(Fl, T))
Next
vBib Fl, FVT(Fl, 0), JETTABLEUSE_INDEX(Fl, 0)
T = 0
Do While Asc(vBC(Fl, T)) <> 0
    SetFields Fl, vBC(Fl, T), vBibTekst(Fl, "#" + vBC(Fl, T) + " #")
    T = T + 1
Loop
On Local Error Resume Next
If Fl = TABLE_VARIOUS Then
    rsMAR(TABLE_VARIOUS)("A000").Value = vBibTekst(TABLE_VARIOUS, "#A000 #")
End If
    
End Function


Sub bClose(Fl As Integer)
Dim T As Integer

If Fl = 99 Then
    For Fl = 0 To NUMBER_TABLES
        TLB_RECORD(Fl) = ""
        GoSub closeTable
    Next
Else
    GoSub closeTable
End If
Exit Sub

closeTable:
If rsMAR(Fl).State = adStateClosed Then
    Return
Else
    Err = 0
    On Local Error Resume Next
    rsMAR(Fl).Close
    Ktrl = Err
    If Err = 3420 Then MsgBox Error
    Return
End If

End Sub

Function bOpen(Fl As Integer) As Integer
Dim DataLijn As String
Dim FlNr As Integer
Dim Dlen As Integer

If rsMAR(Fl).State = adStateClosed Then
Else
    bOpen = 0
    Exit Function
End If

Err = 0
On Error Resume Next
If InStr(SQLConnect, "Provider=SQL") Then
    rsMAR(Fl).CursorLocation = adUseClient
    If Fl = TABLE_COUNTERS Then
        rsMAR(Fl).Open bstNaam(Fl), adntDBSQLS, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    Else
        rsMAR(Fl).Open bstNaam(Fl), adntDBSQLS, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    End If
Else
    rsMAR(Fl).CursorLocation = adUseServer
    If Fl = TABLE_COUNTERS Then
        rsMAR(Fl).Open bstNaam(Fl), adntDB, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    Else
        rsMAR(Fl).Open bstNaam(Fl), adntDB, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    End If
End If
'ntRS(Fl).LockEdits = False
If Err Then
    MsgBox Error
    Ktrl = Err
Else
    bOpen = 0
End If

End Function


Sub bGet(Fl As Integer, fIndex As Integer, fSleutel As String)
Dim ADOCtrl As Integer
Dim ProbeerTellertje As Integer

ProbeerTellertje = 0

bGetNogEens:
On Local Error Resume Next
'adStateOpen and adStateExecuting
If rsMAR(Fl).State = adStateClosed Then
    Ktrl = bOpen(Fl)
End If
If VsoftLog Then
    SchrijfLog "GET   ", Fl, fIndex, fSleutel
End If

Err = 0
If InStr(SQLConnect, "SQLOLEDB") Then
    If ADO_GET(Fl, fIndex, "=", fSleutel) Then
        Ktrl = 0
    Else
        Ktrl = 4
    End If
Else
    Err = 0
    fSleutel = vSet(fSleutel, FLINDEX_LEN(Fl, fIndex))
    If rsMAR(Fl).Index = FLINDEX_CAPTION(Fl, fIndex) Then
    Else
        rsMAR(Fl).Index = FLINDEX_CAPTION(Fl, fIndex)
    End If
    If Err Then
        SnelHelpPrint Error, BL_LOGGING
        bClose Fl
        ADOCtrl = 1
        ProbeerTellertje = ProbeerTellertje + 1
        If ProbeerTellertje > 5 Then
            Ktrl = 99
            Exit Sub
        Else
            GoTo bGetNogEens
            Exit Sub
        End If
    End If
    rsMAR(Fl).Seek fSleutel, adSeekFirstEQ
    If rsMAR(Fl).EOF Then
        Ktrl = 4
    ElseIf Err Then
        MsgBox Error & vbCr & vbCr & "Hierna wordt foutcode 4 doorgegeven", vbExclamation, "bGet routine"
        Ktrl = 4
    Else
        Ktrl = 0
    End If
    KEY_BUF(Fl) = vSet(fSleutel, FLINDEX_LEN(Fl, fIndex))
    KEY_INDEX(Fl) = fIndex
End If

End Sub

Sub bUpdate(Fl As Integer, fIndex As Integer)
Dim XXXXX As Integer

On Error Resume Next
Err = 0
'ntRS(Fl).Edit
'If Err Then
'    MsgBox Error
'End If
XXXXX = VeldToRecord(Fl)
If Ktrl = 32000 Then Exit Sub
KEY_BUF(Fl) = FVT(Fl, fIndex)
KEY_INDEX(Fl) = fIndex
'MsgBox "Stop voor dnnsync"
rsMAR(Fl)("dnnsync") = False
If VsoftLog Then
    SchrijfLog "UPDATE", Fl, fIndex, ""
End If

AccessUpdate:
On Local Error GoTo ErrorUpdate
'ntRS(Fl).Update
rsMAR(Fl).Update
Exit Sub

ErrorUpdate:
If Err = 3197 Then
    MsgBox "Andere gebruiker heeft bewerking uitgevoerd !"
    bGet Fl, fIndex, KEY_BUF(Fl)
    If Ktrl Then
    Else
        RecordToVeld Fl
    End If
    Exit Sub
Else
    Msg = "Database stopkode " + Format(Err) + vbCrLf + vbCrLf
    Msg = Msg + "Mededeling:" + vbCrLf
    Msg = Msg + Error + vbCrLf + vbCrLf
    Msg = Msg + "Steeds opnieuw proberen ?"
    KtrlBox = MsgBox(Msg, vbYesNo)
    If KtrlBox = vbYes Then
        Resume AccessUpdate
    Else
        Ktrl = 99
        Exit Sub
    End If
End If

End Sub

Sub RecordToVeld(Fl As Integer)
Dim T As Integer
Dim b As String

TLB_RECORD(Fl) = ""
On Local Error Resume Next
T = 0
If Fl = TABLE_VARIOUS Then
    'TLB_RECORD(Fl) = ntRS(Fl).Fields("MEMO")
    TLB_RECORD(Fl) = rsMAR(Fl).fields("MEMO")
ElseIf Fl = TABLE_DUMMY Then
    'TLB_RECORD(Fl) = ntRS(Fl).Fields("MEMO")
    TLB_RECORD(Fl) = rsMAR(Fl).fields("MEMO")
Else
    Do While Asc(vBC(Fl, T)) <> 0
        'vBib Fl, Trim$(ntRS(Fl).Fields(vBC(Fl, T))), vBC(Fl, T)
        vBib Fl, rsMAR(Fl).fields(vBC(Fl, T)), vBC(Fl, T)
        T = T + 1
    Loop
End If
'If Err Then MsgBox Error
For T = 0 To FL_NUMBEROFINDEXEN(Fl)
    FVT(Fl, T) = vBibTekst(Fl, "#" + JETTABLEUSE_INDEX(Fl, T) + "#")
Next
If Err Then MsgBox Error

End Sub


Sub vBib(Fl As Integer, StringTekst1 As String, StringTekst2 As String)
Dim TBLen As Long
Dim TBStart As Long
Dim TBStop As Long
Dim TBCode As String * 7

TBCode = "#     #"
Mid(TBCode, 2, 5) = StringTekst2

If StringTekst1 = "" Then
    StringTekst1 = " "
ElseIf InStr(StringTekst1, "#") Then
    SnelHelpPrint "U gebruikte het verboden '#' teken !!!", False
    Exit Sub
End If

jump:
If InStr(TLB_RECORD(Fl), TBCode) = 0 Then
    TLB_RECORD(Fl) = TLB_RECORD(Fl) + TBCode + StringTekst1 + "#"
Else
    If RTrim$(vBibTekst(Fl, TBCode)) = StringTekst1 Then
        Exit Sub
    Else
        TBLen = Len(TLB_RECORD(Fl))
        TBStart = InStr(TLB_RECORD(Fl), TBCode)
        TBStop = InStr(TBStart + 7, TLB_RECORD(Fl), "#")
        TLB_RECORD(Fl) = Left(TLB_RECORD(Fl), TBStart - 1) + Right(TLB_RECORD(Fl), TBLen - TBStop)
        GoTo jump
    End If
End If

End Sub


Function vBibTekst(Fl As Integer, TBS As String) As String

Dim tbsHier As String * 7
If Left(TBS, 1) = "#" Then
    tbsHier = TBS
Else
    tbsHier = "#     #"
    Mid(tbsHier, 2) = TBS
End If
Err = 0
On Error Resume Next
If TLB_RECORD(Fl) = "" Then
    vBibTekst = ""
Else
    vBibTekst = Mid(TLB_RECORD(Fl), InStr(TLB_RECORD(Fl), tbsHier) + 7, InStr(InStr(TLB_RECORD(Fl), tbsHier) + 7, TLB_RECORD(Fl), "#") - (InStr(TLB_RECORD(Fl), tbsHier) + 7))
End If
If Err Then vBibTekst = ""

End Function




Function ADOBIB_TEXT(adoField As ADODB.Field, TBS As String) As String

    Err = 0
    On Error Resume Next
    If adoField = "" Then
        ADOBIB_TEXT = ""
    Else
        ADOBIB_TEXT = Mid(adoField, InStr(adoField, TBS) + 7, InStr(InStr(adoField, TBS) + 7, adoField, "#") - (InStr(adoField, TBS) + 7))
    End If
    If Err Then ADOBIB_TEXT = ""

End Function

Function ADO_GET(iTabel As Integer, iIndex As Integer, sZoals As String, sZoek As Variant) As Boolean

    Dim MsgHier As String

    On Error Resume Next
    If rsMAR(iTabel).State = adStateClosed Then
        Ktrl = bOpen(iTabel)
    End If
    ADO_GET = False
    Err = 0

    rsMAR(iTabel).Index = FLINDEX_CAPTION(iTabel, iIndex)
    If sZoals = "=" Then
        rsMAR(iTabel).Seek sZoek
    ElseIf sZoals = ">=" Then
        rsMAR(iTabel).Seek sZoek, adSeekAfterEQ
    Else
        MsgBox sZoals & " nog niet beschikbaar"
    End If
    If rsMAR(iTabel).EOF Then
    Else
        ADO_GET = True
    End If

End Function


Function RV(adoRecord As ADODB.Recordset, TBS As String) As Variant

    If IsNull(adoRecord(TBS)) Or IsEmpty(adoRecord(TBS)) Then
        RV = ""
    Else
        RV = adoRecord(TBS)
    End If

End Function


Function ADOJOURNAL_OK() As Boolean

    ADOJOURNAL_OK = False
    If Val(RV(rsJournaal, "v068")) = 0 Then
        SnelHelpPrint "BoekBedrag is 0", BL_LOGGING
    Else
        DKTRL_CUMUL = DKTRL_CUMUL + Val(RV(rsJournaal, "v068"))
        DKTRL_BEF = DKTRL_BEF + Round(Val(RV(rsJournaal, "v068")) * EURO, 0)
        DKTRL_EUR = DKTRL_EUR + Round(Val(RV(rsJournaal, "v068")), 2)
        On Error Resume Next
        rsJournaal("dece068") = Val(RV(rsJournaal, "v068"))
        Err = 0
        On Error Resume Next
    
        Dim Pipo As Variant
        Pipo = rsJournaal("v019") & vbTab & rsJournaal("v067") & vbTab
        If Val(rsJournaal("v068")) < 0 Then
            Pipo = Pipo & "" & vbTab & Format(-Val(rsJournaal("v068")), "#,##0.00")
            Pipo = Pipo & vbTab & "" & vbTab & Format(Round(-Val(rsJournaal("v068")) * EURO), "#,##0.00")
        Else
            Pipo = Pipo & Format(Val(rsJournaal("v068")), "#,##0.00") & vbTab & ""
            Pipo = Pipo & vbTab & Format(Round(Val(rsJournaal("v068")) * EURO), "#,##0.00") & vbTab & ""
        End If
        frmBoeking.mshfBoekLijst.AddItem Pipo, frmBoeking.mshfBoekLijst.Rows - 1
    End If

    rsJournaal("v070") = vSet(rsJournaal("v019"), 7) + rsJournaal("v066")
    bGet TABLE_LEDGERACCOUNTS, 0, rsJournaal("v019")
    If Ktrl Then
        MsgBox "Rekening " + rsJournaal("v019") + " niet te vinden." + vbCrLf + "Eerst SETUPrekening inbrengen a.u.b. !"
        DKTRL_CUMUL = DKTRL_CUMUL + 99
        Exit Function
    ElseIf ACTIVE_BOOKYEAR Then
        RecordToVeld TABLE_LEDGERACCOUNTS
        vBib TABLE_LEDGERACCOUNTS, Str$(Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#e023 #")) + Val(RV(rsJournaal, "v068"))), "e023"
        rsMAR(TABLE_LEDGERACCOUNTS)("dece023") = rsMAR(TABLE_LEDGERACCOUNTS)("dece023") + RV(rsJournaal, "dece068")
    Else
        RecordToVeld TABLE_LEDGERACCOUNTS
        vBib TABLE_LEDGERACCOUNTS, Str$(Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#e022 #")) + Val(RV(rsJournaal, "v068"))), "e022"
        rsMAR(TABLE_LEDGERACCOUNTS)("dece022") = rsMAR(TABLE_LEDGERACCOUNTS)("dece022") + RV(rsJournaal, "dece068")
    End If
    rsMAR(TABLE_LEDGERACCOUNTS)("dnnsync") = False
    bUpdate TABLE_LEDGERACCOUNTS, 0
    Err = 0
    On Error Resume Next
    rsJournaal.Update
    If Err Then
        MsgBox Error
    Else
        ADOJOURNAL_OK = True
    End If
    
End Function


