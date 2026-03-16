Attribute VB_Name = "modDataBaseRoutines"
DefInt A-Z
Option Explicit

Dim X   As Integer

' Copy a file's contents into a BLOB field.
Function FileToBlob(fld As ADODB.Field, filename As String, _
    Optional ChunkSize As Long = 8192)
    Dim fnum As Integer, bytesLeft As Long, bytes As Long
    Dim tmp() As Byte
    
    ' Raise an error if the field doesn't support GetChunk.
    If (fld.Attributes And adFldLong) = 0 Then
        Err.Raise 1001, , "Field doesn't support the GetChunk method."
    End If
    ' Open the file; raise an error if the file doesn't exist.
    If Dir$(filename) = "" Then Err.Raise 53, , "File not found"
    
    fnum = FreeFile
    Open filename For Binary As fnum
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


Sub SnelHelpPrint(PrintTekst As String, Logging As Boolean)


If Left(PrintTekst, 1) = "~" Then
Else
    Mim!SnelHelpTijd.Enabled = True
End If
Mim.SnelHelp.Panels(4).text = PrintTekst
If Logging = True Then
    LOG_PRINT = LOG_PRINT & PrintTekst & vbCrLf
End If

End Sub


Function vSet(fTekst As String, fLengte As Integer) As String
Dim b As String

b = Left(fTekst, fLengte)
vSet = b + Space$(fLengte - Len(b))

End Function


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


Function TeleBibPagina(Fl As Integer)
Dim FlInput As Integer
Dim FFDefinitie As String
Dim DummyString As String
Dim DummYtje As String
Dim KtrlAantal As Long

Dim CurBedrag As Currency
Dim T As Integer

TeleBibPagina = False

On Local Error GoTo TeleBibError

'code= 1 - 1 : Poliskontrole 1 = ON
'      2 - 3 : Selektiekeuze waarde (ListIDX)
'      4 - 4 :
'      5 - 8 : TeleBib
'      9 - 9 : Volgnummer eventueel 1 tot 9
'      10    : * voor verplichte invulling
'Inlaadvolgorde 1 : 000     'hoofddefinitie (VSOFT)
'               2 : xxxG    'uitbreiding gebruiker
'               2 : xxxM    'uitbreiding makelaar

Dim LokaalBestand As String

If Fl <> TABLE_COUNTERS Then
    LokaalBestand = Left(TABLEDEF_ONT(Fl), 3)
Else
    LokaalBestand = "00"
End If

If Dir$(PROGRAM_LOCATION + "Def\" + LokaalBestand + ".Def") = "" Then
    MsgBox "Geen VsoftBib definitie " + LokaalBestand + ".Def"
    Exit Function
End If

EerstEnVooral:
If Dir$(PROGRAM_LOCATION + "Def\" + LokaalBestand + "U.Def") = "" Then
    GoTo GeenUserVoorkeur
Else
    T = 0
    FlInput = FreeFile
    Open PROGRAM_LOCATION + "Def\" + LokaalBestand + "U.Def" For Input As FlInput
    Do While Not EOF(FlInput)
            Input #FlInput, TELEBIB_CODE(T), TELEBIB_TEXT(T), TELEBIB_TYPE(T), TELEBIB_LENGTH(T)
            vBC(Fl, T) = Mid(TELEBIB_CODE(T), 5, 4)
            If TELEBIB_TYPE(T) = "D" And DecimalKTRL = True Then
                GoSub JumpToTheBEAT
            End If
            T = T + 1
    Loop
    Close FlInput
    TELEBIB_CODE(T) = ""
    TELEBIB_LAST = T - 1
    TeleBibPagina = True
    Exit Function
End If

GeenUserVoorkeur:
FlInput = FreeFile
Open PROGRAM_LOCATION + "Def\" + LokaalBestand + ".Def" For Input As FlInput
T = 0
Do While Not EOF(FlInput)
    Input #FlInput, TELEBIB_CODE(T), TELEBIB_TEXT(T), TELEBIB_TYPE(T), TELEBIB_LENGTH(T)
    vBC(Fl, T) = Mid(TELEBIB_CODE(T), 5, 4)
    If TELEBIB_TYPE(T) = "D" And DecimalKTRL = True Then
        GoSub JumpToTheBEAT
    End If
    T = T + 1
Loop
Close FlInput
TELEBIB_LAST = T - 1
TELEBIB_CODE(T) = ""
TeleBibPagina = True

MakelaarIn:
If ProducentNummer = Space$(8) Then
ElseIf Dir(PROGRAM_LOCATION + "Def\" + LokaalBestand + "M.Def") = "" Then
Else
    FlInput = FreeFile
    Open PROGRAM_LOCATION + "Def\" + LokaalBestand + "M.Def" For Input As FlInput
    Do While Not EOF(FlInput)
            Input #FlInput, TELEBIB_CODE(T), TELEBIB_TEXT(T), TELEBIB_TYPE(T), TELEBIB_LENGTH(T)
            vBC(Fl, T) = Mid(TELEBIB_CODE(T), 5, 4)
            If TELEBIB_TYPE(T) = "D" And DecimalKTRL = True Then
                GoSub JumpToTheBEAT
            End If
            T = T + 1
    Loop
    Close FlInput
    TELEBIB_LAST = T - 1
    TELEBIB_CODE(T) = ""
End If
Exit Function

JumpToTheBEAT:
On Error Resume Next
Err = 0
DummYtje = rsMAR(Fl)("dec" & vBC(Fl, T)).Name
If Err Then
    bClose Fl
    If adxKolom(bstNaam(Fl), "dec" & vBC(Fl, T), adCurrency, 0) Then
        SnelHelpPrint "Extra SQL Server compatibel Decimal veld : " & "dec" & vBC(Fl, T) & " met succes bijgevoegd in tabel : " & bstNaam(Fl), BL_LOGGING
        Msg = "UPDATE " & bstNaam(Fl) & " SET " & "dec" & vBC(Fl, T) & "=val(" & vBC(Fl, T) & ")"
        SnelHelpPrint "Cijfers van " & vBC(Fl, T) & " worden overgedragen naar dec" & vBC(Fl, T) & Msg, BL_LOGGING
        GoSub DoeDeUpdate
    End If
Else
    DecimalKTRL = False
End If
Return

DoeDeUpdate:
bFirst Fl, 0
Do While Ktrl = 0
    On Error Resume Next
    Err = 0
    CurBedrag = 0
    CurBedrag = Val(rsMAR(Fl)(vBC(Fl, T)))
    If Err Then
        SnelHelpPrint Error & " ter hoogte van " & rsMAR(Fl)(0) & " in bestand " & bstNaam(Fl), BL_LOGGING
    End If
    rsMAR(Fl)("dec" & vBC(Fl, T)) = CurBedrag
    SnelHelpPrint rsMAR(Fl)(0) & " " & (CurBedrag), BL_LOGGING
    rsMAR(Fl).Update
    bNext Fl
Loop
Return

TeleBibError:
MsgBox "Telebibinlaadfout" + Str$(T) + " error:" + Error$
Close FlInput
TeleBibPagina = False
Exit Function
Resume

End Function


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


Sub SchrijfLog(BtrieveAktie As String, FlNummer As Integer, IndexNummer As Integer, IndexSleutel As String)
Dim FlLog As Integer
Dim RecordLijn As String

RecordLijn = BtrieveAktie + Format(IndexNummer)
Select Case FlNummer
    Case TABLE_VARIOUS
        RecordLijn = RecordLijn + "ALLERLEI"
    Case TABLE_CUSTOMERS
        RecordLijn = RecordLijn + "KLANTEN "
    Case TABLE_SUPPLIERS
        RecordLijn = RecordLijn + "LEVERANC"
    Case TABLE_LEDGERACCOUNTS
        RecordLijn = RecordLijn + "REKENING"
    Case TABLE_PRODUCTS
        RecordLijn = RecordLijn + "PRODUKT "
    Case TABLE_INVOICES
        RecordLijn = RecordLijn + "dokument"
    Case TABLE_JOURNAL
        RecordLijn = RecordLijn + "JOURNAAL"
    Case TABLE_CONTRACTS
        RecordLijn = RecordLijn + "POLISSEN"
    Case TABLE_COUNTERS
        RecordLijn = RecordLijn + "TELLERS "
End Select

Select Case BtrieveAktie
    Case "INSERT", "UPDATE"
        RecordLijn = RecordLijn + TLB_RECORD(FlNummer)
    Case Else
        RecordLijn = RecordLijn + IndexSleutel
End Select

FlLog = FreeFile
Open PROGRAM_LOCATION + "NTIMPORT.LOG" For Append Lock Write As FlLog
    Print #FlLog, RecordLijn; vbCrLf;
Close FlLog

End Sub


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
