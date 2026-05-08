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


