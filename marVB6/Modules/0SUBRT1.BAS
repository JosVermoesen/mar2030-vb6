Attribute VB_Name = "Module4"
DefInt A-Z


Function BankOk(fRekString As String) As Integer
Dim dPip As Double
Dim dPip2 As Double

BankOk = False
If Len(Trim$(fRekString)) = 14 Then
    fRekString = Mid(fRekString, 1, 3) + Mid(fRekString, 5, 7) + Mid(fRekString, 13, 2)
ElseIf Len(fRekString) <> 12 Then
    Exit Function
End If
dPip = Val(Mid(fRekString, 1, 3) + Mid(fRekString, 4, 7))
If Mid(fRekString, 11, 2) = "00" Then
    Exit Function
ElseIf dPip - Int(dPip / 97) * 97 = 0 And Mid(fRekString, 11, 2) = "97" Then
    BankOk = True
    Exit Function
ElseIf dPip - Int(dPip / 97) * 97 = Val(Mid(fRekString, 11, 2)) Then
    BankOk = True
End If

End Function

Function BtwKontrole(BtwString As String, bStrip As Boolean) As String
ReDim dPip(3) As Double
Dim ipip As Integer

On Local Error GoTo ErrorBtw

dPip(1) = Val(Mid(BtwString, 1, 8)) 'voorheen Val(Mid(BtwString, 1, 3) + Mid(BtwString, 5, 3) + Mid(BtwString$, 9, 1))
dPip(2) = dPip(1) / 97
dPip(3) = dPip(2) - Int(dPip(2))
ipip = 97 - (dPip(3) * 97)
If ipip <> Val(Right(BtwString, 2)) Then
    BtwKontrole = ""
Else
    BtwKontrole = BtwString
End If
Exit Function

ErrorBtw:
BtwKontrole = ""
Exit Function

End Function

Function CopyFile(SourcePath As String, TargetPath As String, FileToCopy As String) As Integer
Dim FlCopy1 As Integer
Dim FlCopy2 As Integer
Dim BigChunks As Integer
Dim LeftOver As Integer
Dim inD As Integer
Dim Aantal As Integer
Dim BufferVar As String
ReDim BestandReeks(512) As String

    On Local Error GoTo CopyError
    Screen.MousePointer = vbHourglass
    If InStr(FileToCopy, "?") Or InStr(FileToCopy, "*") Then
        Aantal = 1
        BestandReeks(1) = Dir$(SourcePath + "\" + FileToCopy)
        If BestandReeks(1) = "" Then
            MsgBox "Stop tijdens het kopieren.  Bestand niet te vinden: """ + FileToCopy + """", 64, "SETUP"
            GoTo CopyError
        Else
            Do
                Aantal = Aantal + 1
                BestandReeks(Aantal) = Dir$
                If BestandReeks(Aantal) = "" Then
                    Aantal = Aantal - 1
                    Exit Do
                End If
            Loop
        End If
    Else
        BestandReeks(1) = FileToCopy
        Aantal = 1
        If Not FileExists(SourcePath + "\" + FileToCopy) Then
            MsgBox "Bestand niet te vinden: """ + FileToCopy + """", 64, "SETUP"
            GoTo CopyError
        End If
    End If

    For COUNT_TO = 1 To Aantal

    If Dir$(TargetPath + "\" + BestandReeks(COUNT_TO)) <> "" Then
        Kill TargetPath + "\" + BestandReeks(COUNT_TO)
    End If
          
    FlCopy1 = FreeFile
    Open SourcePath + "\" + BestandReeks(COUNT_TO) For Binary As FlCopy1

    FlCopy2 = FreeFile
    Open TargetPath + "\" + BestandReeks(COUNT_TO) For Binary As FlCopy2

    BufferVar = String$(3000, " ")

    BigChunks = LOF(FlCopy1) \ Len(BufferVar)
    LeftOver = LOF(FlCopy1) Mod Len(BufferVar)

    For inD = 1 To BigChunks
        Get FlCopy1, , BufferVar
        Put FlCopy2, , BufferVar
    Next

    If LeftOver Then
        BufferVar = ""
        BufferVar = String$(LeftOver, " ")
        Get FlCopy1, , BufferVar
        Put FlCopy2, , BufferVar
    End If

    If LOF(FlCopy1) = LOF(FlCopy2) Then
        Close FlCopy1
        Close FlCopy2
    Else
        GoTo CopyError
    End If

Next
CopyFile = -1
Screen.MousePointer = vbNormal
Exit Function

CopyError:
    MsgBox "Stop tijdens het kopieren van " + FileToCopy + """", 64, "SETUP"
    MsgBox Error$
    Close FlCopy1
    Close FlCopy2
    CopyFile = False
    Screen.MousePointer = vbNormal
    Exit Function

End Function


'---------------------------------------------
'Create the path contained in DestPath$
'First char must be drive letter, followed by
'a ":\" followed by the path, if any.
'---------------------------------------------
Function CreatePath(destpath As String) As Boolean

Dim BackPos As Integer
Dim ForePos As Integer
Dim Temp As String

    Screen.MousePointer = vbHourglass
    CreatePath = False
    '---------------------------------------------
    'Add slash to end of path if not there already
    '---------------------------------------------
    If Right(destpath$, 1) <> "\" Then
        destpath$ = destpath$ + "\"
    End If


    '-----------------------------------
    'Change to the root dir of the drive
    '-----------------------------------
    On Error Resume Next
    ChDrive destpath$
    If Err <> 0 Then GoTo errorOut
    ChDir "\"

    '-------------------------------------------------
    'Attempt to make each directory, then change to it
    '-------------------------------------------------
    BackPos = 3
    ForePos = InStr(4, destpath$, "\")
    Do While ForePos <> 0
        Temp = Mid(destpath, BackPos + 1, ForePos - BackPos - 1)

        Err = 0
        MkDir Temp$
        If Err <> 0 And Err <> 75 Then GoTo errorOut

        Err = 0
        ChDir Temp$
        If Err <> 0 Then GoTo errorOut

        BackPos = ForePos
        ForePos = InStr(BackPos + 1, destpath$, "\")
    Loop

    CreatePath = True
    Screen.MousePointer = vbNormal
    Exit Function

errorOut:
    MsgBox "Stop tijdens aanmaak van inhoudsopgaves op de doeldisk.", 48, "SETUP"
    CreatePath = False
    Screen.MousePointer = vbNormal

End Function

Function DATE_INVALID(fDatum As String) As Integer
Dim Dag As Integer
Dim Maand As Integer
Dim Jaar As Integer

If Len(fDatum) <> 10 Then
    DATE_INVALID = True
Else
    Dag = Val(Mid(fDatum, 1, 2))
    Maand = Val(Mid(fDatum, 4, 2))
    Jaar = Val(Mid(fDatum, 7, 4))
    If Dag < 32 And Dag > 0 And Maand < 13 And Maand > 0 And Jaar > 1985 And Jaar < 2062 Then
        DATE_INVALID = False
    Else
        DATE_INVALID = True
        Beep
    End If
End If

End Function

Function DATE_KEY(DatumfTXT As String) As String
Dim Dag As String * 2
Dim Maand As String * 2
Dim Jaar As String * 4

Dag = Mid(DatumfTXT, 1, 2)
Maand = Mid(DatumfTXT, 4, 2)
Jaar = Mid(DatumfTXT, 7, 4)

DATE_KEY = Jaar + Maand + Dag

End Function

Function DATE_CHECK(fDatum As String, fVlag As Integer) As Boolean

Dim Dag As String * 2
Dim Maand As String * 2
Dim Jaar As String * 4
Dim gDatum As String
Dim gPos As Integer

gDatum = fDatum
Do While InStr(gDatum, "/")
    gPos = InStr(gDatum, "/")
    gDatum = Left(gDatum, gPos - 1) + Mid(gDatum, gPos + 1)
Loop

Select Case fVlag
    Case PERIODAS_TEXT, BOOKYEARAS_TEXT
        Dag = Mid(gDatum, 1, 2)
        Maand = Mid(gDatum, 3, 2)
        Jaar = Mid(gDatum, 5, 4)
    Case PERIODAS_KEY, BOOKYEAR_KEY
        Jaar = Mid(gDatum, 1, 4)
        Maand = Mid(gDatum, 5, 2)
        Dag = Mid(gDatum, 7, 2)
    Case Else
        MsgBox "Datum onjuist !"
End Select

Select Case fVlag
    Case PERIODAS_TEXT, PERIODAS_KEY
        If (Jaar + Maand + Dag) < Left(PERIOD_FROMTO, 8) Or (Jaar + Maand + Dag) > Right(PERIOD_FROMTO, 8) Then
            DATE_CHECK = False
        Else
            DATE_CHECK = True
        End If
    Case BOOKYEARAS_TEXT, BOOKYEAR_KEY
        If (Jaar + Maand + Dag) < Left(BOOKYEAR_FROMTO, 8) Or (Jaar + Maand + Dag) > Right(BOOKYEAR_FROMTO, 8) Then
            DATE_CHECK = False
        Else
            DATE_CHECK = True
        End If
End Select

End Function


Function DATE_TEXT(dateAsKey As String) As String
    
    Dim Day As String * 2
    Dim Month As String * 2
    Dim Year As String * 4

    Day = Mid(dateAsKey, 7, 2)
    Month = Mid(dateAsKey, 5, 2)
    Year = Mid(dateAsKey, 1, 4)
    DATE_TEXT = Day + "/" + Month + "/" + Year

End Function

'----------------------------------------------------------
' Check for the existence of a file by attempting an OPEN.
'----------------------------------------------------------
Function FileExists(path As String) As Integer
Dim X As Integer

    X = FreeFile

    On Error Resume Next
    Open path$ For Input As X
    If Err = 0 Then
        FileExists = True
    Else
        FileExists = False
    End If
    Close X

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

'------------------------
' Get the size of the file
'------------------------
Function GetFileSize(Source$) As Long
    X = FreeFile
    Open Source$ For Binary Access Read As X
    GetFileSize = LOF(X)
    Close X
End Function


Sub PrintObject(ObjectNummer As Integer, Tabul As Integer, Tekst As String)

On Local Error Resume Next

If ObjectNummer Then
    Afbeeldingen.Picture2.Print Tab(Tabul); Tekst;
Else
    Printer.Print Tab(Tabul); Tekst;
End If

End Sub

Function SleutelDok(fRecordNr As Integer) As String
Dim VoorLetter As String * 2

FL99_RECORD = String99(READING_LOCK, fRecordNr)
Select Case fRecordNr
    Case 1
        VoorLetter = "A0"
    Case 3
        VoorLetter = "A1"
    Case 11
        VoorLetter = "V0"
    Case 13
        VoorLetter = "V1"
    Case 73
        VoorLetter = "B0"
    Case 59
        VoorLetter = "F0"
    Case 121
        VoorLetter = "Q0"
    Case 188
        VoorLetter = "PF"
    Case Else
        MsgBox "Ongeldige record : " + Str$(fRecordNr)
End Select
SleutelDok = VoorLetter + Mid(PERIOD_FROMTO, 1, 4) + Format(Val(FL99_RECORD) + 1, "00000")

End Function


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
    MsgBox "Tellers Stopkode " + Format(Ktrl) + ", voor setup-tellersleutel " + TlString + vbCrLf + vbCrLf + "Overloop ALLE setup instellingen vooraleer énig boekjaar op te starten !" + vbCrLf + "Wij staan tot uw beschikking om U hierbij te helpen."
Else
    RecordToVeld TABLE_COUNTERS
    'String99 = ntRS(TABLE_COUNTERS).Fields("v217")
    On Error Resume Next
    Err = 0
    String99 = rsMAR(TABLE_COUNTERS).Fields("v217")
    If Err Then MsgBox Error
End If
bClose TABLE_COUNTERS

End Function



Function VValdag(rDat1 As String, rvv As String) As String
Dim irdg43, irmd43, irjr43, avd43 As Integer
On Local Error GoTo handlerVVDag

irdg43 = Val(Left(rDat1, 2))
irmd43 = Val(Mid(rDat1, 4, 2))
irjr43 = Val(Right(rDat1, 4)) - 1990
avd43 = Val(rvv)

If avd43 Then
Else
    VValdag = rDat1
    Exit Function
End If

adm1 = DAYS_IN_MONTH(irmd43)
While irdg43 + avd43 > adm1
    avd43 = avd43 - (adm1 - irdg43)
    irdg43 = 0
    If irmd43 = 12 Then
        irmd43 = 1
        irjr43 = irjr43 + 1
    Else
        irmd43 = irmd43 + 1
    End If
    adm1 = DAYS_IN_MONTH(irmd43)
Wend

irdg43 = irdg43 + avd43
If InStr(UCase$(rvv), "E") Then
    irdg43 = adm1
End If
VValdag = Format(irdg43, "00") + "/" + Format(irmd43, "00") + "/" + Format(irjr43 + 1990, "0000")
Exit Function

handlerVVDag:
VValdag = ""
Exit Function

End Function


Function ValidateNumeric(strText As String) As Boolean
    ValidateNumeric = CBool(strText = "" _
        Or strText = "-" _
        Or strText = "-." _
        Or strText = "." _
        Or IsNumeric(strText))
End Function


Public Function IsSchrikkelJaar(intJaar As Integer) As Boolean
    IsSchrikkelJaar = IsDate("2/29/" & intJaar)
End Function


