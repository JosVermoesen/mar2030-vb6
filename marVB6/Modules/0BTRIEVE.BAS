Attribute VB_Name = "modMaybeToIgnore"
DefInt A-Z
Option Explicit

Dim X   As Integer

Declare Function SHGetFolderPath Lib "SHFOLDER" _
      Alias "SHGetFolderPathA" _
      (ByVal hwnd As Long, _
      ByVal CSIDL As Long, _
      ByVal TOKENHANDLE As Long, _
      ByVal FLAGS As Long, _
      ByVal lpPath As String) As Long

'shFolder
Global Const rvDesktop = 0
Global Const rvStartMenuProgrammas = 2
Global Const rvMijnDocumenten = 5
Global Const rvFavorieten = 6
Global Const rvOpstarten = 7
Global Const rvRecent = 8
Global Const rvVerzondenNaar = 9
Global Const rvStartMenu = 11
Global Const rvDesktop2 = 16
Global Const rvNetHood = 19
Global Const rvFonts = 20
Global Const rvShellNew = 21
Global Const rvAllUsersDesktop = 25
Global Const rvProfilesThisUserApplicationData = 26
Global Const rvPrintHood = 27
Global Const rvProfilesThisUserLocalSettingsApplicationData = 28
Global Const rvTemporaryInternetFiles = 32
Global Const rvProfilesThisUserCookies = 33
Global Const rvProfilesThisUserHistory = 34
Global Const rvProfilesAllUsersApplicationData = 35
Global Const rvWindows = 36
Global Const rvWindowsSysteem = 37
Global Const rvProgramFiles = 38
Global Const rvMijnAfbeeldingen = 39
Global Const rvCommonFiles = 43


' Define your custom errors here.  Be sure to use numbers
' greater than 512, to avoid conflicts with OLE error numbers.
Public Const MyObjectError1 = 1000
Public Const MyObjectError2 = 1010
Public Const MyObjectErrorN = 1234
Public Const MyUnhandledError = 9999

Public Const BATCH_SIZE As Long = 5000

Public Function ConvertSuppliersSEPAs() As Boolean
    On Error GoTo EH

    ConvertSuppliersSEPAs = False
    Screen.MousePointer = vbHourglass

    Dim rs As ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim lastKey As String
    Dim Bezig As Long
    Dim bTrans As Boolean

    Dim sepa As String
    Dim oldAcc As String
    Dim IBAN As String

    Dim numberOldValid As Long
    Dim numberNewValid As Long
    Dim numberOldInvalid As Long
    Dim numberNewInvalid As Long

    '-----------------------------------------
    ' UPDATE-command voorbereiden
    '-----------------------------------------
    Set cmd.ActiveConnection = adntDB
    cmd.CommandText = "UPDATE Leveranciers SET v259 = ?, A170 = ? WHERE A110 = ?"

    cmd.Parameters.Append cmd.CreateParameter("pV259", adVarChar, adParamInput, 50)
    cmd.Parameters.Append cmd.CreateParameter("pA170", adVarChar, adParamInput, 50)
    cmd.Parameters.Append cmd.CreateParameter("pA110", adVarChar, adParamInput, 20)

    lastKey = ""

ReopenRecordset:
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer

    Dim sql As String
    If lastKey = "" Then
        sql = "SELECT A110, v259, A170 FROM Leveranciers WHERE v150='BE' ORDER BY A110"
    Else
        sql = "SELECT A110, v259, A170 FROM Leveranciers WHERE v150='BE' AND A110 > '" & lastKey & "' ORDER BY A110"
    End If

    rs.Open sql, adntDB, adOpenForwardOnly, adLockReadOnly

    If rs.EOF Then GoTo Finished

    '-----------------------------------------
    ' Nieuwe transactie starten
    '-----------------------------------------
    adntDB.BeginTrans
    bTrans = True

    '-----------------------------------------
    ' Verwerking van deze batch
    '-----------------------------------------
    Do While Not rs.EOF
        Dim key As String
        key = rs!A110
        lastKey = key

        sepa = Trim$(rs!v259 & "")
        oldAcc = Trim$(rs!A170 & "")

        '-----------------------------------------
        ' CASE 1: SEPA bestaat al
        '-----------------------------------------
        If sepa <> "" Then
            IBAN = IbanCheck(sepa, True, False)

            If IBAN = "invalid" Then
                numberNewInvalid = numberNewInvalid + 1
                cmd.Parameters("pV259").Value = Null
                cmd.Parameters("pA170").Value = Null
            Else
                numberNewValid = numberNewValid + 1
                cmd.Parameters("pV259").Value = IBAN
                cmd.Parameters("pA170").Value = IIf(oldAcc = "", Null, oldAcc)
            End If

        Else
            '-----------------------------------------
            ' CASE 2: Geen SEPA ? probeer oud nummer
            '-----------------------------------------
            If oldAcc <> "" Then
                IBAN = IbanCheck(oldAcc, True, False)

                If IBAN = "invalid" Then
                    numberOldInvalid = numberOldInvalid + 1
                    cmd.Parameters("pV259").Value = Null
                    cmd.Parameters("pA170").Value = Null
                Else
                    numberOldValid = numberOldValid + 1
                    numberNewValid = numberNewValid + 1
                    cmd.Parameters("pV259").Value = IBAN
                    cmd.Parameters("pA170").Value = oldAcc
                End If
            Else
                GoTo NextRecord
            End If
        End If

        '-----------------------------------------
        ' UPDATE uitvoeren
        '-----------------------------------------
        cmd.Parameters("pA110").Value = key
        cmd.Execute

        Bezig = Bezig + 1

        If (Bezig Mod 100) = 0 Then
            SnelHelpPrint "Bezig met " & Bezig, BL_LOGGING
        End If

        If (Bezig Mod BATCH_SIZE) = 0 Then
            adntDB.CommitTrans
            bTrans = False
            rs.Close
            GoTo ReopenRecordset
        End If

NextRecord:
        rs.MoveNext
    Loop

    adntDB.CommitTrans
    bTrans = False
    rs.Close
    GoTo ReopenRecordset

Finished:
    MsgBox _
        "Aantal oude nummers geldig ? SEPA: " & numberOldValid & vbCrLf & _
        "Aantal nieuwe SEPA geldig: " & numberNewValid & vbCrLf & _
        "Aantal oude nummers ongeldig: " & numberOldInvalid & vbCrLf & _
        "Aantal nieuwe SEPA ongeldig: " & numberNewInvalid, _
        vbInformation, "SEPA Check Leveranciers"

    ConvertSuppliersSEPAs = True

CleanExit:
    Screen.MousePointer = vbNormal
    Exit Function

EH:
    If bTrans Then adntDB.RollbackTrans
    MsgBox "Fout: " & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume CleanExit
End Function





Public Function ConvertDocNumbersToOGM() As Boolean
    On Error GoTo EH

    ConvertDocNumbersToOGM = False
    Screen.MousePointer = vbHourglass

    Dim rs As ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim lastKey As String
    Dim DocumentKey As String
    Dim groupCode As String
    Dim ReferteTxt As String
    Dim ReferteTxtNoFormat As String
    Dim dPip As Double
    Dim chk As String
    Dim Bezig As Long
    Dim bTrans As Boolean

    '-----------------------------------------
    ' UPDATE-command voorbereiden
    '-----------------------------------------
    Set cmd.ActiveConnection = adntDB
    cmd.CommandText = "UPDATE Dokumenten SET v413 = ? WHERE v033 = ?"
    cmd.Parameters.Append cmd.CreateParameter("pV413", adVarChar, adParamInput, 50)
    cmd.Parameters.Append cmd.CreateParameter("pV033", adVarChar, adParamInput, 20)

    '-----------------------------------------
    ' Startpositie
    '-----------------------------------------
    lastKey = ""

ReopenRecordset:
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer

    Dim sql As String
    If lastKey = "" Then
        sql = "SELECT v033 FROM Dokumenten ORDER BY v033 DESC"
    Else
        sql = "SELECT v033 FROM Dokumenten WHERE v033 < '" & lastKey & "' ORDER BY v033 DESC"
    End If

    rs.Open sql, adntDB, adOpenForwardOnly, adLockReadOnly

    If rs.EOF Then GoTo Finished

    '-----------------------------------------
    ' Nieuwe transactie starten
    '-----------------------------------------
    adntDB.BeginTrans
    bTrans = True

    '-----------------------------------------
    ' Verwerking van deze batch
    '-----------------------------------------
    Do While Not rs.EOF
        DocumentKey = rs!v033
        lastKey = DocumentKey   ' <-- voortgang bewaren

        Select Case Left$(DocumentKey, 2)
            Case "V0": groupCode = "8"
            Case "V1": groupCode = "7"
            Case "B0": groupCode = "6"
            Case "F0": groupCode = "5"
            Case "A0": groupCode = "4"
            Case "A1": groupCode = "3"
            Case "Q0": groupCode = "2"
            Case Else
                GoTo NextRecord
        End Select

        ReferteTxtNoFormat = groupCode & Mid$(DocumentKey, 3, 2) & _
                             Mid$(DocumentKey, 5, 4) & Mid$(DocumentKey, 9) & "xx"

        ReferteTxt = "+++" & groupCode & Mid$(DocumentKey, 3, 2) & "/" & _
                     Mid$(DocumentKey, 5, 4) & "/" & Mid$(DocumentKey, 9) & "xx+++"

        dPip = Val(Mid$(ReferteTxt, 4, 3) & Mid$(ReferteTxt, 8, 4) & Mid$(ReferteTxt, 13, 3))
        chk = Format$(dPip - Int(dPip / 97) * 97, "00")
        If chk = "00" Then chk = "97"

        Mid$(ReferteTxt, 16, 2) = chk
        Mid$(ReferteTxtNoFormat, 11, 2) = chk

        cmd.Parameters("pV413").Value = ReferteTxtNoFormat
        cmd.Parameters("pV033").Value = DocumentKey
        cmd.Execute

        Bezig = Bezig + 1

        If (Bezig Mod 100) = 0 Then
            SnelHelpPrint "Bezig met " & Bezig, BL_LOGGING
        End If

        If (Bezig Mod BATCH_SIZE) = 0 Then
            adntDB.CommitTrans
            bTrans = False
            rs.Close
            GoTo ReopenRecordset
        End If

NextRecord:
        rs.MoveNext
    Loop

    ' Batch klaar
    adntDB.CommitTrans
    bTrans = False
    rs.Close
    GoTo ReopenRecordset

Finished:
    MsgBox "Aantal documenten bijgewerkt: " & Bezig, vbInformation
    ConvertDocNumbersToOGM = True

CleanExit:
    Screen.MousePointer = vbNormal
    Exit Function

EH:
    If bTrans Then adntDB.RollbackTrans
    MsgBox "Fout: " & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume CleanExit
End Function


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



Sub GridSchoon(GelijkWelkRooster As MSFlexGrid)
Dim T As Integer

GelijkWelkRooster.Rows = 2
GelijkWelkRooster.Row = 1
For T = 0 To GelijkWelkRooster.Cols - 1
    GelijkWelkRooster.Col = T
    GelijkWelkRooster.text = ""
Next

End Sub

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


Sub BeginPlaySound(ByVal ResourceId As Integer)
Dim ret As Variant
    
On Local Error Resume Next
Err = 0

#If Win32 Then
    ' Important: The returned string is converted to Unicode
    SoundBuffer = StrConv(LoadResData(ResourceId, "VSOFT_SOUND"), vbUnicode)
#Else
    SoundBuffer = LoadResData(ResourceId, "VSOFT_SOUND")
#End If

If Err Then
    Exit Sub
Else
    ret = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
    ' Important: This function is neccessary for playing sound asynchronously
    DoEvents
End If

End Sub

Sub EndPlaySound()
    Dim ret As Variant
    ret = sndPlaySound(0&, 0&)
End Sub



Sub PrintAnywhere(Src As Object, Dest As Object)

On Error Resume Next
Dest.PaintPicture Src.Picture, Dest.Width / 2, Dest.Height / 2
If Dest Is Printer Then
    Printer.EndDoc
End If

End Sub




Private Function GetErrorTextFromResource(ErrorNum As Long) _
          As String
      Dim strMsg As String
      

      ' this function will retrieve an error description from a resource
      ' file (.RES).  The ErrorNum is the index of the string
      ' in the resource file.  Called by RaiseError


      On Error GoTo GetErrorTextFromResourceError
      

      ' get the string from a resource file
      GetErrorTextFromResource = LoadResString(ErrorNum)
      

      Exit Function
      

GetErrorTextFromResourceError:
      

      If Err.Number <> 0 Then
            GetErrorTextFromResource = "An unknown error has occurred!"
      End If
      

End Function


Public Sub RaiseError(ErrorNumber As Long, Source As String)
      Dim strErrorText As String


      'there are a number of methods for retrieving the error
      'message.  The following method uses a resource file to
      'retrieve strings indexed by the error number you are
      'raising.
      strErrorText = GetErrorTextFromResource(ErrorNumber)


      'raise an error back to the client
      Err.Raise vbObjectError + ErrorNumber, Source, strErrorText


End Sub




Function GetNextClassDebugID() As Long
    'class ID generator
    Static lClassDebugID As Long
    lClassDebugID = lClassDebugID + 1
    GetNextClassDebugID = lClassDebugID
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
If AantalOptiesMogelijk > 10 Then
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


