Attribute VB_Name = "thisMimEnvironmentClass"
Option Explicit
DefInt A-Z

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


Sub AutoUnloadBedrijf()
Dim T           As Integer
Dim LastUsed As String

Screen.MousePointer = vbHourglass
Ktrl = 100
Unload BYPERDAT
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
Mim.Caption = appTitleAndVersion
Screen.MousePointer = vbNormal

On Error Resume Next
adntDB.Close
Set adntDB = Nothing
Mim.Basis(3).Enabled = False
Close

LOCATION_COMPANYDATA = ""

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

Mim.SnelHelp.Panels(2).text = "---"
XisEuroWisBEF = False

Static ProducentKopij As String * 8

Err = 0
On Error Resume Next

'TODO: Menuopties toegang geven
For COUNT_TO = 1 To 5
    Mim.MenuTitel(COUNT_TO).Enabled = True
Next
For T = 1 To 4
    BasisB(T).Enabled = True
Next
Mim.Basis(11).Enabled = True

BYPERDAT.Hide
BYPERDAT.DatumVerwerking.Value = MIM_GLOBAL_DATE
Mim.DatumVerwerking.Value = MIM_GLOBAL_DATE
BYPERDAT.PeriodeBoekjaar.Clear

Ktrl = netVoorbereiden
If Ktrl = True Then
    MsgBox LOCATION_COMPANYDATA & "is voorbereid voor MarIntegraal 2026"
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
            BYPERDAT.Boekjaar.AddItem XX, 0
        End If
    Next
    Get FlTemp, 3, aa
    AktievePeriode = Val(aa)
Close FlTemp
BYPERDAT.Boekjaar.ListIndex = ACTIVE_BOOKYEAR
bstNaam(TABLE_COUNTERS) = "jr" & BYPERDAT.Boekjaar.text

FlTemp = FreeFile
Open LOCATION_COMPANYDATA + "DEF" + Format(ACTIVE_BOOKYEAR, "00") + ".OCT" For Random As FlTemp Len = 16
    For T = 1 To 99
        Get FlTemp, T, A
        If A = Space$(16) Then
            BYPERDAT.PeriodeBoekjaar.ListIndex = 0
            YY = BYPERDAT.PeriodeBoekjaar.text
            BOOKYEAR_FROMTO = Mid(YY, 7, 4) + Mid(YY, 4, 2) + Mid(YY, 1, 2) + Mid(XX, 20, 4) + Mid(XX, 17, 2) + Mid(XX, 14, 2)
            Exit For
        Else
            XX = Mid(A, 7, 2) + "/" + Mid(A, 5, 2) + "/" + Left(A, 4) + " - " + Right(A, 2) + "/" + Mid(A, 13, 2) + "/" + Mid(A, 9, 4)
            BYPERDAT.PeriodeBoekjaar.AddItem XX
        End If
    Next T
    If AktievePeriode - 1 > BYPERDAT.PeriodeBoekjaar.ListCount Then
        MsgBox "Het hoogste boekjaar wordt automatisch ingeladen.  Laatste bewerking gebeurde in een boekjaar met meer periodes dan nu mogelijk.  De eerste periode van het hoogste boekjaar wordt hierna automatisch geaktiveerd"
        AktievePeriode = 1
    End If
    BYPERDAT.PeriodeBoekjaar.ListIndex = AktievePeriode - 1
    AT = BYPERDAT.PeriodeBoekjaar.text
    PERIOD_FROMTO = Mid(AT, 7, 4) + Mid(AT, 4, 2) + Left(AT, 2) + Right(AT, 4) + Mid(AT, 17, 2) + Mid(AT, 14, 2)
Close FlTemp
BYPERDAT.Show

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
        
        TQString = OpenSchemeAsString("TABLE")
        
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

Dim ktrlfl As Integer
Dim ptel As Integer


'Checking for new fields needed
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

'Example ALTER TABLE Dokumenten ADD COLUMN v405 MEMO
'Example ALTER TABLE Dokumenten ADD COLUMN v406 MEMO
'bClose TABLE_INVOICES
'Err = 0
'On Error Resume Next
'Dim recDokument As ADODB.Recordset
'Set recDokument = New ADODB.Recordset
'Msg = "SELECT TOP 1 v405 FROM Dokumenten"
'recDokument.CursorLocation = adUseClient
'recDokument.Open Msg, adntDB, adOpenForwardOnly, adLockOptimistic
'If Err Then
'    bClose TABLE_INVOICES
'    Err = 0
'    On Error Resume Next
'    Msg = "ALTER TABLE Dokumenten ADD COLUMN v405 MEMO"
'    adntDB.Execute Msg
'    If Err = -2147217887 Then
'        'Stop
'    ElseIf Err = -2147217900 Then
'        'Stop
'        MsgBox "De database kan niet bijgewerkt worden wanneer in gebruik (netwerk?)" & vbCrLf & "Veiligheidshalve wordt het programma beëindigd.", vbInformation
'        End
'    ElseIf Err = 0 Then
'        Msg = "ALTER TABLE Dokumenten ADD COLUMN v406 MEMO"
'        adntDB.Execute Msg
'        If Err = 0 Then
'        Else
'            MsgBox Err.Description
'        End If
'    Else
'        MsgBox "Een foutmelding buiten controle:" & vbCrLf & Err.Description & vbCrLf & "Veiligheidshalve wordt het programma beëindigd.", vbInformation
'        End
'    End If
'End If

'Here example for fields 20250809 editing for klanten- en leveranciersbestand
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
Mim.SnelHelp.Panels(1).text = Mid(fmarBoxText("899", "2", Left(String99(READING, 20), 1)), 4)
If Mim.SnelHelp.Panels(1).text = "marIntegraal NT Light" Then
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
        If BYPERDAT.Boekjaar.ListIndex <> 0 Then MsgBox "Draai de Light versie enkel voor het hoogste boekjaar a.u.b. !"
        VsoftLog = 0
    Case "1"
        If BYPERDAT.Boekjaar.ListIndex <> 0 Then MsgBox "Draai de Light versie enkel voor het hoogste boekjaar a.u.b. !"
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
    Mim.SnelHelp.Panels(2).text = "EUR"
End If

'checking directories
'0. Coda
CodaDirCheck:
Dim codaDir As String
codaDir = LOCATION_COMPANYDATA + "coda"
On Local Error Resume Next
ChDir codaDir

If Err = 76 Then
    'Msg = "Vanaf versie 11.5.012 dient de datainhoudsopgave over submap 'coda' te beschikken." & vbCrLf & vbCrLf
    'Msg = Msg & "Deze map wordt hierna aangemaakt"
    'MsgBox Msg, vbInformation
    MkDir (LOCATION_COMPANYDATA + "coda")
    ChDir codaDir
    MkDir (LOCATION_COMPANYDATA + "coda\in")
    ChDir codaDir
    MkDir (LOCATION_COMPANYDATA + "coda\out")
    ChDir LOCATION_COMPANYDATA
Else
    ChDir LOCATION_COMPANYDATA
End If
On Local Error GoTo 0

'1. Peppol
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

'2. QR
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

'3. VPE templates check
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

'4. xlsx directory
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

'5. vat directory
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

Dim X As Integer

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

Dim resultDir As Boolean
Dim destpath As String

destpath = LOCATION_COMPANYDATA & "sjb"
Err = 0
On Local Error Resume Next
If Dir(destpath & "\sBrief.doc") = "" Then
    resultDir = CreatePath(destpath)
    If Not CopyFile(App.path, destpath, "sbrief.doc") Then
       MsgBox "Fout tijdens installatie briefsjabloon: sbrief.doc", vbExclamation
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

DocCheckForOGM:
If String99(READING, 100) = "" Then
    Msg = "Database extra veld OGM Documentnummers" & vbCrLf & vbCrLf
    Msg = Msg & "* Alle 'A0~' en 'A1~' aankoopdocumenten," & vbCrLf
    Msg = Msg & "* Alle 'V0~', 'V1~', 'B0~' en 'F0~' verkoopdocumenten," & vbCrLf
    Msg = Msg & "* Alle 'Q0~' verzekeringsdocumenten" & vbCrLf & vbCrLf
    Msg = Msg & "Er wijzigt niets aan de documentnummering" & vbCrLf
    Msg = Msg & "Functionaliteit uitsluitend voor betaalverrichtingen" & vbCrLf & vbCrLf
    Msg = Msg & "Hierna bijwerken, akkoord?"
    KtrlBox = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2)
    If KtrlBox = vbYes Then
        Dim returnResult As Boolean
        returnResult = ConvertDocNumbersToOGM
        If returnResult = False Then
            MsgBox "Er ging iets fout, probeer later opnieuw", vbExclamation
        Else
            Msg = "Omwerking met succes" & vbCrLf & vbCrLf
            Msg = Msg & "Duidt nu eerst in setup uw voorkeur aan:" & vbCrLf & vbCrLf
            Msg = Msg & "* ofwel OGM betaalreferte op basis document" & vbCrLf
            Msg = Msg & "* ofwel verkoopdocument als betaalreferte" & vbCrLf
            Msg = Msg & "  zoals voorheen"
            MsgBox Msg, vbInformation
        End If
    Else
        Msg = "Vergeet niet ten gepaste tijd deze functionaliteit beschikbaar te maken."
        MsgBox Msg, vbExclamation
    End If
End If

Screen.MousePointer = vbNormal
Exit Sub

ErrorOpvang:
MsgBox vBC(ktrlfl, ptel) + " " + Error
Resume Next

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

Function InitBestanden() As Boolean
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


Sub InitFirst()

    FULL_LINE = String$(128, 173)

    TABLEDEF_ONT(TABLE_VARIOUS) = "0000000.ONT"        '00
    TABLEDEF_ONT(TABLE_CUSTOMERS) = "0010000.ONT"      '01
    TABLEDEF_ONT(TABLE_SUPPLIERS) = "0020000.ONT"      '02
    TABLEDEF_ONT(TABLE_LEDGERACCOUNTS) = "0030000.ONT" '03
    TABLEDEF_ONT(TABLE_PRODUCTS) = "0040000.ONT"       '04
    TABLEDEF_ONT(TABLE_JOURNAL) = "0600000.ONT"        '05
    TABLEDEF_ONT(TABLE_INVOICES) = "0200000.ONT"       '06
    TABLEDEF_ONT(TABLE_CONTRACTS) = "0700000.ONT"      '07
    TABLEDEF_ONT(TABLE_DUMMY) = "90DUMMY.ONT"          '08
    TABLEDEF_ONT(TABLE_COUNTERS) = "00.ONT"            '09

    bstNaam(TABLE_VARIOUS) = "Allerlei"           '00
    bstNaam(TABLE_CUSTOMERS) = "Klanten"          '01
    bstNaam(TABLE_SUPPLIERS) = "Leveranciers"     '02
    bstNaam(TABLE_LEDGERACCOUNTS) = "Rekeningen"  '03
    bstNaam(TABLE_PRODUCTS) = "Produkten"         '04
    bstNaam(TABLE_JOURNAL) = "Journalen"          '05
    bstNaam(TABLE_INVOICES) = "dokumenten"        '06
    bstNaam(TABLE_CONTRACTS) = "Polissen"         '07
    bstNaam(TABLE_DUMMY) = "TmpBestand"           '08
    bstNaam(TABLE_COUNTERS) = "Tell"              '09

    DAYS_IN_MONTH(1) = 31
    DAYS_IN_MONTH(2) = 29
    DAYS_IN_MONTH(3) = 31
    DAYS_IN_MONTH(4) = 30
    DAYS_IN_MONTH(5) = 31
    DAYS_IN_MONTH(6) = 30
    DAYS_IN_MONTH(7) = 31
    DAYS_IN_MONTH(8) = 31
    DAYS_IN_MONTH(9) = 30
    DAYS_IN_MONTH(10) = 31
    DAYS_IN_MONTH(11) = 30
    DAYS_IN_MONTH(12) = 31

    MONTH_AS_TEXT(1) = "Januari  "
    MONTH_AS_TEXT(2) = "Februari "
    MONTH_AS_TEXT(3) = "Maart    "
    MONTH_AS_TEXT(4) = "April    "
    MONTH_AS_TEXT(5) = "Mei      "
    MONTH_AS_TEXT(6) = "Juni     "
    MONTH_AS_TEXT(7) = "Juli     "
    MONTH_AS_TEXT(8) = "Augustus "
    MONTH_AS_TEXT(9) = "September"
    MONTH_AS_TEXT(10) = "October  "
    MONTH_AS_TEXT(11) = "November "
    MONTH_AS_TEXT(12) = "December "

End Sub

Function SaveFormProperties(frmWindow As Form) As Boolean

    Err = 0: SaveFormProperties = True
    On Error Resume Next
    SaveSetting App.ProductName, frmWindow.Name, "Top", frmWindow.Top
    SaveSetting App.ProductName, frmWindow.Name, "Links", frmWindow.Left
    SaveSetting App.ProductName, frmWindow.Name, "Breedte", frmWindow.Width
    SaveSetting App.ProductName, frmWindow.Name, "Hoogte", frmWindow.Height
    If Err Then SaveFormProperties = False

End Function


Sub LoadFormProperties(frmWindow As Form)

    On Local Error GoTo ErrorATLoading

    frmWindow.Top = GetSetting(App.ProductName, frmWindow.Name, "Top")
    frmWindow.Left = GetSetting(App.ProductName, frmWindow.Name, "Links")
    frmWindow.Width = GetSetting(App.ProductName, frmWindow.Name, "Breedte")
    frmWindow.Height = GetSetting(App.ProductName, frmWindow.Name, "Hoogte")
    Exit Sub

ErrorATLoading:
    Exit Sub

End Sub


Sub BeWaarTekst(Onderdeel As String, SubDeel As String, Element As String)

    SaveSetting App.ProductName, Onderdeel, SubDeel, Element

End Sub

Function LaadTekst(Onderdeel As String, SubDeel As String) As Variant

    On Local Error Resume Next
    If InStr(Onderdeel, ";") Then
        LaadTekst = GetSetting(Left(Onderdeel, InStr(Onderdeel, ";") - 1), Mid(Onderdeel, InStr(Onderdeel, ";") + 1), SubDeel)
    Else
        LaadTekst = GetSetting(App.ProductName, Onderdeel, SubDeel)
    End If

End Function

