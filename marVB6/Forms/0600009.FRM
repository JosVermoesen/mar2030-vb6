VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ProefEnSaldi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proef- en SaldiBalans"
   ClientHeight    =   3060
   ClientLeft      =   1260
   ClientTop       =   735
   ClientWidth     =   5670
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
   ScaleHeight     =   3060
   ScaleWidth      =   5670
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cbXMLExport 
      Caption         =   "XML Export Boekjaar"
      Height          =   735
      Left            =   4320
      TabIndex        =   16
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox cbLijstPrinterHier 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   15
      ToolTipText     =   "Dokumenten printer"
      Top             =   2640
      Width           =   3795
   End
   Begin VB.CommandButton cbTogglePrinter 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2640
      Width           =   315
   End
   Begin VB.CheckBox chkAfdrukLiggend 
      Alignment       =   1  'Right Justify
      Caption         =   "Liggende Printerafdruk"
      Height          =   195
      Left            =   1800
      TabIndex        =   13
      Top             =   2400
      Width           =   2430
   End
   Begin VB.TextBox txtTekstLijn 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   3
      Left            =   1410
      TabIndex        =   5
      Text            =   "7999999"
      Top             =   990
      Width           =   1200
   End
   Begin VB.TextBox txtTekstLijn 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   2
      Left            =   90
      TabIndex        =   4
      Text            =   "1"
      Top             =   990
      Width           =   1200
   End
   Begin VB.CheckBox chkDetailJournaal 
      Alignment       =   1  'Right Justify
      Caption         =   "Detail Journaal"
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CheckBox chkAfdrukInVenster 
      Caption         =   "Afdruk in venster"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Value           =   1  'Checked
      Width           =   2400
   End
   Begin VB.TextBox TxtVolgNummer 
      Height          =   315
      Left            =   2190
      TabIndex        =   11
      Top             =   1650
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   0
      Left            =   60
      TabIndex        =   7
      Top             =   1350
      Width           =   4185
   End
   Begin VB.CommandButton cmdSluiten 
      Cancel          =   -1  'True
      Caption         =   "&Sluiten"
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
      Left            =   4320
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1200
   End
   Begin VB.TextBox txtTekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   1
      Left            =   1380
      TabIndex        =   1
      Top             =   210
      Width           =   1230
   End
   Begin VB.CommandButton Drukken 
      Caption         =   "Af&drukken"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfSQL 
      Height          =   2175
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3836
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblRekeningVanTot 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Rekening Van-Tot"
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   3
      Top             =   690
      Width           =   2520
   End
   Begin VB.Label LblVolgNummer 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Manuele VolgNummer"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblPeriodeVanTot 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Periode Van-Tot"
      Height          =   255
      Index           =   0
      Left            =   2790
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Datu&m Drukken"
      Height          =   465
      Index           =   1
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   1185
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "ProefEnSaldi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim datPrimaryRS As ADODB.Recordset

Dim psTekst(5) As String        'Koptekstinfo
Dim ErrorTekst As String

Dim VeldTXT(20) As String

Dim Lijn As Long

Dim SubTotaalD As Currency
Dim SubTotaalC As Currency

Dim CumTotaalD As Currency
Dim CumTotaalC As Currency

Dim TotaalD As Currency
Dim TotaalC As Currency

Dim LFontSize(20)           As Single
Dim LAantalL(20)            As Integer
Dim FontDefChanged          As Integer

Dim PlGrensVan                  As String * 8
Dim PlGrensTot                  As String * 8

Dim SaldiKontrole               As Boolean
Dim KontroleLijst               As String
Dim aa                          As String

'Msg = "ALTER TABLE Journalen ADD COLUMN v248 CURRENCY;"
'Msg = "ALTER TABLE Journalen DROP COLUMN v248"
'Msg = "UPDATE Journalen SET v248=Str$(Cdbl(Format(Val(v068)/40.3399," + Chr$(34) + "0.0000" + Chr$(34) + "))) WHERE v066 <= '" + Right(BOOKYEAR_FROMTO, 8) + "';"
'adntDB.Execute Msg, AantalRecordsHier
'If Err Then
'    MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
'End If
    

Sub DetailJournaalDrukken()
Dim SubTitelTekst As String
Dim BeginSleutel As String * 8
Dim EindSleutel As String * 8
Dim DCBedrag As Long

Dim DCKontrole As Currency
Dim DCDatum As String * 8

'On Local Error GoTo PrtHandler2

ErrorTekst = ""
Lijn = 0
KontroleLijst = ""
BeginSleutel = PlGrensVan
EindSleutel = PlGrensTot

psTekst(2) = "Algemeen Journaal (Systeem OT) " + Mid(Mim.Caption, InStr(Mim.Caption, "["))
psTekst(0) = txtTekstLijn(1).text
psTekst(3) = "Boekjaar aanvang : " + Left(BOOKYEAR_FROMTO, 4) + ", " + txtTekstLijn(0).text
    InitVelden

bFirst TABLE_JOURNAL, 4
bGetOrGreater TABLE_JOURNAL, 4, BeginSleutel
If Ktrl Then
    Beep
    Exit Sub
End If
    ProefEnSaldi.Enabled = False
    PAGE_COUNTER = Val(TxtVolgNummer.text)
    If chkAfdrukInVenster = 0 Then
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
    TotaalD = 0
    TotaalC = 0
    
Screen.MousePointer = vbHourglass
Do
    GoSub PrintInfo
    bNext TABLE_JOURNAL
    XDoEvents = DoEvents
    If Ktrl Or KEY_BUF(TABLE_JOURNAL) > EindSleutel Then
        Exit Do
    End If
Loop
    EindTotaal
    If chkAfdrukInVenster Then
    Else
        Printer.NewPage
        Printer.FontSize = Printer.FontSize
        Printer.Print " "
        Printer.FontSize = Printer.FontSize
    End If
    If chkAfdrukInVenster Then
    Else
        Printer.EndDoc
    End If
    Screen.MousePointer = vbNormal
    ProefEnSaldi.Enabled = True
    cmdSluiten_Click
Exit Sub

PrintInfo:
RecordToVeld TABLE_JOURNAL
If vBibTekst(TABLE_JOURNAL, "#v066 #") >= PlGrensVan And vBibTekst(TABLE_JOURNAL, "#v066 #") <= PlGrensTot Then
    Lijn = Lijn + 1
    If DCDatum <> vBibTekst(TABLE_JOURNAL, "#v066 #") Then
        DCDatum = vBibTekst(TABLE_JOURNAL, "#v066 #")
        If ErrorTekst = "j" Then
        Else
            If TotaalD - TotaalC <> 0 Then
                ErrorTekst = "j"
                Msg = "DC ongelijkheid vanaf laatste dag vóór : " & ErrorTekst & vBibTekst(TABLE_JOURNAL, "#v066 #")
                MsgBox Msg
            End If
        End If
    End If
        
    bGet TABLE_LEDGERACCOUNTS, 0, vSet(vBibTekst(TABLE_JOURNAL, "#v019 #"), 7)
    If Ktrl Then
        VeldTXT(3) = "-"
    Else
        RecordToVeld TABLE_LEDGERACCOUNTS
        VeldTXT(3) = vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #")
        SnelHelpPrint vBibTekst(TABLE_LEDGERACCOUNTS, "#v019 #") + " " + vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #") + " " + VeldTXT(3), BL_LOGGING
    End If
    VeldTXT(0) = Format(Lijn, "00000")
    VeldTXT(1) = DATE_TEXT(vBibTekst(TABLE_JOURNAL, "#v066 #"))
    VeldTXT(2) = vBibTekst(TABLE_JOURNAL, "#v019 #")
    VeldTXT(4) = vBibTekst(TABLE_JOURNAL, "#v067 #")
    
    Select Case rsMAR(TABLE_JOURNAL)("dece068").Value
        Case Is < 0
            VeldTXT(5) = ""
            VeldTXT(6) = Dec$(Abs(rsMAR(TABLE_JOURNAL)("dece068").Value), MASK_2002)
            TotaalC = TotaalC + Val(VeldTXT(6))
        Case Else
            VeldTXT(5) = Dec$(rsMAR(TABLE_JOURNAL)("dece068").Value, MASK_2002)
            VeldTXT(6) = ""
            TotaalD = TotaalD + Val(VeldTXT(5))
    End Select
    
    VeldTXT(7) = vBibTekst(TABLE_JOURNAL, "#v033 #")
    PrintVelden
End If
Return

End Sub

Sub ProefEnSaldiDrukken()
Dim SubTitelTekst As String
Dim BeginSleutel As String * 15
Dim EindSleutel As String * 15
Dim VorigeSleutel As String * 15
Dim DCBedrag As Currency
Dim recVeldje As String
Dim rkVeldje As String

'On Local Error GoTo PrtHandler2

Lijn = 0
VorigeSleutel = ""
KontroleLijst = ""
BeginSleutel = vSet(txtTekstLijn(2), 7) + PlGrensVan
EindSleutel = vSet(txtTekstLijn(3), 7) + PlGrensTot

psTekst(2) = "Proef- en Saldibalans " + Mid(Mim.Caption, InStr(Mim.Caption, "["))
psTekst(0) = txtTekstLijn(1).text
psTekst(3) = "Boekjaar aanvang : " + Left(BOOKYEAR_FROMTO, 4) + ", " + txtTekstLijn(0).text
    InitVelden

bFirst TABLE_JOURNAL, 0
bGetOrGreater TABLE_JOURNAL, 0, BeginSleutel
If Ktrl Then
        Beep
        Exit Sub
End If
    ProefEnSaldi.Enabled = False
    PAGE_COUNTER = 0
    If chkAfdrukInVenster = 0 Then
        Set Printer = Printers(LijstPrinterNr)
        On Error Resume Next
        Printer.PaperBin = LaadTekst(App.Title, "LIJSTPRINTER")
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
            Printer.FontBold = True
        End If
    End If
    PrintTitel

Screen.MousePointer = vbHourglass
If Right(KEY_BUF(TABLE_JOURNAL), 8) >= PlGrensVan And Right(KEY_BUF(TABLE_JOURNAL), 8) <= PlGrensTot Then
    GoSub PrintInfo
End If

Do
    bNext TABLE_JOURNAL
    XDoEvents = DoEvents
    If Ktrl Then
        Exit Do
    ElseIf Right(KEY_BUF(TABLE_JOURNAL), 8) >= PlGrensVan And Right(KEY_BUF(TABLE_JOURNAL), 8) <= PlGrensTot Then
        If VorigeSleutel = Space$(15) Then
            GoSub PrintInfo
        ElseIf Left(VorigeSleutel, 7) = Left(KEY_BUF(TABLE_JOURNAL), 7) Then
            If Mid(VorigeSleutel, 12, 2) <> Mid(KEY_BUF(TABLE_JOURNAL), 12, 2) Then
                RekeningTotaal
                GoSub PrintInfo
            Else
                GoSub PrintInfo
            End If
        Else
            RekeningTotaal
            GoSub PrintInfo
        End If
    End If
Loop
    GoSub RekSaldoKTRL
    RekeningTotaal
    EindTotaal
    If chkAfdrukInVenster Then
    Else
        Printer.NewPage
        Printer.FontSize = Printer.FontSize
        Printer.Print " "
        Printer.FontSize = Printer.FontSize
    End If
    If KontroleLijst <> "" Then
        Printer.CurrentY = 400
        Printer.Print KontroleLijst
        Printer.EndDoc
    ElseIf chkAfdrukInVenster Then
    Else
        Printer.EndDoc
    End If
    Screen.MousePointer = vbNormal
    ProefEnSaldi.Enabled = True
    cmdSluiten_Click
Exit Sub

PrintInfo:
RecordToVeld TABLE_JOURNAL
If Left(VorigeSleutel, 7) <> Left(KEY_BUF(TABLE_JOURNAL), 7) Then
    If VorigeSleutel = Space(15) Then
    ElseIf SaldiKontrole = True Then
        GoSub RekSaldoKTRL
    End If
    
    CumTotaalD = 0
    CumTotaalC = 0
    bGet TABLE_LEDGERACCOUNTS, 0, vSet(vBibTekst(TABLE_JOURNAL, "#v019 #"), 7)
    VeldTXT(0) = vBibTekst(TABLE_JOURNAL, "#v019 #")
    If Ktrl Then
        VeldTXT(1) = "Rekening reeds vernietigd..."
        VeldTXT(2) = ""
    Else
        RecordToVeld TABLE_LEDGERACCOUNTS
        VeldTXT(1) = vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #")
        
        If bhEuro Then
            rkVeldje = "#e" + Format(22 + BYPERDAT.Boekjaar.ListIndex, "000") + " #"
        Else
            rkVeldje = "#v" + Format(22 + BYPERDAT.Boekjaar.ListIndex, "000") + " #"
        End If
        
        VeldTXT(2) = Str$(Val(vBibTekst(TABLE_LEDGERACCOUNTS, rkVeldje)))
        Select Case Val(VeldTXT(2))
            Case Is < 0
                VeldTXT(2) = "CS:" + Dec$(Abs(Val(VeldTXT(2))), MASK_2002)
            Case Else
                VeldTXT(2) = "DS:" + Dec$(Val(VeldTXT(2)), MASK_2002)
        End Select
    End If
End If
VorigeSleutel = KEY_BUF(TABLE_JOURNAL)
VeldTXT(3) = MONTH_AS_TEXT(Val(Mid(KEY_BUF(TABLE_JOURNAL), 12, 2)))
SnelHelpPrint vBibTekst(TABLE_LEDGERACCOUNTS, "#v019 #") + " " + vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #") + " " + Trim$(VeldTXT(3)) + " " + Mid(KEY_BUF(TABLE_JOURNAL), 8, 4), BL_LOGGING

If IsNull(rsMAR(TABLE_JOURNAL)("dece068").Value) Then
    DCBedrag = 0
Else
    DCBedrag = rsMAR(TABLE_JOURNAL)("dece068").Value
End If
Select Case DCBedrag
    Case Is < 0
        SubTotaalC = SubTotaalC + DCBedrag
        CumTotaalC = CumTotaalC + DCBedrag
        TotaalC = TotaalC + DCBedrag
    Case Else
        SubTotaalD = SubTotaalD + DCBedrag
        CumTotaalD = CumTotaalD + DCBedrag
        TotaalD = TotaalD + DCBedrag
End Select
Return

RekSaldoKTRL:
If bhEuro Then
    rkVeldje = "#e" + Format(22 + BYPERDAT.Boekjaar.ListIndex, "000") + " #"
Else
    rkVeldje = "#v" + Format(22 + BYPERDAT.Boekjaar.ListIndex, "000") + " #"
End If
If Abs(CumTotaalD + CumTotaalC) - Abs(Val(vBibTekst(TABLE_LEDGERACCOUNTS, rkVeldje))) = 0 Then
Else
    Msg = "KontroleStop bijwerking rekeningsaldo !!" + vbCrLf + vbCrLf
    Msg = Msg + "Rekening : " + KEY_BUF(TABLE_LEDGERACCOUNTS) + vbCrLf
    If bhEuro Then
        Msg = Msg + Str$(CumTotaalD + CumTotaalC) + " <> " + Str$(Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#e" + Format(22 + BYPERDAT.Boekjaar.ListIndex, "000") + " #")))
    Else
        Msg = Msg + Str$(CumTotaalD + CumTotaalC) + " <> " + Str$(Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#v" + Format(22 + BYPERDAT.Boekjaar.ListIndex, "000") + " #")))
    End If
    MsgBox Msg & vbCrLf & vbCrLf & "Na deze P&S, nogmaals opnieuw P&S a.u.b. samenstellen", vbExclamation
                
    KontroleLijst = KontroleLijst + vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #") + " " + Dec$(CumTotaalD + CumTotaalC, "#########.00") + " " + Dec$(Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#v" + Format(22 + BYPERDAT.Boekjaar.ListIndex, "000") + " #")), "#########.00") + vbCrLf
    If bhEuro Then
        vBib TABLE_LEDGERACCOUNTS, Str$(CumTotaalD + CumTotaalC), "e" + Format(22 + BYPERDAT.Boekjaar.ListIndex, "000")
        rsMAR(TABLE_LEDGERACCOUNTS)("dece" + Format(22 + BYPERDAT.Boekjaar.ListIndex, "000")) = CumTotaalD + CumTotaalC
    Else
        vBib TABLE_LEDGERACCOUNTS, Str$(CumTotaalD + CumTotaalC), "v" + Format(22 + BYPERDAT.Boekjaar.ListIndex, "000")
    End If
    bUpdate TABLE_LEDGERACCOUNTS, 0
End If
Return

PrtHandler2:
MsgBox "Kontroleer de printer."
Resume

End Sub


Private Sub cbDocumentPrinterHier_Change()

End Sub


Private Sub cbLijstPrinterHier_Click()

    LijstPrinterNr = cbLijstPrinterHier.ListIndex
        
End Sub


Private Sub cbTogglePrinter_Click()
    
    Me.cbLijstPrinterHier.Enabled = Not Me.cbLijstPrinterHier.Enabled
    
End Sub

Private Sub cbXMLExport_Click()

Set datPrimaryRS = New ADODB.Recordset
datPrimaryRS.CursorLocation = adUseClient

Dim txtSQL As String
Dim txtSQLLeft As String
Dim txtSQLMid As String
Dim txtSQLRight As String

Dim txtStartBJ As String
Dim txtEndBJ As String

MsgBox "Journalen ophalen.  Even geduld a.u.b.", vbInformation
txtStartBJ = Mid(BOOKYEAR_FROMTO, 1, 8)
txtEndBJ = Mid(BOOKYEAR_FROMTO, 9, 8)

txtSQLLeft = "SELECT DISTINCT Rekeningen.v019 AS [Nummer], " & _
    "Rekeningen.v020 AS [Naam Rekening], "

txtSQLMid = "Mid(Journalen.v066,7,2)+""/"" + Mid(Journalen.v066,5,2) + ""/"" + Mid(Journalen.v066,1,4) AS [Datum], "

txtSQLRight = "Journalen.v067 AS [Omschrijving], Journalen.v033 AS [Document], Journalen.v038 AS [FinDoc], Journalen.dece068 AS [Bedrag], Journalen.v069 AS [Tegenrekening] FROM Journalen, Rekeningen WHERE Trim(Journalen.v019) = Trim(Rekeningen.v019) AND Trim(Journalen.v066) >= '" & txtStartBJ & "' AND Trim(Journalen.v066) <= '" & txtEndBJ & "' ORDER BY Rekeningen.v019"
txtSQL = txtSQLLeft & txtSQLMid & txtSQLRight

Screen.MousePointer = vbHourglass
On Error Resume Next
msfSQL.Recordset = Nothing

datPrimaryRS.Close
On Error Resume Next
datPrimaryRS.Open txtSQL, adntDB
If Err Then
    MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
    msfSQL.Refresh
Else
    Set msfSQL.DataSource = datPrimaryRS
End If

Dim lblRecordCount As Long
lblRecordCount = datPrimaryRS.RecordCount
MsgBox "Er zijn " & lblRecordCount & " journaallijnen", vbInformation

Screen.MousePointer = vbNormal

On Error GoTo CancelError
    Mim.Teken.fileName = ""
    Mim.Teken.CancelError = True
    Mim.Teken.Filter = "Alle bestanden (*.xml)|*.xml"
    Mim.Teken.ShowSave
    Dim BestandHier As String
    BestandHier = Mim.Teken.fileName
    If Not Dir(BestandHier) = "" Then
        Kill BestandHier
    End If
    datPrimaryRS.Save Mim.Teken.fileName, adPersistXML
    MsgBox "Klaar!" & vbCrLf & "Start uw spreadsheet programmma (bvb Excel) en open het xml bestand (of sleep het in een leeg werkblad)." & vbCrLf & vbCrLf & "Aanvaardt het voorstel van het spreadsheet programma om automatisch een schema aan te maken.", vbInformation
    
    
    Unload Me
    Exit Sub

CancelError:
Screen.MousePointer = vbNormal

End Sub

Private Sub chkAfdrukLiggend_Click()

Set Printer = Printers(LijstPrinterNr)
On Error Resume Next
Printer.PaperBin = LaadTekst(App.Title, "LIJSTPRINTER")
If chkAfdrukLiggend.Value = 1 Then
    Printer.Orientation = vbPRORLandscape
Else
    Printer.Orientation = vbPRORPortrait
End If

End Sub

Private Sub chkDetailJournaal_Click()

If chkDetailJournaal.Value Then
    LblVolgNummer.Visible = True
    TxtVolgNummer.Visible = True
    txtTekstLijn(0).text = BYPERDAT!PeriodeBoekjaar.text
Else
    LblVolgNummer.Visible = False
    TxtVolgNummer.Visible = False
    txtTekstLijn(0).text = Left(BYPERDAT!PeriodeBoekjaar.List(0), 10) + " - " + Right(BYPERDAT!PeriodeBoekjaar.List(BYPERDAT.PeriodeBoekjaar.ListCount - 1), 10)
End If
If chkDetailJournaal.Value Then
    BeginPlaySound 9
    Msg = "Sinds 1993 zijn er aanbevelingen om na elk boek "
    Msg = Msg & "een centralisatieblad af te drukken en te bewaren "
    Msg = Msg & "in een grootboek.  Dit heeft de voorkeur boven het "
    Msg = Msg & "omslachtig en papierverslindend systeem van algemeen "
    Msg = Msg & "journaal afdrukken welke U zopas hebt gekozen..." & vbCr & vbCr
    Msg = Msg & "U KAN OOK voor detailjournaal kiezen om kontroles uit te voeren.  "
    Msg = Msg & "Laat dan best in venster afdrukken (bvb. vaststelling "
    Msg = Msg & "van Debet <> Credit verschil de juiste dag van feiten "
    Msg = Msg & "op te zoeken"
    MsgBox Msg
End If

End Sub

Private Sub cmdSluiten_Click()

Unload Me

End Sub


Private Sub Drukken_Click()

Dim lTeller As Long

If PlGrensVan = Left(BOOKYEAR_FROMTO, 8) And PlGrensTot = Right(BOOKYEAR_FROMTO, 8) Then
    SaldiKontrole = True
Else
    SaldiKontrole = False
End If

If chkDetailJournaal.Value Then
    DetailJournaalDrukken
Else
    ProefEnSaldiDrukken
End If

End Sub

Private Sub EindTotaal()

'On Local Error GoTo PrtHandler3

For T = 0 To 7:
    VeldTXT(T) = ""
Next
VeldTXT(1) = "Totalen :"
If chkDetailJournaal Then
    VeldTXT(5) = Dec$((TotaalD), MASK_2002)
    VeldTXT(6) = Dec$((Abs(TotaalC)), MASK_2002)
Else
    VeldTXT(4) = Dec$((TotaalD), MASK_2002)
    VeldTXT(5) = Dec$((Abs(TotaalC)), MASK_2002)
End If
T = 0
If chkAfdrukInVenster Then
Else
    Printer.Print vbCrLf; FULL_LINE
End If

aa = ""
Do While REPORT_TAB(T) <> 0
    If chkAfdrukInVenster Then
        aa = aa + VeldTXT(T) + vbTab
    Else
        Printer.Print Tab(REPORT_TAB(T));
        Printer.Print VeldTXT(T);
    End If
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        If chkAfdrukInVenster = 0 Then
            Printer.Print vbCrLf;
        End If
    End If
    T = T + 1
Loop
If chkAfdrukInVenster Then
    Xlog.X.AddItem aa, Xlog.X.Rows - 1
End If

If Round(TotaalD, 2) + Round(TotaalC, 2) Then
    KontroleLijst = KontroleLijst + vbCrLf + vbCrLf + "Katastrofale fout : Algemene cumul Debet <> cumul credit.  Kontakteer ons 053/21.59.25 !" + vbCrLf + "Indien U niet beschikt over veiligheidskopij dient recuperatie van bestanden door ons te gebeuren."
End If

Lijn = 0
SubTotaalD = 0
SubTotaalC = 0
TotaalD = 0
TotaalC = 0

If chkAfdrukInVenster Then
    Xlog.X.Row = 1
    Xlog.X.Col = 0
    If chkDetailJournaal Then
        With Xlog.X
            .ColWidth(0) = 705
            .ColWidth(1) = 990
            .ColWidth(2) = 750
            .ColWidth(3) = 2865
            .ColWidth(4) = 1830
            .ColWidth(5) = 915
            .ColWidth(6) = 915
            .ColWidth(7) = 1140
        End With
        Xlog.WindowState = 2
    Else
        With Xlog.X
            .ColWidth(0) = 700
            .ColWidth(1) = 2955
            .ColWidth(2) = 1250
            .ColAlignment(2) = flexAlignRightTop
            .ColWidth(3) = 855
            .ColWidth(4) = 1200
            .ColAlignment(4) = flexAlignRightTop
            .ColWidth(5) = 1200
            .ColAlignment(5) = flexAlignRightTop
            .ColWidth(6) = 1200
            .ColAlignment(6) = flexAlignRightTop
            .ColWidth(7) = 1200
            .ColAlignment(7) = flexAlignRightTop
        End With
    End If
    Screen.MousePointer = vbNormal
    '---------------------------------------------------
    'TODO everywhere when choosen show print in a window
    Xlog.WindowState = vbNormal
    Xlog.statusForm = ""
    Xlog.Width = 11160
    '---------------------------------------------------
    Xlog.WijzigenLijn.Visible = False
    Xlog.Afsluiten.Enabled = False
    Xlog.Afsluiten.TabStop = False
    Xlog.cbAfbeelding.Visible = False
    XLogKey = ""
    Xlog.SSTab1.TabVisible(1) = False
    Xlog.Show 1
    Unload Xlog
End If
Exit Sub

PrtHandler3:
MsgBox "Kontroleer de printer."
Resume

End Sub

Private Sub Form_Load()

Dim PrinterTekst As String

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If

cbLijstPrinterHier.Clear
If Printers.Count = 0 Then MsgBox "Installeer eerst een printerdriver a.u.b. (ook al beschikt U nog over geen printer.  Bij blijvende problemen in netwerk gelieve op te starten met het command '/P=FALSE' en uw netwerkbeheerder te informeren)", vbExclamation: End
For COUNT_TO = 0 To Printers.Count - 1
    PrinterTekst = Printers(COUNT_TO).Port + " " + Printers(COUNT_TO).DeviceName
    Me.cbLijstPrinterHier.AddItem PrinterTekst
Next
cbLijstPrinterHier.ListIndex = LijstPrinterNr

Dim TempB_MODUS As Integer

SubTotaalD = 0
SubTotaalC = 0
TotaalD = 0
TotaalC = 0

txtTekstLijn(0).text = Left(BYPERDAT.PeriodeBoekjaar.List(0), 10) + " - " + Right(BYPERDAT.PeriodeBoekjaar.List(BYPERDAT.PeriodeBoekjaar.ListCount - 1), 10)
PlGrensVan = Mid(txtTekstLijn(0).text, 7, 4) + Mid(txtTekstLijn(0).text, 4, 2) + Mid(txtTekstLijn(0).text, 1, 2)
PlGrensTot = Mid(txtTekstLijn(0).text, 20, 4) + Mid(txtTekstLijn(0).text, 17, 2) + Mid(txtTekstLijn(0).text, 14, 2)
txtTekstLijn(1).text = MIM_GLOBAL_DATE
Caption = Caption + " " + Left(BOOKYEAR_FROMTO, 4)

Set Printer = Printers(LijstPrinterNr)
On Error Resume Next
Printer.PaperBin = LaadTekst(App.Title, "LIJSTPRINTER")

End Sub


Private Sub InitVelden()
Dim T As Integer
Dim VolgTab As Integer

If chkDetailJournaal Then
REPORT_FIELD(0) = "Nr.Ln"
        REPORT_TAB(0) = 2

REPORT_FIELD(1) = "Datum"
        REPORT_TAB(1) = 8

REPORT_FIELD(2) = "RNummer"
        REPORT_TAB(2) = 19

REPORT_FIELD(3) = "Naam Rekening"
        REPORT_TAB(3) = 27

REPORT_FIELD(4) = "Boekingsomschrijving"
        REPORT_TAB(4) = 68

REPORT_FIELD(5) = "    Debet"
        REPORT_TAB(5) = 104

REPORT_FIELD(6) = "   Credit"
        REPORT_TAB(6) = 115

REPORT_FIELD(7) = "dokument"
        REPORT_TAB(7) = 126
Else
REPORT_FIELD(0) = "Nummer"
        REPORT_TAB(0) = 2

REPORT_FIELD(1) = "Omschrijving Rekening"
        REPORT_TAB(1) = 10

REPORT_FIELD(2) = "Saldo"
        REPORT_TAB(2) = 51

REPORT_FIELD(3) = "Maand"
        REPORT_TAB(3) = 65

REPORT_FIELD(4) = "     Debet"
        REPORT_TAB(4) = 76

REPORT_FIELD(5) = "    Credit"
        REPORT_TAB(5) = 87

REPORT_FIELD(6) = "Mnd Saldo"
        REPORT_TAB(6) = 98

REPORT_FIELD(7) = "D/C Cumul"
        REPORT_TAB(7) = 111
End If
REPORT_TAB(8) = 0

If chkAfdrukInVenster Then
    Me.Hide
    Unload Xlog
    Xlog.Hide
    Xlog.Caption = "Proef- en Saldibalans"
    Xlog.X.Cols = 8
    Xlog.X.Row = 0
    For T = 0 To 7
        Xlog.X.Col = T
        Xlog.X.text = REPORT_FIELD(T)
    Next
    Me.Show
End If


End Sub

Private Sub PrintTitel()
Dim T As Integer

'On Local Error GoTo PrtHandler1

If chkAfdrukInVenster Then Exit Sub

  If usrLicentieInfo <> "" Then
        Printer.CurrentX = 50
        Printer.CurrentY = 50
        Printer.Print usrLicentieInfo;
    End If
PAGE_COUNTER = PAGE_COUNTER + 1
Printer.CurrentY = 400
Printer.Print Tab(1); psTekst(2); Tab(108); "Pagina : " + Dec$((PAGE_COUNTER), "##########");
Printer.Print Tab(108); "Datum  : "; psTekst(0); vbCrLf$; vbCrLf$;
Printer.Print Tab(1); UCase$(psTekst(3));

Printer.Print vbCrLf; FULL_LINE
Do While REPORT_TAB(T) <> 0
    Printer.Print Tab(REPORT_TAB(T));
    Printer.Print REPORT_FIELD(T);
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        Printer.Print vbCrLf;
    End If
    T = T + 1
Loop

Printer.Print FULL_LINE; vbCrLf; vbCrLf;
Exit Sub

PrtHandler1:
MsgBox "Kontroleer de printer."
Resume

End Sub

Private Sub PrintVelden()
Dim T As Integer
Dim VeldTekst As String

On Local Error GoTo PrtHandler4

aa = ""
Do While REPORT_TAB(T) <> 0
    If chkAfdrukInVenster Then
        aa = aa + VeldTXT(T) + vbTab
    Else
        Printer.Print Tab(REPORT_TAB(T));
        Printer.Print VeldTXT(T);
    End If
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        If chkAfdrukInVenster Then
        Else
            Printer.Print vbCrLf;
        End If
    End If
    T = T + 1
Loop

If chkAfdrukInVenster Then Xlog.X.AddItem aa, Xlog.X.Rows - 1: Exit Sub

If Printer.CurrentY >= Printer.Height - 1200 Then
    Printer.NewPage
    Printer.FontSize = Printer.FontSize
    Printer.Print " "
    Printer.FontSize = Printer.FontSize
    PrintTitel
End If
Exit Sub

PrtHandler4:
MsgBox "Kontroleer de printer."
Resume

End Sub

Private Sub RekeningTotaal()

'On Local Error GoTo PrtHandler5
Dim TotaalDC As Double
Dim CumTotaalDC As Double

VeldTXT(4) = Dec$((SubTotaalD), MASK_2002)
VeldTXT(5) = Dec$((Abs(SubTotaalC)), MASK_2002)
'VeldTXT(6) = Dec$((CumTotaalD), MASK_SY(0))
TotaalDC = SubTotaalD + SubTotaalC
If TotaalDC < 0 Then
    VeldTXT(6) = "C:" + Dec$(Abs(TotaalDC), MASK_2002)
ElseIf TotaalDC > 0 Then
    VeldTXT(6) = "D:" + Dec$(TotaalDC, MASK_2002)
Else
    VeldTXT(6) = "  " + Dec$(TotaalDC, MASK_2002)
End If

CumTotaalDC = CumTotaalD + CumTotaalC
If CumTotaalDC < 0 Then
    VeldTXT(7) = "C:" + Dec$(Abs(CumTotaalDC), MASK_2002)
ElseIf CumTotaalDC > 0 Then
    VeldTXT(7) = "D:" + Dec$(CumTotaalDC, MASK_2002)
Else
    VeldTXT(7) = "  " + Dec$(CumTotaalDC, MASK_2002)
End If

T = 0
aa = ""
Do While REPORT_TAB(T) <> 0
    If chkAfdrukInVenster Then
        aa = aa + VeldTXT(T) + vbTab
    Else
        Printer.Print Tab(REPORT_TAB(T));
        Printer.Print VeldTXT(T);
        If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
            Printer.Print vbCrLf;
        End If
    End If
    T = T + 1
Loop

SubTotaalD = 0
SubTotaalC = 0

If chkAfdrukInVenster Then
    Xlog.X.AddItem aa, Xlog.X.Rows - 1
ElseIf Printer.CurrentY >= Printer.Height - 1200 Then
    Printer.NewPage
    Printer.FontSize = Printer.FontSize
    Printer.Print " "
    Printer.FontSize = Printer.FontSize
    PrintTitel
End If

For T = 0 To 7:
        VeldTXT(T) = ""
Next
Exit Sub

PrtHandler5:
MsgBox "Kontroleer de printer."
Resume

End Sub



Private Sub txtTekstLijn_LostFocus(Index As Integer)

Select Case Index
    Case 0
        If DATE_INVALID(Right(txtTekstLijn(0).text, 10)) Then
            MsgBox "Respecteer : " + vbCrLf + vbCrLf + "DD/MM/EEJJ - DD/MM/EEJJ a.u.b. !"
            txtTekstLijn(0).text = BYPERDAT.PeriodeBoekjaar.text
            txtTekstLijn(0).SetFocus
            Exit Sub
        ElseIf Len(txtTekstLijn(0).text) <> 23 Then
            MsgBox "Respecteer : " + vbCrLf + vbCrLf + "DD/MM/EEJJ - DD/MM/EEJJ a.u.b. !"
            txtTekstLijn(0).text = BYPERDAT.PeriodeBoekjaar.text
            txtTekstLijn(0).SetFocus
            Exit Sub
        Else
            PlGrensVan = Mid(txtTekstLijn(0).text, 7, 4) + Mid(txtTekstLijn(0).text, 4, 2) + Mid(txtTekstLijn(0).text, 1, 2)
            PlGrensTot = Mid(txtTekstLijn(0).text, 20, 4) + Mid(txtTekstLijn(0).text, 17, 2) + Mid(txtTekstLijn(0).text, 14, 2)
            If PlGrensVan < Left(BOOKYEAR_FROMTO, 8) Or PlGrensTot > Right(BOOKYEAR_FROMTO, 8) Then
                'MsgBox "Geen selectie BUITEN het boekjaar a.u.b. !"
                'txtTekstLijn(0).Text = BYPERDAT.PeriodeBoekjaar.Text
                'txtTekstLijn(0).SetFocus
            End If
        End If
    
    Case 1
        If DATE_INVALID((txtTekstLijn(1).text)) Then
            Beep
            txtTekstLijn(1).text = MIM_GLOBAL_DATE
            txtTekstLijn(1).SetFocus
        End If
End Select

End Sub

