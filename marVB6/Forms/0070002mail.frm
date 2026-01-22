VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form KwijtingMailOptIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BetalingsVerzoeken Mail & OptIn"
   ClientHeight    =   5415
   ClientLeft      =   285
   ClientTop       =   405
   ClientWidth     =   10515
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
   LinkMode        =   1  'Source
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5415
   ScaleWidth      =   10515
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox cbTB2Info 
      Caption         =   "TB2 Detail op blad 2"
      Height          =   240
      Left            =   8160
      TabIndex        =   29
      Top             =   120
      Value           =   1  'Checked
      Width           =   2160
   End
   Begin VB.CheckBox cbPdfExport 
      Caption         =   "Pdf Exporteren"
      Height          =   255
      Left            =   4200
      TabIndex        =   28
      Top             =   360
      Width           =   1695
   End
   Begin VB.CheckBox cbPdfDetail 
      Caption         =   "Pdf Genereren"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   27
      Top             =   120
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox ckMetDatum 
      Alignment       =   1  'Right Justify
      Caption         =   "Met datum"
      Height          =   195
      Left            =   120
      TabIndex        =   25
      Top             =   900
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.ComboBox KeuzeInfo 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   2
      Left            =   150
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   420
      Visible         =   0   'False
      Width           =   3480
   End
   Begin VB.CheckBox cbIBANLayout 
      Caption         =   "IBAN Lay Out"
      Height          =   240
      Left            =   6120
      TabIndex        =   23
      Top             =   360
      Width           =   2520
   End
   Begin VB.CheckBox DetailVeld 
      Caption         =   "Polisinfo beperking tot Veld: DETAIL"
      Height          =   600
      Left            =   3720
      TabIndex        =   22
      Top             =   960
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.CheckBox Post 
      Caption         =   "P&ost"
      Height          =   240
      Left            =   3840
      TabIndex        =   13
      Top             =   720
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton TekstBewaren 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Opslaan"
      Height          =   330
      Left            =   4320
      TabIndex        =   11
      Top             =   3780
      Width           =   1335
   End
   Begin VB.CheckBox MetOverschrijving 
      Caption         =   "Overschrijvingsstrook"
      Height          =   240
      Left            =   5880
      TabIndex        =   12
      Top             =   120
      Value           =   1  'Checked
      Width           =   2265
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   5
      Left            =   7200
      TabIndex        =   5
      Top             =   900
      Visible         =   0   'False
      Width           =   3225
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   2130
      TabIndex        =   3
      Top             =   1230
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.TextBox TekstInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   7020
      TabIndex        =   15
      Text            =   "471.00"
      Top             =   4950
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.TextBox TekstInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   5370
      TabIndex        =   14
      Text            =   "145.56"
      Top             =   4950
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.TextBox TekstInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1200
      Index           =   1
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   4140
      Width           =   10215
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   2130
      TabIndex        =   1
      Top             =   870
      Width           =   1440
   End
   Begin VB.ComboBox KeuzeInfo 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3780
      Width           =   4080
   End
   Begin VB.ComboBox KeuzeInfo 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   150
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   3480
   End
   Begin VB.CommandButton Afsluiten 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Drukken"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9180
      TabIndex        =   8
      Top             =   1800
      Width           =   1230
   End
   Begin VB.CommandButton Annuleren 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   375
      Left            =   9180
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2820
      Width           =   1245
   End
   Begin VB.CommandButton Schoonvegen 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Schoon"
      Height          =   375
      Left            =   9180
      TabIndex        =   16
      Top             =   3270
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid PolisDetail 
      Height          =   2025
      Left            =   60
      TabIndex        =   6
      Top             =   1710
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   3572
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      FocusRect       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.OLE OLE1 
      Height          =   1095
      Left            =   3780
      TabIndex        =   26
      Top             =   2820
      Width           =   3015
   End
   Begin VB.Label LabelInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Polisdetail"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   7980
      TabIndex        =   7
      Top             =   1440
      Width           =   1035
   End
   Begin VB.Label CRLFCaption 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "0"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   9660
      TabIndex        =   21
      Top             =   3840
      Width           =   720
   End
   Begin VB.Label LabelInfo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Referte aanbod"
      Height          =   285
      Index           =   6
      Left            =   8760
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label LabelInfo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Kostprijs aanbod"
      Height          =   315
      Index           =   5
      Left            =   90
      TabIndex        =   2
      Top             =   1230
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label LabelInfo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lijnen"
      Height          =   240
      Index           =   4
      Left            =   8880
      TabIndex        =   20
      Top             =   3840
      Width           =   720
   End
   Begin VB.Label LabelInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "BRAND ER IDX"
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   3
      Left            =   6570
      TabIndex        =   19
      Top             =   4950
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label LabelInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "BA Priv‚ IDX"
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   2
      Left            =   3810
      TabIndex        =   18
      Top             =   4950
      Visible         =   0   'False
      Width           =   1440
   End
End
Attribute VB_Name = "KwijtingMailOptIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim SDLijn(3) As String
Dim SDTitel As String
Dim KlantNummer As String * 12
Dim TaalTekst As String * 2
Dim TotaalBEF As Currency
Dim TotaalEUR As Currency
Dim ReferteTxt As String * 20
Dim FlLogBestand As Integer

Dim VsoftVanaf As String
Dim VsoftTot As String

Dim BeginXbox As Integer
Dim BeginYbox As Integer
Dim OVSStrooklijnen

Dim KontaktPersoon As Integer
Dim IndexBM As String * 6
Dim TaksEnKost As Currency
Dim BedragEUR As Currency
Dim BrutoPremie As Currency
Dim PaperLess As Currency

Dim TotaalTeBetalen As Currency
Dim datKwijting As String

'Dim wApp As Word.Application

Dim pdfY As Double
Dim pdfOVSStrook As Double

Dim sharedPolis As String
Dim sharedTotaal As Currency

Dim isDOM As Boolean

Function TB2info2017() As String
Dim rsTMP As New ADODB.Recordset
Set rsTMP = New ADODB.Recordset

Dim tb2infoHier As String
Dim dummyHier As String
Dim buildUp As String
    
    
On Error Resume Next
Err = 0
buildUp = ""
rsTMP.CursorLocation = adUseClient
Msg = "SELECT * FROM Dokumenten WHERE v033 Like 'Q%' AND A000 = '" & sharedPolis & "' AND Val(v249) = " & Str(sharedTotaal - PaperLess) & " ORDER BY v033 DESC"
SnelHelpPrint Msg, blLogging
Screen.MousePointer = vbHourglass
rsTMP.Open Msg, adntDB, adOpenForwardOnly, adLockReadOnly
If Err Then
    MsgBox Error
Else
    If rsTMP.RecordCount Then
        If rsTMP.RecordCount <> 1 Then
            'MsgBox "recordcount is :" + Str(rsTMP.RecordCount)
        End If
        dummyHier = rodCheck(rsTMP("rvxmltb2"))
        If dummyHier <> "" Then
            buildUp = dummyHier
        End If
        dummyHier = icdCheck(rsTMP("rvxmltb2"))
        If dummyHier <> "" Then
            buildUp = buildUp + vbCrLf + dummyHier
        End If
        dummyHier = ftxCheck(rsTMP("rvxmltb2"))
        If dummyHier <> "" Then
            buildUp = buildUp + vbCrLf + vbCrLf + dummyHier
        End If
        'rsTMP("rvxmltb2")
    End If
End If
TB2info2017 = buildUp

End Function


Function TB2infoblad() As Boolean

Dim rsTMP As New ADODB.Recordset
Set rsTMP = New ADODB.Recordset

Dim tb2infoHier As String

    
On Error Resume Next
Err = 0
rsTMP.CursorLocation = adUseClient
Msg = "SELECT * FROM Dokumenten WHERE v033 Like 'Q%' AND A000 = '" & sharedPolis & "' AND Val(v249) = " & Str(sharedTotaal) & " ORDER BY v033 DESC"
SnelHelpPrint Msg, blLogging
Screen.MousePointer = vbHourglass
rsTMP.Open Msg, adntDB, adOpenForwardOnly, adLockReadOnly
If Err Then
    MsgBox Error
Else
    'Debug.Print Msg
    If rsTMP.RecordCount Then
        Do While Not rsTMP.EOF
            'Stop
            'TODO: info op voorkant mogelijk?
            '      boekdatum, hoofdvervaldag, mededeling (o.a. tariefaanpassing?), verzekeringsnemer, verzekerde periode, type risicoobject, waarborgen,
            Mim.Report.PageBreak
            With Mim.Report
                .FontName = "Courier New"
                .FontSize = 8
                .TextBold = True
                .TextColor = 0 'zwart
            End With
            tb2infoHier = tb2Indent(rsTMP("rvxmltb2"))
            If InStr(tb2infoHier, "Commissie ") Then
                tb2infoHier = Replace(tb2infoHier, "Commissie", "Commercialisatiekost")
            End If
            pdfY = Mim.Report.VpePrint(1, 1, tb2infoHier & vbCrLf)
            If rsTMP.RecordCount > 1 Then 'best verfijnen!
                SnelHelpPrint "Stop recordcount>1, nog te verfijnen voor " & sharedPolis & " :" & rsTMP.RecordCount, False
                Exit Do
            End If
            rsTMP.MoveNext
        Loop
    Else
        'meer dan een record te controleren eerst
        'nu gewoon blanco blad
        Mim.Report.PageBreak
        With Mim.Report
            .FontName = "Courier New"
            .FontSize = 8
            .TextBold = False
            .TextColor = 0 'zwart
        End With
    End If
End If
rsTMP.Close
Set rsTMP = Nothing

End Function

Private Sub Afsluiten_Click()
Dim T As Integer
Dim MaxGrens As Integer
Dim TempoFL As Integer
Dim PolisNummer As String * 12
Dim CrText As String
Dim dPip As Double
Dim XX As Integer
Dim BedragBEF As Currency
Dim result As Boolean
Dim ktrlBOOL As Boolean

'If Me.Detail.Value <> 0 Then
'    SnelHelpPrint "Word Document Object Inladen...", blLogging
'    Screen.MousePointer = vbHourglass
'    Me.Enabled = False
'    Err = 0
'    On Error Resume Next
'    'Set wApp = New Word.Application
'    'wApp.Visible = True
'    Mim.Enabled = False
'    Me.Enabled = False
'    If Err Then MsgBox Error
'    SnelHelpPrint "Klaar.", blLogging
'    Screen.MousePointer = vbNormal
'ElseIf Me.cbPdfDetail.Value = vbChecked Then
    With Mim.Report
        .OpenDoc
        .Author = "Jos Vermoesen"
        .GUILanguage = 3 'Nederlands
        .Title = "dnnInning4Brokers"
        .PenSize = 0.01
    End With
'Else
'    Set Printer = Printers(dokumentPrinterNr)
'    On Error Resume Next
'    Printer.PaperBin = LaadTekst(App.Title, "dokumentPRINTER")
'    If Printer.Orientation = vbPRORLandscape Then
'        Printer.Orientation = vbPRORPortrait
'        DoEvents
'    End If
'End If
MaxGrens = 10
KlantNummer = ""

If Left(KeuzeInfo(0).Text, 1) = "1" Then
    FlLogBestand = FreeFile
    Err = 0
    On Error Resume Next
    Open ProgrammaLokatie + "VVLijst.LOG" For Append As FlLogBestand
    If Err Then MsgBox Error
End If

Screen.MousePointer = vbHourglass
For T = 1 To PolisDetail.Rows - 2
    PolisDetail.Row = T
    PolisDetail.Col = 0
    PolisNummer = PolisDetail.Text
    bGet FlPolis, 0, PolisNummer
    If Ktrl Then
        MsgBox "Stop"
    Else
        RecordToVeld FlPolis
    End If
    
    If Left(KeuzeInfo(0).Text, 1) = "1" Then
        'If Detail.Value <> 0 Then
        '    teldetail = 0
        '    TotaalBEF = 0
        '    TotaalEUR = 0
        'End If
    End If
    If KlantNummer = vSet(vBibTekst(FlPolis, "#A110 #"), 12) Then
        teldetail = teldetail + 1
        
        '2009: een per blad forceren!
        'If TelDetail > 3 Then
            XX = DrukAf
            teldetail = 0
            TotaalBEF = 0
            TotaalEUR = 0
            If Me.cbPdfDetail.Value = vbChecked Then
                ktrlBOOL = Mim.Report.PageBreak
                'Stop
            End If
        'End If
    ElseIf KlantNummer = Space$(12) Then
        teldetail = 0
        TotaalBEF = 0
        TotaalEUR = 0
    Else
       DrukAf
       teldetail = 0
       TotaalBEF = 0
       TotaalEUR = 0
       If Me.cbPdfDetail.Value = vbChecked Then
           ktrlBOOL = Mim.Report.PageBreak
       End If

    End If

    bGet FlKlant, 0, vBibTekst(FlPolis, "#A110 #")
    If Ktrl Then
        MsgBox "stop"
    Else
        RecordToVeld FlKlant
        If Val(vBibTekst(FlKlant, "#A10C #")) <> KeuzeInfo(1).ListIndex + 1 Then
            KeuzeInfo(1).ListIndex = Val(vBibTekst(FlKlant, "#A10C #")) - 1
        End If
    End If

    KlantNummer = vBibTekst(FlKlant, "#A110 #")
    bGet FlLeverancier, 0, "CO" + vBibTekst(FlPolis, "#A010 #")
    If Ktrl Then
        Beep
    Else
        RecordToVeld FlLeverancier
    End If
    
    Select Case KeuzeInfo(0).ListIndex
        Case 1 To 4
            Select Case Left(KeuzeInfo(1).Text, 1)
                Case "1"
                              SDTitel = "Compagnie   " + Space$(14)
                    SDTitel = SDTitel + "Déscription " + Space$(19)
                    SDTitel = SDTitel + "BM/INDX    "
                    SDTitel = SDTitel + "Splitsing" + Space$(4)
                    SDTitel = SDTitel + " Prime en BEF" + vbCrLf
                    SDTitel = SDTitel + "Contrat n°   " + Space$(13)
                    SDTitel = SDTitel + "Détail" + Space$(25)
                    SDTitel = SDTitel + "Echéan.    " + Space$(13)
                    SDTitel = SDTitel + " Prime EUR"
                    SDTitel = UCase$(SDTitel)
                Case "2", "3", "4"
                              SDTitel = "Maatschappij" + Space$(14)
                    SDTitel = SDTitel + "Omschrijving" + Space$(19)
                    SDTitel = SDTitel + "BM/INDX    "
                    SDTitel = SDTitel + "Splitsing" + Space$(4)
                    SDTitel = SDTitel + "Premie BEF" + vbCrLf
                    SDTitel = SDTitel + "Kontrakt nr. " + Space$(13)
                    SDTitel = SDTitel + "Detail" + Space$(25)
                    SDTitel = SDTitel + "Vv.dag     " + Space$(13)
                    SDTitel = SDTitel + "Premie EUR"
                    SDTitel = UCase$(SDTitel)
            End Select

            bGet FlAllerlei, 1, "25" + vSet(vBibTekst(FlPolis, "#A010 #"), 4) + vBibTekst(FlPolis, "#A000 #")
            IndexBM = "  --  "
            TaksEnKost = 0
            PaperLess = 0
            If Ktrl Then
            Else
                RecordToVeld FlAllerlei
                If Val(vBibTekst(FlPolis, "#B010 #")) = Val(vBibTekst(FlAllerlei, "#B010 #")) Then
                    TaksEnKost = Val(vBibTekst(FlAllerlei, "#B011 #"))
                    BrutoPremie = Val(vBibTekst(FlAllerlei, "#B013 #"))
                    PaperLess = Val(vBibTekst(FlPolis, "#v403 #"))
                End If
                If Left(vBibTekst(FlPolis, "#v223 #"), 1) = "5" Then
                    CrText = vBibTekst(FlAllerlei, "#5315 #")
                    'BM voor huidige premie BA
                    If Trim$(CrText) = "" Then
                    Else
                        IndexBM = Format(Val(CrText), "00")
                    End If
                Else
                    'controle aanwezigheid index
                    CrText = vBibTekst(FlAllerlei, "#AW.R #")
                    On Error Resume Next
                    If CDbl(CrText) > 99 Then
                        IndexBM = CrText
                    Else
                        IndexBM = Format(Val(CrText) / 100, "000.00")
                    End If
                End If
            End If
                 
                                SDLijn(teldetail) = vSet(vBibTekst(FlLeverancier, "#A100 #"), 25) + " "
            SDLijn(teldetail) = SDLijn(teldetail) + vSet(vBibTekst(FlPolis, "#vs99 #"), 30) + " "
            SDLijn(teldetail) = SDLijn(teldetail) + IndexBM + "     "
            SDLijn(teldetail) = SDLijn(teldetail) + vSet(Mid(fmarBoxText("914", Val(KeuzeInfo(1)), vBibTekst(FlPolis, "#A325 #")), 4), 14)
            If Val(vBibTekst(FlPolis, "#A325 #")) > 7 Then
                isDOM = True
            Else
                isDOM = False
            End If
       
            If IsNull(rsMAR(FlPolis)("e069")) Then
                BedragEUR = Val(rsMAR(FlPolis)("B010"))
            Else
                BedragEUR = Val(rsMAR(FlPolis)("B010")) + Val(rsMAR(FlPolis)("e069")) + Val(vBibTekst(FlPolis, "#v403 #"))
            End If
            BedragBEF = Round(BedragEUR * Euro)
            datKwijting = vSet(Mid(vBibTekst(FlPolis, "#AW_2 #"), 7, 2), 2) + "/" + vSet(vBibTekst(FlPolis, "#v164 #"), 2) + "/" + Right(Rdt, 4)
            
            SDLijn(teldetail) = SDLijn(teldetail) + Dec$((BedragBEF), MaskerSy(0)) + vbCrLf   'Premie in BEF
                        SDLijn(teldetail) = SDLijn(teldetail) + vSet(vBibTekst(FlPolis, "#A000 #"), 12) + Space$(14)
            SDLijn(teldetail) = SDLijn(teldetail) + vSet(vBibTekst(FlPolis, "#vs98 #"), 30) + " "
            SDLijn(teldetail) = SDLijn(teldetail) + datKwijting
            SDLijn(teldetail) = SDLijn(teldetail) + "              " + Dec((BedragEUR), MaskerEUR) 'Premie in EUR
            
            TotaalBEF = TotaalBEF + BedragBEF
            TotaalEUR = TotaalEUR + BedragEUR

            If DetailVeld.Visible = True And DetailVeld.Value = 1 Then
                ReferteTxt = vBibTekst(FlPolis, "#vs98 #")
            Else
                ReferteTxt = "+++" + Format(teldetail + 1) + Format(Val(vBibTekst(FlPolis, "#v164 #")), "00") + "/" + Mid(KlantNummer, 1, 4) + "/0" + Mid(KlantNummer, 5, 2) + "xx+++"
                dPip = Val(Mid(ReferteTxt, 4, 3) + Mid(ReferteTxt, 8, 4) + Mid(ReferteTxt, 13, 3))
                Mid(ReferteTxt, 16, 2) = Format(dPip - Int(dPip / 97) * 97, "00")
                If Mid(ReferteTxt, 16, 2) = "00" Then
                    Mid(ReferteTxt, 16, 2) = "97"
                End If
            End If
        Case 5
            If bhEuro Then
                TotaalEUR = Val(TekstInfo(4))
                TotaalBEF = Round(TotaalEUR * Euro)
            Else
                TotaalBEF = Val(TekstInfo(4))
                TotaalEUR = Round(TotaalBEF / Euro, 2)
            End If
            ReferteTxt = TekstInfo(5)
        Case Else
            MsgBox "STOP"
    End Select
    If Left(KeuzeInfo(0).Text, 1) = "1" Then
        'If Detail.Value <> 0 Then
        '    result = wordPREMIE(rsMAR(FlPolis)("A010"), Left(KeuzeInfo(1).Text, 1), Right(KeuzeInfo(0).Text, Len(KeuzeInfo(0).Text) - 3), TekstInfo(1).Text)
        'End If
    End If
NextLijn:
Next
If Left(KeuzeInfo(0).Text, 1) = "1" Then
    'If Detail.Value <> 0 Then
    'Else
        If KlantNummer <> Space$(12) Then
            XX = DrukAf
            TotaalBEF = 0
            TotaalEUR = 0
            'Printer.NewPage
        End If
    'End If
Else
    If KlantNummer <> Space$(12) Then
        XX = DrukAf
        TotaalBEF = 0
        TotaalEUR = 0
        If Me.cbPdfDetail.Value = vbChecked Then
            If Me.PolisDetail.Rows > 3 Then
                ktrlBOOL = Mim.Report.PageBreak
            End If
        End If
        'Printer.NewPage
    End If
End If
Const VMAIL_TO = 1

If Me.cbPdfDetail.Value = vbChecked Then
    Mim.Report.WriteDoc (ProgrammaLokatie & Format(Now, "YYYYMMDDHHMMSS") & "-dnnInning4Brokers.pdf")
    'ktrlBOOL = Mim.Report.AddMailReceiver("[FAX: +32 53781922]", VMAIL_TO)
    ' Doc.AddMailReceiver("[FAX: +49 1234 12345678]", VMAIL_TO)


    If Me.cbPdfExport.Value = vbChecked Then
        MsgBox "exporteren naar DNN folder of andere nog te definiëren"
        Mim.Report.CloseDoc
    Else
        Mim.Report.Preview
    End If
Else
    Printer.EndDoc
End If

If Left(KeuzeInfo(0).Text, 1) = "1" Then
    Close FlLogBestand
End If

On Error Resume Next
'wApp.Visible = False
Mim.Enabled = True
Me.Enabled = True
'Set wApp = Nothing

Screen.MousePointer = vbNormal
Schoon
KeuzeInfo(0).SetFocus

End Sub

Private Sub InstallMij()
Dim T As Integer

KeuzeInfo(2).Clear
KeuzeInfo(2).AddItem "Alle maatschappijen"
T = -1
bGetOrGreater FlLeverancier, 0, vSet("CO", 12)
If Ktrl Then
    Beep
    Exit Sub
Else
    RecordToVeld FlLeverancier
End If
Screen.MousePointer = vbHourglass
If Left(KeyBuf(FlLeverancier), 2) <> "CO" Then
    Beep
Else
    T = T + 1
    KeuzeInfo(2).AddItem Mid(vBibTekst(FlLeverancier, "#A110 #"), 3, 4) + ": " + vBibTekst(FlLeverancier, "#A100 #") + "/" + vSet(vBibTekst(FlLeverancier, "#A400 #"), 8)
    Do
        bNext FlLeverancier
        If Ktrl Or Left(KeyBuf(FlLeverancier), 2) <> "CO" Then
            Exit Do
        Else
            RecordToVeld FlLeverancier
            T = T + 1
            KeuzeInfo(2).AddItem Mid(vBibTekst(FlLeverancier, "#A110 #"), 3, 4) + ": " + vBibTekst(FlLeverancier, "#A100 #") + "/" + vSet(vBibTekst(FlLeverancier, "#A400 #"), 8)
        End If
    Loop
End If
Screen.MousePointer = vbNormal

End Sub


Private Sub Annuleren_Click()

If PolisDetail.Rows > 2 Then
    Msg$ = "Aangeduide verrichtingen negeren." + vbCrLf + vbCrLf + "Bent U zeker ?"
    Ktrl = MsgBox(Msg$, 292, "Voortijdig stoppen...")
    If Ktrl = 6 Then
    Else
        Exit Sub
    End If
End If
GridText = ""
Unload Me

End Sub

Private Function DrukAf()
Dim hTAB    As Integer
Dim dokumentType As String
ReDim rSip(6) As String
ReDim sSip(6) As String
ReDim VeldInfo(9) As String
ReDim dVeldInfo(7) As Double
Dim BtwTekst As String * 4
Dim OMSTab As Integer
Dim TempoEvent As Integer

Dim BedragTxt As String
Dim rNTTxt As String
Dim rNTTxt2 As String

Dim T                   As Integer
Dim TT                  As Integer
Dim Taal                As String
Dim FlTemp              As Integer
ReDim rft(10) As String
Dim rnr                 As String * 13
Dim sy                  As String
Dim sy2                 As String

Dim XVan                As Single
Dim XTot                As Single
Dim YVan                As Single
Dim YTot                As Single
Dim X                   As Integer
Dim MeerLijn            As Integer
Dim ktrlHier As Integer

Dim Nog As Integer
Dim NogString As String

'If Me.Detail.Value <> 0 Then Exit Function
If Me.cbPdfDetail.Value = vbChecked Then
    pdfDrukAf
    Exit Function
End If

OVSStrooklijnen = 5100
Taal = vBibTekst(FlKlant, "#A10C #")
FlTemp = FreeFile

GoSub KopBalk
GoSub DetailLijnen
Select Case Left(KeuzeInfo(0).Text, 1)
    Case "1", "2", "3", "5"
        If MetOverschrijving.Value Then
            GoSub Overschrijvingsstrook
        End If
    Case Else
End Select
Printer.NewPage
'Printer.EndDoc
XDoEvents = DoEvents
KlantNummer = ""
Exit Function

KopBalk:
pfTl = 0
Pagina = Pagina + 1
dokumentType = Right(KeuzeInfo(0).Text, Len(KeuzeInfo(0).Text) - 3)
If Val(vBibTekst(FlKlant, "#A102 #")) = 0 Then
    rSip(0) = vBibTekst(FlKlant, "#A100 #") + " " + vBibTekst(FlKlant, "#A101 #")
Else
    rSip(0) = Mid(fmarBoxText("003", Taal, vBibTekst(FlKlant, "#A102 #")), 4, 10) + " " + vBibTekst(FlKlant, "#A100 #") + " " + vBibTekst(FlKlant, "#A101 #")
End If
If KontaktPersoon = 1 Then
    If Val(vBibTekst(FlKlant, "#vs01 #")) = 0 Then
        rSip(1) = vBibTekst(FlKlant, "#A125 #") + " " + vBibTekst(FlKlant, "#A127 #")
    Else
        rSip(1) = Mid(fmarBoxText("003", Taal, vBibTekst(FlKlant, "#vs01 #")), 4, 10) + " " + vBibTekst(FlKlant, "#A125 #") + " " + vBibTekst(FlKlant, "#A127 #")
    End If
Else
    rSip(1) = ""
End If
rSip(2) = vBibTekst(FlKlant, "#A104 #") + " " + vBibTekst(FlKlant, "#A105 #") + " " + vBibTekst(FlKlant, "#A106 #")
rSip(4) = vBibTekst(FlKlant, "#A109 #") + " " + vBibTekst(FlKlant, "#A107 #") + " " + vBibTekst(FlKlant, "#A108 #")
GoSub PrintKopTekst

Printer.FontSize = 10
Printer.FontBold = True

Printer.CurrentY = Val(VsoftVanaf)
Printer.CurrentX = 0

If ckMetDatum Then
Printer.Print Tab(7);
    XVan = Printer.CurrentX
    YVan = Printer.CurrentY
Printer.Print dokumentType;

    XTot = Printer.CurrentX
    YTot = Printer.CurrentY

Printer.DrawWidth = 2
Printer.Line (XVan - 50, YVan - 50)-(XTot + 50, YTot + 250), , B
    
    Printer.CurrentX = XVan     'ander geen afdruk met vb4 32bit
    Printer.CurrentY = YVan     'ander geen afdruk met vb4 32bit
    Printer.Print dokumentType; 'ander geen afdruk met vb4 32bit
    Printer.CurrentX = XTot     'ander geen afdruk met vb4 32bit
    Printer.CurrentY = YTot     'ander geen afdruk met vb4 32bit
End If

Printer.FontBold = False
If ckMetDatum Then
    Printer.Print Tab(60); Format(Now, "DDDD dd mmmm yyyy"); vbCrLf;
Else
    Printer.Print Tab(60); " "; vbCrLf;
End If
Printer.FontName = "Courier New"
Printer.FontSize = 10
Return

DetailLijnen:
If DetailVeld.Visible = True And DetailVeld.Value = 1 Then
    GoTo InfoTekst
Else
    Printer.FontBold = False
    Printer.Print vbCrLf; Tab(2);
    XVan = Printer.CurrentX
    YVan = Printer.CurrentY
    Printer.Print SDTitel;
    
    XTot = Printer.CurrentX
    YTot = Printer.CurrentY
    Printer.DrawWidth = 1
    Printer.Line (XVan - 50, YVan - 50)-(XTot + 50, YTot + 250), , B
    
    Printer.CurrentX = XVan 'ander geen afdruk met vb4 32bit
    Printer.CurrentY = YVan 'ander geen afdruk met vb4 32bit
    Printer.Print SDTitel;  'ander geen afdruk met vb4 32bit
    Printer.CurrentX = XTot 'ander geen afdruk met vb4 32bit
    Printer.CurrentY = YTot 'ander geen afdruk met vb4 32bit
    Printer.Print vbCrLf; vbCrLf; 'ander geen afdruk met vb4 32bit
    Printer.FontBold = True
End If

If Left(KeuzeInfo(0).Text, 1) = "5" Then
    TotMax = 0
Else
    TotMax = 3
End If
For T = 0 To TotMax
    If SDLijn(T) <> "" Then
        Printer.Print Tab(2); SDLijn(T); vbCrLf; vbCrLf;
        SDLijn(T) = ""
    Else
        Printer.Print vbCrLf; vbCrLf; vbCrLf;
    End If
    pfTl = pfTl + 3
Next
Printer.FontBold = False

InfoTekst:
'If Left(KeuzeInfo(0).Text, 1) = "5" Then
'    MsgBox "Stop"
    'Printer.Print pVerkleind(PrinterType);
'End If
Printer.FontName = "Arial"
Printer.FontSize = 9
If DetailVeld.Visible = True And DetailVeld.Value = 1 Then
    InsPos = InStr(TekstInfo(1).Text, "<##>")
    If InsPos Then
        LenTekst = Len(TekstInfo(1).Text)
        TekstInfo(1).Text = Mid(TekstInfo(1).Text, 1, InsPos - 1) + UCase$(ReferteTxt) + Right(TekstInfo(1).Text, LenTekst - (InsPos + 3))
    End If
End If
Printer.Print TekstInfo(1).Text;

'If Left(KeuzeInfo(0).Text, 1) = "5" Then
'    MsgBox "stop"
'    'Printer.Print pNormaal(PrinterType);
'End If
Return


PrintKopTekst:
PrintUserDef "1" + Format(Taal) + "0"
Printer.CurrentX = BeginXbox + 350
Printer.CurrentY = BeginYbox + 350
Printer.FontSize = 10
Printer.FontBold = True
For tSip = 0 To 4
    Printer.Print UCase$(rSip(tSip)); vbCrLf;
    Printer.CurrentX = BeginXbox + 350
Next
Return

Overschrijvingsstrook:
If Dir$(BedrijfsLokatie + "LicMarnt.###") = "" Then
    Beep
    MsgBox BedrijfsLokatie + "LicMarnt.### niet te vinden !  Hierna wordt kladblok opgestart.  Breng uw eigen gegevens in a.u.b. !", 0, "Foutieve Installatie ?"

    On Local Error Resume Next
    X = Shell("notepad.exe " + BedrijfsLokatie + "LicMarnt.###", 1)
    Return
Else
    FlTemp = FreeFile
    Open BedrijfsLokatie + "LicMarnt.###" For Input As FlTemp
        Line Input #FlTemp, sSip(0)
        Line Input #FlTemp, sSip(1)
        Line Input #FlTemp, sSip(2)
        Line Input #FlTemp, sSip(3)
        Line Input #FlTemp, sSip(4)
        Line Input #FlTemp, sSip(5)
        Line Input #FlTemp, A$
        Err = 0
        On Local Error Resume Next
        Line Input #FlTemp, NogString
        If Err Then
            Nog = 0
        Else
            Nog = Val(NogString)
        End If
    Close FlTemp
    hTAB = Val(A$)
End If

Printer.FontName = "Courier New"
Printer.FontSize = 12
Printer.FontBold = True
Printer.CurrentY = 11000
Printer.CurrentX = 0
For T = 1 To 1 + Nog
    Printer.Print vbCrLf;
Next
If Me.cbIBANLayout.Value = vbUnchecked Then
    BedragTxt = "*" + Format(TotaalEUR, MaskerEUR) + "*"
Else
    MsgBox "stop voor BEF"
    BedragTxt = "*" + Format(TotaalBEF, MaskerSy(0)) + "*"
End If

NieuwFormulier:
If Me.cbIBANLayout.Value = vbUnchecked Then
    rNTTxt = Dec((TotaalEUR), "#######0.00")
    GoSub Spatieren
    Mid(rNTTxt2, 17, 1) = " "
Else
    MsgBox "stop voor BEF"
    rNTTxt = Dec((TotaalBEF), "##########")
    GoSub Spatieren
End If
    
    Printer.Print Tab(10 + hTAB); BedragTxt;
    Printer.Print Tab(52 + hTAB); rNTTxt2; vbCrLf; vbCrLf;

rNTTxt = Mid(sSip(0), 1, 3) + Mid(sSip(0), 5, 7) + Mid(sSip(0), 13, 2)
    GoSub Spatieren
    rNTTxt2 = Mid(rNTTxt2, 1, 6) + " " + Mid(rNTTxt2, 7, 14) + " " + Mid(rNTTxt2, 21, 3)
    Printer.Print Tab(52 + hTAB); rNTTxt2;
    Printer.FontBold = False
    Printer.Print vbCrLf; vbCrLf;

For T = 0 To 4
    Printer.Print Tab(25 + hTAB); Left(UCase$(rSip(T)), 26); Tab(52 + hTAB); UCase$(sSip(T + 1)); vbCrLf;
Next

rNTTxt = ReferteTxt
    GoSub Spatieren
    Printer.FontBold = True
    Printer.Print vbCrLf; Tab(1 + hTAB); ReferteTxt; Tab(25 + hTAB); rNTTxt2;
    Printer.FontBold = False
Return

Spatieren:
iRNTxt = Len(rNTTxt)
rNTTxt2 = ""
For Teltxt = 1 To iRNTxt
    rNTTxt2 = rNTTxt2 + Mid(rNTTxt, Teltxt, 1) + " "
Next
Return

End Function
Private Function pdfDrukAf()

Dim hTAB    As Integer
Dim dokumentType As String
ReDim rSip(6) As String
ReDim sSip(6) As String
ReDim VeldInfo(9) As String
ReDim dVeldInfo(7) As Double
Dim BtwTekst As String * 4
Dim OMSTab As Integer
Dim TempoEvent As Integer

Dim BedragTxt As String
Dim rNTTxt As String
Dim rNTTxt2 As String

Dim T                   As Integer
Dim TT                  As Integer
Dim Taal                As String
Dim FlTemp              As Integer
ReDim rft(10) As String
Dim rnr                 As String * 13
Dim sy                  As String
Dim sy2                 As String

Dim XVan                As Single
Dim XTot                As Single
Dim YVan                As Single
Dim YTot                As Single
Dim X                   As Integer
Dim MeerLijn            As Integer

Dim Nog As Integer
Dim NogString As String

Dim adresString As String
Dim ovsDefinitie As String

'If Me.Detail.Value <> 0 Then Exit Function


pdfOVSStrook = 8.2
If Me.cbIBANLayout.Value = vbChecked Then
    ovsDefinitie = "FORM-IBAN.TXT"
Else
    'ovsDefinitie = "FORM-EURO.TXT"
    'vanaf 02/2021 enkel nog IBAN
    'check dient voor evaluatie papier/pdf zonder overschrijvingsstrook
    ovsDefinitie = "FORM-IBAN.TXT"
End If
Taal = vBibTekst(FlKlant, "#A10C #")
FlTemp = FreeFile
    
            
GoSub KopBalk
GoSub DetailLijnen
Select Case Left(KeuzeInfo(0).Text, 1)
    Case "1", "2", "3", "5"
        If MetOverschrijving.Value Then
            GoSub Overschrijvingsstrook
        End If
    Case Else
End Select
XDoEvents = DoEvents

'If Me.cbTB2Info.Value = vbChecked Then
'    Ktrl = TB2infoblad
'End If
'KlantNummer = ""
Exit Function

KopBalk:
pfTl = 0
Pagina = Pagina + 1
dokumentType = Right(KeuzeInfo(0).Text, Len(KeuzeInfo(0).Text) - 3)
If Val(vBibTekst(FlKlant, "#A102 #")) = 0 Then
    rSip(0) = vBibTekst(FlKlant, "#A100 #") + " " + vBibTekst(FlKlant, "#A101 #")
Else
    rSip(0) = Mid(fmarBoxText("003", Taal, vBibTekst(FlKlant, "#A102 #")), 4, 10) + " " + vBibTekst(FlKlant, "#A100 #") + " " + vBibTekst(FlKlant, "#A101 #")
End If
If KontaktPersoon = 1 Then
    If Val(vBibTekst(FlKlant, "#vs01 #")) = 0 Then
        rSip(1) = vBibTekst(FlKlant, "#A125 #") + " " + vBibTekst(FlKlant, "#A127 #")
    Else
        rSip(1) = Mid(fmarBoxText("003", Taal, vBibTekst(FlKlant, "#vs01 #")), 4, 10) + " " + vBibTekst(FlKlant, "#A125 #") + " " + vBibTekst(FlKlant, "#A127 #")
    End If
Else
    rSip(1) = ""
End If
rSip(2) = vBibTekst(FlKlant, "#A104 #") + " " + vBibTekst(FlKlant, "#A105 #") + " " + vBibTekst(FlKlant, "#A106 #")
rSip(4) = vBibTekst(FlKlant, "#A109 #") + " " + vBibTekst(FlKlant, "#A107 #") + " " + vBibTekst(FlKlant, "#A108 #")

With Mim.Report
    .FontName = "Courier New"
    .FontSize = 10
    .TextBold = True
    .TextColor = 0 'zwart
End With
GoSub PrintKopTekst

'pdfVsoftVanaf
If ckMetDatum Then
    'pdfY = Mim.Report.VpePrint(1.7, pdfVsoftVanaf, dokumentType)
    Mim.Report.PenSize = 0.01
    pdfY = Mim.Report.PrintBox(1.7, pdfVsoftVanaf, UCase(dokumentType))
End If

Mim.Report.TextBold = False
If ckMetDatum Then
    pdfY = Mim.Report.VpePrint(10, pdfVsoftVanaf, Format(Now, "DDDD dd mmmm yyyy") & vbCrLf & vbCrLf)
Else
    pdfY = Mim.Report.VpePrint(10, pdfvsofvanaf, " ")
End If
Mim.Report.FontName = "Courier New"
Mim.Report.FontSize = 10
Return

DetailLijnen:
If DetailVeld.Visible = True And DetailVeld.Value = 1 Then
    MsgBox "stop"
    GoTo InfoTekst
Else
    Mim.Report.TextBold = False
    Mim.Report.PenSize = 0.01
    pdfY = Mim.Report.PrintBox(0.7, pdfY, SDTitel)
    pdfY = Mim.Report.VpePrint(0.7, pdfY, " ")
    Mim.Report.TextBold = True
End If

'If Left(KeuzeInfo(0).Text, 1) = "5" Then
    TotMax = 0
'Else
'    TotMax = 3
'End If
For T = 0 To TotMax
    If SDLijn(T) <> "" Then
        pdfY = Mim.Report.VpePrint(0.8, pdfY, SDLijn(T)) 'telt voor 2 lijnen!!
        pdfY = Mim.Report.VpePrint(0.8, pdfY, " ")
        If T = 0 Then
            sharedPolis = Trim(Mid(SDLijn(0), 94, 12))
            sharedTotaal = Val(Right(SDLijn(0), 10))
        End If
        SDLijn(T) = ""
    Else
        pdfY = Mim.Report.VpePrint(0.9, pdfY, " ")
        pdfY = Mim.Report.VpePrint(0.9, pdfY, " ")
        pdfY = Mim.Report.VpePrint(0.9, pdfY, " ")
    End If
    pfTl = pfTl + 3
Next

Mim.Report.FontName = "Arial"
Mim.Report.FontSize = 8
Mim.Report.TextBold = False

Msg = TB2info2017()
If Msg = "" Then
    pdfY = Mim.Report.VpePrint(0.9, pdfY, " ")
    pdfY = Mim.Report.VpePrint(0.9, pdfY, " ")
    pdfY = Mim.Report.VpePrint(0.9, pdfY, " ")
Else
    pdfY = Mim.Report.VpePrint(0.6, pdfY, Msg)
End If

InfoTekst:
Mim.Report.FontName = "Arial"
Mim.Report.FontSize = 9
If DetailVeld.Visible = True And DetailVeld.Value = 1 Then
    InsPos = InStr(TekstInfo(1).Text, "<##>")
    If InsPos Then
        LenTekst = Len(TekstInfo(1).Text)
        TekstInfo(1).Text = Mid(TekstInfo(1).Text, 1, InsPos - 1) + UCase$(ReferteTxt) + Right(TekstInfo(1).Text, LenTekst - (InsPos + 3))
    End If
End If

'pdfY = pdfY + 2 '+2= 15,57398 ; +3= 16,57398
pdfY = 15.57398
pdfY = Mim.Report.VpePrint(0.6, pdfY, TekstInfo(1).Text)
Return

PrintKopTekst:
pdfPrintUserDef "1" + Format(Taal) + "0", pdfOVSStrook
With Mim.Report
    .FontSize = 10
    .TextBold = True
    .TextItalic = False
    .TextUnderline = False
    .TextAlignment = 0
    .TextColor = 0 'zwart
End With

For tSip = 0 To 4
    adresString = adresString & UCase(rSip(tSip)) & vbCrLf
    'Printer.Print UCase$(rSip(tSip)); vbCrLf;
    'Printer.CurrentX = BeginXbox + 350
Next
ktrlHier = Mim.Report.Write(pdfadresXpos, pdfadresYpos, pdfadresXpos2, pdfadresYpos2, adresString)
Return

Overschrijvingsstrook:
On Local Error GoTo 0
If Dir(BedrijfsLokatie & ovsDefinitie) = "" Then
    MsgBox BedrijfsLokatie & ovsDefinitie & " niet te vinden !  Hierna wordt kladblok opgestart.  Breng uw eigen gegevens in a.u.b. !", 0, "Foutieve Installatie ?"
    On Local Error Resume Next
    X = Shell("notepad.exe " + BedrijfsLokatie + ovsDefinitie, 1)
    Return
Else
    FlTemp = FreeFile
    Open BedrijfsLokatie + ovsDefinitie For Input As FlTemp
        Line Input #FlTemp, sSip(0)
        Line Input #FlTemp, sSip(1)
        Line Input #FlTemp, sSip(2)
        Line Input #FlTemp, sSip(3)
        Line Input #FlTemp, sSip(4)
        Line Input #FlTemp, sSip(5)
        Line Input #FlTemp, sSip(6)
    Close FlTemp
End If
Mim.Report.FontName = "Courier New"
Mim.Report.FontSize = 12
Mim.Report.TextBold = True

If isDOM Then
    rNTTxt = "********.**"
Else
    rNTTxt = Dec((TotaalEUR), "#######0.00")
End If

If Me.cbIBANLayout.Value = vbUnchecked Then
    Dim betaalinfo As String
    'geen ovs strook
    If isDOM Then
        betaalinfo = "Betalingsinformatie" & vbCrLf & "(domiciliëring enkel ter info)" & vbCrLf & vbCrLf
    Else
        betaalinfo = "Betalingsinformatie" & vbCrLf & vbCrLf
    End If
    betaalinfo = betaalinfo & "Bedrag: " & TotaalEUR & " EUR" & vbCrLf
    betaalinfo = betaalinfo & "Rekeningnummer: " & sSip(0) & vbCrLf
    betaalinfo = betaalinfo & "BIC: " & sSip(1) & vbCrLf
    betaalinfo = betaalinfo & "Op naam: " & sSip(2) & vbCrLf
    betaalinfo = betaalinfo & "Gestructureerde mededeling: " & ReferteTxt & vbCrLf
    Mim.Report.PenSize = 0.02
    Mim.Report.TextBold = True
    Mim.Report.TextColor = 0 ' 0 'zwart
    pdfY = Mim.Report.PrintBox(1.7, 22.5, UCase(betaalinfo))
    Mim.Report.TextBold = False
    pdfY = Mim.Report.VpePrint(1.7, pdfY, vbCrLf & "Liever overschrijvingsstrook?" & vbCrLf & "Laat het ons weten!")
    Return
End If

GoSub Spatieren
Mid(rNTTxt2, 17, 1) = " "
Dim tmppdfY As Double
If Me.cbIBANLayout.Value = vbUnchecked Then
    BedragTxt = "*" + Format(TotaalEUR, MaskerEUR) + "*"
    pdfY = Mim.Report.VpePrint(3, 22.5, BedragTxt)
    pdfY = Mim.Report.VpePrint(13.8, 22.5, rNTTxt2 & vbCrLf & vbCrLf)
    rNTTxt = Mid(sSip(0), 1, 3) + Mid(sSip(0), 5, 7) + Mid(sSip(0), 13, 2)
        GoSub Spatieren
    
    'rNTTxt2 = Mid(rNTTxt2, 1, 6) + " " + Mid(rNTTxt2, 7, 14) + " " + Mid(rNTTxt2, 21, 3)
    rNTTxt2 = Mid(rNTTxt2, 1, 6) + Mid(rNTTxt2, 7, 14) + Mid(rNTTxt2, 21, 3)
    pdfY = Mim.Report.VpePrint(13.8, pdfY, rNTTxt2 & vbCrLf & vbCrLf)
    
    Mim.Report.TextBold = False
    
    For T = 0 To 4
        tmppdfY = pdfY
        pdfY = Mim.Report.VpePrint(7, tmppdfY, Left(UCase(rSip(T)), 26))
        pdfY = Mim.Report.VpePrint(13.8, tmppdfY, UCase(sSip(T + 2)))
    Next
    pdfY = Mim.Report.VpePrint(13.8, pdfY, " ")
   
    rNTTxt = ReferteTxt
        GoSub Spatieren
        Mim.Report.TextBold = True
        tmppdfY = pdfY
    
        pdfY = Mim.Report.VpePrint(0.5, tmppdfY, ReferteTxt)
        pdfY = Mim.Report.VpePrint(6.7, tmppdfY, rNTTxt2)
    
    Mim.Report.TextBold = False
Else
    pdfY = Mim.Report.VpePrint(15, 22, rNTTxt2)     'bedrag
    
    Mim.Report.TextBold = False
    rNTTxt = Left(UCase(rSip(0)), 26) 'Klant naam1
        GoSub Spatieren
        pdfY = Mim.Report.VpePrint(3.6, 23.7, rNTTxt2 & vbCrLf)
    
    rNTTxt = Left(UCase(rSip(1)), 26) 'Klant naam2
        GoSub Spatieren
        pdfY = Mim.Report.VpePrint(3.6, pdfY, rNTTxt2 & vbCrLf)
    
    rNTTxt = Left(UCase(rSip(4)), 26) 'Klant plaats
        GoSub Spatieren
        pdfY = Mim.Report.VpePrint(3.6, pdfY, rNTTxt2 & vbCrLf)
        
    Mim.Report.TextBold = True
        If isDOM = True Then
            rNTTxt = "*domiciliëring: enkel ter info*"
        Else
            rNTTxt = sSip(0)                                'IBANbedrijf
        End If
        GoSub Spatieren
        pdfY = Mim.Report.VpePrint(3.6, 25.3, rNTTxt2 & vbCrLf & vbCrLf)
    
    rNTTxt = sSip(1)                                'BICbedrijf
        GoSub Spatieren
        pdfY = Mim.Report.VpePrint(3.6, pdfY, rNTTxt2 & vbCrLf & vbCrLf)
    
    Mim.Report.TextBold = False
    For T = 2 To 4
        rNTTxt = sSip(T) 'ADRESbedrijf
            GoSub Spatieren
        pdfY = Mim.Report.VpePrint(3.6, pdfY, UCase(rNTTxt2))
    Next
        
    rNTTxt = ReferteTxt
        GoSub Spatieren
        Mim.Report.TextBold = True
        pdfY = Mim.Report.VpePrint(3.6, pdfY, rNTTxt2)
End If
Return

Spatieren:
iRNTxt = Len(rNTTxt)
rNTTxt2 = ""
For Teltxt = 1 To iRNTxt
    rNTTxt2 = rNTTxt2 + Mid(rNTTxt, Teltxt, 1) + " "
Next
Return

End Function


Private Sub PrintUserDef(TypeEnTaal As String)
Dim T               As Integer
Dim Tabulatie       As Integer
Dim FlFree          As Integer
Dim Teller          As Integer

ReDim psTekst(50) As String
ReDim psX(50) As Single
ReDim psY(50) As Single
ReDim psFontSize(50) As Single
ReDim psFontName(50) As String
ReDim psFontBold(50) As Integer
ReDim psFontItalic(50) As Integer
ReDim psFontUnderLine(50) As Integer
ReDim psColor(50) As Long
Dim MaxPslokatie        As Integer

Dim BoxTeller As Integer
ReDim Box(10) As Integer
ReDim BeginX(10) As Integer
ReDim BeginY(10) As Integer
ReDim TotX(10) As Integer
ReDim TotY(10) As Integer
Dim PsLokatie As Integer

On Local Error GoTo ErrorLoad

If Dir$(BedrijfsLokatie + "DDEF" + TypeEnTaal + ".Txt") = "" Then
    Beep
    Exit Sub
Else
    FlFree = FreeFile
    Open BedrijfsLokatie + "DDEF" + TypeEnTaal + ".Txt" For Input As FlFree
    Input #FlFree, Dummy
    Input #FlFree, VsoftVanaf
    Input #FlFree, VsoftTot
    For Teller = 0 To 10
        Input #FlFree, BeginX(Teller), BeginY(Teller), TotX(Teller), TotY(Teller), Box(Teller)
    Next
    Teller = 0
    While Not EOF(FlFree)
        Input #FlFree, psTekst(Teller), psX(Teller), psY(Teller), psFontSize(Teller), psFontName(Teller), psFontBold(Teller), psFontItalic(Teller), psFontUnderLine(Teller), psColor(Teller)
        Teller = Teller + 1
    Wend
    MaxPslokatie = Teller
    Close FlFree
End If

BeginXbox = BeginX(0)
BeginYbox = BeginY(0)

PsLokatie = 0
Do While PsLokatie <= MaxPslokatie
    If psTekst(PsLokatie) <> "" Then
        Printer.FontName = psFontName(PsLokatie)
        Printer.FontItalic = psFontItalic(PsLokatie)
        Printer.FontSize = psFontSize(PsLokatie)
        Printer.FontBold = psFontBold(PsLokatie)
        Printer.FontUnderline = psFontUnderLine(PsLokatie)
        Printer.ForeColor = psColor(PsLokatie)
    End If
    Printer.CurrentX = psX(PsLokatie)
    If psY(PsLokatie) > 10500 Then
        If Left(KeuzeInfo(0).Text, 1) <> "4" Then
            Printer.CurrentY = psY(PsLokatie) - OVSStrooklijnen
        Else
            Printer.CurrentY = psY(PsLokatie)
        End If
    Else
        Printer.CurrentY = psY(PsLokatie)
    End If
    Printer.Print psTekst(PsLokatie);
    PsLokatie = PsLokatie + 1
Loop
For T = 0 To 10
    If BeginX(T) = 0 Then
    ElseIf Box(T) Then
        Printer.Line (BeginX(T), BeginY(T))-(TotX(T), TotY(T)), , B
    Else
        'Printer.Line (BeginX(T), BeginY(T))-(TotX(T), TotY(T))
    End If
Next

Dim FiguurX As Integer
Dim FiguurY As Integer
Dim FiguurName As String
Dim FigBestandsnaam As String
Dim Hoogte As Variant
Dim Breedte As Variant

If Dir$(BedrijfsLokatie + "DDEF" + TypeEnTaal + "G.Txt") = "" Then
Else
    FlFree = FreeFile
    Open BedrijfsLokatie + "DDEF" + TypeEnTaal + "G.Txt" For Input As FlFree
    While Not EOF(FlFree)
        Input #FlFree, FiguurX, FiguurY, FiguurName
        If ScrLeesTekstBestand(Msg, BedrijfsLokatie & FiguurName & ".mfd") Then
            FigBestandsnaam = Left(Msg, InStr(Msg, vbTab) - 1)
            Mim.imgFiguur.Picture = LoadPicture(FigBestandsnaam)
            Msg = Mid(Msg, InStr(Msg, vbTab) + 1)
            Hoogte = Left(Msg, InStr(Msg, vbTab) - 1)
            Msg = Mid(Msg, InStr(Msg, vbTab) + 1)
            Breedte = Msg
            If FiguurY + Hoogte > 10500 Then
                Printer.PaintPicture Mim.imgFiguur.Picture, FiguurX, FiguurY - OVSStrooklijnen, Breedte, Hoogte
            Else
                Printer.PaintPicture Mim.imgFiguur.Picture, FiguurX, FiguurY, Breedte, Hoogte
            End If
        Else
            MsgBox "afdrukprobleemke figuurobject: " & FiguurName
        End If
    Wend
    Close FlFree
End If
Exit Sub

ErrorLoad:
MsgBox "Stop tijdens inladen dokumentdefinitie" + vbCrLf + vbCrLf + Error$
Exit Sub
Resume

End Sub



Private Sub cbPdfDetail_Click()

    If Me.cbPdfDetail.Value = vbChecked Then
        'Me.Detail.Value = vbUnchecked
        'Me.Detail.Enabled = False
        Me.cbPdfExport.Visible = True
    Else
        'Me.Detail.Visible = True
        'Me.Detail.Enabled = True
        Me.cbPdfExport.Visible = False
    End If
    
End Sub



Private Sub Form_Load()

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If
'If String99(Lees, 299) = "2" Then
'    Me.cbPdfDetail.Value = vbUnchecked
'Else
'    Me.cbPdfDetail.Value = vbChecked
'End If

PaginaTeller = 0
TaalTekst = ""

KeuzeInfo(0).AddItem "0: Kies bewerkingskode !"
KeuzeInfo(0).AddItem "1: Bericht van vervaldag"
KeuzeInfo(0).AddItem "2: Kontante Afrekening Premie"
KeuzeInfo(0).AddItem "3: Memo Vervaldagbericht"
KeuzeInfo(0).AddItem "4: Kwijting"
KeuzeInfo(0).AddItem "5: Verzekeringsaanbod"
KeuzeInfo(0).ListIndex = 0

KeuzeInfo(1).AddItem "1: Frans"
KeuzeInfo(1).AddItem "2: Nederlands"
KeuzeInfo(1).AddItem "3: Engels"
KeuzeInfo(1).AddItem "4: Duits"
KeuzeInfo(1).ListIndex = 1

InstallMij
KeuzeInfo(2).ListIndex = 0

PolisDetail.Cols = 7
With PolisDetail
    .Col = 0
    .Row = 0
    .Text = "Polis"
    .ColWidth(0) = 1215
    
    .Col = 1
    .Text = "Vervaldag"
    .ColWidth(1) = 1005
    
    .Col = 2
    .Text = "Premie"
    .ColWidth(2) = 1035
'    .ColAlignment(2) = 1
        
    .Col = 3
    .Text = "Taksen"
    .ColWidth(3) = 795
'    .ColAlignment(3) = 1
    
    .Col = 4
    .Text = "Klant (Naam 1)"
    .ColWidth(4) = 3945
    
    .Col = 5
    .Text = "Com.%"
    .ColWidth(5) = 645
    
    .Col = 6
    .Text = "TB2"
    .ColWidth(5) = 645
End With
KontaktPersoon = Val(String99(Lees, 201))
Schoon
'If String99(Lees, 299) = "2" Then
'    Me.cbPDF.Value = vbUnchecked
'Else
'    Me.cbPDF.Value = vbChecked
'End If

End Sub


Private Sub InstallPolissen()
Dim Zoekstring As String * 24
Dim Dummy As String * 30
Dim tb2Dummy As String
Dim PostOfBezoek As String
Dim X As Integer
Dim TaksEnKost As Currency
Dim comPercentage As Single

Select Case Post.Value
    Case 1
        PostOfBezoek = "1"
    Case Else
        PostOfBezoek = "2"
End Select

Zoekstring = Mid(PeriodeVanTot, 5, 2)
bGetOrGreater FlPolis, 3, Zoekstring
If Ktrl Or Left(KeyBuf(FlPolis), 2) > Left(Zoekstring, 2) Then
    Beep
    Exit Sub
Else
    Screen.MousePointer = vbHourglass
    KwijtingDrukken.Enabled = False
    Do
        XDoEvents = DoEvents
        PolisDetail.Refresh
        RecordToVeld FlPolis
        Select Case Left(vBibTekst(FlPolis, "#vs97 #"), 1)
            Case PostOfBezoek
                bGet FlKlant, 0, vBibTekst(FlPolis, "#A110 #")
                If Ktrl Then
                    Dummy = "KlantLink onmogelijk !!! Kontroleer !!!"
                Else
                    RecordToVeld FlKlant
                    Dummy = Trim$(vBibTekst(FlKlant, "#A100 #") & " " & vBibTekst(FlKlant, "#A101 #"))
                End If
                GoSub Checkdokument
                bGet FlAllerlei, 1, "25" + vSet(vBibTekst(FlPolis, "#A010 #"), 4) + vBibTekst(FlPolis, "#A000 #")
                comPercentage = CommissieCheck(rsMAR(FlPolis)("A010"), rsMAR(FlPolis)("A000"))
                TaksEnKost = 0
                If Ktrl Then
                Else
                    RecordToVeld FlAllerlei
                    If Val(vBibTekst(FlPolis, "#B010 #")) = Val(vBibTekst(FlAllerlei, "#B010 #")) Then
                        TaksEnKost = Val(vBibTekst(FlAllerlei, "#B011 #"))
                    End If
                End If
                If KeuzeInfo(2).ListIndex <> 0 Then
                    If Left(KeuzeInfo(2).Text, 4) = vBibTekst(FlPolis, "#A010 #") Then
                        PolisDetail.AddItem vBibTekst(FlPolis, "#A000 #") & vbTab & Mid(vBibTekst(FlPolis, "#v165 #"), 1, 2) & "/" & Mid(vBibTekst(FlPolis, "#v164 #"), 1, 2) & "/" & Mid(PeriodeVanTot, 1, 4) & vbTab & Format(Val(vBibTekst(FlPolis, "#B010 #")), "#,##0.00") & vbTab & Format(TaksEnKost, "#,##0.00") & vbTab & Dummy & vbTab & Format(comPercentage, "0.00"), PolisDetail.Rows - 1
                        SnelHelpPrint Dummy, blLogging
                    End If
                Else
                    PolisDetail.AddItem vBibTekst(FlPolis, "#A000 #") & vbTab & Mid(vBibTekst(FlPolis, "#v165 #"), 1, 2) & "/" & Mid(vBibTekst(FlPolis, "#v164 #"), 1, 2) & "/" & Mid(PeriodeVanTot, 1, 4) & vbTab & Format(Val(vBibTekst(FlPolis, "#B010 #")) + Val(vBibTekst(FlPolis, "#v403 #")), "#,##0.00") & vbTab & Format(TaksEnKost, "#,##0.00") & vbTab & Dummy & vbTab & Format(comPercentage, "0.00") & vbTab & tb2Dummy, PolisDetail.Rows - 1
                    SnelHelpPrint Dummy, blLogging
                End If
        End Select
        bNext FlPolis
        If Ktrl = 0 And Left(KeyBuf(FlPolis), 2) = Left(Zoekstring, 2) Then
        Else
            Screen.MousePointer = vbNormal
            KwijtingDrukken.Enabled = True
            Exit Do
        End If
    Loop
    If PolisDetail.Rows > 2 Then
        PolisDetail.Row = 2
    End If
End If
Exit Sub

Checkdokument:
Ktrl2 = 99
'bGet Fldokument, 1, vSet("K" + vBibTekst(FlPolis, "#A110 #"), 13)
bGet Fldokument, 2, vBibTekst(FlPolis, "#A000 #")
If Ktrl Then
    Dummy = "!" & Dummy
    Return
Else
    Do
        If vSet(KeyBuf(Fldokument), 12) = vSet(vBibTekst(FlPolis, "#A000 #"), 12) Then
            RecordToVeld Fldokument
            'If Left(vBibTekst(Fldokument, "#v035 #"), 6) = Left(PeriodeVanTot, 6) Then
            If Val(vBibTekst(Fldokument, "#B010 #")) = Val(vBibTekst(FlPolis, "#B010 #")) And Left(vBibTekst(Fldokument, "#v035 #"), 6) = Left(PeriodeVanTot, 6) Then
                Ktrl2 = 0
                Exit Do
            Else
                bNext Fldokument
                If Ktrl Then Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
End If
If Ktrl2 <> 0 Then
    Dummy = "!" & Dummy
    tb2Dummy = ""
Else
    tb2Dummy = rsMAR(Fldokument)("rvxmltb2")
End If
Return

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Mim.Report.IsOpen = True Then
        MsgBox "Sluit eerst het PDF venster a.u.b.", vbInformation
        Cancel = True
   End If

End Sub

Private Sub KeuzeInfo_Click(Index As Integer)
Dim T As Integer
Dim aa As String
Dim CRLFTeller As Integer

Select Case Index
    Case 0
        If Left(KeuzeInfo(0).Text, 1) = "1" Then
            'Detail.Visible = True
            'Detail.Value = vbUnchecked
            Post.Visible = True
            KeuzeInfo(2).Visible = True
        Else
            'Detail.Visible = False
            'Detail.Value = vbUnchecked
            Post.Visible = False
            KeuzeInfo(2).Visible = False
        End If
        If Left(KeuzeInfo(0).Text, 1) = "5" Then
            LabelInfo(5).Visible = True
            LabelInfo(6).Visible = True
            TekstInfo(4).Visible = True
            TekstInfo(5).Visible = True
            DetailVeld.Visible = True
        Else
            LabelInfo(5).Visible = False
            LabelInfo(6).Visible = False
            TekstInfo(4).Visible = False
            TekstInfo(5).Visible = False
            DetailVeld.Visible = False
        End If

        If KeuzeInfo(0).ListIndex = 0 Then
            PolisDetail.Enabled = False
            Afsluiten.Enabled = False
            TekstInfo(1).Text = ""
            TaalTekst = ""
        Else
            PolisDetail.Enabled = True
            Afsluiten.Enabled = True
            TaalTekst = Left(KeuzeInfo(1).Text, 1) + Left(KeuzeInfo(0).Text, 1)
            If Dir$(BedrijfsLokatie + "kwijt" + TaalTekst + ".DEF") = "" Then
                TekstInfo(1).Text = ""
            Else
                TempoFL = FreeFile
                TekstInfo(1).Text = ""
                Open BedrijfsLokatie + "kwijt" + TaalTekst + ".DEF" For Input As TempoFL
                    aa = ""
                    CRLFTeller = 0
                    Do While Not EOF(TempoFL)
                        Line Input #TempoFL, A$
                        aa = aa + A$ + vbCrLf
                        CRLFTeller = CRLFTeller + 1
                    Loop
                    TekstInfo(1).Text = Left(aa, Len(aa) - 2)
                    CRLFCaption.Caption = Dec$((CRLFTeller), "##0")
                Close TempoFL
            End If
        End If
    Case 1
        If TaalTekst = "  " Then
            Exit Sub
        End If
        TaalTekst = Left(KeuzeInfo(1).Text, 1) + Left(KeuzeInfo(0).Text, 1)
       
        If Dir$(BedrijfsLokatie + "kwijt" + TaalTekst + ".DEF") = "" Then
            TekstInfo(1).Text = ""
        Else
            TempoFL = FreeFile
            TekstInfo(1).Text = ""
            Open BedrijfsLokatie + "kwijt" + TaalTekst + ".DEF" For Input As TempoFL
                aa = ""
                CRLFTeller = 0
                Do While Not EOF(TempoFL)
                    Line Input #TempoFL, A$
                    aa = aa + A$ + vbCrLf
                    CRLFTeller = CRLFTeller + 1
                Loop
                TekstInfo(1).Text = Left(aa, Len(aa) - 2)
                CRLFCaption.Caption = Dec$((CRLFTeller), "##0")
            Close TempoFL
        End If
End Select

End Sub

Private Sub KeuzeInfo_LostFocus(Index As Integer)
Dim Msg2 As String

Select Case Index
    Case 0
        If KeuzeInfo(0).ListIndex = 1 And PolisDetail.Rows = 2 Then
            If Post.Value Then
                Msg2 = "(post)"
            Else
                Msg2 = "(bezoek)"
            End If
            Msg = "Kwijtingen " + Msg2 + " met vervaldag voor de" + vbCrLf + Mid(PeriodeVanTot, 5, 2) + "e maand inladen ?"
            Ktrl = MsgBox(Msg, 292)
            If Ktrl = vbYes Then
                InstallPolissen
                PolisDetail.SetFocus
            End If
        End If
End Select

End Sub

Private Sub MetOverschrijving_Click()

    If Me.MetOverschrijving.Value = vbChecked Then
        Me.cbIBANLayout.Value = vbChecked
        Me.cbIBANLayout.Enabled = True
        Me.cbIBANLayout.Visible = True
    Else
        Me.cbIBANLayout.Visible = False
    End If
    
End Sub

Private Sub PolisDetail_Click()

On Local Error Resume Next
With PolisDetail
    .Col = 0
    '.SelStartCol = 0
    '.SelEndCol = 4
    '.SelStartRow = .Row
    '.SelEndRow = .Row
End With


End Sub


Private Sub PolisDetail_GotFocus()

SnelHelpPrint "[+] lijn bijvoegen, [-] verwijderen, [Enter] om te wijzigen", blLogging

End Sub

Private Sub PolisDetail_KeyDown(KeyCode As Integer, Shift As Integer)

PolisDetail_Click

End Sub

Private Sub PolisDetail_KeyPress(KeyAscii As Integer)
Dim Positie As Integer
Dim TempoPremie As Long
Dim TempoLoon As Long

Positie = PolisDetail.Row
Select Case KeyAscii
    Case 13, 43
        If KeyAscii = 43 Then
            PolisDetail.AddItem vbTab & vbTab & vbTab & vbTab & vbTab, PolisDetail.Rows - 1
            PolisDetail.Row = PolisDetail.Rows - 2
        ElseIf KeyAscii = 13 Then
            If Positie = PolisDetail.Rows - 1 Then
                MsgBox "Druk [+] toets om bij te voegen !"
                Exit Sub
            End If
        End If
        Load KwijtingEdit
        KwijtingEdit.Caption = Left(KeuzeInfo(0).Text, 4) + ":" + KwijtingEdit.Caption
        KwijtingEdit.Show 1
        KwijtingDrukken.SetFocus
        If KeyAscii = 43 And GridText = "ESC" Then
            PolisDetail.RemoveItem PolisDetail.Rows - 1
        End If
                        
    Case 43
    
    Case 45, 127
        If PolisDetail.Rows = 2 Then
            Exit Sub
        ElseIf Positie < 0 Then
            MsgBox "Eerst een lijn selekteren !", 0, "Lijn wijzigen"
            Exit Sub
        End If
        On Local Error Resume Next
        PolisDetail.RemoveItem Positie
        If PolisDetail.Rows - 1 > Positie + 1 Then
            PolisDetail.Row = Positie - 1
        End If
        KwijtingDrukken!PolisDetail.SetFocus
        On Local Error GoTo 0
    Case Else
End Select

If PolisDetail.Rows = 2 Then
    Afsluiten.Enabled = False
Else
    Afsluiten.Enabled = True
End If

End Sub

Private Sub RasterSchoon()

With PolisDetail
    .Rows = 2
    .Row = 1
    .Col = 0
    .Text = ""
    
    .Col = 1
    .Text = ""
    
    .Col = 2
    .Text = ""
    
    .Col = 3
    .Text = ""
    
    .Col = 4
    .Text = ""
    
    .Col = 5
    .Text = ""
End With

End Sub

Private Sub Schoon()
Dim T As Integer

Afsluiten.Enabled = False
KeuzeInfo(2).Visible = False
KeuzeInfo(2).ListIndex = 0
TekstInfo(0).Text = Rdt
RasterSchoon
KeuzeInfo(0).ListIndex = 0
CRLFCaption.Caption = Dec$(0, "##0")
On Local Error Resume Next

End Sub

Private Sub PolisDetail_KeyUp(KeyCode As Integer, Shift As Integer)

'Select Case KeyCode
'    Case 37 To 40
        PolisDetail_Click
'End Select

End Sub


Private Sub SchoonVegen_Click()

RasterSchoon
Schoon

End Sub

Private Sub TekstBewaren_Click()

TempoFL = FreeFile
Open BedrijfsLokatie + "kwijt" + Left(KeuzeInfo(1).Text, 1) + Left(KeuzeInfo(0).Text, 1) + ".DEF" For Output As TempoFL
    Print #TempoFL, TekstInfo(1).Text
Close TempoFL

End Sub

Private Sub TekstInfo_KeyPress(Index As Integer, KeyAscii As Integer)
Dim Search As String
Dim Start As Integer
Dim CRLFTeller As Integer

Select Case Index
    Case 1
        Select Case KeyAscii
            Case 13
                CRLFCaption.Caption = Dec$(Val(CRLFCaption.Caption) + 1, "##0")
            Case 8
                Search = Chr$(13)
                CRLFTeller = 0
                Start = 1
                Do
                    Start = InStr(Start, TekstInfo(1).Text, Search)
                    If Start = 0 Then
                        Exit Do
                    Else
                        CRLFTeller = CRLFTeller + 1
                        Start = Start + 1
                    End If
                Loop
                CRLFCaption.Caption = Dec$((CRLFTeller + 1), "##0")
        End Select
End Select



End Sub

Private Sub TekstInfo_LostFocus(Index As Integer)

Select Case Index
    Case 0
        If DatumFout((TekstInfo(Index).Text)) Then
            TekstInfo(Index).Text = Rdt
            Beep
            TekstInfo(Index).SetFocus
        End If
    Case 2, 3
        TekstInfo(Index).Text = Format(Val(TekstInfo(Index).Text), "000.00")
    Case 4
        'TekstInfo(Index).Text = Dec$(Val(TekstInfo(Index).Text), "#######")
End Select

End Sub


Function CommissieCheck(Maatschappij As String, PolisNummer As String) As Single

Dim DummyText As String
Dim TempoX As Integer
Dim Ktrl99 As Integer

Dim NettoPremie As Currency
Dim Commissie As Currency

bGet FlAllerlei, 1, "25" + vSet(Maatschappij, 4) + PolisNummer
If Ktrl Then
    CommissieCheck = 0
Else
    RecordToVeld FlAllerlei
    NettoPremie = Val(vBibTekst(FlAllerlei, "#B013 #"))
    Commissie = Val(vBibTekst(FlAllerlei, "#B014 #"))
    If NettoPremie = 0 Then
        CommissieCheck = 0
    Else
        CommissieCheck = (Commissie / NettoPremie) * 100
    End If
End If

End Function


Function wordPREMIE(strMij As String, strTaal As String, strType As String, strInfo As String) As Boolean

    On Error Resume Next
 
    Dim TT As Integer
    Dim volNaam As String
    Dim Bestandsnaam As String
    Dim Book() As Variant
    ReDim Book(42, 2)
    Dim sSip() As String
    ReDim sSip(10)
    Dim NogString As String
    Dim Nog As Integer
    Dim rNTTxt As String
    Dim rNTTxt2 As String
    Dim hTAB As String
    Dim BedragTxt As String
    
    Book(1, 1) = Trim$(Mid(fmarBoxText("003", strTaal, oWaarde(rsMAR(FlKlant)("A102"))), 4, 10))
    Book(2, 1) = oWaarde(rsMAR(FlKlant)("A100"))
    Book(3, 1) = oWaarde(rsMAR(FlKlant)("A101"))
    If Trim$(xrsMar(FlKlant, "vs01")) = "" Then
        Book(4, 1) = ""
    Else
        Book(4, 1) = Trim$(Mid(fmarBoxText("003", strTaal, oWaarde(rsMAR(FlKlant)("vs01"))), 4, 10))
    End If
    Book(5, 1) = oWaarde(rsMAR(FlKlant)("A125"))
    Book(6, 1) = oWaarde(rsMAR(FlKlant)("A127"))
    Book(7, 1) = oWaarde(rsMAR(FlKlant)("A104"))
    Book(8, 1) = oWaarde(rsMAR(FlKlant)("A105"))
    If Trim$(oWaarde(rsMAR(FlKlant)("A106"))) = "" Then
        Book(9, 1) = ""
    Else
        Book(9, 1) = " " & oWaarde(rsMAR(FlKlant)("A106"))
    End If
    Book(10, 1) = Trim$(oWaarde(rsMAR(FlKlant)("A109")))
    Book(11, 1) = oWaarde(rsMAR(FlKlant)("A107"))
    Book(12, 1) = oWaarde(rsMAR(FlKlant)("A108"))
    
    Book(13, 1) = Trim$(oWaarde(rsMAR(FlPolis)("A000")))
    Book(14, 1) = oWaarde(rsMAR(FlLeverancier)("A100"))
    Book(15, 1) = Trim$(oWaarde(rsMAR(FlKlant)("A110")))
    Book(16, 1) = Format(TaksEnKost, "#,##0.00")
    Book(17, 1) = Format(BrutoPremie, "#,##0.00")
    
    On Local Error Resume Next
    Book(18, 1) = Format(Val(oWaarde(rsMAR(FlPolis)("e069"))), "#,##0.00")
    Book(19, 1) = IndexBM
    Book(20, 1) = "EUR " & Format(BedragEUR, "#,##0.00")
    Book(21, 1) = oWaarde(rsMAR(FlPolis)("vs98"))
    Book(22, 1) = oWaarde(rsMAR(FlPolis)("vs99"))
    Book(23, 1) = datKwijting
    Book(24, 1) = datKwijting
    Book(25, 1) = Mid(fmarBoxText("914", Val(KeuzeInfo(1)), oWaarde(rsMAR(FlPolis)("A325"))), 4)
    Book(26, 1) = Mid(ReferteTxt, 4, 14)
    Book(27, 1) = Mid(ReferteTxt, 4, 14)
        
    On Local Error GoTo 0
    
    If Dir(BedrijfsLokatie + "LicMarnt.###") = "" Then
        MsgBox BedrijfsLokatie + "LicMarnt.### niet te vinden !  Hierna wordt kladblok opgestart.  Breng uw eigen gegevens in a.u.b. !", vbExclamation
        On Local Error Resume Next
        X = Shell("notepad.exe " & BedrijfsLokatie & "LicMarnt.###", vbNormalFocus)
        Exit Function
    Else
        FlTemp = FreeFile
        Open BedrijfsLokatie & "LicMarnt.###" For Input As FlTemp
            Line Input #FlTemp, sSip(0)
            Line Input #FlTemp, sSip(1)
            Line Input #FlTemp, sSip(2)
            Line Input #FlTemp, sSip(3)
            Line Input #FlTemp, sSip(4)
            Line Input #FlTemp, sSip(5)
            Line Input #FlTemp, Msg
            Err = 0
            On Local Error Resume Next
            Line Input #FlTemp, NogString
            If Err Then
                Nog = 0
            Else
                Nog = Val(NogString)
            End If
        Close FlTemp
        hTAB = Space(Val(Msg))
    End If
    BedragTxt = "*" & Format(TotaalEUR, MaskerEUR) & "*"
    Book(28, 1) = hTAB & BedragTxt
    rNTTxt = Dec((TotaalEUR), "#######0.00")
    GoSub Spatieren
    Mid(rNTTxt2, 17, 1) = " "
    Book(29, 1) = hTAB & rNTTxt2
    
    rNTTxt = Mid(sSip(0), 1, 3) & Mid(sSip(0), 5, 7) & Mid(sSip(0), 13, 2)
        GoSub Spatieren
    rNTTxt2 = Mid(rNTTxt2, 1, 6) & " " & Mid(rNTTxt2, 7, 14) & " " & Mid(rNTTxt2, 21, 3)
    Book(30, 1) = hTAB & rNTTxt2
    
    For T = 0 To 4
        Book(31 + T, 1) = hTAB & sSip(T + 1)
    Next
    rNTTxt = ReferteTxt
        GoSub Spatieren
    Book(36, 1) = hTAB & ReferteTxt
    Book(37, 1) = hTAB & rNTTxt2
        
    Book(1, 2) = "Aan1"
    Book(2, 2) = "Naam1"
    Book(3, 2) = "Voornaam1"
    Book(4, 2) = "Aan2"
    Book(5, 2) = "Naam2"
    Book(6, 2) = "Voornaam2"
    Book(7, 2) = "Straat"
    Book(8, 2) = "Nr"
    Book(9, 2) = "Bus"
    Book(10, 2) = "LND"
    Book(11, 2) = "PC"
    Book(12, 2) = "Plaats"
        
    Book(13, 2) = "A000"
    Book(14, 2) = "A010"
    Book(15, 2) = "A110"
    Book(16, 2) = "B011"
    Book(17, 2) = "B013"
    Book(18, 2) = "e069"
    Book(19, 2) = "rv5315"
    Book(20, 2) = "v249"
    Book(21, 2) = "vs98"
    Book(22, 2) = "vs99"
    Book(23, 2) = "VVDAG"
    Book(24, 2) = "VVDAG2"
    Book(25, 2) = "A325"
    Book(26, 2) = "betaalreferte"
    Book(27, 2) = "betaalreferte2"
    Book(28, 2) = "bedragLinks"
    Book(29, 2) = "bedragRechts"
    Book(30, 2) = "onsRekNr"
    Book(31, 2) = "onsNaam1"
    Book(32, 2) = "onsNaam2"
    Book(33, 2) = "onsStraat"
    Book(34, 2) = "onsStraat2"
    Book(35, 2) = "onsPlaats"
    Book(36, 2) = "onsRefLinks"
    Book(37, 2) = "onsRefRechts"
    Book(38, 2) = "klantNaam1"
    Book(39, 2) = "klantNaam2"
    Book(40, 2) = "klantStraat"
    Book(41, 2) = "klantPlaats"
    Book(42, 2) = "detailTekst"
        
    Book(38, 1) = UCase(Left(Book(1, 1) & " " & Book(2, 1) & " " & Book(3, 1), 28))
    Book(39, 1) = UCase(Left(Book(4, 1) & " " & Book(5, 1) & " " & Book(6, 1), 28))
    Book(40, 1) = UCase(Left(Book(7, 1) & " " & Book(8, 1) & " " & Book(9, 1), 28))
    Book(41, 1) = UCase(Left(Book(11, 1) & " " & Book(12, 1), 28))
    Book(42, 1) = "" '"Testlijn1" & vbTab & "Testlijn2"
        
    volNaam = BedrijfsLokatie & "sjb\" & strMij & strTaal & "Termijn.doc"
    If Dir(volNaam) = "" Then
        volNaam = BedrijfsLokatie & "sjb\00002Termijn.doc"
    End If
    'wApp.Documents.Open (volNaam)
    'With wApp.Documents(1)
    '    For TT = 1 To 42
    '        .Bookmarks(Book(TT, 2)).Select
    '        .Bookmarks(Book(TT, 2)).Range.InsertAfter (Book(TT, 1))
    '    Next
    'End With
    DoEvents
    volNaam = BedrijfsLokatie & "Archief\" & Book(15, 1) & "$" & Format(Now, "yyyymmddhhmmss") & ".doc"
    'volNaam = BedrijfsLokatie & "Archief\" & Book(15, 1) & "$" & Format(Now, "yyyymmddhhmmss") & ".pdf"
    'wApp.Documents(1).PrintOut ' , , , volNaam & ".pdf", , , , , , "PDF"
    'wApp.Documents(1).PrintOut .PrintOut, , , volNaam & ".pdf", , , , , , "PDF"
    
    'Dim cnvWrdPdf As FileConverter
    
    'Look for WordPdf file converter
    'And save document using the converter
    'For the FileFormat converter value
    
    'For Each cnvWrdPdf In Application.FileConverters
     '   Debug.Print cnvWrdPdf.ClassName
      '  If Left(cnvWrdPdf.ClassName, 3) = "PDF" Then
       '     Stop
            'wApp.Documents(1).SaveAs volNaam ', cnvWrdPdf.SaveFormat
        'End If
    'Next cnvWrdPdf
    DoEvents
    SnelHelpPrint "Bezig aan: " & Trim$(Book(2, 1) & " " & Book(3, 1) & " " & Book(5, 1) & " " & Book(6, 1)), False
    'wApp.Documents(1).Close SaveChanges:=False
    DoEvents
    Exit Function
        
Spatieren:
iRNTxt = Len(rNTTxt)
rNTTxt2 = ""
For Teltxt = 1 To iRNTxt
    rNTTxt2 = rNTTxt2 + Mid(rNTTxt, Teltxt, 1) + " "
Next
Return
    
End Function


