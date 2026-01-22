VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form BalansKL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Balans"
   ClientHeight    =   2745
   ClientLeft      =   495
   ClientTop       =   2925
   ClientWidth     =   7245
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
   ScaleHeight     =   2745
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   300
      Width           =   4320
   End
   Begin VB.CheckBox chkAfdrukLiggend 
      Alignment       =   1  'Right Justify
      Caption         =   "Liggende Printerafdruk"
      Height          =   255
      Left            =   4890
      TabIndex        =   27
      Top             =   2190
      Width           =   2310
   End
   Begin VB.CheckBox chkAfdrukInVenster 
      Alignment       =   1  'Right Justify
      Caption         =   "Afdruk in venster"
      Height          =   255
      Left            =   5370
      TabIndex        =   26
      Top             =   2460
      Width           =   1830
   End
   Begin VB.ListBox BetalingenVoorNa 
      CausesValidation=   0   'False
      Height          =   255
      Index           =   3
      Left            =   2610
      TabIndex        =   25
      Top             =   330
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.ListBox BetalingenVoorNa 
      CausesValidation=   0   'False
      Height          =   255
      Index           =   2
      Left            =   2310
      TabIndex        =   24
      Top             =   300
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.ListBox BetalingenVoorNa 
      CausesValidation=   0   'False
      Height          =   255
      Index           =   1
      Left            =   2010
      TabIndex        =   23
      Top             =   300
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.ListBox BetalingenVoorNa 
      CausesValidation=   0   'False
      Height          =   255
      Index           =   0
      Left            =   1470
      TabIndex        =   22
      Top             =   330
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.CommandButton CmdStandaardBetaling 
      Caption         =   "Standaard Betalingskontrole"
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
      Left            =   1800
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1740
      Width           =   2415
   End
   Begin VB.CheckBox Selektie 
      Alignment       =   1  'Right Justify
      Caption         =   "Betalingsjournaal &weergeven"
      Height          =   240
      Index           =   5
      Left            =   4380
      TabIndex        =   13
      Top             =   1920
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.CommandButton CmdBewaar 
      Caption         =   "Bewaar instellingen"
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
      Left            =   1800
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1755
   End
   Begin VB.CommandButton CmdStandaard 
      Caption         =   "Standaard Boekhoudkontrole"
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
      Left            =   1800
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   60
      TabIndex        =   15
      Top             =   2160
      Width           =   3555
   End
   Begin VB.CheckBox Selektie 
      Alignment       =   1  'Right Justify
      Caption         =   "Dokumenten -1994 &uitsluiten"
      Height          =   240
      Index           =   4
      Left            =   4380
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1440
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   4440
      TabIndex        =   3
      Top             =   300
      Width           =   1260
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   60
      TabIndex        =   5
      Top             =   960
      Width           =   1605
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   60
      TabIndex        =   6
      Top             =   1260
      Width           =   1605
   End
   Begin VB.CheckBox Selektie 
      Alignment       =   1  'Right Justify
      Caption         =   "&Kontroleren op vervaldag"
      Height          =   240
      Index           =   0
      Left            =   3240
      TabIndex        =   8
      Top             =   660
      Width           =   2475
   End
   Begin VB.CheckBox Selektie 
      Alignment       =   1  'Right Justify
      Caption         =   "Betalingen &buiten periode uitsluiten"
      Height          =   240
      Index           =   1
      Left            =   3810
      TabIndex        =   9
      Top             =   960
      Value           =   1  'Checked
      Width           =   3345
   End
   Begin VB.CheckBox Selektie 
      Alignment       =   1  'Right Justify
      Caption         =   "Enkel dokum&enten deze periode"
      Height          =   195
      Index           =   2
      Left            =   540
      TabIndex        =   16
      Top             =   2460
      Value           =   1  'Checked
      Width           =   3075
   End
   Begin VB.CheckBox Selektie 
      Alignment       =   1  'Right Justify
      Caption         =   "&Niet betaalde dokumenten"
      Height          =   240
      Index           =   3
      Left            =   4590
      TabIndex        =   10
      Top             =   1200
      Value           =   1  'Checked
      Width           =   2565
   End
   Begin VB.CommandButton Drukken 
      Caption         =   "Af&druk"
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   60
      Width           =   1275
   End
   Begin VB.CommandButton Annuleren 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   420
      Width           =   1275
   End
   Begin Threed.SSCheck ChkFinancieelDetail 
      Height          =   195
      Left            =   4320
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2835
      _Version        =   65536
      _ExtentX        =   5001
      _ExtentY        =   344
      _StockProps     =   78
      Caption         =   "&Financiëel detail via journaal"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Value           =   -1  'True
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Periode van - tot"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   4
      Left            =   180
      TabIndex        =   14
      Top             =   1920
      Width           =   1530
   End
   Begin VB.Label Label1 
      Caption         =   "Sub&Titel"
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Datu&m"
      Height          =   240
      Index           =   1
      Left            =   5100
      TabIndex        =   2
      Top             =   60
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "&Van"
      Height          =   240
      Index           =   2
      Left            =   180
      TabIndex        =   4
      Top             =   690
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Tot"
      Height          =   240
      Index           =   3
      Left            =   720
      TabIndex        =   18
      Top             =   690
      Width           =   480
   End
End
Attribute VB_Name = "BalansKL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit

Dim psTekst(5) As String
Dim RapportVeld(23) As String
Dim RapportTab(23) As Integer

Dim RapportSelektie As String
Dim RapportTitel As String
Dim RapportDefinitie As String
Dim LijstNaam As String
Dim VeldTXT(20) As String
Dim FlPartij As Integer

Dim dttot As Double
Dim dtbtw As Double
Dim dvtot As Double
Dim dTrb As Double

Dim TotaalBTW As Double
Dim TotaalGOED As Double
Dim TotaalALBETAALD As Double
Dim TotaalNOGTEBETALEN As Double
Dim TotaalVOOR As Double
Dim TotaalNA          As Double

Dim PlGrensVan As String * 8
Dim PlGrensTot As String * 8

Dim BHDetailTekst(3) As String
Dim AantalBovenPeriode As Integer
Dim AantalOnderPeriode As Integer
Dim TotaalDokBovenPeriode As Double
Dim TotaalDokOnderPeriode As Double

Private Sub Annuleren_Click()

Unload Me

End Sub


Private Sub chkAfdrukLiggend_Click()

If chkAfdrukLiggend.Value = 1 Then
    Printer.Orientation = vbPRORLandscape
Else
    Printer.Orientation = vbPRORPortrait
End If
XDoEvents = DoEvents

End Sub

Private Sub ChkFinancieelDetail_Click(Value As Integer)

If ChkFinancieelDetail.Value = False Then
    'Selektie(1).Enabled = False
    Selektie(1).Value = False
    Selektie(5).Enabled = False
    Selektie(5).Value = False
Else
    'Selektie(1).Enabled = True
    Selektie(5).Enabled = True
End If

End Sub





Private Sub CmdBewaar_Click()

BeWaarTekst "DokumentenBalans", "KontroleVervaldag", Selektie(0)
BeWaarTekst "DokumentenBalans", "GeenBetalingHogerBoekjaar", Selektie(1)
BeWaarTekst "DokumentenBalans", "PeriodeBegrenzen", Selektie(2)
BeWaarTekst "DokumentenBalans", "EnkelNietBetaaldeDokumenten", Selektie(3)
BeWaarTekst "DokumentenBalans", "-Dokumenten-1994Uitsluiten", Selektie(4)
BeWaarTekst "DokumentenBalans", "BetalingsKontrole", Selektie(5)
BeWaarTekst "DokumentenBalans", "FinancieelDetailViaJournaal", ChkFinancieelDetail
BeWaarTekst "DokumentenBalans", "AfdrukInRooster", chkAfdrukInVenster

End Sub


Private Sub CmdStandaard_Click()

Selektie(0) = 0
PlGrensVan = Left$(BoekjaarVanTot, 8)
PlGrensTot = Right$(BoekjaarVanTot, 8)
TekstLijn(4).Text = fDatumText(Left(BoekjaarVanTot, 8)) + " - " + fDatumText(Right(BoekjaarVanTot, 8))
Selektie(2) = 1

Selektie(1) = 1
Selektie(3) = 1
ChkFinancieelDetail = True
Selektie(5) = 1

End Sub

Private Sub CmdStandaardBetaling_Click()

Selektie(0) = 1
Selektie(2) = 0

Selektie(1) = 0
Selektie(3) = 1
ChkFinancieelDetail = False
TekstLijn(1) = Rdt

End Sub

Private Sub Drukken_Click()
Dim BeginSleutel As String * 13
Dim EindSleutel As String * 13
Dim rdtemp As String * 8
Dim syMasker As String
Dim KopBuf As String
Dim XX As Integer
Dim Tel As Integer
Dim Teller As Integer
Dim ReedsMetBetalingen As Boolean
Dim TDS As String

Dim BetaaldBedrag As Double
Dim dBtw As Double
Dim dTot As Double
Dim dnt As Double
Dim dTnt As Double
Dim drb  As Double
Dim drbet As Double

Dim LengteSleutel As Integer
Dim sleutelIndex As Integer
Dim xOmgeving As Integer

TotaalBTW = 0
TotaalGOED = 0
TotaalALBETAALD = 0
TotaalNOGTEBETALEN = 0
TotaalVOOR = 0
TotaalNA = 0

dttot = 0
dtbtw = 0
dvtot = 0
dTrb = 0
dTnt = 0

BetalingenVoorNa(2).Clear
BetalingenVoorNa(3).Clear

Select Case FlPartij
    Case FlLeverancier
        BeginSleutel = "L" + TekstLijn(2).Text
        EindSleutel = "L" + TekstLijn(3).Text
    Case FlKlant
        BeginSleutel = "K" + TekstLijn(2).Text
        EindSleutel = "K" + TekstLijn(3).Text
End Select

rdtemp = DatumKey((TekstLijn(1).Text))

psTekst(2) = BalansKL.Caption + " " + Mid(Mim.Caption, InStr(Mim.Caption, "["))
psTekst(0) = TekstLijn(1).Text
psTekst(3) = TekstLijn(0).Text
    InitVelden

TDS = "Geen journalen voor : " + vbCrLf
bGetOrGreater FlDokument, 1, BeginSleutel
If Ktrl Or UCase$(vSet(KeyBuf(FlDokument), 13)) > UCase$(EindSleutel) Then
    Beep
    Exit Sub
ElseIf Left$(KeyBuf(FlDokument), 1) <> Left$(BeginSleutel, 1) Then
    Beep
    Exit Sub
Else
    Screen.MousePointer = 11
    BalansKL.Enabled = False
    PaginaTeller = 0
    If chkAfdrukInVenster = 0 Then
        Set Printer = Printers(LijstPrinterNr)
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
End If

Do
    XDoEvents = DoEvents
    RecordToVeld FlDokument
    GoSub KontroleVoorwaarden
    bNext FlDokument
    If Ktrl Or UCase$(vSet(KeyBuf(FlDokument), 13)) > UCase$(EindSleutel) Then
        Exit Do
    End If
Loop
    TotaalNOGTEBETALEN = TotaalGOED + TotaalBTW - TotaalALBETAALD
    EindTotaal
    Screen.MousePointer = 0
    BalansKL.Enabled = True
    If chkAfdrukInVenster Then
    Else
        Printer.NewPage
        Printer.EndDoc
    End If
    If Len(TDS) = 24 Then
    Else
        TDS = TDS + vbCrLf
        TDS = TDS + "De betalingen voor bovenvermelde dokumenten kunnen dus niet" + vbCrLf
        TDS = TDS + "in detail weergegeven worden.  Enkel het algemeen totaal" + vbCrLf
        TDS = TDS + "van het dokument en laatste financieel uittreksel..."
        MsgBox TDS
        Msg = "Mededeling eveneens op papier ?"
        KtrlBox = MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton2)
        If KtrlBox = vbYes Then
            Printer.Print TDS
            Printer.NewPage
            Printer.EndDoc
        End If
    End If
 If BetalingenVoorNa(0).ListCount + BetalingenVoorNa(1).ListCount + BetalingenVoorNa(2).ListCount + BetalingenVoorNa(3).ListCount <> 0 Then
        Msg = "Betalingsdetail dokumenten lagere/hogere boekjaren..." + vbCrLf
        Msg = Msg + "Op papier ?"
        KtrlBox = MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton1)
        If KtrlBox = vbYes Then
            BHDetailTekst(0) = "Detail betalingen dokumenten van lagere boekjaren in dit boekjaar" + vbCrLf + vbCrLf
            BHDetailTekst(1) = "Detail betalingen dokumenten van hogere boekjaren in dit boekjaar" + vbCrLf + vbCrLf
            BHDetailTekst(2) = "Detail betalingen dokumenten van dit boekjaar in lagere boekjaren" + vbCrLf + vbCrLf
            BHDetailTekst(3) = "Detail betalingen dokumenten van dit boekjaar in hogere boekjaren" + vbCrLf + vbCrLf
            For Tel = 0 To 3
                If BetalingenVoorNa(Tel).ListCount <> 0 Then
                    Printer.Print BHDetailTekst(Tel)
                    For Teller = 0 To BetalingenVoorNa(Tel).ListCount - 1
                        Printer.Print BetalingenVoorNa(Tel).List(Teller); vbCrLf;
                    Next
                    Printer.Print vbCrLf; vbCrLf;
                    If Tel = 0 Then
                        Printer.Print "Totaal financiële bewegingen: " + Format(TotaalDokOnderPeriode, Masker2002)
                    ElseIf Tel = 1 Then
                        Printer.Print "Totaal financiële bewegingen: " + Format(TotaalDokBovenPeriode, Masker2002)
                    ElseIf Tel = 2 Then
                        Printer.Print "Totaal financiële bewegingen: " + Format(TotaalVOOR, Masker2002)
                    ElseIf Tel = 3 Then
                        Printer.Print "Totaal financiële bewegingen: " + Format(TotaalNA, Masker2002)
                    End If
                    Printer.NewPage
                End If
            Next
            Printer.EndDoc
        End If
    End If
    SnelHelpPrint "Klaar"
    BalansKL!TekstLijn(0).SetFocus
Exit Sub

KontroleVoorwaarden:
If Selektie(0).Value = 1 Then
    'vervaldag
    If vBibTekst(FlDokument, "#v036 #") > rdtemp Then
        Return
    End If
End If

If Selektie(4).Value = 1 Then
    '-1994 dokumenten uitsluiten
    If Left$(vBibTekst(FlDokument, "#v035 #"), 4) < "1994" Then
        Return
    End If
End If

If Selektie(2).Value = 1 Then
    'dokumenten enkel van deze periode
    If vBibTekst(FlDokument, "#v035 #") < PlGrensVan Or vBibTekst(FlDokument, "#v035 #") > PlGrensTot Then
        Return
    End If
End If

Select Case aIndex
    Case FlKlant
        Select Case Left$(vBibTekst(FlDokument, "#v033 #"), 1)
            Case "V"
                dBtw = Val(vBibTekst(FlDokument, "#v064 #"))
                dTot = 0
                For Teller = 55 To 63
                    dTot = dTot + Val(vBibTekst(FlDokument, "#v" + Format$(Teller, "000") + " #"))
                Next
                dTot = dTot + Val(vBibTekst(FlDokument, "#v089 #")) + dBtw
            Case "Q"
                dTot = Val(vBibTekst(FlDokument, "#B010 #")) - Val(vBibTekst(FlDokument, "#B090 #")) - Val(vBibTekst(FlDokument, "#v065 #"))
                dBtw = 0
        End Select
    Case FlLeverancier
        dBtw = Val(vBibTekst(FlDokument, "#v045 #")) - Val(vBibTekst(FlDokument, "#v044 #")) - Val(vBibTekst(FlDokument, "#v043 #"))
        If Left$(vBibTekst(FlDokument, "#v035 #"), 4) > "1992" Then
            dBtw = dBtw - Val(vBibTekst(FlDokument, "#v042 #"))
        End If
        dTot = 0
        For Teller = 46 To 49
            dTot = dTot + Val(vBibTekst(FlDokument, "#v" + Format$(Teller, "000") + " #"))
        Next
        dTot = dTot + dBtw
End Select

syMasker = MaskerEUR
If Mid$(vBibTekst(FlDokument, "#v033 #"), 2, 1) = "1" Then
    dTot = -dTot
    dBtw = -dBtw
    drb = -Val(vBibTekst(FlDokument, "#v037 #"))
Else
    drb = Val(vBibTekst(FlDokument, "#v037 #"))
End If

VeldTXT(1) = vBibTekst(FlDokument, "#v033 #")
VeldTXT(2) = fDatumText(vBibTekst(FlDokument, "#v035 #"))
VeldTXT(3) = vBibTekst(FlDokument, "#vs03 #")
VeldTXT(4) = Dec$((dTot - dBtw) / Val(vBibTekst(FlDokument, "#v040 #")), syMasker)
VeldTXT(5) = Dec$((dBtw), Masker2002)
VeldTXT(9) = fDatumText(vBibTekst(FlDokument, "#v036 #"))

If vBibTekst(FlDokument, "#v034 #") <> KopBuf Then
    dTnt = 0
    bGet aIndex, 0, vSet(Mid$(vBibTekst(FlDokument, "#v034 #"), 2, 12), 12)
    KopBuf = vBibTekst(FlDokument, "#v034 #")
    VeldTXT(0) = RTrim(Mid(vBibTekst(FlDokument, "#v034 #"), 2, 12))
    If Ktrl Then
        VeldTXT(0) = VeldTXT(0) + " * niet meer aanwezig *"
    Else
        RecordToVeld aIndex
        VeldTXT(0) = Left(VeldTXT(0) + " " + RTrim(vBibTekst(aIndex, "#A100 #")) + " " + RTrim(vBibTekst(aIndex, "#A101 #")), 29)
    End If
    SnelHelpPrint VeldTXT(0)
End If

bGet FlJournaal, 1, vBibTekst(FlDokument, "#v033 #")
If Ktrl Or Not ChkFinancieelDetail Then
    'Er zijn geen journalen voor dit dokument !
    'Verwittigen via afdruk laatste blz.
    If ChkFinancieelDetail Then
        TDS = TDS + vBibTekst(FlDokument, "#v033 #") + "  ...  " + vBibTekst(FlPartij, "#A110 #") + " " + vBibTekst(FlPartij, "#A100 #") + vbCrLf
    End If
    If dTot = drb And Selektie(3).Value = 1 Then
    Else
        If InStr(VeldTXT(0), Mid(vBibTekst(FlDokument, "#v034 #"), 2)) = 1 Then
            If chkAfdrukInVenster = 0 Then
                Printer.Print vbCrLf;
                If Printer.CurrentY >= Printer.Height - 1200 Then
                    On Local Error Resume Next
                    Printer.NewPage
                    Printer.FontSize = Printer.FontSize
                    Printer.Print " "
                    Printer.FontSize = Printer.FontSize
                    PrintTitel
                End If
            End If
        End If
        
        dTnt = dTnt + dTot - drb 'Algemeen Totaal Cumul te betalen
        VeldTXT(6) = Dec$((drb), Masker2002)
        VeldTXT(7) = vBibTekst(FlDokument, "#v038 #")
        VeldTXT(8) = Dec$((dTot - drb), Masker2002)
        VeldTXT(10) = Dec$((dTnt), Masker2002)
    
        TotaalBTW = TotaalBTW + dBtw
        TotaalGOED = TotaalGOED + dTot - dBtw
        TotaalALBETAALD = TotaalALBETAALD + drb
        PrintVelden
        VeldTXT(0) = ""
    End If
Else
    GoSub AfdrukDetailReedsBetaald
End If
Return

AfdrukDetailReedsBetaald:
GoSub TotaalBETAALD
If dTot = drb And Selektie(3).Value = 1 Then
    Return
Else
    bGet FlJournaal, 1, vBibTekst(FlDokument, "#v033 #")
End If
dTnt = dTnt + dTot
ReedsMetBetalingen = False

drb = 0
Do
    RecordToVeld FlJournaal
    If RTrim(vBibTekst(FlJournaal, "#v038 #")) <> "" Then
        If Left(vBibTekst(FlJournaal, "#v019 #"), 1) <> "4" Then
        Else
            ReedsMetBetalingen = True
            GoSub BetalingErBij
        End If
    End If
    bNext FlJournaal
    If Ktrl Or KeyBuf(FlJournaal) <> vBibTekst(FlDokument, "#v033 #") Then
        Exit Do
    End If
Loop
If Not ReedsMetBetalingen Then
    drb = 0
    vBib FlJournaal, "0", "v068"
    GoSub BetalingErBij
End If
Return

BetalingErBij:
BetaaldBedrag = Val(vBibTekst(FlJournaal, "#v068 #"))
If FlPartij = FlKlant Then
    BetaaldBedrag = -BetaaldBedrag
End If
If Selektie(1).Value = 1 Then
    'Betalingen buiten periode uitsluiten
    If vBibTekst(FlJournaal, "#v066 #") < PlGrensVan Then
        TotaalVOOR = TotaalVOOR + BetaaldBedrag
        Msg = vSet(Mid(vBibTekst(FlDokument, "#v034 #"), 2), 12) + " "
        Msg = Msg + vSet(vBibTekst(FlJournaal, "#v067 #"), 20) + " "
        Msg = Msg + fDatumText(vBibTekst(FlDokument, "#v035 #")) + " "
        Msg = Msg + vBibTekst(FlDokument, "#v033 #") + " "
        Msg = Msg + Dec(BetaaldBedrag, Masker2002) + " "
        Msg = Msg + vSet(vBibTekst(FlJournaal, "#v019 #"), 7) + " "
        Msg = Msg + fDatumText(vBibTekst(FlJournaal, "#v066 #")) + " "
        Msg = Msg + vBibTekst(FlJournaal, "#v038 #") + " "
        Msg = Msg + vBibTekst(FlJournaal, "#v069 #")
        BetalingenVoorNa(2).AddItem Msg
        BetaaldBedrag = 0
    ElseIf vBibTekst(FlJournaal, "#v066 #") > PlGrensTot Then
        TotaalNA = TotaalNA + BetaaldBedrag
        Msg = vSet(Mid(vBibTekst(FlDokument, "#v034 #"), 2), 12) + " "
        Msg = Msg + vSet(vBibTekst(FlJournaal, "#v067 #"), 20) + " "
        Msg = Msg + fDatumText(vBibTekst(FlDokument, "#v035 #")) + " "
        Msg = Msg + vBibTekst(FlDokument, "#v033 #") + " "
        Msg = Msg + Dec(BetaaldBedrag, Masker2002) + " "
        Msg = Msg + vSet(vBibTekst(FlJournaal, "#v019 #"), 7) + " "
        Msg = Msg + fDatumText(vBibTekst(FlJournaal, "#v066 #")) + " "
        Msg = Msg + vBibTekst(FlJournaal, "#v038 #") + " "
        Msg = Msg + vBibTekst(FlJournaal, "#v069 #")
        BetalingenVoorNa(3).AddItem Msg
        BetaaldBedrag = 0
    End If
End If

If InStr(VeldTXT(0), Mid(vBibTekst(FlDokument, "#v034 #"), 2)) = 1 Then
    dTnt = dTot: drb = 0
    If chkAfdrukInVenster = 0 Then
        Printer.Print vbCrLf;
        If Printer.CurrentY >= Printer.Height - 1200 Then
            On Local Error Resume Next
            Printer.NewPage
            Printer.FontSize = Printer.FontSize
            Printer.Print " "
            Printer.FontSize = Printer.FontSize
            PrintTitel
        End If
    End If
End If

drb = drb + BetaaldBedrag
dTnt = dTnt - BetaaldBedrag
dTrb = dTrb + drb           'Totaal al betaald
dttot = dttot + dTot - dBtw 'Totaal excl. btw
dtbtw = dtbtw + dBtw        'Totaal btw aftrekbaar

VeldTXT(6) = Dec$(BetaaldBedrag, Masker2002)
VeldTXT(7) = vBibTekst(FlJournaal, "#v038 #")
VeldTXT(8) = Dec$((dTot - drb), Masker2002)
VeldTXT(10) = Dec$((dTnt), Masker2002)

If Val(VeldTXT(4)) + Val(VeldTXT(5)) <> 0 Then
    TotaalBTW = TotaalBTW + dBtw
    TotaalGOED = TotaalGOED + dTot - dBtw
    TotaalNOGTEBETALEN = TotaalNOGTEBETALEN + dTot
End If
TotaalALBETAALD = TotaalALBETAALD + BetaaldBedrag
PrintVelden

VeldTXT(0) = ""
VeldTXT(1) = ""
VeldTXT(2) = ""
VeldTXT(3) = ""
VeldTXT(4) = ""
VeldTXT(5) = ""
VeldTXT(9) = ""
Return

TotaalBETAALD:
drb = 0
bGet FlJournaal, 1, vBibTekst(FlDokument, "#v033 #")
If Ktrl Then
    MsgBox "onlogische situatie"
Else
    Do
        RecordToVeld FlJournaal
        If RTrim(vBibTekst(FlJournaal, "#v038 #")) <> "" Then
            If Left(vBibTekst(FlJournaal, "#v019 #"), 1) <> "4" Then
            Else
                BetaaldBedrag = Val(vBibTekst(FlJournaal, "#v068 #"))
                If FlPartij = FlKlant Then
                    BetaaldBedrag = -BetaaldBedrag
                End If
                If Selektie(1).Value = 1 Then
                    'Betalingen buiten periode uitsluiten
                    If vBibTekst(FlJournaal, "#v066 #") < PlGrensVan Or vBibTekst(FlJournaal, "#v066 #") > PlGrensTot Then
                    Else
                        drb = drb + BetaaldBedrag
                    End If
                Else
                    drb = drb + BetaaldBedrag
                End If
            End If
        'ElseIf FlPartij = FlKlant Then
        '    If vBibTekst(FlJournaal, "#v033 #") = "Q0800286" Then
        '        Stop
        '    End If
        End If
        bNext FlJournaal
        If Ktrl Or KeyBuf(FlJournaal) <> vBibTekst(FlDokument, "#v033 #") Then
            Exit Do
        End If
    Loop
End If
Return

PrtHandler:
MsgBox "Kontroleer de printer."
Resume

End Sub

Private Sub EindTotaal()
Dim T As Integer
Dim aa As String

JumpEindTotaal:
For T = 0 To 10:
    VeldTXT(T) = ""
Next
VeldTXT(0) = "Totalen :"
VeldTXT(4) = Dec$((TotaalGOED), Masker2002)
VeldTXT(5) = Dec$((TotaalBTW), Masker2002)
VeldTXT(6) = Dec$((TotaalALBETAALD), Masker2002)
VeldTXT(10) = Dec$((TotaalNOGTEBETALEN), Masker2002)
If chkAfdrukInVenster = 0 Then Printer.Print vbCrLf; Lijntje; vbCrLf;

T = 0
aa = ""
Do While T < 12
    If chkAfdrukInVenster Then
        aa = aa + VeldTXT(T) + vbTab
    Else
        Printer.Print ; Tab(RapportTab(T));
        Printer.Print ; VeldTXT(T);
    End If
    If RapportTab(T + 1) < RapportTab(T) Then
        If chkAfdrukInVenster Then
        Else
            Printer.Print vbCrLf;
        End If
    End If
    T = T + 1
Loop
If chkAfdrukInVenster Then Xlog.X.AddItem aa, Xlog.X.Rows - 1

Dim GroepSelektie As String
Dim TotaalBedragGroep As Double
Dim AantalInGroep As Integer
Dim BedragZonderDokument As Double
Dim AantalZonderDokument As Integer
Dim BetaaldBedragBB As Double
Dim GroepRekening4 As String * 7

If Selektie(5) = 1 Then
    If chkAfdrukInVenster Then
    Else
        If Printer.CurrentY >= Printer.Height - 1500 Then
            On Local Error Resume Next
            Printer.NewPage
            Printer.FontSize = Printer.FontSize
            Printer.Print " "
            Printer.FontSize = Printer.FontSize
            PrintTitel
            GoTo JumpEindTotaal
        End If
    End If
    
    If TotaalVOOR <> 0 Then
        For T = 0 To 10:
            VeldTXT(T) = ""
        Next
        VeldTXT(0) = "Dok. reeds betaald voor :"
        VeldTXT(2) = Left(TekstLijn(4), 10)
        VeldTXT(6) = Dec$((TotaalVOOR), Masker2002)
        T = 0
        aa = ""
        Do While T < 12
            If chkAfdrukInVenster Then
                aa = aa + VeldTXT(T) + vbTab
            Else
                Printer.Print ; Tab(RapportTab(T));
                Printer.Print ; VeldTXT(T);
            End If
            If RapportTab(T + 1) < RapportTab(T) Then
                If chkAfdrukInVenster Then
                Else
                    Printer.Print vbCrLf;
                End If
            End If
            T = T + 1
        Loop
        If chkAfdrukInVenster Then Xlog.X.AddItem aa, Xlog.X.Rows - 1
    End If
    If TotaalNA <> 0 Then
        For T = 0 To 10:
            VeldTXT(T) = ""
        Next
        VeldTXT(0) = "Dok. reeds betaald na :"
        VeldTXT(2) = Right(TekstLijn(4), 10)
        VeldTXT(6) = Dec$((TotaalNA), Masker2002)
        T = 0
        aa = ""
        Do While T < 12
            If chkAfdrukInVenster Then
                aa = aa + VeldTXT(T) + vbTab
            Else
                Printer.Print ; Tab(RapportTab(T));
                Printer.Print ; VeldTXT(T);
            End If
            If RapportTab(T + 1) < RapportTab(T) Then
                If chkAfdrukInVenster Then
                Else
                    Printer.Print vbCrLf;
                End If
            End If
            T = T + 1
        Loop
        If chkAfdrukInVenster Then Xlog.X.AddItem aa, Xlog.X.Rows - 1
    End If

    If chkAfdrukInVenster = 0 Then Printer.Print vbCrLf; Lijntje; vbCrLf;
    TotaalBedragGroep = 0
    AantalInGroep = 0
    BedragZonderDokument = 0
    AantalZonderDokument = 0
    AantalBovenPeriode = 0
    AantalOnderPeriode = 0
    TotaalDokBovenPeriode = 0
    TotaalDokOnderPeriode = 0
    
    GroepSelektie = RTrim(String99(Lees, FlPartij + 296))
    If RTrim(GroepSelektie) = "" Then
        GroepSelektie = Left(String99(Lees, FlPartij + 8), 4)
    End If
    If Len(GroepSelektie) > 7 Then
        GroepSelektie = Left(GroepSelektie, 7)
    End If
    BetalingenVoorNa(0).Clear
    BetalingenVoorNa(1).Clear
    bGetOrGreater FlRekening, 0, GroepSelektie
    If Ktrl Then
        MsgBox "onlogika"
    Else
        RecordToVeld FlRekening
        Do While InStr(KeyBuf(FlRekening), GroepSelektie)
            GroepRekening4 = KeyBuf(FlRekening)
            SnelHelpPrint "Journalen boekjaar voor rek. " + KeyBuf(FlRekening) + " worden gekontroleerd.  Ogenblik a.u.b."
            AantalInGroep = AantalInGroep + 1
            TotaalBedragGroep = TotaalBedragGroep + Val(vBibTekst(FlRekening, "#v" + Format(22 + AktiefBoekjaar, "000") + " #"))
            bGetOrGreater FlJournaal, 0, vSet(KeyBuf(FlRekening), 7) + PlGrensVan
            If Ktrl Then
                MsgBox "Geen journalen voor deze periode..."
            Else
                RecordToVeld FlJournaal
                Do While Right(KeyBuf(FlJournaal), 8) <= PlGrensTot
                    SnelHelpPrint "Alle journalen voor rekening " + KeyBuf(FlRekening) + " worden gekontroleerd.  Bezig aan :" + Right(KeyBuf(FlJournaal), 8)
                    XDoEvents = DoEvents
                    If RTrim(vBibTekst(FlJournaal, "#v033 #")) = "" Or Left(vBibTekst(FlJournaal, "#v033 #"), 1) = "D" Then
                        AantalZonderDokument = AantalZonderDokument + 1
                        BedragZonderDokument = BedragZonderDokument + Val(vBibTekst(FlJournaal, "#v068 #"))
                    ElseIf RTrim(vBibTekst(FlJournaal, "#v038 #")) <> "" Then
                        bGet FlDokument, 0, vBibTekst(FlJournaal, "#v033 #")
                        If Ktrl Then
                        Else
                            RecordToVeld FlDokument
                            If vBibTekst(FlDokument, "#v035 #") < PlGrensVan Then
                                BetaaldBedragBB = Val(vBibTekst(FlJournaal, "#v068 #"))
                                If FlPartij = FlKlant Then
                                    BetaaldBedragBB = -BetaaldBedragBB
                                End If
                                TotaalDokOnderPeriode = TotaalDokOnderPeriode + BetaaldBedragBB
                                AantalOnderPeriode = AantalOnderPeriode + 1
                                Msg = vSet(Mid(vBibTekst(FlDokument, "#v034 #"), 2), 12) + " "
                                Msg = Msg + vSet(vBibTekst(FlJournaal, "#v067 #"), 20) + " "
                                Msg = Msg + fDatumText(vBibTekst(FlDokument, "#v035 #")) + " "
                                Msg = Msg + vBibTekst(FlDokument, "#v033 #") + " "
                                Msg = Msg + Dec(BetaaldBedragBB, Masker2002) + " "
                                Msg = Msg + vSet(vBibTekst(FlJournaal, "#v019 #"), 7) + " "
                                Msg = Msg + fDatumText(vBibTekst(FlJournaal, "#v066 #")) + " "
                                Msg = Msg + vBibTekst(FlJournaal, "#v038 #") + " "
                                Msg = Msg + vBibTekst(FlJournaal, "#v069 #")
                                BetalingenVoorNa(0).AddItem Msg
                            ElseIf vBibTekst(FlDokument, "#v035 #") > PlGrensTot Then
                                BetaaldBedragBB = Val(vBibTekst(FlJournaal, "#v068 #"))
                                If FlPartij = FlKlant Then
                                    BetaaldBedragBB = -BetaaldBedragBB
                                End If
                                TotaalDokBovenPeriode = TotaalDokBovenPeriode + BetaaldBedragBB
                                AantalBovenPeriode = AantalBovenPeriode + 1
                                Msg = vSet(Mid(vBibTekst(FlDokument, "#v034 #"), 2), 12) + " "
                                Msg = Msg + vSet(vBibTekst(FlJournaal, "#v067 #"), 20) + " "
                                Msg = Msg + fDatumText(vBibTekst(FlDokument, "#v035 #")) + " "
                                Msg = Msg + vBibTekst(FlDokument, "#v033 #") + " "
                                Msg = Msg + Dec(BetaaldBedragBB, Masker2002) + " "
                                Msg = Msg + vSet(vBibTekst(FlJournaal, "#v019 #"), 7) + " "
                                Msg = Msg + fDatumText(vBibTekst(FlJournaal, "#v066 #")) + " "
                                Msg = Msg + vBibTekst(FlJournaal, "#v038 #") + " "
                                Msg = Msg + vBibTekst(FlJournaal, "#v069 #")
                                BetalingenVoorNa(1).AddItem Msg
                            End If
                        End If
                    End If
                    bNext FlJournaal
                    If Ktrl Or KeyBuf(FlJournaal) > GroepRekening4 + PlGrensTot Then
                        Exit Do
                    Else
                        RecordToVeld FlJournaal
                    End If
                Loop
            End If
            bNext FlRekening
            If Ktrl Or InStr(KeyBuf(FlRekening), GroepSelektie) = 0 Then
                Exit Do
            Else
                RecordToVeld FlRekening
            End If
        Loop
    End If

    For T = 0 To 10:
        VeldTXT(T) = ""
    Next
    VeldTXT(0) = "Stand van " + Format(AantalInGroep, "00") + " " + GroepSelektie + "-rekeningen"
    VeldTXT(6) = Dec$((TotaalBedragGroep), Masker2002)
    T = 0
    aa = ""
    Do While T < 12
        If chkAfdrukInVenster Then
            aa = aa + VeldTXT(T) + vbTab
        Else
            Printer.Print ; Tab(RapportTab(T));
            Printer.Print ; VeldTXT(T);
        End If
        If RapportTab(T + 1) < RapportTab(T) Then
            If chkAfdrukInVenster Then
            Else
                Printer.Print vbCrLf;
            End If
        End If
        T = T + 1
    Loop
    If chkAfdrukInVenster Then Xlog.X.AddItem aa, Xlog.X.Rows - 1
    
    For T = 0 To 10:
        VeldTXT(T) = ""
    Next
    VeldTXT(0) = Format(AantalZonderDokument, "00") + " verr. zonder dokument"
    VeldTXT(6) = Dec$((BedragZonderDokument), Masker2002)
    T = 0
    aa = ""
    Do While T < 12
        If chkAfdrukInVenster Then
            aa = aa + VeldTXT(T) + vbTab
        Else
            Printer.Print ; Tab(RapportTab(T));
            Printer.Print ; VeldTXT(T);
        End If
        If RapportTab(T + 1) < RapportTab(T) Then
            If chkAfdrukInVenster Then
            Else
                Printer.Print vbCrLf;
            End If
        End If
        T = T + 1
    Loop
    If chkAfdrukInVenster Then Xlog.X.AddItem aa, Xlog.X.Rows - 1

    For T = 0 To 10:
        VeldTXT(T) = ""
    Next
    VeldTXT(0) = Format(AantalOnderPeriode, "#") + " betalingen dok. < boekjaar"
    VeldTXT(6) = Dec$((TotaalDokOnderPeriode), Masker2002)
    T = 0
    aa = ""
    Do While T < 12
        If chkAfdrukInVenster Then
            aa = aa + VeldTXT(T) + vbTab
        Else
            Printer.Print ; Tab(RapportTab(T));
            Printer.Print ; VeldTXT(T);
        End If
        If RapportTab(T + 1) < RapportTab(T) Then
            If chkAfdrukInVenster Then
            Else
                Printer.Print vbCrLf;
            End If
        End If
        T = T + 1
    Loop
    If chkAfdrukInVenster Then Xlog.X.AddItem aa, Xlog.X.Rows - 1
    
    For T = 0 To 10:
        VeldTXT(T) = ""
    Next
    VeldTXT(0) = Format(AantalBovenPeriode, "#") + " betalingen dok. > boekjaar"
    VeldTXT(6) = Dec$((TotaalDokBovenPeriode), Masker2002)
    T = 0
    aa = ""
    Do While T < 12
        If chkAfdrukInVenster Then
            aa = aa + VeldTXT(T) + vbTab
        Else
            Printer.Print ; Tab(RapportTab(T));
            Printer.Print ; VeldTXT(T);
        End If
        If RapportTab(T + 1) < RapportTab(T) Then
            If chkAfdrukInVenster Then
            Else
                Printer.Print vbCrLf;
            End If
        End If
        T = T + 1
    Loop
    If chkAfdrukInVenster Then Xlog.X.AddItem aa, Xlog.X.Rows - 1
End If

If chkAfdrukInVenster Then
    Xlog!Kopie(1).Checked = True
    Xlog.X.Row = 1
    Xlog.X.Col = 0
    Xlog.X.ColWidth(0) = 2385
    Xlog.X.ColWidth(1) = 840
    Xlog.X.ColWidth(2) = 990
    Xlog.X.ColWidth(3) = 390
    Xlog.X.ColWidth(4) = 945
    Xlog.X.ColWidth(5) = 855
    Xlog.X.ColWidth(6) = 855
    Xlog.X.ColWidth(7) = 855
    Xlog.X.ColWidth(8) = 915
    Xlog.X.ColWidth(9) = 975
    Xlog.X.ColWidth(10) = 960
    Xlog.WindowState = 2

    Xlog.X.ColAlignment(4) = flexAlignRightTop
    Xlog.X.ColAlignment(5) = flexAlignRightTop
    Xlog.X.ColAlignment(6) = flexAlignRightTop
    Xlog.X.ColAlignment(8) = flexAlignRightTop
    Xlog.X.ColAlignment(10) = flexAlignRightTop
    
    Screen.MousePointer = 0
    Xlog.WijzigenLijn.Visible = False
    Xlog.Afsluiten.TabStop = False
    XLogKey = ""
    Xlog.Show 1
    Xlog.WindowState = 0
    Unload Xlog
End If

End Sub
Private Sub Form_Load()
Dim TempbModus As Integer

PlGrensVan = Left$(BoekjaarVanTot, 8)
PlGrensTot = Right$(BoekjaarVanTot, 8)

Set Printer = Printers(LijstPrinterNr)
If Printer.Orientation = vbPRORLandscape Then
    chkAfdrukLiggend.Value = 1
Else
    chkAfdrukLiggend.Value = 0
End If

On Local Error Resume Next
Err = 0
Selektie(5) = LaadTekst("DokumentenBalans", "BetalingsKontrole")
If Err Then
Else
    Selektie(0) = LaadTekst("DokumentenBalans", "KontroleVervaldag")
    Selektie(1) = LaadTekst("DokumentenBalans", "GeenBetalingHogerBoekjaar")
    Selektie(2) = LaadTekst("DokumentenBalans", "PeriodeBegrenzen")
    Selektie(3) = LaadTekst("DokumentenBalans", "EnkelNietBetaaldeDokumenten")
    Selektie(4) = LaadTekst("DokumentenBalans", "-Dokumenten-1994Uitsluiten")
    ChkFinancieelDetail = LaadTekst("DokumentenBalans", "FinancieelDetailViaJournaal")
    chkAfdrukInVenster = LaadTekst("DokumentenBalans", "AfdrukInRooster")
End If
TekstLijn(1).Text = Rdt
TekstLijn(4).Text = fDatumText(Left$(BoekjaarVanTot, 8)) + " - " + fDatumText(Right$(BoekjaarVanTot, 8))

Select Case aIndex
    Case FlLeverancier
        LijstNaam = "Balans Leveranciers"
        FlPartij = FlLeverancier
    Case FlKlant
        LijstNaam = "Balans Klanten"
        FlPartij = FlKlant
    Case Else
        MsgBox "stop balans partijen!"
End Select
BalansKL.Caption = LijstNaam
TekstLijn(2).Text = "0"
TekstLijn(3).Text = String$(12, "z")

End Sub


Private Sub InitVelden()
Dim T As Integer
Dim VolgTab As Integer

RapportVeld(0) = "Identiteit"
    RapportTab(0) = 1

RapportVeld(1) = "Dokument"
    RapportTab(1) = 31

RapportVeld(2) = "Datum"
    RapportTab(2) = 40

RapportVeld(3) = "Munt"
    RapportTab(3) = 51

RapportVeld(4) = "      Goed"
    RapportTab(4) = 55

RapportVeld(5) = "    B.T.W."
    RapportTab(5) = 66

RapportVeld(6) = "   Betaald"
    RapportTab(6) = 77

RapportVeld(7) = "Fin.stuk"
    RapportTab(7) = 88

RapportVeld(8) = "      Rest"
    RapportTab(8) = 97

RapportVeld(9) = "Vervaldag"
    RapportTab(9) = 108

RapportVeld(10) = "     Cumul"
    RapportTab(10) = 119

RapportTab(11) = 0

If chkAfdrukInVenster Then
    Me.Hide
    Unload Xlog
    Xlog.Hide
    Xlog.Caption = "DokumentenBalans"
    Xlog.X.Cols = 11
    Xlog.X.Row = 0
    For T = 0 To 10
        Xlog.X.Col = T
        Xlog.X.Text = RapportVeld(T)
    Next
    Me.Show
End If

End Sub

Private Sub PrintTitel()
Dim T As Integer

If chkAfdrukInVenster Then Exit Sub

    PaginaTeller = PaginaTeller + 1
    Printer.CurrentY = 400
    Printer.Print Tab(1); psTekst(2);
    Printer.Print Tab(108); "Pagina : " + Dec$((PaginaTeller), "##########"); vbCrLf;
    Printer.Print Tab(108); "Datum  : " + psTekst(0); vbCrLf; vbCrLf;
    Printer.Print Tab(1); UCase$(psTekst(3)); vbCrLf;

Printer.Print Tab(1); Lijntje; vbCrLf;

Do While RapportTab(T) <> 0
    Printer.Print Tab(RapportTab(T)); RapportVeld(T);
    If RapportTab(T + 1) < RapportTab(T) Then
        Printer.Print vbCrLf;
    End If
    T = T + 1
Loop

Printer.Print Tab(1); Lijntje; vbCrLf; vbCrLf;

End Sub

Private Sub PrintVelden()
Dim T As Integer
Dim VeldTekst As String
Dim aa As String

T = 0
aa = ""
Do While RapportTab(T) <> 0
    If chkAfdrukInVenster Then
        aa = aa + VeldTXT(T) + vbTab
    Else
        Printer.Print Tab(RapportTab(T));
        Printer.Print VeldTXT(T);
    End If
    If RapportTab(T + 1) < RapportTab(T) Then
        If chkAfdrukInVenster Then
        Else
            Printer.Print vbCrLf;
        End If
    End If
    T = T + 1
Loop
If Printer.CurrentY >= Printer.Height - 1200 Then
    On Local Error Resume Next
    Printer.NewPage
    Printer.FontSize = Printer.FontSize
    Printer.Print " "
    Printer.FontSize = Printer.FontSize
    PrintTitel
End If

If chkAfdrukInVenster Then Xlog.X.AddItem aa, Xlog.X.Rows - 1

End Sub

Private Sub Selektie_Click(Index As Integer)

Select Case Index
    Case 0
        If Selektie(Index).Value = 1 Then
            TekstLijn(1) = fDatumText(Right(BoekjaarVanTot, 8))
        Else
            TekstLijn(1) = Rdt
        End If
    Case 4
        If Selektie(4).Value = 1 Then
            BeginPlaySound 9
            MsgBox "Schakel uitsluitend aan indien U problemen ondervindt met sommige geimporteerde DOS-dokumenten van voor 1994 (o.a. BTW 33 %, 8 % luxetaks)." + vbCrLf + vbCrLf + "Indien U alle mogelijkheden van marIntegraal Windows versie met uw oude data wenst te benutten, gelieve ons pér bedrijf een veiligheidskopij te bezorgen." + vbCrLf + vbCrLf + "Binnen uw servicekontrakt werken wij deze kosteloos om in onze lokalen."
        End If
End Select

End Sub

Private Sub TekstLijn_GotFocus(Index As Integer)

TekstLijn(Index).SelLength = Len(TekstLijn(Index).Text)

End Sub

Private Sub TekstLijn_LostFocus(Index As Integer)

Select Case Index
    Case 1
        If DatumFout((TekstLijn(1).Text)) Then
            Beep
            TekstLijn(1).Text = Rdt
        End If
    Case 4
        If DatumFout(Right$(TekstLijn(4).Text, 10)) Then
            MsgBox "Respecteer : " + vbCrLf + vbCrLf + "DD/MM/EEJJ - DD/MM/EEJJ a.u.b. !"
            TekstLijn(4).Text = fDatumText(Left$(BoekjaarVanTot, 8)) + " - " + fDatumText(Right$(BoekjaarVanTot, 8))
            TekstLijn(4).SetFocus
            Exit Sub
        ElseIf Len(TekstLijn(4).Text) <> 23 Then
            MsgBox "Respecteer : " + vbCrLf + vbCrLf + "DD/MM/EEJJ - DD/MM/EEJJ a.u.b. !"
            TekstLijn(4).Text = fDatumText(Left$(BoekjaarVanTot, 8)) + " - " + fDatumText(Right$(BoekjaarVanTot, 8))
            TekstLijn(4).SetFocus
            Exit Sub
        Else
            PlGrensVan = Mid$(TekstLijn(4).Text, 7, 4) + Mid$(TekstLijn(4).Text, 4, 2) + Mid$(TekstLijn(4).Text, 1, 2)
            PlGrensTot = Mid$(TekstLijn(4).Text, 20, 4) + Mid$(TekstLijn(4).Text, 17, 2) + Mid$(TekstLijn(4).Text, 14, 2)
            If BoekjaarVanTot = PlGrensVan + PlGrensTot Then
                Selektie(5) = 1
                Selektie(5).Visible = True
            Else
                Selektie(5) = 0
                Selektie(5).Visible = False
            End If
        End If
End Select

End Sub

