VERSION 5.00
Begin VB.Form BalansKL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Balans"
   ClientHeight    =   3165
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
   ScaleHeight     =   3165
   ScaleWidth      =   7245
   StartUpPosition =   1  'CenterOwner
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
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   30
      ToolTipText     =   "Lijst printer"
      Top             =   2760
      Width           =   4635
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
      Left            =   2280
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2760
      Width           =   315
   End
   Begin VB.CommandButton cmdEuroCheck 
      Caption         =   "EURO check !"
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
      Left            =   1800
      TabIndex        =   28
      Top             =   720
      Width           =   1430
   End
   Begin VB.CheckBox ChkFinancieelDetail 
      Alignment       =   1  'Right Justify
      Caption         =   "&Financiëel Detail Journaal"
      Height          =   240
      Left            =   4500
      TabIndex        =   27
      Top             =   1680
      Value           =   1  'Checked
      Width           =   2685
   End
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
      Left            =   4920
      TabIndex        =   26
      Top             =   2160
      Width           =   2310
   End
   Begin VB.CheckBox chkAfdrukInVenster 
      Alignment       =   1  'Right Justify
      Caption         =   "Afdruk in venster"
      Height          =   255
      Left            =   5370
      TabIndex        =   25
      Top             =   2460
      Width           =   1830
   End
   Begin VB.ListBox BetalingenVoorNa 
      CausesValidation=   0   'False
      Height          =   255
      Index           =   3
      Left            =   2610
      TabIndex        =   24
      Top             =   330
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.ListBox BetalingenVoorNa 
      CausesValidation=   0   'False
      Height          =   255
      Index           =   2
      Left            =   2310
      TabIndex        =   23
      Top             =   300
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.ListBox BetalingenVoorNa 
      CausesValidation=   0   'False
      Height          =   255
      Index           =   1
      Left            =   2010
      TabIndex        =   22
      Top             =   300
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.ListBox BetalingenVoorNa 
      CausesValidation=   0   'False
      Height          =   255
      Index           =   0
      Left            =   1470
      TabIndex        =   21
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
      Left            =   1920
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CheckBox Selektie 
      Alignment       =   1  'Right Justify
      Caption         =   "Betalingsjournaal &weergeven"
      Height          =   240
      Index           =   5
      Left            =   4410
      TabIndex        =   12
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1110
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
      TabIndex        =   18
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
      TabIndex        =   14
      Top             =   2160
      Width           =   3555
   End
   Begin VB.CheckBox Selektie 
      Alignment       =   1  'Right Justify
      Caption         =   "Documenten -1994 &uitsluiten"
      Height          =   240
      Index           =   4
      Left            =   4410
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
      Caption         =   "&Controleren op vervaldag"
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
      Left            =   3840
      TabIndex        =   9
      Top             =   960
      Value           =   1  'Checked
      Width           =   3345
   End
   Begin VB.CheckBox Selektie 
      Alignment       =   1  'Right Justify
      Caption         =   "Enkel docum&enten deze periode"
      Height          =   195
      Index           =   2
      Left            =   540
      TabIndex        =   15
      Top             =   2460
      Value           =   1  'Checked
      Width           =   3075
   End
   Begin VB.CheckBox Selektie 
      Alignment       =   1  'Right Justify
      Caption         =   "&Niet betaalde documenten"
      Height          =   240
      Index           =   3
      Left            =   4620
      TabIndex        =   10
      Top             =   1200
      Value           =   1  'Checked
      Width           =   2565
   End
   Begin VB.CommandButton Drukken 
      Caption         =   "Af&druk"
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
      Left            =   5760
      TabIndex        =   7
      Top             =   60
      Width           =   1430
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
      Height          =   300
      Left            =   5790
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   480
      Width           =   1430
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Periode van - tot"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   4
      Left            =   180
      TabIndex        =   13
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
      TabIndex        =   17
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

Dim RapportSelektie As String
Dim RapportTitel As String
Dim RapportDefinitie As String
Dim LijstNaam As String
Dim VeldTXT(20) As String
Dim FlPartij As Integer
Dim DocTeller As Integer

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

Private Sub cbLijstPrinterHier_Click()
    
    LijstPrinterNr = cbLijstPrinterHier.ListIndex
    Set Printer = Printers(LijstPrinterNr)
    
End Sub

Private Sub cbTogglePrinter_Click()

    Me.cbLijstPrinterHier.Enabled = Not Me.cbLijstPrinterHier.Enabled
    
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


Private Sub ChkFinancieelDetail_Click()

If ChkFinancieelDetail.Value = 0 Then
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

BeWaarTekst "dokumentenBalans", "KontroleVervaldag", Selektie(0)
BeWaarTekst "dokumentenBalans", "GeenBetalingHogerBoekjaar", Selektie(1)
BeWaarTekst "dokumentenBalans", "PeriodeBegrenzen", Selektie(2)
BeWaarTekst "dokumentenBalans", "EnkelNietBetaaldedokumenten", Selektie(3)
BeWaarTekst "dokumentenBalans", "-dokumenten-1994Uitsluiten", Selektie(4)
BeWaarTekst "dokumentenBalans", "BetalingsKontrole", Selektie(5)
BeWaarTekst "dokumentenBalans", "FinancieelDetailViaJournaal", ChkFinancieelDetail
BeWaarTekst "dokumentenBalans", "AfdrukInRooster", chkAfdrukInVenster

End Sub


Private Sub cmdEuroCheck_Click()

Dim dBTW As Currency
Dim dTOT As Currency

Dim Teller As Integer

Dim TotaalAantalBEF As Long
Dim TotaalAantalEUR As Long
Dim TotaalGewijzigd As Long

    Exit Sub

    Msg = "Hierna kan U de totaalbedragen pér dokument terug"
    Msg = Msg & vbCr & "hertotaliseren in BEF voor vorige boekjaren"
    Msg = Msg & vbCr & "en vervolgens in EUR beschikbaar maken."
    Msg = Msg & vbCr & vbCr
    Msg = Msg & "Mag de herrekening gestart worden."
    KtrlBox = MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton2)
    If KtrlBox = vbNo Then
        Exit Sub
    End If
    bFirst TABLE_INVOICES, 1
    If Ktrl Then
        MsgBox "stop"
        Exit Sub
    Else
        Screen.MousePointer = vbHourglass
        BalansKL.Enabled = False
        Do
            XDoEvents = DoEvents
            RecordToVeld TABLE_INVOICES
            If DATE_CHECK(vBibTekst(TABLE_INVOICES, "#v035 #"), BOOKYEAR_KEY) = False Then
                TotaalAantalBEF = TotaalAantalBEF + 1
                Select Case Left(vBibTekst(TABLE_INVOICES, "#v034 #"), 1)
                    Case "K"
                        Select Case Left(vBibTekst(TABLE_INVOICES, "#v033 #"), 1)
                            Case "V"
                                dBTW = Val(vBibTekst(TABLE_INVOICES, "#v064 #"))
                                dTOT = 0
                                For Teller = 55 To 63
                                    dTOT = dTOT + Val(vBibTekst(TABLE_INVOICES, "#v" + Format(Teller, "000") + " #"))
                                Next
                                dTOT = dTOT + Val(vBibTekst(TABLE_INVOICES, "#v089 #")) + dBTW
                            Case "Q"
                                dTOT = Val(vBibTekst(TABLE_INVOICES, "#B010 #")) - Val(vBibTekst(TABLE_INVOICES, "#B090 #")) - Val(vBibTekst(TABLE_INVOICES, "#v065 #"))
                                dBTW = 0
                        End Select
                    Case "L"
                        dBTW = Val(vBibTekst(TABLE_INVOICES, "#v045 #")) - Val(vBibTekst(TABLE_INVOICES, "#v044 #")) - Val(vBibTekst(TABLE_INVOICES, "#v043 #"))
                        If Left(vBibTekst(TABLE_INVOICES, "#v035 #"), 4) > "1992" Then
                            dBTW = dBTW - Val(vBibTekst(TABLE_INVOICES, "#v042 #"))
                        End If
                        dTOT = 0
                        For Teller = 46 To 49
                            dTOT = dTOT + Val(vBibTekst(TABLE_INVOICES, "#v" + Format(Teller, "000") + " #"))
                        Next
                        dTOT = dTOT + dBTW
                End Select
                If Round(dTOT / EURO, 2) <> Val(vBibTekst(TABLE_INVOICES, "#v249 #")) Then
                    MsgBox vBibTekst(TABLE_INVOICES, "#v034 #") & ": " & Round(dTOT / EURO, 2) & " <> " & Val(vBibTekst(TABLE_INVOICES, "#v249 #")) & vbCr & " en zogezegd reeds betaald : " & Val(vBibTekst(TABLE_INVOICES, "#v037 #")) & vbCr & vbCr & "Wordt bijgewerkt...", vbInformation, "Bijwerking voor:"
                    If Val(vBibTekst(TABLE_INVOICES, "#v249 #")) = Val(vBibTekst(TABLE_INVOICES, "#v037 #")) Then
                        vBib TABLE_INVOICES, Str$(Round(dTOT / EURO, 2)), "v249"
                        vBib TABLE_INVOICES, Str$(Round(dTOT / EURO, 2)), "v037"
                    Else
                        vBib TABLE_INVOICES, Str$(Round(dTOT / EURO, 2)), "v249"
                        If Val(vBibTekst(TABLE_INVOICES, "#v037 #")) = 0 Then
                        Else
                            MsgBox "Totaal betaald dient U manueel nog te verbeteren.  Staat nu op " & Val(vBibTekst(TABLE_INVOICES, "#v037 #")), vbExclamation, vBibTekst(TABLE_INVOICES, "#v034 #")
                        End If
                    End If
                    bUpdate TABLE_INVOICES, 1
                    TotaalGewijzigd = TotaalGewijzigd + 1
                ElseIf dTOT = Val(vBibTekst(TABLE_INVOICES, "#v037 #")) Then
                    '100 % zeker totaal BEF staat nog als betaald
                    SnelHelpPrint vBibTekst(TABLE_INVOICES, "#v033 #") & " was al volledig betaald in BEF voor " & Val(vBibTekst(TABLE_INVOICES, "#v037 #")) & " Wordt nu als betaald geplaatst € " & Round(dTOT / EURO, 2), BL_LOGGING
                    vBib TABLE_INVOICES, Str$(Round(dTOT / EURO, 2)), "v037"
                    bUpdate TABLE_INVOICES, 1
                ElseIf Val(vBibTekst(TABLE_INVOICES, "#v249 #")) = Val(vBibTekst(TABLE_INVOICES, "#v037 #")) Then
                    '100 % zeker totaal en betaald in EUR reeds aangeduid
                ElseIf Val(vBibTekst(TABLE_INVOICES, "#v037 #")) = 0 Then
                    'nog geen betalingen, dus ok
                ElseIf InStr(vBibTekst(TABLE_INVOICES, "#v037 #"), ".") = 0 Then
                    'een totaal bedrag betaald zonder aanwezigheid van decimaal punt
                    Msg = vBibTekst(TABLE_INVOICES, "#v033 #") & " met totaal € " & Val(vBibTekst(TABLE_INVOICES, "#v249 #")) & vbCr
                    Msg = Msg & "heeft een bedrag als betaling : " & Val(vBibTekst(TABLE_INVOICES, "#v037 #")) & vbCr & vbCr
                    Msg = Msg & "mag omgewerkt worden als betaald € " & Round(Val(vBibTekst(TABLE_INVOICES, "#v037 #")) / EURO, 2) & vbCr & vbCr & vbCr
                    Msg = Msg & "AANDACHT: ENKEL WIJZIGEN INDIEN U 100 % ZEKER BENT !!!" & vbCr & vbCr
                    Msg = Msg & "Ja om te wijzigen, Nee om over te slaan"
                    KtrlBox = vbCancel
                    Do While KtrlBox = vbCancel
                        KtrlBox = MsgBox(Msg, vbYesNoCancel + vbDefaultButton3 + vbQuestion)
                    Loop
                    If KtrlBox = vbYes Then
                        vBib TABLE_INVOICES, Str$(Round(Val(vBibTekst(TABLE_INVOICES, "#v037 #")) / EURO, 2)), "v037"
                        bUpdate TABLE_INVOICES, 1
                    End If
                Else
                    SnelHelpPrint "Vermoedelijk alles ok voor " & vBibTekst(TABLE_INVOICES, "#v033 #"), BL_LOGGING
                End If
            Else
                TotaalAantalEUR = TotaalAantalEUR + 1
            End If
            bNext TABLE_INVOICES
            If Ktrl Then
                Exit Do
            End If
        Loop
        Screen.MousePointer = vbNormal
        BalansKL.Enabled = True
        Msg = "Totaal dokumenten in BEF " & TotaalAantalBEF
        Msg = Msg & vbCr
        Msg = Msg & "Totaal dokumenten in EUR " & TotaalAantalEUR
        Msg = Msg & vbCr
        Msg = Msg & "Totaal aantal gewijzigd " & TotaalGewijzigd
        MsgBox Msg, vbInformation
    End If
    
End Sub

Private Sub CmdStandaard_Click()

Selektie(0) = 0
PlGrensVan = Left(BOOKYEAR_FROMTO, 8)
PlGrensTot = Right(BOOKYEAR_FROMTO, 8)
TekstLijn(4).text = DATE_TEXT(Left(BOOKYEAR_FROMTO, 8)) + " - " + DATE_TEXT(Right(BOOKYEAR_FROMTO, 8))
Selektie(2) = 1

Selektie(1) = 1
Selektie(3) = 1
ChkFinancieelDetail = 1
Selektie(5) = 1
TekstLijn(0).text = "Boekhoudcontrole " & TekstLijn(4).text

End Sub

Private Sub CmdStandaardBetaling_Click()

Selektie(0) = 1
Selektie(2) = 0

Selektie(1) = 0
Selektie(3) = 1
ChkFinancieelDetail = 0
TekstLijn(1) = MIM_GLOBAL_DATE
TekstLijn(0).text = "Betalingscontrole"

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
Dim dBTW As Double
Dim dTOT As Double
Dim dnt As Double
Dim dTnt As Double
Dim drb  As Double
Dim drbet As Double

Dim LengteSleutel As Integer
Dim sleutelIndex As Integer
Dim xOmgeving As Integer
Dim MerkOp As Boolean

TotaalBTW = 0
TotaalGOED = 0
TotaalALBETAALD = 0
TotaalNOGTEBETALEN = 0
TotaalVOOR = 0
TotaalNA = 0
DocTeller = 0

dttot = 0
dtbtw = 0
dvtot = 0
dTrb = 0
dTnt = 0

BetalingenVoorNa(2).Clear
BetalingenVoorNa(3).Clear

Select Case FlPartij
    Case TABLE_SUPPLIERS
        BeginSleutel = "L" + TekstLijn(2).text
        EindSleutel = "L" + TekstLijn(3).text
    Case TABLE_CUSTOMERS
        BeginSleutel = "K" + TekstLijn(2).text
        EindSleutel = "K" + TekstLijn(3).text
End Select

rdtemp = DATE_KEY((TekstLijn(1).text))

psTekst(2) = BalansKL.Caption + " " + Mid(Mim.Caption, InStr(Mim.Caption, "["))
psTekst(0) = TekstLijn(1).text
psTekst(3) = TekstLijn(0).text
    InitVelden

TDS = "Geen journalen voor : " + vbCrLf
bGetOrGreater TABLE_INVOICES, 1, BeginSleutel
If Ktrl Or UCase$(vSet(KEY_BUF(TABLE_INVOICES), 13)) > UCase$(EindSleutel) Then
    Beep
    Exit Sub
ElseIf Left(KEY_BUF(TABLE_INVOICES), 1) <> Left(BeginSleutel, 1) Then
    Beep
    Exit Sub
Else
    Screen.MousePointer = vbHourglass
    BalansKL.Enabled = False
    PAGE_COUNTER = 0
    If chkAfdrukInVenster = 0 Then
        If Printer.Width > 12000 Then
            Printer.FontSize = 10
            Printer.FontName = "Courier New"
            Printer.Print " "
            Printer.FontSize = 10
        Else
            On Local Error Resume Next
            Printer.FontSize = 7.2
            Printer.FontName = "Courier New"
            Printer.Print " "
            Printer.FontSize = 7.2
            Printer.FontBold = True
        End If
    End If
    PrintTitel
End If

MerkOp = False
If PlGrensVan = Left(BOOKYEAR_FROMTO, 8) And PlGrensTot = Right(BOOKYEAR_FROMTO, 8) Then
    If Left(PlGrensVan, 4) = Left(PlGrensTot, 4) Then
        MerkOp = True
    End If
End If
    

Do
    XDoEvents = DoEvents
    RecordToVeld TABLE_INVOICES
    GoSub KontroleVoorwaarden
    bNext TABLE_INVOICES
    If Ktrl Or UCase$(vSet(KEY_BUF(TABLE_INVOICES), 13)) > UCase$(EindSleutel) Then
        Exit Do
    End If
Loop
    TotaalNOGTEBETALEN = TotaalGOED - TotaalALBETAALD
    EindTotaal
    Screen.MousePointer = vbNormal
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
    BalansKL.SetFocus
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
                        Printer.Print "Totaal financiële bewegingen: " + Format(TotaalDokOnderPeriode, MASK_2002)
                    ElseIf Tel = 1 Then
                        Printer.Print "Totaal financiële bewegingen: " + Format(TotaalDokBovenPeriode, MASK_2002)
                    ElseIf Tel = 2 Then
                        Printer.Print "Totaal financiële bewegingen: " + Format(TotaalVOOR, MASK_2002)
                    ElseIf Tel = 3 Then
                        Printer.Print "Totaal financiële bewegingen: " + Format(TotaalNA, MASK_2002)
                    End If
                    Printer.NewPage
                End If
            Next
            Printer.EndDoc
        End If
    End If
    SnelHelpPrint "Klaar", BL_LOGGING
    Unload BalansKL
Exit Sub

KontroleVoorwaarden:
If Selektie(0).Value = 1 Then
    'vervaldag
    If vBibTekst(TABLE_INVOICES, "#v036 #") > rdtemp Then
        Return
    End If
End If

If Selektie(4).Value = 1 Then
    '-1994 dokumenten uitsluiten
    If Left(vBibTekst(TABLE_INVOICES, "#v035 #"), 4) < "1994" Then
        Return
    End If
End If

If Selektie(2).Value = 1 Then
    'dokumenten enkel van deze periode
    If MerkOp = True Then
        If Left(vBibTekst(TABLE_INVOICES, "#v035 #"), 4) <> Mid(vBibTekst(TABLE_INVOICES, "#v033 #"), 3, 4) Then
            MsgBox "Opgelet, noteer/controleer a.u.b.:" & vbCrLf & "Datum document: " & vBibTekst(TABLE_INVOICES, "#v035 #") & " onlogisch voor document nummer " & vBibTekst(TABLE_INVOICES, "#v033 #"), vbExclamation
        End If
    End If
    If vBibTekst(TABLE_INVOICES, "#v035 #") < PlGrensVan Or vBibTekst(TABLE_INVOICES, "#v035 #") > PlGrensTot Then
        Return
    End If
End If

If Selektie(5).Value = 1 And FlPartij = TABLE_SUPPLIERS Then
    bGet TABLE_JOURNAL, 1, vBibTekst(TABLE_INVOICES, "#v033 #")
    If Ktrl Then
        Return
    Else
        RecordToVeld TABLE_JOURNAL
        If vBibTekst(TABLE_JOURNAL, "#v035 #") < PlGrensVan Or vBibTekst(TABLE_INVOICES, "#v035 #") > PlGrensTot Then
            MsgBox "dokumentdatum (" & vBibTekst(TABLE_INVOICES, "#v035 #") & ") <> boekdatum journaal (" & vBibTekst(TABLE_JOURNAL, "#v035 #") & ")" & vbCr & vbCr & "Wordt uit boekhoudcontrole geweerd.  Kontroleer eventueel manueel", vbInformation, vBibTekst(TABLE_INVOICES, "#v033 #")
            Return
        End If
    End If
End If

dTOT = Val(vBibTekst(TABLE_INVOICES, "#v249 #"))
If XisEuroWisBEF = True Then
    dTOT = Round(dTOT * EURO)
End If

syMasker = MASK_EUR
If Mid(vBibTekst(TABLE_INVOICES, "#v033 #"), 2, 1) = "1" Then
    dTOT = -dTOT
    drb = -Val(vBibTekst(TABLE_INVOICES, "#v037 #"))
    If XisEuroWisBEF = True Then
        drb = Round(-Val(vBibTekst(TABLE_INVOICES, "#v037 #")) * EURO)
    Else
        drb = -Val(vBibTekst(TABLE_INVOICES, "#v037 #"))
    End If
Else
    If XisEuroWisBEF = True Then
        drb = Round(Val(vBibTekst(TABLE_INVOICES, "#v037 #")) * EURO)
    Else
        drb = Val(vBibTekst(TABLE_INVOICES, "#v037 #"))
    End If
End If

VeldTXT(1) = vBibTekst(TABLE_INVOICES, "#v033 #")
VeldTXT(2) = DATE_TEXT(vBibTekst(TABLE_INVOICES, "#v035 #"))
VeldTXT(3) = vBibTekst(TABLE_INVOICES, "#vs03 #")
VeldTXT(4) = Dec$(dTOT / Val(vBibTekst(TABLE_INVOICES, "#v040 #")), syMasker)
VeldTXT(5) = "" 'Dec$((dBtw), MASK_2002)
VeldTXT(9) = DATE_TEXT(vBibTekst(TABLE_INVOICES, "#v036 #"))

If Trim$(vBibTekst(TABLE_INVOICES, "#v034 #")) <> KopBuf Then
    dTnt = 0
    bGet aIndex, 0, vSet(Mid(vBibTekst(TABLE_INVOICES, "#v034 #"), 2, 12), 12)
    KopBuf = Trim$(vBibTekst(TABLE_INVOICES, "#v034 #"))
    VeldTXT(0) = RTrim$(Mid(vBibTekst(TABLE_INVOICES, "#v034 #"), 2, 12))
    If Ktrl Then
        VeldTXT(0) = VeldTXT(0) + " * niet meer aanwezig *"
    Else
        RecordToVeld aIndex
        VeldTXT(0) = Left(VeldTXT(0) + " " + RTrim$(vBibTekst(aIndex, "#A100 #")) + " " + RTrim$(vBibTekst(aIndex, "#A101 #")), 27)
    End If
    SnelHelpPrint VeldTXT(0), BL_LOGGING
End If

bGet TABLE_JOURNAL, 1, vBibTekst(TABLE_INVOICES, "#v033 #")
If Ktrl Or ChkFinancieelDetail = 0 Then
    'Er zijn geen journalen voor dit dokument !
    'Verwittigen via afdruk laatste blz.
    If ChkFinancieelDetail Then
        TDS = TDS + vBibTekst(TABLE_INVOICES, "#v033 #") + "  ...  " + vBibTekst(FlPartij, "#A110 #") + " " + vBibTekst(FlPartij, "#A100 #") + vbCrLf
    End If
    If dTOT = drb And Selektie(3).Value = 1 Then
    Else
        If InStr(VeldTXT(0), Mid(vBibTekst(TABLE_INVOICES, "#v034 #"), 2)) = 1 Then
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
        
        dTnt = dTnt + dTOT - drb 'Algemeen Totaal Cumul te betalen
        VeldTXT(6) = Dec$((drb), MASK_2002)
        VeldTXT(7) = vBibTekst(TABLE_INVOICES, "#v038 #")
        VeldTXT(8) = "" 'Dec$((dTot - drb), MASK_2002)
        VeldTXT(10) = Dec$((dTnt), MASK_2002)
    
        'TotaalBTW = TotaalBTW + dBtw
        TotaalGOED = TotaalGOED + dTOT '- dBtw
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
If Round(dTOT, 2) = Round(drb, 2) And Selektie(3).Value = 1 Then
    Return
Else
    bGet TABLE_JOURNAL, 1, vBibTekst(TABLE_INVOICES, "#v033 #")
End If
dTnt = dTnt + dTOT
ReedsMetBetalingen = False

drb = 0
Do
    RecordToVeld TABLE_JOURNAL
    If Trim$(vBibTekst(TABLE_JOURNAL, "#v038 #")) <> "" Then
        If Left(vBibTekst(TABLE_JOURNAL, "#v019 #"), 1) <> "4" Then
        Else
            ReedsMetBetalingen = True
            GoSub BetalingErBij
        End If
    End If
    bNext TABLE_JOURNAL
    If Ktrl Or Trim$(KEY_BUF(TABLE_JOURNAL)) <> Trim$(vBibTekst(TABLE_INVOICES, "#v033 #")) Then
        Exit Do
    End If
Loop
If Not ReedsMetBetalingen Then
    drb = 0
    vBib TABLE_JOURNAL, "0", "v068"
    GoSub BetalingErBij
End If
Return

BetalingErBij:
BetaaldBedrag = Val(vBibTekst(TABLE_JOURNAL, "#v068 #"))
If FlPartij = TABLE_CUSTOMERS Then
    BetaaldBedrag = -BetaaldBedrag
End If
If Selektie(1).Value = 1 Then
    'Betalingen buiten periode uitsluiten
    If vBibTekst(TABLE_JOURNAL, "#v066 #") < PlGrensVan Then
        TotaalVOOR = TotaalVOOR + BetaaldBedrag
        Msg = vSet(Mid(vBibTekst(TABLE_INVOICES, "#v034 #"), 2), 12) + " "
        Msg = Msg + vSet(vBibTekst(TABLE_JOURNAL, "#v067 #"), 20) + " "
        Msg = Msg + DATE_TEXT(vBibTekst(TABLE_INVOICES, "#v035 #")) + " "
        Msg = Msg + vBibTekst(TABLE_INVOICES, "#v033 #") + " "
        Msg = Msg + Dec(BetaaldBedrag, MASK_2002) + " "
        Msg = Msg + vSet(vBibTekst(TABLE_JOURNAL, "#v019 #"), 7) + " "
        Msg = Msg + DATE_TEXT(vBibTekst(TABLE_JOURNAL, "#v066 #")) + " "
        Msg = Msg + vBibTekst(TABLE_JOURNAL, "#v038 #") + " "
        Msg = Msg + vBibTekst(TABLE_JOURNAL, "#v069 #")
        BetalingenVoorNa(2).AddItem Msg
        BetaaldBedrag = 0
    ElseIf vBibTekst(TABLE_JOURNAL, "#v066 #") > PlGrensTot Then
        TotaalNA = TotaalNA + BetaaldBedrag
        Msg = vSet(Mid(vBibTekst(TABLE_INVOICES, "#v034 #"), 2), 12) + " "
        Msg = Msg + vSet(vBibTekst(TABLE_JOURNAL, "#v067 #"), 20) + " "
        Msg = Msg + DATE_TEXT(vBibTekst(TABLE_INVOICES, "#v035 #")) + " "
        Msg = Msg + vBibTekst(TABLE_INVOICES, "#v033 #") + " "
        Msg = Msg + Dec(BetaaldBedrag, MASK_2002) + " "
        Msg = Msg + vSet(vBibTekst(TABLE_JOURNAL, "#v019 #"), 7) + " "
        Msg = Msg + DATE_TEXT(vBibTekst(TABLE_JOURNAL, "#v066 #")) + " "
        Msg = Msg + vBibTekst(TABLE_JOURNAL, "#v038 #") + " "
        Msg = Msg + vBibTekst(TABLE_JOURNAL, "#v069 #")
        BetalingenVoorNa(3).AddItem Msg
        BetaaldBedrag = 0
    End If
End If

If InStr(VeldTXT(0), Mid(vBibTekst(TABLE_INVOICES, "#v034 #"), 2)) = 1 Then
    dTnt = dTOT: drb = 0
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
dttot = dttot + dTOT - dBTW 'Totaal excl. btw
dtbtw = dtbtw + dBTW        'Totaal btw aftrekbaar

VeldTXT(6) = Dec$(BetaaldBedrag, MASK_2002)
VeldTXT(7) = vBibTekst(TABLE_JOURNAL, "#v038 #")
VeldTXT(8) = Dec$((dTOT - drb), MASK_2002)
VeldTXT(10) = Dec$((dTnt), MASK_2002)

If Val(VeldTXT(4)) + Val(VeldTXT(5)) <> 0 Then
    'TotaalBTW = TotaalBTW + dBtw
    TotaalGOED = TotaalGOED + dTOT '- dBtw
    TotaalNOGTEBETALEN = TotaalNOGTEBETALEN + dTOT
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
bGet TABLE_JOURNAL, 1, vBibTekst(TABLE_INVOICES, "#v033 #")
If Ktrl Then
    MsgBox "onlogische situatie"
Else
    Do
        RecordToVeld TABLE_JOURNAL
        If Trim$(vBibTekst(TABLE_JOURNAL, "#v038 #")) <> "" Then
            If Left(vBibTekst(TABLE_JOURNAL, "#v019 #"), 1) <> "4" Then
            Else
                BetaaldBedrag = Val(vBibTekst(TABLE_JOURNAL, "#v068 #"))
                If FlPartij = TABLE_CUSTOMERS Then
                    BetaaldBedrag = -BetaaldBedrag
                End If
                If XisEuroWisBEF = True Then
                    If vBibTekst(TABLE_JOURNAL, "#v066 #") > Right(BOOKYEAR_FROMTO, 8) Then
                        BetaaldBedrag = Round(BetaaldBedrag * EURO)
                    End If
                End If
                If Selektie(1).Value = 1 Then
                    'Betalingen buiten periode uitsluiten
                    If vBibTekst(TABLE_JOURNAL, "#v066 #") < PlGrensVan Or vBibTekst(TABLE_JOURNAL, "#v066 #") > PlGrensTot Then
                    Else
                        drb = drb + BetaaldBedrag
                    End If
                Else
                    drb = drb + BetaaldBedrag
                End If
            End If
        'ElseIf FlPartij = TABLE_CUSTOMERS Then
        '    If vBibTekst(TABLE_JOURNAL, "#v033 #") = "Q0800286" Then
        '        Stop
        '    End If
        End If
        bNext TABLE_JOURNAL
        If Ktrl Or Trim$(KEY_BUF(TABLE_JOURNAL)) <> Trim$(vBibTekst(TABLE_INVOICES, "#v033 #")) Then
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
VeldTXT(4) = Dec$((TotaalGOED), MASK_2002)
VeldTXT(5) = "" 'Dec$((TotaalBTW), MASK_2002)
VeldTXT(6) = Dec$((TotaalALBETAALD), MASK_2002)
VeldTXT(10) = Dec$((TotaalNOGTEBETALEN), MASK_2002)
If chkAfdrukInVenster = 0 Then Printer.Print vbCrLf; FULL_LINE; vbCrLf;

T = 0
aa = ""
Do While T < 12
    If chkAfdrukInVenster Then
        aa = aa + VeldTXT(T) + vbTab
    Else
        Printer.Print ; Tab(REPORT_TAB(T));
        Printer.Print ; VeldTXT(T);
    End If
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        If chkAfdrukInVenster Then
        Else
            Printer.Print vbCrLf;
        End If
    End If
    T = T + 1
Loop
If chkAfdrukInVenster Then mdiXlog.X.AddItem aa, mdiXlog.X.Rows - 1

Dim GroepSelektie As String
Dim TotaalBedragGroep As Double
Dim AantalInGroep As Integer
Dim BedragZonderdokument As Double
Dim AantalZonderdokument As Integer
Dim BetaaldBedragBB As Double
Dim GroepRekening4 As String

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
        VeldTXT(6) = Dec$((TotaalVOOR), MASK_2002)
        T = 0
        aa = ""
        Do While T < 12
            If chkAfdrukInVenster Then
                aa = aa + VeldTXT(T) + vbTab
            Else
                Printer.Print ; Tab(REPORT_TAB(T));
                Printer.Print ; VeldTXT(T);
            End If
            If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
                If chkAfdrukInVenster Then
                Else
                    Printer.Print vbCrLf;
                End If
            End If
            T = T + 1
        Loop
        If chkAfdrukInVenster Then mdiXlog.X.AddItem aa, mdiXlog.X.Rows - 1
    End If
    If TotaalNA <> 0 Then
        For T = 0 To 10:
            VeldTXT(T) = ""
        Next
        VeldTXT(0) = "Dok. reeds betaald na :"
        VeldTXT(2) = Right(TekstLijn(4), 10)
        VeldTXT(6) = Dec$((TotaalNA), MASK_2002)
        T = 0
        aa = ""
        Do While T < 12
            If chkAfdrukInVenster Then
                aa = aa + VeldTXT(T) + vbTab
            Else
                Printer.Print ; Tab(REPORT_TAB(T));
                Printer.Print ; VeldTXT(T);
            End If
            If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
                If chkAfdrukInVenster Then
                Else
                    Printer.Print vbCrLf;
                End If
            End If
            T = T + 1
        Loop
        If chkAfdrukInVenster Then mdiXlog.X.AddItem aa, mdiXlog.X.Rows - 1
    End If

    If chkAfdrukInVenster = 0 Then Printer.Print vbCrLf; FULL_LINE; vbCrLf;
    TotaalBedragGroep = 0
    AantalInGroep = 0
    BedragZonderdokument = 0
    AantalZonderdokument = 0
    AantalBovenPeriode = 0
    AantalOnderPeriode = 0
    TotaalDokBovenPeriode = 0
    TotaalDokOnderPeriode = 0
    
    GroepSelektie = String99(READING, FlPartij + 296)
    If Trim(GroepSelektie) = "" Then
        GroepSelektie = Left(String99(READING, FlPartij + 8), 4) + "999"
    End If
    If Len(GroepSelektie) > 7 Then
        MsgBox "Groep bestaat uit meer dan 7 tekens: " & GroepSelektie, vbExclamation
        GroepSelektie = Left(GroepSelektie, 7)
    ElseIf Len(GroepSelektie) <> 7 Then
        GroepSelektie = GroepSelektie + String(7 - Len(GroepSelektie), "9")
    End If
    BetalingenVoorNa(0).Clear
    BetalingenVoorNa(1).Clear
    bGetOrGreater TABLE_LEDGERACCOUNTS, 0, vSet(String99(READING, FlPartij + 8), 7)
    If Ktrl Then
        MsgBox "onlogika"
    Else
        RecordToVeld TABLE_LEDGERACCOUNTS
        Do While KEY_BUF(TABLE_LEDGERACCOUNTS) <= GroepSelektie
            GroepRekening4 = KEY_BUF(TABLE_LEDGERACCOUNTS)
            SnelHelpPrint "Journalen boekjaar voor rek. " + KEY_BUF(TABLE_LEDGERACCOUNTS) + " worden gekontroleerd.  Ogenblik a.u.b.", BL_LOGGING
            AantalInGroep = AantalInGroep + 1
            If bhEuro Then
                TotaalBedragGroep = TotaalBedragGroep + Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#e" + Format(22 + ACTIVE_BOOKYEAR, "000") + " #"))
            Else
                TotaalBedragGroep = TotaalBedragGroep + Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#v" + Format(22 + ACTIVE_BOOKYEAR, "000") + " #"))
            End If
            bGetOrGreater TABLE_JOURNAL, 0, vSet(KEY_BUF(TABLE_LEDGERACCOUNTS), 7) + PlGrensVan
            If Ktrl Then
                MsgBox "Geen journalen voor deze periode..."
            Else
                RecordToVeld TABLE_JOURNAL
                If Left(KEY_BUF(TABLE_JOURNAL), 7) <= GroepSelektie Then
                Else
                    Exit Do
                End If
                Do While Right(KEY_BUF(TABLE_JOURNAL), 8) <= PlGrensTot
                    SnelHelpPrint "Alle journalen voor rekening " + KEY_BUF(TABLE_LEDGERACCOUNTS) + " worden gekontroleerd.  Bezig aan :" + Right(KEY_BUF(TABLE_JOURNAL), 8), BL_LOGGING
                    XDoEvents = DoEvents
                    If Trim$(vBibTekst(TABLE_JOURNAL, "#v033 #")) = "" Or Left(vBibTekst(TABLE_JOURNAL, "#v033 #"), 1) = "D" Then
                        'Debug.Print rsMAR(TABLE_JOURNAL)("v033"), Val(rsMAR(TABLE_JOURNAL)("v068")), rsMAR(TABLE_JOURNAL)("v067"), rsMAR(TABLE_JOURNAL)("v019")
                        AantalZonderdokument = AantalZonderdokument + 1
                        
                        BedragZonderdokument = BedragZonderdokument + Val(vBibTekst(TABLE_JOURNAL, "#v068 #"))
                    ElseIf RTrim$(vBibTekst(TABLE_JOURNAL, "#v038 #")) <> "" Then
                        bGet TABLE_INVOICES, 0, vBibTekst(TABLE_JOURNAL, "#v033 #")
                        If Ktrl Then
                        Else
                            RecordToVeld TABLE_INVOICES
                            If vBibTekst(TABLE_INVOICES, "#v035 #") < PlGrensVan Then
                                BetaaldBedragBB = Val(vBibTekst(TABLE_JOURNAL, "#v068 #"))
                                If FlPartij = TABLE_CUSTOMERS Then
                                    BetaaldBedragBB = -BetaaldBedragBB
                                End If
                                If XisEuroWisBEF = True Then
                                    If vBibTekst(TABLE_JOURNAL, "#v066 #") > Right(BOOKYEAR_FROMTO, 8) Then
                                        BetaaldBedragBB = Round(BetaaldBedragBB * EURO)
                                    End If
                                End If
                                
                                TotaalDokOnderPeriode = TotaalDokOnderPeriode + BetaaldBedragBB
                                AantalOnderPeriode = AantalOnderPeriode + 1
                                Msg = vSet(Mid(vBibTekst(TABLE_INVOICES, "#v034 #"), 2), 12) + " "
                                Msg = Msg + vSet(vBibTekst(TABLE_JOURNAL, "#v067 #"), 20) + " "
                                Msg = Msg + DATE_TEXT(vBibTekst(TABLE_INVOICES, "#v035 #")) + " "
                                Msg = Msg + vBibTekst(TABLE_INVOICES, "#v033 #") + " "
                                Msg = Msg + Dec(BetaaldBedragBB, MASK_2002) + " "
                                Msg = Msg + vSet(vBibTekst(TABLE_JOURNAL, "#v019 #"), 7) + " "
                                Msg = Msg + DATE_TEXT(vBibTekst(TABLE_JOURNAL, "#v066 #")) + " "
                                Msg = Msg + vBibTekst(TABLE_JOURNAL, "#v038 #") + " "
                                Msg = Msg + vBibTekst(TABLE_JOURNAL, "#v069 #")
                                BetalingenVoorNa(0).AddItem Msg
                            ElseIf vBibTekst(TABLE_INVOICES, "#v035 #") > PlGrensTot Then
                                BetaaldBedragBB = Val(vBibTekst(TABLE_JOURNAL, "#v068 #"))
                                If FlPartij = TABLE_CUSTOMERS Then
                                    BetaaldBedragBB = -BetaaldBedragBB
                                End If
                                If XisEuroWisBEF = True Then
                                    If vBibTekst(TABLE_JOURNAL, "#v066 #") > Right(BOOKYEAR_FROMTO, 8) Then
                                        BetaaldBedragBB = Round(BetaaldBedragBB * EURO)
                                    End If
                                End If
                                
                                TotaalDokBovenPeriode = TotaalDokBovenPeriode + BetaaldBedragBB
                                AantalBovenPeriode = AantalBovenPeriode + 1
                                Msg = vSet(Mid(vBibTekst(TABLE_INVOICES, "#v034 #"), 2), 12) + " "
                                Msg = Msg + vSet(vBibTekst(TABLE_JOURNAL, "#v067 #"), 20) + " "
                                Msg = Msg + DATE_TEXT(vBibTekst(TABLE_INVOICES, "#v035 #")) + " "
                                Msg = Msg + vBibTekst(TABLE_INVOICES, "#v033 #") + " "
                                Msg = Msg + Dec(BetaaldBedragBB, MASK_2002) + " "
                                Msg = Msg + vSet(vBibTekst(TABLE_JOURNAL, "#v019 #"), 7) + " "
                                Msg = Msg + DATE_TEXT(vBibTekst(TABLE_JOURNAL, "#v066 #")) + " "
                                Msg = Msg + vBibTekst(TABLE_JOURNAL, "#v038 #") + " "
                                Msg = Msg + vBibTekst(TABLE_JOURNAL, "#v069 #")
                                BetalingenVoorNa(1).AddItem Msg
                            End If
                        End If
                    End If
                    bNext TABLE_JOURNAL
                    If Ktrl Or KEY_BUF(TABLE_JOURNAL) > GroepRekening4 + PlGrensTot Then
                        Exit Do
                    Else
                        RecordToVeld TABLE_JOURNAL
                    End If
                Loop
            End If
            bNext TABLE_LEDGERACCOUNTS
            If Ktrl Or KEY_BUF(TABLE_LEDGERACCOUNTS) > GroepSelektie Then
                Exit Do
            Else
                RecordToVeld TABLE_LEDGERACCOUNTS
            End If
        Loop
    End If

    For T = 0 To 10:
        VeldTXT(T) = ""
    Next
    VeldTXT(0) = "Stand " + Format(AantalInGroep, "00") + " " + GroepSelektie + "-rekeningen"
    VeldTXT(6) = Dec$((TotaalBedragGroep), MASK_2002)
    T = 0
    aa = ""
    Do While T < 12
        If chkAfdrukInVenster Then
            aa = aa + VeldTXT(T) + vbTab
        Else
            Printer.Print ; Tab(REPORT_TAB(T));
            Printer.Print ; VeldTXT(T);
        End If
        If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
            If chkAfdrukInVenster Then
            Else
                Printer.Print vbCrLf;
            End If
        End If
        T = T + 1
    Loop
    If chkAfdrukInVenster Then mdiXlog.X.AddItem aa, mdiXlog.X.Rows - 1
    
    For T = 0 To 10:
        VeldTXT(T) = ""
    Next
    VeldTXT(0) = Format(AantalZonderdokument, "00") + " verr. zonder dokument"
    VeldTXT(6) = Dec$((BedragZonderdokument), MASK_2002)
    T = 0
    aa = ""
    Do While T < 12
        If chkAfdrukInVenster Then
            aa = aa + VeldTXT(T) + vbTab
        Else
            Printer.Print ; Tab(REPORT_TAB(T));
            Printer.Print ; VeldTXT(T);
        End If
        If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
            If chkAfdrukInVenster Then
            Else
                Printer.Print vbCrLf;
            End If
        End If
        T = T + 1
    Loop
    If chkAfdrukInVenster Then mdiXlog.X.AddItem aa, mdiXlog.X.Rows - 1

    For T = 0 To 10:
        VeldTXT(T) = ""
    Next
    VeldTXT(0) = Format(AantalOnderPeriode, "#") + " betalingen dok. < boekjaar"
    VeldTXT(6) = Dec$((TotaalDokOnderPeriode), MASK_2002)
    T = 0
    aa = ""
    Do While T < 12
        If chkAfdrukInVenster Then
            aa = aa + VeldTXT(T) + vbTab
        Else
            Printer.Print ; Tab(REPORT_TAB(T));
            Printer.Print ; VeldTXT(T);
        End If
        If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
            If chkAfdrukInVenster Then
            Else
                Printer.Print vbCrLf;
            End If
        End If
        T = T + 1
    Loop
    If chkAfdrukInVenster Then mdiXlog.X.AddItem aa, mdiXlog.X.Rows - 1
    
    For T = 0 To 10:
        VeldTXT(T) = ""
    Next
    VeldTXT(0) = Format(AantalBovenPeriode, "#") + " betalingen dok. > boekjaar"
    VeldTXT(6) = Dec$((TotaalDokBovenPeriode), MASK_2002)
    T = 0
    aa = ""
    Do While T < 12
        If chkAfdrukInVenster Then
            aa = aa + VeldTXT(T) + vbTab
        Else
            Printer.Print ; Tab(REPORT_TAB(T));
            Printer.Print ; VeldTXT(T);
        End If
        If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
            If chkAfdrukInVenster Then
            Else
                Printer.Print vbCrLf;
            End If
        End If
        T = T + 1
    Loop
    If chkAfdrukInVenster Then mdiXlog.X.AddItem aa, mdiXlog.X.Rows - 1
End If

If chkAfdrukInVenster Then
    'mdiXlog!Kopie(1).Checked = True
    mdiXlog.X.Row = 1
    mdiXlog.X.Col = 0
    mdiXlog.X.ColWidth(0) = 2385
    mdiXlog.X.ColWidth(1) = 1185
    mdiXlog.X.ColWidth(2) = 990
    mdiXlog.X.ColWidth(3) = 390
    mdiXlog.X.ColWidth(4) = 945
    mdiXlog.X.ColWidth(5) = 855
    mdiXlog.X.ColWidth(6) = 855
    mdiXlog.X.ColWidth(7) = 855
    mdiXlog.X.ColWidth(8) = 915
    mdiXlog.X.ColWidth(9) = 975
    mdiXlog.X.ColWidth(10) = 960
    mdiXlog.WindowState = 2

    mdiXlog.X.ColAlignment(4) = flexAlignRightTop
    mdiXlog.X.ColAlignment(5) = flexAlignRightTop
    mdiXlog.X.ColAlignment(6) = flexAlignRightTop
    mdiXlog.X.ColAlignment(8) = flexAlignRightTop
    mdiXlog.X.ColAlignment(10) = flexAlignRightTop
    
    Screen.MousePointer = vbNormal
    '---------------------------------------------------
    'TODO everywhere when choosen show print in a window
    mdiXlog.WindowState = vbNormal
    'mdiXlog.statusForm = ""
    mdiXlog.Width = 11745
    '---------------------------------------------------
        
    mdiXlog.Show '1
    mdiXlog.WindowState = 0
    'Unload mdiXlog
End If

End Sub
Private Sub Form_Load()

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If

If bhEuro Then
    cmdEuroCheck.Visible = True
Else
    cmdEuroCheck.Visible = False
End If

Dim TempB_MODUS As Integer

PlGrensVan = Left(BOOKYEAR_FROMTO, 8)
PlGrensTot = Right(BOOKYEAR_FROMTO, 8)

cbLijstPrinterHier.Clear
If Printers.Count = 0 Then MsgBox "Installeer eerst minstens één printerdriver a.u.b. (bvb. cutePDF)", vbExclamation: End
For COUNT_TO = 0 To Printers.Count - 1
    Me.cbLijstPrinterHier.AddItem Printers(COUNT_TO).Port + " " + Printers(COUNT_TO).DeviceName
Next
cbLijstPrinterHier.ListIndex = LijstPrinterNr
Set Printer = Printers(LijstPrinterNr)

On Error Resume Next
Printer.PaperBin = LaadTekst(App.Title, "LIJSTPRINTER")
If Printer.Orientation = vbPRORLandscape Then
    chkAfdrukLiggend.Value = 1
Else
    chkAfdrukLiggend.Value = 0
End If
chkAfdrukLiggend_Click

On Local Error Resume Next
Err = 0
Selektie(5) = LaadTekst("dokumentenBalans", "BetalingsKontrole")
If Err Then
Else
    Selektie(0) = LaadTekst("dokumentenBalans", "KontroleVervaldag")
    Selektie(1) = LaadTekst("dokumentenBalans", "GeenBetalingHogerBoekjaar")
    Selektie(2) = LaadTekst("dokumentenBalans", "PeriodeBegrenzen")
    Selektie(3) = LaadTekst("dokumentenBalans", "EnkelNietBetaaldedokumenten")
    Selektie(4) = LaadTekst("dokumentenBalans", "-dokumenten-1994Uitsluiten")
    ChkFinancieelDetail = LaadTekst("dokumentenBalans", "FinancieelDetailViaJournaal")
    chkAfdrukInVenster = LaadTekst("dokumentenBalans", "AfdrukInRooster")
End If
TekstLijn(1).text = MIM_GLOBAL_DATE
TekstLijn(4).text = DATE_TEXT(Left(BOOKYEAR_FROMTO, 8)) + " - " + DATE_TEXT(Right(BOOKYEAR_FROMTO, 8))

Select Case aIndex
    Case TABLE_SUPPLIERS
        LijstNaam = "Balans Leveranciers"
        FlPartij = TABLE_SUPPLIERS
    Case TABLE_CUSTOMERS
        LijstNaam = "Balans Klanten"
        FlPartij = TABLE_CUSTOMERS
    Case Else
        MsgBox "stop balans partijen!"
End Select

If XisEuroWisBEF = True Then
    LijstNaam = LijstNaam & " (Speciale modus: Alle cijfers in BEF !)"
End If
BalansKL.Caption = LijstNaam

TekstLijn(2).text = "0"
TekstLijn(3).text = String$(12, "z")

End Sub


Private Sub InitVelden()
Dim T As Integer
Dim VolgTab As Integer

REPORT_FIELD(0) = "Identiteit"
    REPORT_TAB(0) = 1

REPORT_FIELD(1) = "Document"
    REPORT_TAB(1) = 29

REPORT_FIELD(2) = "Datum"
    REPORT_TAB(2) = 41

REPORT_FIELD(3) = "Mdoc"
    REPORT_TAB(3) = 52

If XisEuroWisBEF = True Then
    REPORT_FIELD(4) = " Goed(BEF)"
Else
    If bhEuro Then
        REPORT_FIELD(4) = " Goed(EUR)"
    Else
        REPORT_FIELD(4) = " Goed(BEF)"
    End If
End If
REPORT_TAB(4) = 56

If XisEuroWisBEF = True Then
    REPORT_FIELD(5) = "  BTW(BEF)"
Else
    If bhEuro Then
        REPORT_FIELD(5) = "  BTW(EUR)"
    Else
        REPORT_FIELD(5) = "  BTW(BEF)"
    End If
End If
REPORT_TAB(5) = 67

REPORT_FIELD(6) = "   Betaald"
    REPORT_TAB(6) = 78

REPORT_FIELD(7) = "Fin.stuk"
    REPORT_TAB(7) = 89

REPORT_FIELD(8) = "      Rest"
    REPORT_TAB(8) = 98

REPORT_FIELD(9) = "Vervaldag"
    REPORT_TAB(9) = 109

If XisEuroWisBEF = True Then
    REPORT_FIELD(10) = " Cum.(BEF)"
Else
    If bhEuro Then
        REPORT_FIELD(10) = " Cum.(EUR)"
    Else
        REPORT_FIELD(10) = " Cum.(BEF)"
    End If
End If
REPORT_TAB(10) = 119

REPORT_TAB(11) = 0

If chkAfdrukInVenster Then
    Me.Hide
    Unload mdiXlog
    mdiXlog.Hide
    mdiXlog.Caption = "documentenBalans"
    mdiXlog.X.Cols = 11
    mdiXlog.X.Row = 0
    For T = 0 To 10
        mdiXlog.X.Col = T
        mdiXlog.X.text = REPORT_FIELD(T)
    Next
    Me.Show
End If

End Sub

Private Sub PrintTitel()
Dim T As Integer

If chkAfdrukInVenster Then Exit Sub

    If usrLicentieInfo <> "" Then
        Printer.CurrentX = 50
        Printer.CurrentY = 50
        Printer.Print usrLicentieInfo;
    End If
    PAGE_COUNTER = PAGE_COUNTER + 1
    Printer.CurrentY = 400
    Printer.Print Tab(1); psTekst(2);
    Printer.Print Tab(108); "Pagina : " + Dec$((PAGE_COUNTER), "##########"); vbCrLf;
    Printer.Print Tab(108); "Datum  : " + psTekst(0); vbCrLf; vbCrLf;
    Printer.Print Tab(1); UCase$(psTekst(3)); vbCrLf;

Printer.Print Tab(1); FULL_LINE; vbCrLf;

Do While REPORT_TAB(T) <> 0
    Printer.Print Tab(REPORT_TAB(T)); REPORT_FIELD(T);
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        Printer.Print vbCrLf;
    End If
    T = T + 1
Loop

Printer.Print Tab(1); FULL_LINE; vbCrLf; vbCrLf;

End Sub

Private Sub PrintVelden()
Dim T As Integer
Dim VeldTekst As String
Dim aa As String

T = 0
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
If Printer.CurrentY >= Printer.Height - 1200 Then
    On Local Error Resume Next
    Printer.NewPage
    Printer.FontSize = Printer.FontSize
    Printer.Print " "
    Printer.FontSize = Printer.FontSize
    PrintTitel
End If

If chkAfdrukInVenster Then mdiXlog.X.AddItem aa, mdiXlog.X.Rows - 1

End Sub

Private Sub Selektie_Click(Index As Integer)

Select Case Index
    Case 0
        If Selektie(Index).Value = 1 Then
            TekstLijn(1) = DATE_TEXT(Right(BOOKYEAR_FROMTO, 8))
        Else
            TekstLijn(1) = MIM_GLOBAL_DATE
        End If
    Case 4
        If Selektie(4).Value = 1 Then
            BeginPlaySound 9
            MsgBox "Schakel uitsluitend aan indien U problemen ondervindt met sommige geimporteerde DOS-dokumenten van voor 1994 (o.a. BTW 33 %, 8 % luxetaks)." + vbCrLf + vbCrLf + "Indien U alle mogelijkheden van marIntegraal Windows versie met uw oude data wenst te benutten, gelieve ons pér bedrijf een veiligheidskopij te bezorgen." + vbCrLf + vbCrLf + "Binnen uw servicecontract werken wij deze kosteloos om in onze lokalen."
        End If
End Select

End Sub

Private Sub TekstLijn_GotFocus(Index As Integer)

TekstLijn(Index).SelLength = Len(TekstLijn(Index).text)

End Sub

Private Sub TekstLijn_LostFocus(Index As Integer)

Select Case Index
    Case 1
        If DATE_INVALID((TekstLijn(1).text)) Then
            Beep
            TekstLijn(1).text = MIM_GLOBAL_DATE
        End If
    Case 4
        If DATE_INVALID(Right(TekstLijn(4).text, 10)) Then
            MsgBox "Respecteer : " + vbCrLf + vbCrLf + "DD/MM/EEJJ - DD/MM/EEJJ a.u.b. !"
            TekstLijn(4).text = DATE_TEXT(Left(BOOKYEAR_FROMTO, 8)) + " - " + DATE_TEXT(Right(BOOKYEAR_FROMTO, 8))
            TekstLijn(4).SetFocus
            Exit Sub
        ElseIf Len(TekstLijn(4).text) <> 23 Then
            MsgBox "Respecteer : " + vbCrLf + vbCrLf + "DD/MM/EEJJ - DD/MM/EEJJ a.u.b. !"
            TekstLijn(4).text = DATE_TEXT(Left(BOOKYEAR_FROMTO, 8)) + " - " + DATE_TEXT(Right(BOOKYEAR_FROMTO, 8))
            TekstLijn(4).SetFocus
            Exit Sub
        Else
            PlGrensVan = Mid(TekstLijn(4).text, 7, 4) + Mid(TekstLijn(4).text, 4, 2) + Mid(TekstLijn(4).text, 1, 2)
            PlGrensTot = Mid(TekstLijn(4).text, 20, 4) + Mid(TekstLijn(4).text, 17, 2) + Mid(TekstLijn(4).text, 14, 2)
            If BOOKYEAR_FROMTO = PlGrensVan + PlGrensTot Then
                Selektie(5) = 1
                Selektie(5).Visible = True
            Else
                Selektie(5) = 0
                Selektie(5).Visible = False
            End If
        End If
End Select

End Sub

