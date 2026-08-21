VERSION 5.00
Begin VB.Form DetailInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DetailInfo"
   ClientHeight    =   2265
   ClientLeft      =   1185
   ClientTop       =   2010
   ClientWidth     =   6570
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
   ScaleHeight     =   2265
   ScaleWidth      =   6570
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBank1 
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
      Left            =   4980
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Sepa Test"
      Top             =   1860
      Width           =   315
   End
   Begin VB.CommandButton cmdBank0 
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
      Left            =   2340
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Sepa Test"
      Top             =   1860
      Width           =   315
   End
   Begin VB.TextBox tbBank1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2700
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1860
      Width           =   2235
   End
   Begin VB.TextBox tbBank0 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   60
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1860
      Width           =   2235
   End
   Begin VB.CommandButton Balans 
      Caption         =   "Bala&nscontrole"
      Default         =   -1  'True
      Height          =   375
      Left            =   90
      TabIndex        =   16
      Top             =   90
      Width           =   1815
   End
   Begin VB.CommandButton OK 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5400
      TabIndex        =   13
      Top             =   60
      Width           =   1140
   End
   Begin VB.CommandButton Annuleren 
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
      Height          =   330
      Left            =   5400
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1140
   End
   Begin VB.TextBox TekstInfo0 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3900
      TabIndex        =   6
      Top             =   270
      Width           =   1395
   End
   Begin VB.TextBox TekstInfo2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3660
      TabIndex        =   8
      Top             =   870
      Width           =   1635
   End
   Begin VB.TextBox TekstInfo1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3660
      TabIndex        =   10
      Top             =   1170
      Width           =   1635
   End
   Begin VB.TextBox TekstInfo3 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2160
      TabIndex        =   12
      Top             =   1500
      Width           =   3135
   End
   Begin VB.TextBox TekstInfo5 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2100
      TabIndex        =   4
      Top             =   270
      Width           =   1740
   End
   Begin VB.CheckBox Bewerking 
      Caption         =   "= Ontvangst"
      Height          =   240
      Left            =   210
      TabIndex        =   0
      Top             =   600
      Value           =   1  'Checked
      Width           =   1635
   End
   Begin VB.CheckBox Dokument 
      Caption         =   "= Document"
      Height          =   240
      Left            =   210
      TabIndex        =   1
      Top             =   840
      Value           =   1  'Checked
      Width           =   1635
   End
   Begin VB.CheckBox Partij 
      Caption         =   "= Klant"
      Height          =   240
      Left            =   210
      TabIndex        =   2
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1635
   End
   Begin VB.CommandButton ZoekDokument 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "ZoekDokument"
      Height          =   240
      Left            =   3240
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1920
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   120
      Top             =   540
      Width           =   1815
   End
   Begin VB.Label LabelInfo5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Doc. nummer"
      Height          =   270
      Left            =   2100
      TabIndex        =   3
      Top             =   0
      Width           =   1230
   End
   Begin VB.Label LabelInfo0 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&TegenRekening"
      Height          =   270
      Left            =   3900
      TabIndex        =   5
      Top             =   0
      Width           =   1425
   End
   Begin VB.Label LabelInfo6 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2100
      TabIndex        =   15
      Top             =   570
      Width           =   3195
   End
   Begin VB.Label LabelInfo2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Betaling"
      Height          =   255
      Left            =   2130
      TabIndex        =   7
      Top             =   900
      Width           =   1425
   End
   Begin VB.Label LabelInfo1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Financ. &Korting"
      Height          =   255
      Left            =   2130
      TabIndex        =   9
      Top             =   1200
      Width           =   1425
   End
   Begin VB.Label LabelInfo3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Omschrijvin&G"
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   1500
      Width           =   1185
   End
End
Attribute VB_Name = "DetailInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim Tegenrekening As String * 7
Dim DefaultKlanten As String * 7
Dim DefaultLeveranciers As String * 7

Private Sub Annuleren_Click()

On Local Error Resume Next

GridText = ""
DetailInfo.Hide
InbrengFinancieel.SetFocus

End Sub

Private Sub Balans_Click()

Balans.Default = True
If Partij.Value = 1 Then
    SharedFl = TABLE_CUSTOMERS
Else
    SharedFl = TABLE_SUPPLIERS
End If
GridText = ""
aIndex = 1
SqlSearch.Show 1
If Ktrl = 0 Then
    RecordToVeld SharedFl
    KTRLBalans SharedFl
End If

End Sub

Private Sub Bewerking_Click()

If Bewerking.Value = 1 Then
    Bewerking.Caption = "= Ontvangst"
    Partij.Value = 1
Else
    Bewerking.Caption = "= Uitgave"
    Partij.Value = 0
End If
Partij_click
On Local Error Resume Next
Dokument.SetFocus

End Sub

Private Sub cmdBank0_Click()

    If Me.tbBank0 = "" Then
    Else
        MsgBox "test voor sepa webservice" '& vbCrLf & vbCrLf & Mim.sepaIUC.SSS(Me.tbBank0.Text), vbInformation
    End If
    
End Sub

Private Sub cmdBank1_Click()

    If Me.tbBank1 = "" Then
    Else
        MsgBox "test voor sepa webservice" '& vbCrLf & vbCrLf & Mim.sepaIUC.SSS(Me.tbBank1.Text), vbInformation
    End If
    
End Sub

Private Sub dokument_Click()

If Dokument.Value = 1 Then
    Dokument.Caption = "= dokument"
    Balans.Enabled = True
    Balans.Default = True
    TekstInfo1.Visible = True
    LabelInfo1.Visible = True
    TekstInfo5.Visible = True
    Partij.Visible = True
    If Partij.Value = 1 Then
        TekstInfo0.text = DefaultKlanten
    Else
        TekstInfo0.text = DefaultLeveranciers
    End If
Else
    Dokument.Caption = "= Geen DOK"
    Partij.Visible = False
    TekstInfo1.Visible = False
    LabelInfo1.Visible = False
    TekstInfo5.Visible = False
    Balans.Enabled = False
    TekstInfo0.text = ""
    TekstInfo1.text = ""
End If

End Sub

Private Sub dokument_KeyPress(KeyAscii As Integer)

On Local Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{TAB}", False 'True
End If

End Sub

Private Sub Form_Load()

DefaultKlanten = String99(READING, 9)
DefaultLeveranciers = String99(READING, 10)

TekstInfo0.text = DefaultKlanten

End Sub

Private Sub KTRLBalans(Fl As Integer)
Dim Cumul       As Double
Dim dTotaal     As Double
Dim dBetaald    As Double
Dim aa          As String
Dim AAA         As String
Dim T           As Integer
Dim VoorLetter  As String * 1

Dim TotaalBedrag As Currency
Dim TotaalBETAALD As Currency

bClose TABLE_INVOICES
Select Case Fl
    Case TABLE_CUSTOMERS
        VoorLetter = "K"
    Case TABLE_SUPPLIERS
        VoorLetter = "L"
End Select
Unload Xlog
Load Xlog
Xlog.Caption = "Betaalbalans voor : " + vBibTekst(Fl, "#A100 #")
Xlog.X.Rows = 1
Xlog.X.Cols = 6
Xlog.X.Col = 0
Xlog.X.Row = 0
Xlog.X.text = "dokument"
Xlog.X.Col = 1
Xlog.X.text = "Totaal"
Xlog.X.Col = 2
Xlog.X.text = "Datum"
Xlog.X.Col = 3
Xlog.X.text = "Fin.Stuk"
Xlog.X.Col = 4
Xlog.X.text = "Betaald"
Xlog.X.Col = 5
Xlog.X.text = "CumulRest"

aa = ""
T = 0
bGetOrGreater TABLE_INVOICES, 1, vSet(VoorLetter + vBibTekst(Fl, "#A110 #"), FLINDEX_LEN(TABLE_INVOICES, 1))
If Ktrl Then
    Beep
    Exit Sub
Else
    RecordToVeld TABLE_INVOICES
End If
If vSet(KEY_BUF(TABLE_INVOICES), 13) <> vSet(VoorLetter + vBibTekst(Fl, "#A110 #"), 13) Then
    Beep
    MsgBox "Geen dokumenten voor " + vBibTekst(Fl, "#A100 #")
    Exit Sub
Else
    Screen.MousePointer = vbHourglass
    GoSub VolgendeLijn
    Do
        bNext TABLE_INVOICES
        If Ktrl Or vSet(KEY_BUF(TABLE_INVOICES), 13) <> vSet(VoorLetter + vBibTekst(Fl, "#A110 #"), 13) Then
            Exit Do
        Else
            RecordToVeld TABLE_INVOICES
            GoSub VolgendeLijn
        End If
    Loop
End If
Screen.MousePointer = vbNormal
If Xlog.X.Rows <> 1 Then
Else
    MsgBox "Alles is betaald voor/door" + vbCrLf + vBibTekst(Fl, "#A100 #"), 0
    GoTo SluitAf
End If

Xlog.X.Row = 1
Xlog.X.Col = 0

Xlog.X.ColWidth(0) = 1215
Xlog.X.ColWidth(1) = 1125
Xlog.X.ColWidth(2) = 930
Xlog.X.ColWidth(3) = 1005
Xlog.X.ColWidth(4) = 975
Xlog.X.ColWidth(5) = 1185
Xlog.X.ColAlignment(1) = flexAlignRightTop
Xlog.X.ColAlignment(4) = flexAlignRightTop
Xlog.X.ColAlignment(5) = flexAlignRightTop
Xlog!Afsluiten.Default = True

XLogShow:
Xlog.WijzigenLijn.Visible = False
Xlog.Afsluiten.TabStop = False
Xlog.cbAfbeelding.Visible = False

XLogKey = ""
Xlog.SSTab1.TabVisible(1) = False
Xlog.Show 1
Unload Xlog
If XLogKey <> "" Then
    bGet TABLE_INVOICES, 0, PartLeft(XLogKey, InStr(XLogKey, Chr$(13)) - 1)
    If Ktrl Then
    Else
        RecordToVeld TABLE_INVOICES
        TekstInfo1.text = ""
        If XisEuroWisBEF = True Then
            TekstInfo2.text = Str(Round(Val(vBibTekst(TABLE_INVOICES, "#v249 #")) * EURO) - Round(Val(vBibTekst(TABLE_INVOICES, "#v037 #")) * EURO))
        Else
            TotaalBedrag = Val(vBibTekst(TABLE_INVOICES, "#v249 #"))
            TotaalBETAALD = Val(vBibTekst(TABLE_INVOICES, "#v037 #"))
            
            TekstInfo2.text = Dec(TotaalBedrag - TotaalBETAALD, "")
        End If
        TekstInfo3.text = vBibTekst(Fl, "#A100 #")
        tbBank0.text = vBibTekst(Fl, "#A170 #")
        tbBank1.text = vBibTekst(Fl, "#v251 #")
        TekstInfo5.text = vBibTekst(TABLE_INVOICES, "#v033 #")
        TekstInfo5.Enabled = False
        Bewerking.Enabled = False
        Dokument.Enabled = False
        Partij.Enabled = False
        If OK.Enabled = False Then
            OK.Enabled = True
        End If
        OK.SetFocus
    End If
End If
Exit Sub

VolgendeLijn:
Select Case PartLeft(vBibTekst(TABLE_INVOICES, "#v033 #"), 2)
    Case "A0", "V1"
        If Bewerking.Value = 1 Then Return
    Case "A1", "V0"
        If Bewerking.Value = 0 Then Return
End Select

AAA = ""
Select Case Fl
    Case TABLE_CUSTOMERS
        AAA = vBibTekst(TABLE_INVOICES, "#v033 #") & vbTab
        dBetaald = Val(vBibTekst(TABLE_INVOICES, "#v037 #"))
        If XisEuroWisBEF = True Then
            dBetaald = Round(dBetaald * EURO)
        End If
        Select Case PartLeft(vBibTekst(TABLE_INVOICES, "#v033 #"), 1)
            Case "V"
                dTotaal = Val(vBibTekst(TABLE_INVOICES, "#v249 #"))
                If XisEuroWisBEF = True Then
                    MsgBox "CTRLstop"
                    dTotaal = Round(dTotaal * EURO)
                End If
                If PartMid(vBibTekst(TABLE_INVOICES, "#v033 #"), 2, 1) = "1" Then
                    dTotaal = -dTotaal
                    dBetaald = -dBetaald
                End If
            Case "Q"
                dTotaal = Val(vBibTekst(TABLE_INVOICES, "#v249 #"))
                If XisEuroWisBEF = True Then
                    dTotaal = Round(dTotaal * EURO)
                End If
        End Select
        AAA = AAA & Format(dTotaal, "#,##0.00") & vbTab
        AAA = AAA & DATE_TEXT(vBibTekst(TABLE_INVOICES, "#v035 #")) & vbTab
        AAA = AAA & vBibTekst(TABLE_INVOICES, "#v038 #") & vbTab
        AAA = AAA & Format(dBetaald, "#,##0.00") & vbTab
        Cumul = Cumul + dTotaal - dBetaald
        AAA = AAA & Format(Cumul, "#,##0.00")

    Case TABLE_SUPPLIERS
        AAA = vBibTekst(TABLE_INVOICES, "#v033 #") & vbTab
        dBetaald = Val(vBibTekst(TABLE_INVOICES, "#v037 #"))
        If XisEuroWisBEF = True Then
            dBetaald = Round(dBetaald * EURO)
        End If
        dTotaal = Val(vBibTekst(TABLE_INVOICES, "#v249 #"))
        If XisEuroWisBEF = True Then
            dTotaal = Round(dTotaal * EURO)
        End If
        
        AAA = AAA & Format(dTotaal, "#,##0.00") & vbTab
        AAA = AAA & DATE_TEXT(vBibTekst(TABLE_INVOICES, "#v035 #")) & vbTab
        AAA = AAA & vBibTekst(TABLE_INVOICES, "#v038 #") & vbTab
        AAA = AAA & Format(dBetaald, "#,##0.00") & vbTab
        If PartLeft(vBibTekst(TABLE_INVOICES, "#v033 #"), 2) = "A1" Then
            dTotaal = -dTotaal
            dBetaald = -dBetaald
        End If
        Cumul = Cumul + dTotaal - dBetaald
        AAA = AAA & Format(Cumul, "#,##0.00")
End Select
If dBetaald = dTotaal Then
Else
    T = T + 1
    Xlog.X.AddItem AAA
End If
Return

SluitAf:
Unload Xlog
Exit Sub

End Sub



Private Sub Ok_Click()
Dim Kolom1 As String * 12
Dim Kolom2 As String * 7
Dim Kolom3 As String * 12
Dim Kolom4 As String * 29
Dim Kolom5 As String * 12
Dim kVlag As String * 1

If TekstInfo0.text = "" Then
    MsgBox "Tegenrekening aanduiden a.u.b. !"
    TekstInfo0.SetFocus
    Exit Sub
End If
If Dokument.Value = 1 Then
    bGet TABLE_INVOICES, 0, vSet((TekstInfo5.text), 11)
    If Ktrl Then
        MsgBox "dokumentnummer onbekend !!!"
        Exit Sub
    End If
End If
If Val(TekstInfo1.text) + Val(TekstInfo2.text) = 0 Then
    MsgBox "Bedrag - of + inbrengen a.u.b. !!!"
    Exit Sub
End If

If Bewerking.Value = 1 Then
    kVlag = "+"
Else
    kVlag = "-"
End If

If Dokument.Value = 0 Then
    Kolom1 = kVlag
Else
    Kolom1 = kVlag + TekstInfo5.text
End If
Kolom2 = TekstInfo0.text
Kolom3 = Dec(Val(TekstInfo2.text), MASK_EURBH)
Kolom4 = TekstInfo3.text
Kolom5 = Dec(Val(TekstInfo1.text), MASK_EURBH)
GridText = Kolom1 + Chr$(124) + Kolom2 + Chr$(124) + Kolom3 + Chr$(124) + Kolom4 + Chr$(124) + Kolom5
DetailInfo.Hide
InbrengFinancieel.SetFocus

End Sub

Private Sub Partij_click()

If Partij.Value = 1 Then
    Partij.Caption = "= Klant"
    TekstInfo0.text = DefaultKlanten
Else
    Partij.Caption = "= Leverancier"
    TekstInfo0.text = DefaultLeveranciers
End If

End Sub

Private Sub TekstInfo0_GotFocus()

    Balans.Default = False
    TekstInfo0.SelStart = 0
    TekstInfo0.SelLength = Len(TekstInfo0.text)
    SnelHelpPrint "[Ctrl] voor geïndexeerd zoeken", BL_LOGGING
    Bewerking.Enabled = False
    Dokument.Enabled = False
    Partij.Enabled = False

End Sub

Private Sub TekstInfo1_GotFocus()

    Balans.Default = False
    TekstInfo1.SelStart = 0
    TekstInfo1.SelLength = Len(TekstInfo1.text)

End Sub

Private Sub TekstInfo2_GotFocus()

    Balans.Default = False
    TekstInfo2.SelStart = 0
    TekstInfo2.SelLength = Len(TekstInfo2.text)

End Sub

Private Sub TekstInfo3_GotFocus()

    Balans.Default = False
    TekstInfo3.SelStart = 0
    TekstInfo3.SelLength = Len(TekstInfo3.text)

End Sub


Private Sub TekstInfo5_GotFocus()

    Balans.Default = False
    TekstInfo5.SelStart = 0
    TekstInfo5.SelLength = Len(TekstInfo5.text)
    ZoekDokument.Enabled = True
    ZoekDokument.Default = True

End Sub


Private Sub TekstInfo0_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 17 Then
        SharedFl = TABLE_LEDGERACCOUNTS
        aIndex = 0
        GridText = TekstInfo0.text
        SqlSearch.Show 1
        If Ktrl = 0 Then
            TekstInfo0.text = FVT(TABLE_LEDGERACCOUNTS, 0)
        Else
            TekstInfo0.text = ""
        End If
    ElseIf KeyCode = 13 Then
        TekstInfo2.SetFocus
    End If

End Sub

Private Sub TekstInfo2_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        TekstInfo3.SetFocus
    End If

End Sub


Private Sub TekstInfo3_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        OK.SetFocus
    End If

End Sub

Private Sub TekstInfo0_KeyPress(KeyAscii As Integer)
    
    Dim Kontrole As String
    Kontrole = "*.0123456789" + Chr$(8)
        
End Sub

Private Sub TekstInfo2_KeyPress(KeyAscii As Integer)
    
    Dim Kontrole As String
    Kontrole = "*.0123456789" + Chr$(8)
        
End Sub

Private Sub TekstInfo3_KeyPress(KeyAscii As Integer)
    
    Dim Kontrole As String
    Kontrole = "*.0123456789" + Chr$(8)
        
End Sub


Private Sub TekstInfo5_KeyPress(KeyAscii As Integer)

    Dim Kontrole As String
    Kontrole = "*.0123456789" + Chr$(8)

    If InStr(Kontrole, Chr$(KeyAscii)) Then
        OK.Enabled = False
    Else
        KeyAscii = 0
        Msg = "Syntax : [*] . xxxxx . [eejj]" + vbCrLf + vbCrLf
        Msg = Msg + "*" + vbCrLf
        Msg = Msg + "'Q0'-Kwijting klanten (makelaars), programma maakt" + vbCrLf
        Msg = Msg + "anders 'V0'-sleutel, ontvangst van klantfaktuur, 'V1'" + vbCrLf
        Msg = Msg + "uitgave creditnota, 'A0' faktuur leverancier of 'A1'" + vbCrLf
        Msg = Msg + "creditnota leverancier" + vbCrLf + vbCrLf
        Msg = Msg + "xxxxx" + vbCrLf
        Msg = Msg + "dokument volgnummer (verplicht) van 1 tot 99999." + vbCrLf + vbCrLf
        Msg = Msg + "eejj" + vbCrLf
        Msg = Msg + "Eeuw dokument van 1900 tot max 2099 (optioneel)" + vbCrLf + vbCrLf
        Msg = Msg + "Opties door '.' scheiden van elkaar a.u.b !" + vbCrLf + vbCrLf
        Msg = Msg + "Vb. *.542.1992 = kwijting Q0199200542 van '1992'" + vbCrLf
        Msg = Msg + "Vb. 542 zijnde aan- of verkoopdokument ??????00542 van 'huidig jaar'"
        MsgBox Msg, 0
    End If
        
End Sub

Private Sub TekstInfo0_LostFocus()
    Dim KontroleTekst As String
    Dim DokType As String * 2
    Dim Nummer As String * 5
    Dim Jaar As String * 1

    If TekstInfo0.text = "" Then
    Else
        bGet TABLE_LEDGERACCOUNTS, 0, vSet((TekstInfo0.text), 7)
        If Ktrl Then
            Beep
            TekstInfo0.text = ""
            LabelInfo6.Caption = ""
        Else
            RecordToVeld TABLE_LEDGERACCOUNTS
            LabelInfo6.Caption = vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #")
        End If
    End If

End Sub

Private Sub TekstInfo1_LostFocus()
    Dim KontroleTekst As String
    Dim DokType As String * 2
    Dim Nummer As String * 5
    Dim Jaar As String * 1
    
    If Val(TekstInfo1.text) <> 0 Then
        Msg = "Financiele korting van bedrag aftrekken."
        Ktrl = MsgBox(Msg, 292)
        If Ktrl = 6 Then
            TekstInfo2.text = Str(Val(TekstInfo2.text) - Val(TekstInfo1.text))
            TekstInfo1.Enabled = False
            TekstInfo2.Enabled = False
        Else
            TekstInfo1.text = ""
        End If
    End If

End Sub

Private Sub TekstInfo5_LostFocus()

    ZoekDokument.Enabled = False

End Sub


Private Sub Zoekdokument_Click()
Dim KontroleTekst As String
Dim DokType As String * 2
Dim Nummer As String * 5
Dim Jaar As String * 4
Dim dBetaald As Double
Dim dTotaal As Double


KontroleTekst = Trim$(TekstInfo5.text)
If KontroleTekst = "" Then
    Exit Sub
ElseIf PartLeft(KontroleTekst, 1) = "*" Then
    DokType = "Q0"
    KontroleTekst = PartRight(KontroleTekst, Len(KontroleTekst) - 1)
    If PartLeft(KontroleTekst, 1) = "." Then
        KontroleTekst = PartRight(KontroleTekst, Len(KontroleTekst) - 1)
    End If
Else
    DokType = ""
End If
If InStr(KontroleTekst, ".") Then
    Nummer = Format(Val(PartLeft(KontroleTekst, InStr(KontroleTekst, ".") - 1)), "00000")
    Jaar = PartRight(KontroleTekst, 4)
Else
    Nummer = Format(Val(KontroleTekst), "00000")
    Jaar = PartRight(MIM_GLOBAL_DATE, 4)
End If

If DokType = "Q0" Then
Else
    Select Case Partij.Value
        Case 1
            If Bewerking.Value = 1 Then
                DokType = "V0"
            Else
                DokType = "V1"
            End If
        Case Else
            If Bewerking.Value = 0 Then
                DokType = "A0"
            Else
                DokType = "A1"
            End If
    End Select
End If

bGet TABLE_INVOICES, 0, DokType + Jaar + Nummer
If Ktrl Then
    TekstInfo1.text = ""
    TekstInfo2.text = ""
    TekstInfo3.text = ""
    TekstInfo5.text = ""
    SnelHelpPrint DokType + Jaar + Nummer + " niet gevonden...", BL_LOGGING
    Exit Sub
Else
    RecordToVeld TABLE_INVOICES
    TekstInfo1.text = ""
    dBetaald = Val(vBibTekst(TABLE_INVOICES, "#v037 #"))
    If XisEuroWisBEF = True Then
        dBetaald = Round(dBetaald * EURO)
        MsgBox "CTRLstop"
    End If
    dTotaal = Val(vBibTekst(TABLE_INVOICES, "#v249 #"))
    If XisEuroWisBEF = True Then
        MsgBox "CTRLstop"
        dTotaal = Round(dTotaal * EURO)
    End If
    Select Case PartLeft(FVT(TABLE_INVOICES, 0), 1)
        Case "Q"
            'dTotaal = Val(FVT(TABLE_INVOICES, 5)) - Val(FVT(TABLE_INVOICES, 7)) - Val(FVT(TABLE_INVOICES, 9))
        Case "A"
            'dTotaal = Val(vBibTekst(TABLE_INVOICES, "#v048 #")) + Val(vBibTekst(TABLE_INVOICES, "#v049 #")) + Val(vBibTekst(TABLE_INVOICES, "#v047 #")) + Val(vBibTekst(TABLE_INVOICES, "#v046 #")) - Val(vBibTekst(TABLE_INVOICES, "#v044 #")) - Val(vBibTekst(TABLE_INVOICES, "#v043 #")) + Val(vBibTekst(TABLE_INVOICES, "#v045 #"))
            'Select Case partLeft(vBibTekst(TABLE_INVOICES, "#v035 #"), 4)
            '    Case Is < "1993"
            '    Case Else
                    'dTotaal = dTotaal - Val(vBibTekst(TABLE_INVOICES, "#v042 #"))
            'End Select
        Case "V"
            'dTotaal = Val(vBibTekst(TABLE_INVOICES, "#v064 #"))
            'dTotaal = dTotaal + Val(vBibTekst(TABLE_INVOICES, "#v056 #")) + Val(vBibTekst(TABLE_INVOICES, "#v057 #")) + Val(vBibTekst(TABLE_INVOICES, "#v058 #")) + Val(vBibTekst(TABLE_INVOICES, "#v062 #")) + Val(vBibTekst(TABLE_INVOICES, "#v055 #")) + Val(vBibTekst(TABLE_INVOICES, "#v059 #"))
            'If partLeft(vBibTekst(TABLE_INVOICES, "#v035 #"), 4) < "1993" Then
            '    dTotaal = dTotaal + Val(vBibTekst(TABLE_INVOICES, "#v060 #")) + Val(vBibTekst(TABLE_INVOICES, "#v063 #"))
            '    'opgelet ook oude 12 zou moeten bijgeteld worden !!
            'Else
            '    dTotaal = dTotaal + Val(vBibTekst(TABLE_INVOICES, "#v063 #"))
            'End If
    End Select
End If
If dTotaal - dBetaald <> 0 Then
    TekstInfo2.text = Str(dTotaal - dBetaald)
Else
           Msg = "Document reeds volledig betaald." + vbCrLf
    Msg$ = Msg + "Bedragen die nu bijgeteld worden," + vbCrLf
    Msg$ = Msg + "einde van het boekjaar rechtzetten !!!"
    MsgBox Msg, 0, "Dubbele betaling..."
End If

If PartLeft(FVT(TABLE_INVOICES, 0), 1) = "A" Then
    SharedFl = TABLE_SUPPLIERS
Else
    SharedFl = TABLE_CUSTOMERS
End If

bGet SharedFl, 0, Mid(vBibTekst(TABLE_INVOICES, "#v034 #"), 2)
If Ktrl Then
    Beep
Else
    RecordToVeld SharedFl
    TekstInfo3.text = vBibTekst(SharedFl, "#A100 #")
    tbBank0.text = vBibTekst(SharedFl, "#A170 #")
    tbBank1.text = vBibTekst(SharedFl, "#v251 #")
End If
TekstInfo5.text = FVT(TABLE_INVOICES, 0)
'TekstInfo5.Enabled = False
Bewerking.Enabled = False
Dokument.Enabled = False
Partij.Enabled = False
OK.Enabled = True
OK.SetFocus

End Sub

