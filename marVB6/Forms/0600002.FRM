VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form DiversePosten 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ctrl+F5 Diverse Posten"
   ClientHeight    =   7770
   ClientLeft      =   1080
   ClientTop       =   1305
   ClientWidth     =   7830
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
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7770
   ScaleWidth      =   7830
   Begin MSComCtl2.DTPicker Datum 
      Height          =   315
      Left            =   6300
      TabIndex        =   14
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   76546051
      CurrentDate     =   36328
      MinDate         =   35796
   End
   Begin VB.ComboBox SoortBoeking 
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
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   2685
   End
   Begin VB.CommandButton VolgendeLijn 
      Caption         =   "Bij&voegen"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4620
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Omschrijving 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Top             =   240
      Width           =   3375
   End
   Begin VB.OptionButton DCkeuze 
      Caption         =   "&Debiteren (+)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   2880
      TabIndex        =   3
      Top             =   540
      Value           =   -1  'True
      Width           =   1320
   End
   Begin VB.OptionButton DCkeuze 
      Caption         =   "&Crediteren (-)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   2880
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   780
      Width           =   1275
   End
   Begin VB.CheckBox TRvlag 
      Caption         =   "&Tegenrekening (/)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2880
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1020
      Width           =   1635
   End
   Begin VB.TextBox RekeningNummer 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   60
      TabIndex        =   7
      Top             =   1320
      Width           =   1200
   End
   Begin VB.TextBox Tegenrekening 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   60
      TabIndex        =   8
      Top             =   1620
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.TextBox Bedrag 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1260
      TabIndex        =   10
      Top             =   1020
      Width           =   1320
   End
   Begin VB.CommandButton Afsluiten 
      Caption         =   "Boeken"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4620
      TabIndex        =   12
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Sluiten 
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
      Height          =   315
      Left            =   4620
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Schoon 
      Caption         =   "Schoo&n"
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
      Left            =   4620
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1095
   End
   Begin VB.ListBox JournaalPost 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5460
      Left            =   0
      TabIndex        =   16
      Top             =   2220
      Width           =   7695
   End
   Begin VB.Label NaamTegenRekening 
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
      Height          =   315
      Left            =   1260
      TabIndex        =   24
      Top             =   1620
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label NaamRekening 
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
      Height          =   315
      Left            =   1260
      TabIndex        =   25
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Datu&m"
      Height          =   240
      Index           =   1
      Left            =   6360
      TabIndex        =   13
      Top             =   0
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "&Bedrag"
      Height          =   240
      Index           =   4
      Left            =   1920
      TabIndex        =   9
      Top             =   780
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "&Rekening"
      Height          =   210
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "&Omschrijving"
      Height          =   240
      Index           =   0
      Left            =   2850
      TabIndex        =   1
      Top             =   0
      Width           =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo nog toe te wijzen"
      Height          =   495
      Index           =   5
      Left            =   6240
      TabIndex        =   19
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "TegenR."
      Height          =   240
      Index           =   8
      Left            =   6540
      TabIndex        =   23
      Top             =   1980
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Bedrag"
      Height          =   240
      Index           =   7
      Left            =   5700
      TabIndex        =   22
      Top             =   1980
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "Naam"
      Height          =   240
      Index           =   6
      Left            =   1320
      TabIndex        =   21
      Top             =   1980
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "Nummer"
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   20
      Top             =   1980
      Width           =   840
   End
   Begin VB.Label Saldo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   315
      Left            =   6240
      TabIndex        =   18
      Top             =   1380
      Width           =   1485
   End
End
Attribute VB_Name = "DiversePosten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function AfschrijfBoeking() As Integer
Dim OmschrijvingsLijn   As String * 40
Dim LijnText As String
Dim Succes As Integer

On Error Resume Next

Succes = True
If String99(READING, 63) = "1" Then
    MsgBox "Afschrijvingsposten reeds gegenereerd voor dit boekjaar.  Bijkomende posten kunnen uitsluitend via 'Diverse post'-optie ingebracht worden !"
    AfschrijfBoeking = False
    Exit Function
ElseIf String99(READING, 63) <> "0" Then
    MsgBox "Setup boekjaar en parameters bevat niet de juiste vlag geboekt of niet geboekt.  Kontroleer"
    AfschrijfBoeking = False
    Exit Function
ElseIf String99(READING, 64) <> "1" Then
    MsgBox "Onlogische situatie.  Dit boekjaar bevat nog geen beginbalans ?  De beginbalans dient aanwezig te zijn.  Mogelijk bevindt U zich in het verkeerde boekjaar ?"
    AfschrijfBoeking = False
    Exit Function
End If

bGetOrGreater TABLE_VARIOUS, 1, vSet("18", 20)
If Ktrl Or Left(KEY_BUF(TABLE_VARIOUS), 2) <> "18" Then
    MsgBox "Er zijn geen investeringsfiches !"
    AfschrijfBoeking = False
    Exit Function
Else
'FIXIT: Keyword 'GoSub' not supported in Visual Basic .NET                                 FixIT90210ae-R6614-H1984
    GoSub AfschrijvingsLijnErBij
    Do
        bNext TABLE_VARIOUS
        If Ktrl Or Left(KEY_BUF(TABLE_VARIOUS), 2) <> "18" Then
            Exit Do
        Else
'FIXIT: Keyword 'GoSub' not supported in Visual Basic .NET                                 FixIT90210ae-R6614-H1984
            GoSub AfschrijvingsLijnErBij
        End If
    Loop
    If Succes Then
        Afsluiten.Enabled = True
        Omschrijving.text = SoortBoeking.text
        SoortBoeking.Enabled = False
        Afsluiten.SetFocus
    End If
    DCkeuze(0).Enabled = False
    DCkeuze(1).Enabled = False
    Bedrag.Visible = False
    RekeningNummer.Visible = False
    TRvlag.Enabled = False
    AfschrijfBoeking = True
End If
Exit Function

AfschrijvingsLijnErBij:
OmschrijvingsLijn = ""
RecordToVeld TABLE_VARIOUS
If Len(vBibTekst(TABLE_VARIOUS, "#v083 #")) <> 8 Then
    OmschrijvingsLijn = "Datumformaat onjuist voor " + vBibTekst(TABLE_VARIOUS, "#v087 #")
    MsgBox "dokumentendatum niet in formaat DDMMJJEE" + vbCrLf + vbCrLf + OmschrijvingsLijn + vbCrLf + vbCrLf + vBibTekst(TABLE_VARIOUS, "#v083 #")
    AfschrijfBoeking = False
    Exit Function
End If
    
If Mid(vBibTekst(TABLE_VARIOUS, "#v083 #"), 5, 4) > Mid(BOOKYEAR_FROMTO, 9, 4) Then
'FIXIT: Return has new meaning in Visual Basic .NET                                        FixIT90210ae-R9642-H1984
    Return
ElseIf Val(vBibTekst(TABLE_VARIOUS, "#v084 #")) = Val(vBibTekst(TABLE_VARIOUS, "#v085 #")) Then
    Msg = "Alles is afgeschreven voor " + vBibTekst(TABLE_VARIOUS, "#v087 #") + vbCrLf + vbCrLf + "Totaal : " + Format(Val(vBibTekst(TABLE_VARIOUS, "#v084 #")), MASK_SY(0)) + vbCrLf + vbCrLf
    If Not ADO_GET(TABLE_LEDGERACCOUNTS, 0, "=", vBibTekst(TABLE_VARIOUS, "#v087 #")) Then
        Msg = Msg + "Rekening bestaat zelfs niet eens..."
    Else
        Msg = Msg + RV(rsMAR(TABLE_LEDGERACCOUNTS), "v020")
    End If
    MsgBox Msg
'FIXIT: Return has new meaning in Visual Basic .NET                                        FixIT90210ae-R9642-H1984
    Return
Else
    Dim dRest As Double
    Dim ipct As Integer
    Dim dTas As Double
    Dim dbdrg As Double
    Dim dRa As Double
    Dim Das As Double
    Dim restKtrl As Double
    Dim MaskerHier As String
        
    ipct = Val(vBibTekst(TABLE_VARIOUS, "#v082 #"))
    dbdrg = Val(vBibTekst(TABLE_VARIOUS, "#v084 #"))
    dRa = Val(vBibTekst(TABLE_VARIOUS, "#v085 #"))
    
    If bhEuro = True Then
        restKtrl = 0.5
        MaskerHier = MASK_EURBH
    Else
        restKtrl = 20
        MaskerHier = MASK_SY(0)
    End If
    
    dRest = dbdrg - (dRa + Val(Dec((dbdrg / ipct), MaskerHier)))

    If dRest < restKtrl Then
        Das = Val(Dec((dbdrg / ipct), MaskerHier)) + dRest
        dTas = dTas + Das
    Else
        Das = Val(Dec((dbdrg / ipct), MaskerHier))
        dTas = dTas + Das
    End If
End If

If Not ADO_GET(TABLE_LEDGERACCOUNTS, 0, "=", vBibTekst(TABLE_VARIOUS, "#v087 #")) Then
    OmschrijvingsLijn = RTrim$(OmschrijvingsLijn) + " Afschr. op!!!"
End If
If Not ADO_GET(TABLE_LEDGERACCOUNTS, 0, "=", vBibTekst(TABLE_VARIOUS, "#v088 #")) Then
    OmschrijvingsLijn = RTrim$(OmschrijvingsLijn) + " Kostrekening !!"
Else
    If OmschrijvingsLijn = Space$(40) Then
        OmschrijvingsLijn = "Ok, " + RV(rsMAR(TABLE_LEDGERACCOUNTS), "v020")
    End If
End If

If Left(OmschrijvingsLijn, 3) = "Ok," Then
Else
    Succes = False
End If
LijnText = vSet(vBibTekst(TABLE_VARIOUS, "#v088 #"), 7) + " " + OmschrijvingsLijn + " " + Dec((Das), MASK_EURBH) + " " + vSet(vBibTekst(TABLE_VARIOUS, "#v087 #"), 7)
JournaalPost.AddItem LijnText
'FIXIT: Return has new meaning in Visual Basic .NET                                        FixIT90210ae-R9642-H1984
Return

End Function

Sub BoekUserDef()

End Sub

Private Sub Afsluiten_Click()

If Trim$(Omschrijving.text) = "" Then
    MsgBox "Omschrijving mag niet leeg zijn, ook géén spaties...", vbInformation
    Omschrijving.SetFocus
    Exit Sub
End If
    
If Not DATE_CHECK(Format(Datum.Value, "dd/mm/yyyy"), PERIODAS_TEXT) Then
    Beep
    Datum.SetFocus
    Exit Sub
ElseIf JournaalPost.ListCount = 0 Then
    Beep
    Exit Sub
End If

Ktrl = MsgBox("Journaalpost bestaande uit" + Str$(JournaalPost.ListCount) + " Lijnen wegboeken.  Bent U zeker ?", 292)
If Ktrl = 6 Then
    bBegin
    Screen.MousePointer = vbHourglass
    If BoekFout() Then
        bAbort
        Me.SetFocus
    Else
        Select Case Left(SoortBoeking.text, 1)
            Case "1"
                EindeAfschrijving
            Case "2"
                SS99 "1", 64
            Case "0"
            Case Else
                MsgBox "Stop"
        End Select
        bEnd
        Schoon_Click
    End If
'FIXIT: The use of 'vbNormal' is not valid for the property being assigned.                FixIT90210ae-R5049-R62328
    Screen.MousePointer = vbNormal
End If

End Sub

Private Sub Bedrag_Change()

If Bedrag.text <> "" Then
    VolgendeLijn.Enabled = True
    VolgendeLijn.Default = True
Else
    VolgendeLijn.Enabled = False
    VolgendeLijn.Default = False
End If

End Sub

Private Sub Bedrag_GotFocus()

If Val(Saldo.Caption) = 0 Then
    Afsluiten.Enabled = True
Else
    Afsluiten.Enabled = False
End If

End Sub

Private Sub Bedrag_KeyPress(KeyAscii As Integer)

If bhEuro Then
ElseIf KeyAscii = 46 Then
    Beep
    KeyAscii = 0
End If

End Sub

'FIXIT: Declare 'BoekBeginBalans' with an early-bound data type                            FixIT90210ae-R1672-R1B8ZE
Private Function BoekBeginBalans()

Dim pHier As String
Dim Verwittigen As Boolean

Verwittigen = False
If String99(READING, 64) = "1" Then
    MsgBox "Beginbalans reeds gegenereerd voor dit boekjaar.  Bijkomende posten kunnen uitsluitend via 'Diverse post'-optie ingebracht worden !"
    BoekBeginBalans = False
    Exit Function
ElseIf String99(READING, 64) <> "0" Then
    MsgBox "Setup beginbalans bevat niet de juiste vlag geboekt of niet geboekt.  Kontroleer"
    BoekBeginBalans = False
    Exit Function
End If

Dim TotaalBalans As Currency
Dim TotaalResultaat As Currency

'kontrole vanaf eerst 1 tot laatste 5
TotaalBalans = 0

'If Not ADO_GET(TABLE_LEDGERACCOUNTS, 0, "=", XLogKey) Then
'Else
'    RekeningNummer = RV(rsMAR(TABLE_LEDGERACCOUNTS), "v019")
'    NaamRekening = RV(rsMAR(TABLE_LEDGERACCOUNTS), "v020")

If Not ADO_GET(TABLE_LEDGERACCOUNTS, 0, ">=", vSet("1", 7)) Then
    MsgBox "Geen rekeningen ???"
    BoekBeginBalans = False
    Exit Function
ElseIf Left(RV(rsMAR(TABLE_LEDGERACCOUNTS), "v019"), 1) < "6" Then
    If bhEuro Then
        TotaalBalans = TotaalBalans + Val(RV(rsMAR(TABLE_LEDGERACCOUNTS), "e" + Format(ACTIVE_BOOKYEAR + 23, "000")))
    Else
        TotaalBalans = TotaalBalans + Val(RV(rsMAR(TABLE_LEDGERACCOUNTS), "v" + Format(ACTIVE_BOOKYEAR + 23, "000")))
    End If
    Do While Not rsMAR(TABLE_LEDGERACCOUNTS).EOF
        rsMAR(TABLE_LEDGERACCOUNTS).MoveNext
        If Left(RV(rsMAR(TABLE_LEDGERACCOUNTS), "v019"), 1) >= "6" Then
            Exit Do
        Else
            If bhEuro Then
                TotaalBalans = TotaalBalans + Val(RV(rsMAR(TABLE_LEDGERACCOUNTS), "e" + Format(ACTIVE_BOOKYEAR + 23, "000")))
            Else
                TotaalBalans = TotaalBalans + Val(RV(rsMAR(TABLE_LEDGERACCOUNTS), "v" + Format(ACTIVE_BOOKYEAR + 23, "000")))
            End If
        End If
    Loop
End If

'+ eerste 6 tot laatste 7
TotaalResultaat = 0
bGetOrGreater TABLE_LEDGERACCOUNTS, 0, vSet("6", 7)
If Ktrl Then
    MsgBox "Geen rekeningen ???"
    BoekBeginBalans = False
    Exit Function
ElseIf Left(KEY_BUF(TABLE_LEDGERACCOUNTS), 1) < "8" Then
    RecordToVeld TABLE_LEDGERACCOUNTS
    If bhEuro Then
        TotaalResultaat = TotaalResultaat + Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#e" + Format(ACTIVE_BOOKYEAR + 23, "000") + " #"))
    Else
        TotaalResultaat = TotaalResultaat + Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#v" + Format(ACTIVE_BOOKYEAR + 23, "000") + " #"))
    End If
    Do
        bNext TABLE_LEDGERACCOUNTS
        If Ktrl Or Left(KEY_BUF(TABLE_LEDGERACCOUNTS), 1) > "7" Then
            Exit Do
        Else
            RecordToVeld TABLE_LEDGERACCOUNTS
            If bhEuro Then
                TotaalResultaat = TotaalResultaat + Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#e" + Format(ACTIVE_BOOKYEAR + 23, "000") + " #"))
            Else
                TotaalResultaat = TotaalResultaat + Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#v" + Format(ACTIVE_BOOKYEAR + 23, "000") + " #"))
            End If
        End If
    Loop
End If

Msg = ""
If TotaalResultaat <> 0 Then
    Msg = Msg + "Resultatenbalans verschil  : " + Dec$((TotaalResultaat), MASK_EURBH) + vbCrLf
End If
If TotaalBalans <> 0 Then
    If Msg = "" Then
        Verwittigen = True
        Msg = "Blijkbaar is het resultaat verleden jaar GOED geboekt maar actief/passief is verschillend.  Straks wordt U de mogelijkheid geboden om de verschillen bij te werken OP EIGEN VERANTWOORDELIJKHEID" & vbCrLf & vbCrLf
    End If
    Msg = Msg + "ACTIVE/PASSIVA verschil   : " + Dec$((TotaalBalans), MASK_EURBH) + vbCrLf
End If
If Msg = "" Then
Else
    Msg = Msg + vbCrLf + "Balansrekeningen Actief <> Passief, Resultaatrekeningen Debet <> Credit" + vbCrLf + vbCrLf + "Eerst resultaatsverwerking vorig jaar boeken a.u.b."
    MsgBox Msg
    'If Verwittigen Then
    'Else
        'BoekBeginBalans = False
        'Exit Function
    'End If
End If

Dim dTotaalKtrl As Currency
Dim dBedrag As Currency
Dim LijnText As String
bGetOrGreater TABLE_LEDGERACCOUNTS, 0, vSet("1", 7)
If Ktrl Then
    MsgBox "Geen rekeningen ???"
    BoekBeginBalans = False
    Exit Function
ElseIf Left(KEY_BUF(TABLE_LEDGERACCOUNTS), 1) < "6" Then
    RecordToVeld TABLE_LEDGERACCOUNTS
    If bhEuro Then
        dBedrag = Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#e" + Format(ACTIVE_BOOKYEAR + 23, "000") + " #"))
    Else
        dBedrag = Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#v" + Format(ACTIVE_BOOKYEAR + 23, "000") + " #"))
    End If
    dTotaalKtrl = dTotaalKtrl + dBedrag
    If dBedrag <> 0 Then
        LijnText = vSet(vBibTekst(TABLE_LEDGERACCOUNTS, "#v019 #"), 7) + " " + vSet(vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #"), 40) + " " + Dec$((dBedrag), MASK_EURBH) + " " + vSet("", 7)
        JournaalPost.AddItem LijnText
    End If
    Do
        bNext TABLE_LEDGERACCOUNTS
        If Ktrl Or Left(KEY_BUF(TABLE_LEDGERACCOUNTS), 1) >= "6" Then
            Exit Do
        Else
            RecordToVeld TABLE_LEDGERACCOUNTS
            If bhEuro Then
                dBedrag = Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#e" + Format(ACTIVE_BOOKYEAR + 23, "000") + " #"))
            Else
                dBedrag = Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#v" + Format(ACTIVE_BOOKYEAR + 23, "000") + " #"))
            End If
            dTotaalKtrl = dTotaalKtrl + dBedrag
            If dBedrag <> 0 Then
                LijnText = vSet(vBibTekst(TABLE_LEDGERACCOUNTS, "#v019 #"), 7) + " " + vSet(vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #"), 40) + " " + Dec$((dBedrag), MASK_EURBH) + " " + vSet("", 7)
                JournaalPost.AddItem LijnText
            End If
        End If
    Loop
End If
Afsluiten.Enabled = True
Omschrijving.text = SoortBoeking.text
Afsluiten.SetFocus
'
'If Verwittigen Then
    JournaalSaldoKTRL
'Else
'    DCkeuze(0).Enabled = False
'    DCkeuze(1).Enabled = False
'    Bedrag.Visible = False
'    RekeningNummer.Visible = False
'    TRvlag.Enabled = False
'End If
BoekBeginBalans = True
BYPERDAT.PeriodeBoekjaar.ListIndex = 0

End Function

Private Function BoekFout() As Boolean

Dim T As Integer
Dim TT As Integer

BoekFout = False
DKTRL_CUMUL = 0: DKTRL_BEF = 0: DKTRL_EUR = 0

Unload frmBoeking
frmBoeking.Hide

For T = 0 To JournaalPost.ListCount - 1
    rsJournaal.AddNew
    rsJournaal("v041") = "0"
    rsJournaal("v066") = Format(Datum.Value, "yyyymmdd")            'Boekdatum
    rsJournaal("v033") = "D0" + Format(Datum.Value, "yyyymmdd")     'dokument
    rsJournaal("v035") = Format(Datum.Value, "yyyymmdd")            'dokumentdatum
    rsJournaal("v067") = Left(Omschrijving, 35)                     'Omschrijving
    rsJournaal("v019") = Mid(JournaalPost.List(T), 1, 7)            'Rekening
    rsJournaal("v068") = Mid(JournaalPost.List(T), 50, 12)          'Bedrag
    rsJournaal("dece068") = Val(Mid(JournaalPost.List(T), 50, 12))  'Bedrag
    rsJournaal("v069") = Mid(JournaalPost.List(T), 63, 7)           'TegenRekening
    If Not ADOJOURNAL_OK Then
        BoekFout = True
        DKTRL_CUMUL = 999
    ElseIf Mid(JournaalPost.List(T), 63, 7) <> Space$(7) Then
        rsJournaal.AddNew
        rsJournaal("v041") = "0"
        rsJournaal("v066") = Format(Datum.Value, "yyyymmdd")        'Boekdatum
        rsJournaal("v033") = "D0" + Format(Datum.Value, "yyyymmdd")   'dokument
        rsJournaal("v035") = Format(Datum.Value, "yyyymmdd")        'dokumentdatum
        rsJournaal("v067") = Left(Omschrijving, 35)                        'Omschrijving
            
        rsJournaal("v019") = Mid(JournaalPost.List(T), 63, 7)       'Rekening
        rsJournaal("v068") = Str$(-Val(Mid(JournaalPost.List(T), 50, 12))) 'Bedrag
        rsJournaal("dece068") = -Val(Mid(JournaalPost.List(T), 50, 12)) 'Bedrag
        rsJournaal("v069") = Mid(JournaalPost.List(T), 1, 7)        'TegenRekening
        If Not ADOJOURNAL_OK Then
            BoekFout = True
            DKTRL_CUMUL = 999
        End If
    End If
Next

'FIXIT: The use of 'vbNormal' is not valid for the property being assigned.                FixIT90210ae-R5049-R62328
Screen.MousePointer = vbNormal
If DKTRL_CUMUL Then
    MsgBox "Fout bij vierkantskontrole journaal." & vbCrLf & vbCrLf & "Deze verrichting wordt genegeerd."
    frmBoeking.cmdBoeken.Enabled = False
    frmBoeking.Show 1
    BoekFout = True
ElseIf JournaalLocked = True Then
    frmBoeking.cmdBoeken.Enabled = False
    frmBoeking.Show 1
    BoekFout = True
Else
    Select Case Left(Mim.cmdWegBoekModus.text, 1)
        Case "0"
        Case "1"
            If DKTRL_BEF Or DKTRL_EUR Then
                frmBoeking.Show 1
            End If
        Case "2"
            frmBoeking.Show 1
        Case Else
            MsgBox "situatie..."
    End Select
    If DKTRL_CUMUL Then BoekFout = True
End If

End Function


Private Sub Bedrag_LostFocus()
    
VolgendeLijn.Enabled = False
VolgendeLijn.Default = False

End Sub


Private Sub Datum_LostFocus()

If Not DATE_CHECK(Format(Datum.Value, "dd/mm/yyyy"), PERIODAS_TEXT) Then
    BYPERDAT.WindowState = 0
    BYPERDAT.PeriodeBoekjaar.SetFocus
End If

End Sub

Private Sub DCkeuze_KeyPress(Index As Integer, KeyAscii As Integer)
Dim Toets As String

'If KeyAscii = 13 Then SendKeys "{TAB}", True

Toets = UCase$(Chr$(KeyAscii))
Select Case Toets
    Case "D", "+"
        DCkeuze(0).Value = True
        RekeningNummer.SetFocus
    Case "C", "-"
        DCkeuze(1).Value = True
        RekeningNummer.SetFocus
    Case "T", "/"
        If TRvlag.Value Then
            TRvlag.Value = False
        Else
            TRvlag.Value = 1
        End If
        TRaanUit
        RekeningNummer.SetFocus
End Select


End Sub

Private Sub EindeAfschrijving()
Dim dRa As Double
Dim DummySleutel As String
Dim T As Integer

If XisEuroWisBEF = True Then
    Msg = "Laatste afschrijving in BEF is geboekt." & vbCr & vbCr
    Msg = Msg & "Uw hoogste boekjaar is in EUR" & vbCr & vbCr
    Msg = Msg & "Hierna worden alle nog openstaande afschrijvingslijnen eveneens omgerekend naar EUR voor toekomstige bewerkingen"
    MsgBox Msg, vbInformation
End If

For T = 0 To JournaalPost.ListCount - 1
    bGetOrGreater TABLE_VARIOUS, 1, vSet("18" + Right(JournaalPost.List(T), 7), 20)
    If Ktrl Then
        MsgBox "Stop"
    Else
        RecordToVeld TABLE_VARIOUS
        dRa = Val(vBibTekst(TABLE_VARIOUS, "#v085 #")) + Val(Mid(JournaalPost.List(T), 50, 12))
        vBib TABLE_VARIOUS, Str$(dRa), "v085"
        If XisEuroWisBEF = True Then
            vBib TABLE_VARIOUS, Str$(Val(Dec(dRa / EURO, "########0.00"))), "v085"
            dRa = Val(vBibTekst(TABLE_VARIOUS, "#v084 #"))
            vBib TABLE_VARIOUS, Str$(Val(Dec(dRa / EURO, "########0.00"))), "v084"
        End If
        bUpdate TABLE_VARIOUS, 1
        If Ktrl Then MsgBox "stop tijdens update investeringsfiche"
    End If
Next
SS99 "1", 63

End Sub


Private Sub Form_Load()

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If

Dim T As Integer

Top = 0
Left = 0
SoortBoeking.AddItem "0: Diverse post"
SoortBoeking.AddItem "1: Afschrijvingspost Eindejaar"
SoortBoeking.AddItem "2: Beginbalans"
'SoortBoeking.AddItem "9: Gebruikerssjabloon"
Datum.Value = MIM_GLOBAL_DATE
Schoon_Click
If XLogKey = "SchrijfAF!" Then
    SoortBoeking.ListIndex = 1
End If

End Sub


Private Sub JournaalPost_KeyPress(KeyAscii As Integer)
Dim Positie As Integer

Positie = JournaalPost.ListIndex
Select Case KeyAscii
    Case 45, 127
        If Positie < 0 Then
            Exit Sub
        Else
            Msg = Format(Positie + 1) + "e Journaallijn verwijderen.  Bent U zeker ?"
            Ktrl = MsgBox(Msg, 292, "Journaallijn verwijderen")
            If Ktrl = vbYes Then
                JournaalPost.RemoveItem Positie
                JournaalSaldoKTRL
            End If
        End If
        JournaalSaldoKTRL
    Case Else
End Select

End Sub

Private Sub Omschrijving_GotFocus()

SoortBoeking.Enabled = False

End Sub

Private Sub OpKuisVolgendeLijn()

TRvlag.Value = 0
TRaanUit
RekeningNummer.text = ""
NaamRekening.Caption = ""
Bedrag.text = ""
Tegenrekening.text = ""
NaamTegenRekening.Caption = ""
DCkeuze(0).Value = True
VolgendeLijn.Enabled = False

End Sub

Private Sub Omschrijving_KeyDown(KeyCode As Integer, Shift As Integer)

On Local Error Resume Next
If KeyCode = 13 Then
    SendKeys "{TAB}", True
    Exit Sub
End If

End Sub

Private Sub Omschrijving_KeyPress(KeyAscii As Integer)

'If KeyAscii = 13 Then SendKeys "{TAB}", True

End Sub

Private Sub Omschrijving_LostFocus()

    If Trim$(Omschrijving.text) = "" Then
        MsgBox "Omschrijving mag niet leeg zijn, ook géén spaties...", vbInformation
    End If
    
End Sub

Private Sub RekeningNummer_DblClick()

RekeningNummer_KeyDown 17, 0

End Sub

Private Sub RekeningNummer_GotFocus()

Afsluiten.Default = False
SnelHelpPrint "Dubbelklikken of [Ctrl] voor geïndexeerd zoeken", BL_LOGGING

End Sub

Private Sub RekeningNummer_KeyDown(KeyCode As Integer, Shift As Integer)

On Local Error Resume Next
If KeyCode = 13 Then
    SendKeys "{TAB}", True
    Exit Sub
End If

'If KeyCode = 17 Then CTRLFlag = True
Select Case KeyCode
    Case 9, 13
        MsgBox "Stop"
                         
    Case 17
        SharedFl = TABLE_LEDGERACCOUNTS
        aIndex = 0
        GridText = Me.RekeningNummer.text
        SqlSearch.Show 1
        If Ktrl Then
            Me.NaamRekening.Caption = ""
        Else
            Me.RekeningNummer.text = RV(rsMAR(TABLE_LEDGERACCOUNTS), "v019")
            Me.NaamRekening.Caption = RV(rsMAR(TABLE_LEDGERACCOUNTS), "v020")
        End If
End Select

End Sub

Private Sub RekeningNummer_KeyPress(KeyAscii As Integer)

'If KeyAscii = 13 Then SendKeys "{TAB}", True

End Sub

Private Sub RekeningNummer_LostFocus()

If Trim$(RekeningNummer) = "" Then
Else
    If Not ADO_GET(TABLE_LEDGERACCOUNTS, 0, "=", RekeningNummer) Then
        RekeningNummer = ""
        NaamRekening = ""
    Else
        NaamRekening = RV(rsMAR(TABLE_LEDGERACCOUNTS), "v020")
    End If
End If

End Sub




Private Sub Schoon_Click()
Dim T As Integer

On Local Error Resume Next

OpKuisVolgendeLijn
Afsluiten.Enabled = False
Omschrijving.text = ""
Saldo.Caption = "0"
JournaalPost.Clear
SoortBoeking.Enabled = True
SoortBoeking.ListIndex = 0
Sluiten.Enabled = True

DCkeuze(0).Enabled = True
DCkeuze(1).Enabled = True
Bedrag.Visible = True
RekeningNummer.Visible = True
TRvlag.Enabled = True
SoortBoeking.SetFocus

End Sub

Private Sub sluiten_Click()

If JournaalPost.ListCount Then
    Msg = Format(JournaalPost.ListCount) + " Journaallijnen negeren.  Bent U zeker ?"
    Ktrl = MsgBox(Msg, 292, "Inbreng Journaalpost")
    If Ktrl <> 6 Then
        Exit Sub
    End If
End If
Unload Me

End Sub

Private Sub SoortBoeking_Click()

Select Case Left(SoortBoeking.text, 1)
    Case "0"
    Case Else
               Msg = SoortBoeking.text + " aktiveren !" + vbCrLf + vbCrLf
        Msg = Msg + "Bent U zeker ?"
        Ktrl = MsgBox(Msg, 292)
        If Ktrl = 6 Then
            Select Case Left(SoortBoeking.text, 1)
                Case "1"
                    If AfschrijfBoeking() Then
                    Else
                        Schoon_Click
                    End If
                Case "2"
                    If BoekBeginBalans() Then
                    Else
                        Schoon_Click
                    End If
                Case "9"
                    BoekUserDef
                Case Else
                    MsgBox "stop"
            End Select
        Else
            Schoon_Click
            Exit Sub
        End If
End Select

End Sub

Private Sub SoortBoeking_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then Omschrijving.SetFocus

End Sub

Private Sub Tegenrekening_DblClick()

TegenRekening_KeyDown 17, 0

End Sub

Private Sub Tegenrekening_GotFocus()

SnelHelpPrint "Dubbelklikken of [Ctrl] voor geïndexeerd zoeken", BL_LOGGING

End Sub

Private Sub TegenRekening_KeyDown(KeyCode As Integer, Shift As Integer)

On Local Error Resume Next
If KeyCode = 13 Then
    SendKeys "{TAB}", True
    Exit Sub
End If

'If KeyCode = 17 Then CTRLFlag = True
Select Case KeyCode
    Case 9, 13
        MsgBox "Stop"
                         
    Case 17
        SharedFl = TABLE_LEDGERACCOUNTS
        aIndex = 0
        GridText = Me.Tegenrekening.text
        SqlSearch.Show 1
        If Ktrl Then
            Me.NaamTegenRekening.Caption = ""
        Else
            Me.Tegenrekening.text = RV(rsMAR(TABLE_LEDGERACCOUNTS), "v019")
            Me.NaamTegenRekening.Caption = RV(rsMAR(TABLE_LEDGERACCOUNTS), "v020")
        End If
End Select

End Sub

Private Sub Tegenrekening_KeyPress(KeyAscii As Integer)

'If KeyAscii = 13 Then SendKeys "{TAB}", True

End Sub

Private Sub Tegenrekening_LostFocus()

JackTop2:
If RTrim$(Tegenrekening) = "" Then
Else
    If Not ADO_GET(TABLE_LEDGERACCOUNTS, 0, "=", Tegenrekening) Then
        Tegenrekening = ""
        NaamTegenRekening = ""
    Else
        NaamTegenRekening = RV(rsMAR(TABLE_LEDGERACCOUNTS), "v020")
    End If
End If

End Sub

Private Sub TRaanUit()

If TRvlag.Value Then
    Tegenrekening.Visible = True
    NaamTegenRekening.Visible = True
Else
    Tegenrekening.Visible = False
    NaamTegenRekening.Visible = False
End If

End Sub

Private Sub TRvlag_Click()

TRaanUit

End Sub

Private Sub TRvlag_KeyPress(KeyAscii As Integer)
Dim Xindex As Integer

DCkeuze_KeyPress Xindex, KeyAscii

End Sub

Private Sub VolgendeLijn_Click()
Dim LijnBedrag As Double
Dim LijnText As String

If RTrim$(RekeningNummer.text) = "" Then
    Beep
    RekeningNummer.SetFocus
    Exit Sub
ElseIf Val(Bedrag.text) = 0 Then
    Beep
    Bedrag.SetFocus
    Exit Sub
ElseIf TRvlag.Value Then
    If RTrim$(Tegenrekening.text) = "" Then
        Beep
        Tegenrekening.SetFocus
        Exit Sub
    End If
End If

LijnBedrag = Val(Bedrag.text)
If DCkeuze(0).Value Then
Else
    LijnBedrag = -LijnBedrag
End If
Saldo.Caption = Format(CDbl(Saldo.Caption) + LijnBedrag, "#,##0.00")
If TRvlag.Value Then
    Saldo.Caption = Format(CDbl(Saldo.Caption) - LijnBedrag, "#,##0.00")
End If

LijnText = vSet((RekeningNummer.text), 7) + " " + vSet((NaamRekening.Caption), 40) + " " + Dec$((LijnBedrag), MASK_EURBH) + " "
If TRvlag.Value Then
    LijnText = LijnText + vSet((Tegenrekening.text), 7)
Else
    LijnText = LijnText + Space$(7)
End If
JournaalPost.AddItem LijnText

If Val(Saldo.Caption) = 0 Then
    Afsluiten.Enabled = True
    Afsluiten.Default = True
Else
    Afsluiten.Enabled = False
End If
OpKuisVolgendeLijn
DCkeuze(0).SetFocus

End Sub

Function JournaalSaldoKTRL() As Boolean

    Dim SaldoKTRL As Currency
    
    SaldoKTRL = 0
    For COUNT_TO = 0 To JournaalPost.ListCount - 1
        SaldoKTRL = SaldoKTRL + Val(Mid(JournaalPost.List(COUNT_TO), 50, 12))
    Next
    Saldo.Caption = Format(SaldoKTRL, "#,##0.00")
    
End Function
