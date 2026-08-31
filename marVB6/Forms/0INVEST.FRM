VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form InvesteringsFiche 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "InvesteringsFiche"
   ClientHeight    =   2715
   ClientLeft      =   1065
   ClientTop       =   1875
   ClientWidth     =   6090
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
   ScaleHeight     =   2715
   ScaleWidth      =   6090
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Versneld 
      Caption         =   "Versneld Afschrijven"
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   2340
      Width           =   2115
   End
   Begin VB.CommandButton Ok 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   405
      Left            =   4440
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1260
      Width           =   1560
   End
   Begin VB.CommandButton Annuleren 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   420
      Left            =   4440
      TabIndex        =   1
      Top             =   1740
      Width           =   1560
   End
   Begin MSMask.MaskEdBox TekstInfo 
      Height          =   285
      Index           =   6
      Left            =   3780
      TabIndex        =   21
      Top             =   210
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   12632256
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TekstInfo 
      Height          =   285
      Index           =   7
      Left            =   3780
      TabIndex        =   20
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   12632256
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TekstInfo 
      Height          =   285
      Index           =   8
      Left            =   2100
      TabIndex        =   19
      Top             =   750
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   12632256
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TekstInfo 
      Height          =   285
      Index           =   5
      Left            =   2100
      TabIndex        =   18
      Top             =   1950
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   12632256
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TekstInfo 
      Height          =   285
      Index           =   4
      Left            =   2100
      TabIndex        =   17
      Top             =   1680
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   12632256
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TekstInfo 
      Height          =   285
      Index           =   3
      Left            =   2100
      TabIndex        =   16
      Top             =   1410
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   12632256
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TekstInfo 
      Height          =   285
      Index           =   2
      Left            =   2100
      TabIndex        =   15
      Top             =   1020
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   12632256
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TekstInfo 
      Height          =   285
      Index           =   1
      Left            =   2100
      TabIndex        =   14
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   12632256
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TekstInfo 
      Height          =   285
      Index           =   0
      Left            =   2100
      TabIndex        =   13
      Top             =   210
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   12632256
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Caption         =   "Huidige bewerking"
      Height          =   240
      Index           =   0
      Left            =   1860
      TabIndex        =   10
      Top             =   0
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "Reeds afgeschr."
      Height          =   240
      Index           =   1
      Left            =   660
      TabIndex        =   3
      Top             =   780
      Width           =   1440
   End
   Begin VB.Label Label1 
      Caption         =   "Lineair n-jaar"
      Height          =   240
      Index           =   2
      Left            =   900
      TabIndex        =   4
      Top             =   1050
      Width           =   1170
   End
   Begin VB.Label Label1 
      Caption         =   "Datum"
      Height          =   240
      Index           =   4
      Left            =   1500
      TabIndex        =   5
      Top             =   240
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "Cumul bewerkingen"
      Height          =   210
      Index           =   3
      Left            =   3840
      TabIndex        =   11
      Top             =   0
      Width           =   1800
   End
   Begin VB.Label Label1 
      Caption         =   "Tot. bedrag"
      Height          =   240
      Index           =   5
      Left            =   1020
      TabIndex        =   6
      Top             =   510
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "InvesteringsRekening"
      Height          =   240
      Index           =   6
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1980
   End
   Begin VB.Label Label1 
      Caption         =   "AfschrijvingsRekening"
      Height          =   240
      Index           =   8
      Left            =   120
      TabIndex        =   8
      Top             =   1710
      Width           =   1980
   End
   Begin VB.Label Label1 
      Caption         =   "jaar"
      Height          =   240
      Index           =   7
      Left            =   2940
      TabIndex        =   12
      Top             =   1050
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "AfschrijfKostRekening"
      Height          =   240
      Index           =   10
      Left            =   120
      TabIndex        =   9
      Top             =   1980
      Width           =   1980
   End
End
Attribute VB_Name = "InvesteringsFiche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit

Dim NietAanwezig As Integer

Private Sub Annuleren_Click()

Msg = "InvesteringFiche negeren !  Bent U zeker ?" & vbCrLf & vbCrLf
Msg = Msg & "Opgelet: U dient dan alle lijnen voor dit document over te slaan."
Ktrl = MsgBox(Msg, vbExclamation + vbYesNo + vbDefaultButton2, "Investeringsfiche overslaan")
If Ktrl = 6 Then
    Unload Me
End If

End Sub

Private Sub Form_Load()
Dim T As Integer
Dim Teller As Integer
Dim RekeningTest As String * 7
Dim RekeningNaam As String * 40
Dim RekeningTest2 As String

TekstInfo(3).text = vBibTekst(TABLE_LEDGERACCOUNTS, "#v019 #")
RekeningTest2 = RTrim$(vBibTekst(TABLE_LEDGERACCOUNTS, "#v019 #"))
TekstInfo(0).text = Mid(GridText, 1, 10)
TekstInfo(1).text = Dec$(Val(Mid(GridText, 11, 12)), MASK_EURBH)
TLB_RECORD(TABLE_VARIOUS) = ""

'PartRight(RekeningTest2, 1) = "9"
RekeningTest2 = PartLeft(RekeningTest2, Len(RekeningTest2) - 1) + "9"
RekeningTest = vSet(RekeningTest2, 7)
TekstInfo(4).text = RekeningTest
bGet TABLE_LEDGERACCOUNTS, 0, RekeningTest
If Ktrl Then
    RekeningNaam = vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #")
        Msg = "Afschrijving op " + RTrim$(RekeningNaam) + vbCrLf
    Msg = Msg + "Rekeningnr. : " + RekeningTest + " bestaat nog niet." + vbCrLf + vbCrLf
    Msg = Msg + "Wordt hierna automatisch aangemaakt..."
    MsgBox Msg, 0, "Aanmaak afschrijfrekening"
        
    TLB_RECORD(TABLE_LEDGERACCOUNTS) = ""
    vBib TABLE_LEDGERACCOUNTS, RekeningTest, "v019"
    vBib TABLE_LEDGERACCOUNTS, "Afschrijving op " + RTrim$(RekeningNaam), "v020"
    vBib TABLE_LEDGERACCOUNTS, "O", "v032"
    bInsert TABLE_LEDGERACCOUNTS, 0
    If Ktrl <> 0 Then
        MsgBox "Fout bij invoegen afschrijvingsrekening"
    End If
End If

bGet TABLE_VARIOUS, 1, vSet("18" + RekeningTest, 20)
NietAanwezig = Ktrl
If Ktrl = 4 Then
    TekstInfo(6).text = ""
    For T = 7 To 8
        TekstInfo(T).text = Dec$(0, MASK_EURBH)
    Next
    TekstInfo(2).text = Dec$(5, "###")
    TekstInfo(5).text = "6300000"
    Versneld.Value = 1
Else
    RecordToVeld TABLE_VARIOUS
    TekstInfo(2).text = Dec$(Val(vBibTekst(TABLE_VARIOUS, "#v082 #")), "###") 'lineair over aantal jaar
    TekstInfo(6).text = (vBibTekst(TABLE_VARIOUS, "#v083 #"))    'datum vorige bewerking
    TekstInfo(7).text = Dec$(Val(vBibTekst(TABLE_VARIOUS, "#v084 #")), MASK_EURBH)   'totaal vorige bewerkingen
    TekstInfo(8).text = Dec$(Val(vBibTekst(TABLE_VARIOUS, "#v085 #")), MASK_EURBH)  'reeds afgeschreven
    Versneld.Value = Val(vBibTekst(TABLE_VARIOUS, "#v086 #"))     'versnelde afschrijving
    TekstInfo(4).text = vBibTekst(TABLE_VARIOUS, "#v087 #")
    TekstInfo(5).text = vBibTekst(TABLE_VARIOUS, "#v088 #")
    If TekstInfo(6).text = TekstInfo(0).text And Val(vBibTekst(TABLE_VARIOUS, "#v084 #")) = Val(Mid(GridText, 11, 12)) Then
              Msg = "Opgelet, laatste bijwerking dezelfde dag én zelfde bedrag" & vbCrLf
        Msg = Msg + "reeds aanwezig.  Vermijdt dubbele verrichtingen." & vbCrLf & vbCrLf
        Msg = Msg + "Kies Sluiten indien U zopas de fiche reeds bijgewerkt hebt."
        MsgBox Msg, 48, "Investeringsfiche éénzelfde datum"
    ElseIf Not DATE_CHECK((TekstInfo(6).text), BOOKYEARAS_TEXT) Then
              Msg = "Opgelet, U probeert een investeringsfiche" + vbCrLf
        Msg = Msg + "van een ander boekjaar bij te werken !" + vbCrLf + vbCrLf
        Msg = Msg + "Duidt EERST in de fiche van de leverancier een investeringsrekening aan geldig voor dit boekjaar en probeer daarna opnieuw."
        MsgBox Msg, 16, "Boekhoudkundige Gebruikersfout"
        investWarning = True
        Unload Me
    End If
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

bClose TABLE_VARIOUS

End Sub

Private Sub Ok_Click()
Dim TempoTel As Integer
Dim Ktrl2 As Integer

vBib TABLE_VARIOUS, (TekstInfo(3).text), "v019"
vBib TABLE_VARIOUS, (TekstInfo(2).text), "v082"
vBib TABLE_VARIOUS, (TekstInfo(0).text), "v083"
vBib TABLE_VARIOUS, Dec$(Val(TekstInfo(1).text) + Val(TekstInfo(7).text), MASK_EURBH), "v084"
'v085 TekstInfo(8).Text reeds afgeschreven mag niet gewijzigd worden
vBib TABLE_VARIOUS, Format(Versneld.Value, "0"), "v086"
vBib TABLE_VARIOUS, (TekstInfo(4).text), "v087"
vBib TABLE_VARIOUS, (TekstInfo(5).text), "v088"
vBib TABLE_VARIOUS, "18" + vBibTekst(TABLE_VARIOUS, "#v087 #"), "v005"
      Msg = "Informatielijn opslaan..." + vbCrLf
Msg = Msg + "Bent U zeker ?"
Ktrl2 = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton1, "Fiche bijwerken/wegschrijven")
If Ktrl2 = vbYes Then
    If NietAanwezig Then
        bInsert TABLE_VARIOUS, 1
    Else
        bUpdate TABLE_VARIOUS, 1
    End If
    Unload Me
End If

End Sub

Private Sub TekstInfo_GotFocus(Index As Integer)

TekstInfo(Index).SelLength = Len(TekstInfo(Index).text)

End Sub

Private Sub TekstInfo_LostFocus(Index As Integer)
Dim TempoBdrg As Long

Select Case Index
    Case 0
        If DATE_INVALID((TekstInfo(Index).text)) Then
            TekstInfo(Index).text = Mid(GridText, 1, 10)
            Beep
        End If
    Case 1, 8
        TekstInfo(Index).text = Dec$(Val(TekstInfo(Index).text), MASK_EURBH)
    Case 2
        TempoBdrg = Val(TekstInfo(Index).text)
        If TempoBdrg < 1 Or TempoBdrg > 50 Then
            Beep
            TempoBdrg = 5
        End If
        TekstInfo(Index).text = Dec$((TempoBdrg), "##0")
    Case 5
        If Left(TekstInfo(Index).text, 3) <> "630" Then
            Beep
            TekstInfo(Index).text = "6300000"
            TekstInfo(Index).SetFocus
            Exit Sub
        End If
End Select

End Sub

