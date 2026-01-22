VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form StandaardKostPrijsKaart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "StandaardKostprijsKaart"
   ClientHeight    =   7680
   ClientLeft      =   300
   ClientTop       =   1770
   ClientWidth     =   11025
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7680
   ScaleWidth      =   11025
   Begin VB.CommandButton Command1 
      Caption         =   "&Lijst alle kaarten"
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
      Left            =   8400
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6960
      Width           =   2235
   End
   Begin VB.CommandButton cmdVerbruik 
      Caption         =   "Totaal Productie Herrekenen"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3420
      TabIndex        =   18
      Top             =   6900
      Width           =   1695
   End
   Begin VB.TextBox txtVerbruik 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   2580
      TabIndex        =   17
      Text            =   "1"
      Top             =   6900
      Width           =   735
   End
   Begin VB.CheckBox cbPrijswijziging 
      Caption         =   "Enkel voor prijswijzigingen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   16
      Top             =   7260
      Width           =   2175
   End
   Begin VB.CommandButton CmdProduktieFiche 
      Caption         =   "&Produktiefiche"
      Height          =   315
      Left            =   5400
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6900
      Width           =   1575
   End
   Begin VB.CommandButton cmdLijst 
      Caption         =   "&Lijst alle kaarten"
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6900
      Width           =   2235
   End
   Begin VB.CommandButton CmdSchoon 
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
      Left            =   7020
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6900
      Width           =   1035
   End
   Begin VB.CommandButton CmdWegschrijven 
      Caption         =   "&Wegschrijven"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5400
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7260
      Width           =   1575
   End
   Begin VB.CommandButton CmdSProdukt 
      Caption         =   "Samengesteld produkt &Zoeken"
      Height          =   345
      Left            =   1320
      TabIndex        =   0
      Top             =   660
      Width           =   2805
   End
   Begin VB.CommandButton CmdAnnuleren 
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
      Left            =   7020
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7260
      Width           =   1035
   End
   Begin MSFlexGridLib.MSFlexGrid grdProduktDetail 
      Height          =   5745
      Left            =   60
      TabIndex        =   20
      ToolTipText     =   "Gebruik [Insert][+] of [Delete][-]"
      Top             =   1110
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   10134
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
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
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Detaillijnen"
      Height          =   255
      Index           =   4
      Left            =   60
      TabIndex        =   1
      Top             =   810
      Width           =   1095
   End
   Begin VB.Label LblProduktInfo 
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
      Height          =   585
      Left            =   30
      TabIndex        =   13
      Top             =   30
      Width           =   4065
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Verkoop"
      Height          =   285
      Index           =   3
      Left            =   5730
      TabIndex        =   12
      Top             =   720
      Width           =   825
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aankoop"
      Height          =   285
      Index           =   2
      Left            =   5730
      TabIndex        =   11
      Top             =   390
      Width           =   825
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nieuwe prijs"
      Height          =   255
      Index           =   1
      Left            =   8730
      TabIndex        =   10
      Top             =   90
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Per eenh. huidige prijs"
      Height          =   255
      Index           =   0
      Left            =   6060
      TabIndex        =   9
      Top             =   90
      Width           =   2025
   End
   Begin VB.Label LblVerkoopNieuw 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   8100
      TabIndex        =   8
      Top             =   690
      Width           =   1755
   End
   Begin VB.Label LblAankoopNieuw 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   8100
      TabIndex        =   7
      Top             =   360
      Width           =   1755
   End
   Begin VB.Label LblVerkoop 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   6570
      TabIndex        =   6
      Top             =   690
      Width           =   1515
   End
   Begin VB.Label LblAankoop 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   6570
      TabIndex        =   5
      Top             =   360
      Width           =   1515
   End
End
Attribute VB_Name = "StandaardKostPrijsKaart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim GrdProduktIsGewijzigd   As Boolean
Dim X                       As Boolean

Dim NietAanwezig            As Integer
Dim AantalDeelProdukten     As Integer
Dim Teller                  As Integer

Dim rsKPK As ADODB.Recordset

Dim verbruikBedrag As Currency
Dim productieBedrag As Currency
Dim productiecumul As Currency

Dim kpkProductOMS As String
Dim kpkProductNR As String

Dim TotaalAankoop As Currency
Dim TotaalVerkoop As Currency

Dim TotaalAankoopNieuw As Currency
Dim TotaalVerkoopNieuw As Currency

Sub KtrlKEYPRess(KeyCode As Integer)
Dim aa As String
Dim Antwoord As String
Dim vkBedrag As Currency
Dim akBedrag As Currency

Select Case KeyCode
    Case 45, 107 'Insert
        If LblProduktInfo.Caption = "" Then Exit Sub
        aa = ""
        aIndex = 0
        SharedFl = TABLE_PRODUCTS
        GridText = ""
        SqlSearch.Show 1
        If Ktrl = 0 Then
            If bhEuro Then
                vkBedrag = Val(vBibTekst(TABLE_PRODUCTS, "#e112 #"))
                akBedrag = Val(vBibTekst(TABLE_PRODUCTS, "#e113 #"))
            Else
                vkBedrag = Val(vBibTekst(TABLE_PRODUCTS, "#v112 #"))
                akBedrag = Val(vBibTekst(TABLE_PRODUCTS, "#v113 #"))
            End If
            If vkBedrag = 0 Then
                MsgBox "Verkoopprijs staat nog op 0 van dit produkt.  Eerst de nodige wijzigingen aanbrengen a.u.b.", , LTrim$(vBibTekst(TABLE_PRODUCTS, "#v105 #"))
                Exit Sub
            ElseIf akBedrag = 0 Then
                MsgBox "Aankoopkoopprijs staat nog op 0 van dit produkt.  Eerst de nodige wijzigingen aanbrengen a.u.b.", , LTrim$(vBibTekst(TABLE_PRODUCTS, "#v105 #"))
                Exit Sub
            End If
            
            aa = fmarBoxText("004", "2", vBibTekst(TABLE_PRODUCTS, "#v106 #"))
            Antwoord = vsfInputBox("Verkoopprijs per eenheid in BEF : " + Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#v112 #")), MASK_SY(4)), "Geef het verbruik aan (in " + Mid(aa, 4) + ")", "", "")
            If Antwoord = "" Then
                Exit Sub
            Else
                aa = vBibTekst(TABLE_PRODUCTS, "#v102 #") & vbTab
                aa = aa & vBibTekst(TABLE_PRODUCTS, "#v105 #") & vbTab
                If bhEuro Then
                    aa = aa & Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#e113 #")), MASK_SY(2) + "0000") & vbTab
                    aa = aa & Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#e112 #")), MASK_SY(2) + "0000") & vbTab
                Else
                    aa = aa & Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#v113 #")), MASK_SY(2) + "0000") & vbTab
                    aa = aa & Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#v112 #")), MASK_SY(2) + "0000") & vbTab
                End If
                aa = aa & Mid(fmarBoxText("004", "2", vBibTekst(TABLE_PRODUCTS, "#v106 #")), 4) & vbTab
                aa = aa & Dec$(Val(Antwoord), MASK_SY(4) + "0")
                grdProduktDetail.AddItem aa, grdProduktDetail.Rows - 1
                X = MaakTotaal
                GrdProduktIsGewijzigd = True
                CmdWegschrijven.Enabled = True
                CmdProduktieFiche.Enabled = True
            End If
        Else
            Exit Sub
        End If
        
    Case 46, 109 'Delete
        grdProduktDetail.Col = 0
        If grdProduktDetail.Text = "" Or grdProduktDetail.Rows < 3 Then
        ElseIf grdProduktDetail.Row > 0 Then
            KtrlBox = MsgBox(grdProduktDetail.Text + "-lijn verwijderen.  Bent U zeker ?", vbYesNo + vbDefaultButton2)
            If KtrlBox = vbYes Then
                grdProduktDetail.RemoveItem grdProduktDetail.Row
                X = MaakTotaal
            End If
        End If
        grdProduktDetail.SetFocus
       
Case Else
    
End Select
SnelHelpPrint "Deelprodukten : [Insert]/[+] om toe te voegen, [Delete]/[-] om te verwijderen", BL_LOGGING

End Sub

Function MaakTotaal() As Boolean

LblAankoopNieuw.Caption = "0"
LblVerkoopNieuw.Caption = "0"
For Teller = 1 To grdProduktDetail.Rows - 2
    grdProduktDetail.Row = Teller
    grdProduktDetail.Col = 0
    bGet TABLE_PRODUCTS, 0, (grdProduktDetail.Text)
    If Ktrl Then
        MsgBox "Onlogische situatie!"
    Else
        RecordToVeld TABLE_PRODUCTS
        grdProduktDetail.Col = 5
        LblAankoopNieuw.Caption = Format(CDbl(LblAankoopNieuw.Caption) + (Val(vBibTekst(TABLE_PRODUCTS, "#e113 #")) * Val(grdProduktDetail.Text)), "#,##0.000000")
        LblVerkoopNieuw.Caption = Format(CDbl(LblVerkoopNieuw.Caption) + (Val(vBibTekst(TABLE_PRODUCTS, "#e112 #")) * Val(grdProduktDetail.Text)), "#,##0.000000")
    End If
Next

End Function

Function MaakLijstTotaal() As Boolean

TotaalAankoopNieuw = 0
TotaalVerkoopNieuw = 0
For Teller = 1 To grdProduktDetail.Rows - 2
    grdProduktDetail.Row = Teller
    grdProduktDetail.Col = 0
    bGet TABLE_PRODUCTS, 0, (grdProduktDetail.Text)
    If Ktrl Then
        MsgBox "Onlogische situatie!"
    Else
        RecordToVeld TABLE_PRODUCTS
        grdProduktDetail.Col = 5
        TotaalAankoopNieuw = TotaalAankoopNieuw + (Val(vBibTekst(TABLE_PRODUCTS, "#e113 #")) * Val(grdProduktDetail.Text))
        TotaalVerkoopNieuw = TotaalVerkoopNieuw + (Val(vBibTekst(TABLE_PRODUCTS, "#e112 #")) * Val(grdProduktDetail.Text))
    End If
Next

End Function


Private Sub CmdAnnuleren_Click()

If GrdProduktIsGewijzigd And grdProduktDetail.Rows <> 2 Then
    KtrlBox = MsgBox("Wijzigingen negeren.  Bent U zeker ?", vbYesNo + vbDefaultButton2)
    If KtrlBox = vbYes Then
        Unload Me
    End If
Else
    Unload Me
End If

End Sub


Private Sub CmdProduktieFiche_Click()
Dim KtrlInput       As Variant

Dim DatumAanmaak    As String * 10
Dim DatumKlaar      As String * 10
Dim dokumentSleutel As String * 11

Dim Aantal          As Single

Dim Aantalv119      As Double
Dim Aantalv120      As Double

Dim Bedragv121      As Currency
Dim Bedragv122      As Currency
Dim TotaalPrijs     As Currency
Dim BedragX         As Currency


InputVanProduktie:
KtrlInput = InputBox("Aantal eenheden", "ProduktieFiche " + CmdSProdukt.Caption, txtVerbruik.Text)
If KtrlInput = "" Then Exit Sub Else Aantal = Val(KtrlInput)

KtrlInput = InputBox("Aanmaakdatum", "ProduktieFiche " + CmdSProdukt.Caption, MIM_GLOBAL_DATE)
If KtrlInput = "" Or Len(KtrlInput) <> 10 Then Exit Sub Else DatumAanmaak = KtrlInput

KtrlInput = InputBox("Klaar tegen", "ProduktieFiche " + CmdSProdukt.Caption, MIM_GLOBAL_DATE)
If KtrlInput = "" Or Len(KtrlInput) <> 10 Then Exit Sub Else DatumKlaar = KtrlInput

      Msg = "Aanmaak van " + Format(Aantal, "#,##0.00") + " eenheden" + vbCrLf + vbCrLf
Msg = Msg + LblProduktInfo.Caption + vbCrLf + vbCrLf
Msg = Msg + "Aanmaakdatum : " + DatumAanmaak + ", klaar tegen : " + DatumKlaar + vbCrLf + vbCrLf
Msg = Msg + "Bent U zeker ?"

dokumentSleutel = SleutelDok(188)

DeAfdrukVraag:
KtrlBox = MsgBox(Msg, vbYesNoCancel + vbQuestion + vbDefaultButton3, "Produktiefiche " + dokumentSleutel + " afdrukken/aanmaken")
Select Case KtrlBox
    Case vbYes
    Case vbNo
        Exit Sub
    Case vbCancel
        GoTo DeAfdrukVraag
    Case Else
        MsgBox "onlogische keuze"
End Select

AfdrukProduktieFiche:
Set Printer = Printers(LijstPrinterNr)
On Error Resume Next
Printer.PaperBin = LaadTekst(App.Title, "LIJSTPRINTER")
Printer.FontName = "Courier New"
Printer.FontSize = 7.2
Printer.Print vbCrLf; vbCrLf; FULL_LINE; vbCrLf;
Printer.Print Tab(10); "PRODUKTIEFICHE : "; dokumentSleutel; Tab(100); MIM_GLOBAL_DATE; vbCrLf;
Printer.Print vbCrLf; Tab(10); CmdSProdukt.Caption + " : "; Mid(LblProduktInfo.Caption, InStr(LblProduktInfo.Caption, vbCrLf) + 2); vbCrLf;
Printer.Print vbCrLf; vbCrLf;
Printer.Print Tab(20); "STANDAARD    WERKELIJK"; vbCrLf;
Printer.Print Tab(10); "AANMAAK : "; DatumAanmaak; "  ../../...."; vbCrLf;
Printer.Print Tab(10); "KLAAR OP: "; DatumKlaar; "  ../../...."; vbCrLf;
Printer.Print Tab(10); "AANTAL  : "; Dec$((Aantal), MASK_SY(2)); "  .........."; vbCrLf;
Printer.Print vbCrLf; vbCrLf;
Printer.Print "Anal. produkt Omschrijving                             Maat   Stand.hoev. Stand.prijs Werk.hoev. Werk.prijs"; vbCrLf;
Printer.Print FULL_LINE

TotaalPrijs = 0
For Teller = 1 To grdProduktDetail.Rows - 2
    grdProduktDetail.Row = Teller
    grdProduktDetail.Col = 0
    bGet TABLE_PRODUCTS, 0, (grdProduktDetail.Text)
    If Ktrl Then
        MsgBox "Onlogische situatie!"
    Else
        RecordToVeld TABLE_PRODUCTS
    End If
    Printer.Print vSet(grdProduktDetail.Text, 13); " ";
    grdProduktDetail.Col = 1
    Printer.Print vSet(grdProduktDetail.Text, 40); " ";
    grdProduktDetail.Col = 4
    Printer.Print vSet(grdProduktDetail.Text, 5); " ";
    grdProduktDetail.Col = 5
    Printer.Print Dec(Val(grdProduktDetail.Text), MASK_SY(4) + "0"); " ";
    BedragX = CDbl(Format(Val(vBibTekst(TABLE_PRODUCTS, "#e113 #")) * Val(grdProduktDetail.Text), "#.##"))
    TotaalPrijs = TotaalPrijs + BedragX
    Printer.Print Dec((BedragX), MASK_SY(4));
    Printer.Print " ........   ........."; Dec(BedragX * Aantal, MASK_SY(4)); vbCrLf;
Next
Printer.Print FULL_LINE; vbCrLf;
Printer.Print Tab(75); Dec((TotaalPrijs), MASK_SY(4)); vbCrLf;
Printer.Print Tab(20); "Totaal voor aantal"; Tab(75); Dec(TotaalPrijs * Aantal, MASK_SY(4));
Printer.EndDoc

DeWegschrijfVraag:
KtrlBox = MsgBox(Msg, vbYesNoCancel + vbQuestion + vbDefaultButton3, "Produktiefiche verwerken")
Select Case KtrlBox
    Case vbYes
    Case vbNo
        Exit Sub
    Case vbCancel
        GoTo DeWegschrijfVraag
    Case Else
        MsgBox "onlogische keuze"
End Select

BijwerkenVanStock:
DKTRL_CUMUL = 0: DKTRL_BEF = 0: DKTRL_EUR = 0
bBegin
For Teller = 1 To grdProduktDetail.Rows - 2
    grdProduktDetail.Row = Teller
    grdProduktDetail.Col = 0
    bGet TABLE_PRODUCTS, 0, (grdProduktDetail.Text)
    If Ktrl Then
        MsgBox "Onlogische situatie!"
    Else
        RecordToVeld TABLE_PRODUCTS
        grdProduktDetail.Col = 5

        If bhEuro Then
            DKTRL_CUMUL = DKTRL_CUMUL + CDbl(Format(Aantal * Val(vBibTekst(TABLE_PRODUCTS, "#e113 #")) * Val(grdProduktDetail.Text), "0"))
            Bedragv122 = Val(vBibTekst(TABLE_PRODUCTS, "#e122 #")) + CDbl(Format(Aantal * Val(vBibTekst(TABLE_PRODUCTS, "#e113 #")) * Val(grdProduktDetail.Text), "0"))
            Aantalv120 = Val(vBibTekst(TABLE_PRODUCTS, "#v120 #")) + (Aantal * Val(grdProduktDetail.Text))
            vBib TABLE_PRODUCTS, Str$(Bedragv122), "e122"
            vBib TABLE_PRODUCTS, Dec(Aantalv120, MASK_EUR), "v120"
        Else
            DKTRL_CUMUL = DKTRL_CUMUL + CDbl(Format(Aantal * Val(vBibTekst(TABLE_PRODUCTS, "#v113 #")) * Val(grdProduktDetail.Text), "0"))
            Bedragv122 = Val(vBibTekst(TABLE_PRODUCTS, "#v122 #")) + CDbl(Format(Aantal * Val(vBibTekst(TABLE_PRODUCTS, "#v113 #")) * Val(grdProduktDetail.Text), "0"))
            Aantalv120 = Val(vBibTekst(TABLE_PRODUCTS, "#v120 #")) + (Aantal * Val(grdProduktDetail.Text))
            vBib TABLE_PRODUCTS, Str$(Bedragv122), "v122"
            vBib TABLE_PRODUCTS, Dec(Aantalv120, MASK_EUR), "v120"
        End If
        bUpdate TABLE_PRODUCTS, 0
    End If
Next

StelDEVraag:
If DKTRL_CUMUL <> CDbl(Format(TotaalPrijs * Aantal, "0")) Then
    Msg = "Vierkantskontrole..." + vbCrLf + Str$(DKTRL_CUMUL) + "/" + Str$(TotaalPrijs * Aantal) + vbCrLf + vbCrLf
    Msg = Msg + "Rekenverschil : " + Format(DKTRL_CUMUL - CDbl(Format(TotaalPrijs * Aantal))) + vbCrLf
    Msg = Msg + "mag genegeerd worden ?"
    KtrlBox = MsgBox(Msg, vbYesNoCancel + vbDefaultButton3 + vbCritical)
    If KtrlBox = vbYes Then
        GoSub SKPKUpdate
    ElseIf KtrlBox = vbNo Then
        bAbort
    ElseIf KtrlBox = vbCancel Then
        GoTo StelDEVraag
    Else
        MsgBox "ongeldige keuzemogelijkheid"
    End If
    Exit Sub
Else
    GoSub SKPKUpdate
End If
Schoon
Exit Sub

SKPKUpdate:
bGet TABLE_PRODUCTS, 0, (CmdSProdukt.Caption)
If Ktrl Then
    MsgBox "Onlogische situatie!"
Else
    RecordToVeld TABLE_PRODUCTS
End If
If bhEuro Then
    Bedragv121 = Val(vBibTekst(TABLE_PRODUCTS, "#e121 #")) + (Aantal * TotaalPrijs)
    Aantalv119 = Val(vBibTekst(TABLE_PRODUCTS, "#v119 #")) + Aantal
    vBib TABLE_PRODUCTS, Str$(Bedragv121), "e121"
Else
    Bedragv121 = Val(vBibTekst(TABLE_PRODUCTS, "#v121 #")) + (Aantal * TotaalPrijs)
    Aantalv119 = Val(vBibTekst(TABLE_PRODUCTS, "#v119 #")) + Aantal
    vBib TABLE_PRODUCTS, Str$(Bedragv121), "v121"
End If
vBib TABLE_PRODUCTS, Dec(Aantalv119, MASK_EUR), "v119"
bUpdate TABLE_PRODUCTS, 0
bEnd
SS99 Right(dokumentSleutel, 5), 188
Return

'BON ZELF EVENEENS BEWAREN !!!
'FOR t.dl=1 TO tl
    'LSET r.nra7=rnr(1)
    'LSET r.om7=som(1)
    'LSET r.bnr7=sleut
    'LSET r.vkpr7=MKD$(dkp(1))
    'LSET r.ant7=MKS$(nant(1))
    'LSET r.btk7=Right(STR$(ibtk(1)),1)
    'LSET r.vol7=MKS$(nvol(1))
    'LSET r.ktr7=MKI$(ikrt(1))
    'LSET r.msf7=Right(STR$(imsf(1)),1)
    'LSET r.oprek7=roprek(1)
    
    'sltl=sleut:fl=7:idx=0:GOSUB 11200
    'prec=FETCHREC(fl)
    'idx=1:sltl=rnr(1)+SPACE$(13-LEN(rnr(1)))
    'GOSUB 11170
'Next

'3540 rdat1=Right(rdat1,2)+Mid(rdat1,4,2)+Left(rdat1,2)
'3550 rdat2=Right(rdat2,2)+Mid(rdat2,4,2)+Left(rdat2,2)
'3560 LSET r.nrk6=snrk:LSET r.dat16=rdat1:LSET r.dat26=rdat2:LSET r.bnr6=sleut
'3570 LSET r.goed6=MKD$(dttb+dnttb+dnnttb):LSET r.tbtw6=MKD$(dtbtw):snrm8=r.nrm8:LSET r.nrm6=snrm8:LSET r.munt6=MKD$(dmunt)
'3580 sltl=snrk:fl=6:idx=0:GOSUB 11200:prec=FETCHREC(fl):sltl=rdat2:idx=1:GOSUB 11170:sltl=sleut:idx=2:GOSUB 11170
'3590 '

End Sub


Private Sub cmdLijst_Click()
Dim aa As String
Dim AAA As Variant
Dim Antwoord As String

Dim PrijsAantepassen As Boolean

If cbPrijswijziging.Value = 1 Then
    PrijsAantepassen = True
Else
    PrijsAantepassen = False
End If

'on error resume next
Err = 0
rsKPK.CursorLocation = adUseClient
Msg = "SELECT * FROM Allerlei WHERE v005 Like '11%'"
SnelHelpPrint Msg, BL_LOGGING
Screen.MousePointer = vbHourglass
rsKPK.Open Msg, adntDB, adOpenForwardOnly, adLockReadOnly
If Err Then
    MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
    Screen.MousePointer = vbNormal
    Exit Sub
ElseIf rsKPK.RecordCount = 0 Then
    MsgBox "Geen standaardkostprijskaarten te vinden.", vbInformation
    Screen.MousePointer = vbNormal
    rsKPK.Close
    Exit Sub
End If
SnelHelpPrint "Lijst samenstellen voor " & rsKPK.RecordCount & " lijnen", BL_LOGGING
Unload frmADOXlog
frmADOXlog.Hide

If PrijsAantepassen Then
    frmADOXlog.Caption = "SKPK's voor Prijsaanpassing"
Else
    frmADOXlog.Caption = "SKPK's Totaallijst"
End If
With frmADOXlog.lvDetail
    ' Add ColumnHeaders. The width of the columns is
    ' the width of the control divided by the number of
    ' ColumnHeader objects.
    .ColumnHeaders.Add , , "Omschrijving", 2625
    .ColumnHeaders.Add , , "Nummer SKPK", 1620
    .ColumnHeaders.Add , , "AK Huidig", 1215, lvwColumnRight
    .ColumnHeaders.Add , , "AK Nieuw", 1215, lvwColumnRight
    .ColumnHeaders.Add , , "VK Huidig", 1215, lvwColumnRight
    .ColumnHeaders.Add , , "VK Nieuw", 1215, lvwColumnRight
    
    ' Set View property to Report.
    .View = lvwReport
End With

Dim itmX As ListItem
Do While Not rsKPK.EOF
    cmdSchoon_Click
    bGet TABLE_PRODUCTS, 0, Mid(rsKPK("v005"), 3)
    If Ktrl Then
    Else
        RecordToVeld TABLE_PRODUCTS
        
        LblProduktInfo.Caption = vBibTekst(TABLE_PRODUCTS, "#v102 #") + vbCr + vBibTekst(TABLE_PRODUCTS, "#v105 #")
        CmdSProdukt.Caption = vBibTekst(TABLE_PRODUCTS, "#v102 #")
        LblAankoop.Caption = Format(Val(vBibTekst(TABLE_PRODUCTS, "#e113 #")), "#,##0.000000")
        LblVerkoop.Caption = Format(Val(vBibTekst(TABLE_PRODUCTS, "#e112 #")), "#,##0.000000")
        GridText = ADOBIB_TEXT(rsKPK("MEMO"), "#v214 #")
        Do While GridText <> ""
            AAA = Mid(GridText, 1, InStr(GridText, vbTab) - 1)
            bGet TABLE_PRODUCTS, 0, (AAA)
            If Ktrl Then
            Else
                RecordToVeld TABLE_PRODUCTS
                aa = vBibTekst(TABLE_PRODUCTS, "#v102 #") & vbTab
                aa = aa & vBibTekst(TABLE_PRODUCTS, "#v105 #") & vbTab
                aa = aa & Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#e113 #")), MASK_SY(2) + "00") & vbTab
                aa = aa & Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#e112 #")), MASK_SY(2) + "00") & vbTab
                aa = aa & Mid(fmarBoxText("004", "2", vBibTekst(TABLE_PRODUCTS, "#v106 #")), 4) & vbTab
            End If
            Antwoord = Mid(GridText, InStr(GridText, vbTab) + 1, InStr(GridText, vbCrLf) - InStr(GridText, vbTab) - 1)
            aa = aa & Dec$(Val(Antwoord), MASK_SY(4) + "0")
            grdProduktDetail.AddItem aa, grdProduktDetail.Rows - 1
            GridText = Mid(GridText, InStr(GridText, vbCrLf) + 2)
        Loop
        X = MaakTotaal
        If Not PrijsAantepassen Then
            ' Add a ListItem object.
            Set itmX = frmADOXlog.lvDetail.ListItems.Add(, , Mid(LblProduktInfo, InStr(LblProduktInfo, vbCr) + 1))
                itmX.SubItems(1) = Left(LblProduktInfo, InStr(LblProduktInfo, vbCr) - 1)
                itmX.SubItems(2) = Format(CDbl(LblAankoop.Caption), "#,##0.00")
                itmX.SubItems(3) = Format(CDbl(LblAankoopNieuw.Caption), "#,##0.00")
                itmX.SubItems(4) = Format(CDbl(LblVerkoop.Caption), "#,##0.00")
                itmX.SubItems(5) = Format(CDbl(LblVerkoopNieuw.Caption), "#,##0.00")
        ElseIf LblAankoop = LblAankoopNieuw And LblVerkoop = LblVerkoopNieuw Then
        Else
            ' Add a ListItem object.
            Set itmX = frmADOXlog.lvDetail.ListItems.Add(, , Mid(LblProduktInfo, InStr(LblProduktInfo, vbCr) + 1))
                itmX.SubItems(1) = Left(LblProduktInfo, InStr(LblProduktInfo, vbCr) - 1)
                itmX.SubItems(2) = Format(CDbl(LblAankoop.Caption), "#,##0.00")
                itmX.SubItems(3) = Format(CDbl(LblAankoopNieuw.Caption), "#,##0.00")
                itmX.SubItems(4) = Format(CDbl(LblVerkoop.Caption), "#,##0.00")
                itmX.SubItems(5) = Format(CDbl(LblVerkoopNieuw.Caption), "#,##0.00")
        End If
    End If
    rsKPK.MoveNext
Loop
rsKPK.Close
Screen.MousePointer = vbNormal
If frmADOXlog.lvDetail.ListItems.Count = 0 Then
    MsgBox "Geen SKPK's in selectie", vbInformation
    cmdSchoon_Click
    StandaardKostPrijsKaart.SetFocus
    Exit Sub
Else
    frmADOXlog.lblAantal(0).Caption = frmADOXlog.lvDetail.ListItems.Count
End If
frmADOXlog.cmdOptie.Caption = "Prijs Wijzigen"
frmADOXlog.Show 1

Dim TempoTekst As String
TempoTekst = GridText
Do While TempoTekst <> ""
    cmdSchoon_Click
    Msg = "11" + Left(TempoTekst, InStr(TempoTekst, vbCr) - 1)
    X = ADO_GET(TABLE_VARIOUS, 1, "=", Msg)
    If Not X Then MsgBox "Onlogische situatie", vbCritical
    bGet TABLE_PRODUCTS, 0, Left(TempoTekst, InStr(TempoTekst, vbCr) - 1)
    If Ktrl Then
        MsgBox "Onlogische situatie", vbCritical
    Else
        RecordToVeld TABLE_PRODUCTS
        LblProduktInfo.Caption = vBibTekst(TABLE_PRODUCTS, "#v102 #") + vbCr + vBibTekst(TABLE_PRODUCTS, "#v105 #")
        CmdSProdukt.Caption = vBibTekst(TABLE_PRODUCTS, "#v102 #")
        If bhEuro Then
            LblAankoop.Caption = Format(Val(vBibTekst(TABLE_PRODUCTS, "#e113 #")), "#,##0.000000")
            LblVerkoop.Caption = Format(Val(vBibTekst(TABLE_PRODUCTS, "#e112 #")), "#,##0.000000")
        Else
            LblAankoop.Caption = Format(Val(vBibTekst(TABLE_PRODUCTS, "#v113 #")), "#,##0.000000")
            LblVerkoop.Caption = Format(Val(vBibTekst(TABLE_PRODUCTS, "#v112 #")), "#,##0.000000")
        End If
        GridText = ADOBIB_TEXT(rsMAR(TABLE_VARIOUS)("MEMO"), "#v214 #")
        Do While GridText <> ""
            AAA = Mid(GridText, 1, InStr(GridText, vbTab) - 1)
            bGet TABLE_PRODUCTS, 0, (AAA)
            If Ktrl Then
            Else
                RecordToVeld TABLE_PRODUCTS
                aa = vBibTekst(TABLE_PRODUCTS, "#v102 #") & vbTab
                aa = aa & vBibTekst(TABLE_PRODUCTS, "#v105 #") & vbTab
                If bhEuro Then
                    aa = aa & Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#e113 #")), MASK_SY(2) + "0000") & vbTab
                    aa = aa & Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#e112 #")), MASK_SY(2) + "0000") & vbTab
                Else
                    aa = aa & Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#v113 #")), MASK_SY(2) + "0000") & vbTab
                    aa = aa & Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#v112 #")), MASK_SY(2) + "0000") & vbTab
                End If
                aa = aa & Mid(fmarBoxText("004", "2", vBibTekst(TABLE_PRODUCTS, "#v106 #")), 4) & vbTab
            End If
            Antwoord = Mid(GridText, InStr(GridText, vbTab) + 1, InStr(GridText, vbCrLf) - InStr(GridText, vbTab) - 1)
            aa = aa & Dec$(Val(Antwoord), MASK_SY(4) + "0")
            grdProduktDetail.AddItem aa, grdProduktDetail.Rows - 1
            GridText = Mid(GridText, InStr(GridText, vbCrLf) + 2)
        Loop
        X = MaakTotaal
        bGet TABLE_PRODUCTS, 0, Left(TempoTekst, InStr(TempoTekst, vbCr) - 1)
        If Ktrl Then
            MsgBox "Onlogische situatie", vbCritical
        Else
            RecordToVeld TABLE_PRODUCTS
            If bhEuro Then
                vBib TABLE_PRODUCTS, Dec$(CDbl(LblVerkoopNieuw.Caption), "#######.000000"), "e112"
                vBib TABLE_PRODUCTS, Dec$(CDbl(LblAankoopNieuw.Caption), "#######.000000"), "e113"
            Else
                vBib TABLE_PRODUCTS, Dec$(CDbl(LblVerkoopNieuw.Caption), "#######.000000"), "v112"
                vBib TABLE_PRODUCTS, Dec$(CDbl(LblAankoopNieuw.Caption), "#######.000000"), "v113"
            End If
            bUpdate TABLE_PRODUCTS, 0
        End If
    End If
    TempoTekst = Mid(TempoTekst, InStr(TempoTekst, vbCr) + 1)
Loop
cmdSchoon_Click
StandaardKostPrijsKaart.SetFocus

End Sub


Private Sub cmdSchoon_Click()

Schoon

End Sub

Private Sub CmdSprodukt_Click()

If GrdProduktIsGewijzigd Then
        Msg = "Huidige wijzigingen negeren." + vbCrLf + vbCrLf
    Msg = Msg + "Bent U zeker ?"
    KtrlBox = MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton2, "Wijziging wegschrijven")
    If KtrlBox = vbYes Then
    Else
        Exit Sub
    End If
End If

aIndex = 0
SharedFl = TABLE_PRODUCTS
SqlSearch.Show 1
If Ktrl = 0 Then
    If vBibTekst(TABLE_PRODUCTS, "#v108 #") <> "S" Then
        MsgBox "Dit is geen produkt van het type 'standaardkostprijs' !"
        Exit Sub
    Else
        X = Instal
        grdProduktDetail.SetFocus
        CmdSProdukt.Enabled = False
    End If
Else
    Schoon
End If
End Sub


Private Sub cmdVerbruik_Click()

    bGet TABLE_PRODUCTS, 0, LblProduktInfo.Caption
    If Ktrl Then
        MsgBox "onlogica"
    Else
        Schoon
        Stop
        X = Instal
    End If
    
        
End Sub

Private Sub CmdWegschrijven_Click()
Dim aa As String
Dim Teller As Integer

Msg = "Huidige definitie wegschrijven." + vbCrLf + vbCrLf
Msg = Msg + "Bent U zeker ?"
KtrlBox = MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton2, "Wegschrijven")
If KtrlBox = vbYes Then
    bGet TABLE_VARIOUS, 1, vSet("11" + CmdSProdukt.Caption, 20)
    If Ktrl Then
    Else
        Do
            bDelete TABLE_VARIOUS
            bNext TABLE_VARIOUS
            If Ktrl Or vSet(KEY_BUF(TABLE_VARIOUS), 20) <> vSet("11" + CmdSProdukt.Caption, 20) Then
                Exit Do
            End If
        Loop
    End If
    aa = ""
    For Teller = 1 To grdProduktDetail.Rows - 2
        grdProduktDetail.Row = Teller
        grdProduktDetail.Col = 0
        aa = aa + Trim$(grdProduktDetail.Text) + vbTab
        grdProduktDetail.Col = 5
        aa = aa + Trim$(grdProduktDetail.Text) + vbCrLf
        If Len(aa) > 32000 Then
            GoSub SchrijfEensWeg
        End If
    Next
    If Len(aa) > 0 Then GoSub SchrijfEensWeg
          
          Msg = "Gewijzigde aankoop- en/of verkoopprijs wegschrijven." + vbCrLf + vbCrLf
    Msg = Msg + "Bent U zeker ?"
    KtrlBox = MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton1, "Prijzen bijwerken")
    If KtrlBox = vbYes Then
        bGet TABLE_PRODUCTS, 0, CmdSProdukt.Caption
        If Ktrl Then
            MsgBox "onlogische situatie"
        Else
            RecordToVeld TABLE_PRODUCTS
            If bhEuro Then
                vBib TABLE_PRODUCTS, Dec$(CDbl(LblVerkoopNieuw.Caption), "#######.000000"), "e112"
                vBib TABLE_PRODUCTS, Dec$(CDbl(LblAankoopNieuw.Caption), "#######.000000"), "e113"
            Else
                vBib TABLE_PRODUCTS, Dec$(CDbl(LblVerkoopNieuw.Caption), "#######.000000"), "v112"
                vBib TABLE_PRODUCTS, Dec$(CDbl(LblAankoopNieuw.Caption), "#######.000000"), "v113"
            End If
            bUpdate TABLE_PRODUCTS, 0
        End If
    End If
    Schoon
End If
Exit Sub

SchrijfEensWeg:
TLB_RECORD(TABLE_VARIOUS) = ""
vBib TABLE_VARIOUS, CmdSProdukt.Caption, "v102"
vBib TABLE_VARIOUS, "11" + CmdSProdukt.Caption, "v005"
vBib TABLE_VARIOUS, aa, "v214"
bInsert TABLE_VARIOUS, 1
Return

End Sub


Private Sub Command1_Click()
Dim aa As String
Dim AAA As Variant
Dim Antwoord As String

Dim PrijsAantepassen As Boolean

If cbPrijswijziging.Value = 1 Then
    PrijsAantepassen = True
Else
    PrijsAantepassen = False
End If

'on error resume next
Err = 0
rsKPK.CursorLocation = adUseClient
Msg = "SELECT * FROM Allerlei WHERE v005 Like '11%'"
SnelHelpPrint Msg, BL_LOGGING
Screen.MousePointer = vbHourglass
rsKPK.Open Msg, adntDB, adOpenForwardOnly, adLockReadOnly
If Err Then
    MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
    Screen.MousePointer = vbNormal
    Exit Sub
ElseIf rsKPK.RecordCount = 0 Then
    MsgBox "Geen standaardkostprijskaarten te vinden.", vbInformation
    Screen.MousePointer = vbNormal
    rsKPK.Close
    Exit Sub
End If
SnelHelpPrint "Lijst samenstellen voor " & rsKPK.RecordCount & " lijnen", BL_LOGGING
Unload frmADOXlog
frmADOXlog.Hide

If PrijsAantepassen Then
    frmADOXlog.Caption = "SKPK's voor Prijsaanpassing"
Else
    frmADOXlog.Caption = "SKPK's Totaallijst"
End If
With frmADOXlog.lvDetail
    ' Add ColumnHeaders. The width of the columns is
    ' the width of the control divided by the number of
    ' ColumnHeader objects.
    .ColumnHeaders.Add , , "Omschrijving", 2625
    .ColumnHeaders.Add , , "Nummer SKPK", 1620
    .ColumnHeaders.Add , , "AK Huidig", 1215, lvwColumnRight
    .ColumnHeaders.Add , , "AK Nieuw", 1215, lvwColumnRight
    .ColumnHeaders.Add , , "VK Huidig", 1215, lvwColumnRight
    .ColumnHeaders.Add , , "VK Nieuw", 1215, lvwColumnRight
    
    ' Set View property to Report.
    .View = lvwReport
End With

Dim itmX As ListItem

Do While Not rsKPK.EOF
    TotaalAankoop = 0
    TotaalVerkoop = 0
    TotaalAankoopNieuw = 0
    TotaalVerkoopNieuw = 0
    
    bGet TABLE_PRODUCTS, 0, Mid(rsKPK("v005"), 3)
    If Ktrl Then
    Else
        RecordToVeld TABLE_PRODUCTS
        X = LijstInstal
        If Not PrijsAantepassen Then
            ' Add a ListItem object.
            Set itmX = frmADOXlog.lvDetail.ListItems.Add(, , kpkProductOMS)
                itmX.SubItems(1) = kpkProductNR
                itmX.SubItems(2) = Format(TotaalAankoop, "#,##0.00")
                itmX.SubItems(3) = Format(TotaalAankoopNieuw, "#,##0.00")
                itmX.SubItems(4) = Format(TotaalVerkoop, "#,##0.00")
                itmX.SubItems(5) = Format(TotaalVerkoopNieuw, "#,##0.00")
        ElseIf TotaalAankoop = TotaalAankoopNieuw And TotaalVerkoop = TotaalVerkoopNieuw Then
        Else
            ' Add a ListItem object.
            Set itmX = frmADOXlog.lvDetail.ListItems.Add(, , kpkProductOMS)
                itmX.SubItems(1) = kpkProductNR
                itmX.SubItems(2) = Format(TotaalAankoop, "#,##0.00")
                itmX.SubItems(3) = Format(TotaalAankoopNieuw, "#,##0.00")
                itmX.SubItems(4) = Format(TotaalVerkoop, "#,##0.00")
                itmX.SubItems(5) = Format(TotaalVerkoopNieuw, "#,##0.00")
        End If
    End If
    rsKPK.MoveNext
Loop
rsKPK.Close
Screen.MousePointer = vbNormal
If frmADOXlog.lvDetail.ListItems.Count = 0 Then
    MsgBox "Geen SKPK's in selectie", vbInformation
    cmdSchoon_Click
    StandaardKostPrijsKaart.SetFocus
    Exit Sub
Else
    frmADOXlog.lblAantal(0).Caption = frmADOXlog.lvDetail.ListItems.Count
End If
frmADOXlog.cmdOptie.Caption = "Prijs Wijzigen"
frmADOXlog.Show 1

Dim TempoTekst As String
TempoTekst = GridText
Do While TempoTekst <> ""
    TotaalAankoop = 0
    TotaalVerkoop = 0
    TotaalAankoopNieuw = 0
    TotaalVerkoopNieuw = 0
    Msg = "11" + Left(TempoTekst, InStr(TempoTekst, vbCr) - 1)
    X = ADO_GET(TABLE_VARIOUS, 1, "=", Msg)
    If Not X Then MsgBox "Onlogische situatie", vbCritical
    bGet TABLE_PRODUCTS, 0, Left(TempoTekst, InStr(TempoTekst, vbCr) - 1)
    If Ktrl Then
        MsgBox "Onlogische situatie", vbCritical
    Else
        RecordToVeld TABLE_PRODUCTS
        X = LijstInstal
        bGet TABLE_PRODUCTS, 0, Left(TempoTekst, InStr(TempoTekst, vbCr) - 1)
        RecordToVeld TABLE_PRODUCTS
        vBib TABLE_PRODUCTS, Dec$((TotaalVerkoopNieuw), "#######.000000"), "e112"
        vBib TABLE_PRODUCTS, Dec$((TotaalAankoopNieuw), "#######.000000"), "e113"
        bUpdate TABLE_PRODUCTS, 0
    End If
    TempoTekst = Mid(TempoTekst, InStr(TempoTekst, vbCr) + 1)
Loop
cmdSchoon_Click
StandaardKostPrijsKaart.SetFocus

End Sub

Private Sub Form_Load()

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If

Top = 0
Left = 0

Set rsKPK = New ADODB.Recordset

AantalDeelProdukten = 0
With grdProduktDetail
    .Rows = AantalDeelProdukten + 2
    .Cols = 8

    .Col = 0
    .Row = 0
    .ColWidth(0) = 1290
    .ColAlignment(0) = flexAlignLeftCenter
    .Text = "Artikelnr."

    .Col = 1
    .ColWidth(1) = 3015
    .ColAlignment(1) = flexAlignLeftCenter
    .Text = "Omschrijving"

    .Col = 2
    .ColWidth(2) = 1095
    .ColAlignment(2) = 1
    .ColAlignment(2) = flexAlignRightCenter
    .Text = "AKprijs/"

    .Col = 3
    .ColWidth(3) = 1095
    .ColAlignment(3) = flexAlignRightCenter
    .Text = "VKprijs/"

    .Col = 4
    .ColWidth(4) = 795
    .ColAlignment(4) = flexAlignLeftCenter
    .Text = "maatstaf"

    .Col = 5
    .ColWidth(5) = 945
    .ColAlignment(5) = flexAlignRightCenter
    .Text = "Verbruik"

    .Col = 6
    .ColWidth(6) = 1045
    .ColAlignment(6) = flexAlignRightCenter
    .Text = "€ Verbruik"

    .Col = 7
    .ColWidth(7) = 1045
    .ColAlignment(7) = flexAlignRightCenter
    .Text = "€ Productie"

End With

GrdProduktIsGewijzigd = False
cmdSchoon_Click

End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    rsKPK.Close
    On Error Resume Next
    Set rsKPK = Nothing

End Sub


Private Sub GrdProduktDetail_GotFocus()

SnelHelpPrint "Deelprodukten : [Insert]/[+] om toe te voegen, [Delete]/[-] om te verwijderen", BL_LOGGING

End Sub

Private Sub GrdProduktDetail_KeyDown(KeyCode As Integer, Shift As Integer)

KtrlKEYPRess KeyCode

End Sub


Private Function Instal() As Boolean

Dim aa As String
Dim Antwoord As String
Dim verbruikBedrag As Currency
Dim productieBedrag As Currency
Dim productiecumul As Currency

LblProduktInfo.Caption = vBibTekst(TABLE_PRODUCTS, "#v102 #") + vbCrLf + vBibTekst(TABLE_PRODUCTS, "#v105 #")
CmdSProdukt.Caption = vBibTekst(TABLE_PRODUCTS, "#v102 #")

bGet TABLE_VARIOUS, 1, vSet("11" + CmdSProdukt.Caption, 20)
NietAanwezig = Ktrl
GrdProduktIsGewijzigd = False

If bhEuro Then
    LblAankoop.Caption = Format(Val(vBibTekst(TABLE_PRODUCTS, "#e113 #")), "#,##0.000000")
    LblVerkoop.Caption = Format(Val(vBibTekst(TABLE_PRODUCTS, "#e112 #")), "#,##0.000000")
Else
    LblAankoop.Caption = Format(Val(vBibTekst(TABLE_PRODUCTS, "#v113 #")), "#,##0.000000")
    LblVerkoop.Caption = Format(Val(vBibTekst(TABLE_PRODUCTS, "#v112 #")), "#,##0.000000")
End If
If NietAanwezig Then
    CmdWegschrijven.Enabled = False
    CmdProduktieFiche.Enabled = False
    AantalDeelProdukten = 0
    LblAankoopNieuw.Caption = Format(0, "#,##0.000000")
    LblVerkoopNieuw.Caption = Format(0, "#,##0.000000")
Else
    Do
        RecordToVeld TABLE_VARIOUS
        GoSub VoegLijnERBij
        bNext TABLE_VARIOUS
        If Ktrl Or vSet(KEY_BUF(TABLE_VARIOUS), 20) <> vSet("11" + CmdSProdukt.Caption, 20) Then
            Exit Do
        End If
    Loop
    CmdWegschrijven.Enabled = True
    CmdProduktieFiche.Enabled = True
    X = MaakTotaal
End If
Exit Function

VoegLijnERBij:
GridText = vBibTekst(TABLE_VARIOUS, "#v214 #")
Do While GridText <> ""
    bGet TABLE_PRODUCTS, 0, Mid(GridText, 1, InStr(GridText, vbTab) - 1)
    If Ktrl Then
        aa = Mid(GridText, 1, InStr(GridText, vbTab) - 1) & vbTab
        aa = aa & "Niet meer aanwezig !!" & vbTab
        aa = aa & "0" & vbTab
        aa = aa & "--" & vbTab
    Else
        RecordToVeld TABLE_PRODUCTS
        aa = vBibTekst(TABLE_PRODUCTS, "#v102 #") & vbTab
        aa = aa & vBibTekst(TABLE_PRODUCTS, "#v105 #") & vbTab
        aa = aa & Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#e113 #")), MASK_SY(2) + "00") & vbTab
        aa = aa & Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#e112 #")), MASK_SY(2) + "00") & vbTab
        aa = aa & Mid(fmarBoxText("004", "2", vBibTekst(TABLE_PRODUCTS, "#v106 #")), 4) & vbTab
    End If
    Antwoord = Mid(GridText, InStr(GridText, vbTab) + 1, InStr(GridText, vbCrLf) - InStr(GridText, vbTab) - 1)
    aa = aa & Dec$(Val(Antwoord), MASK_SY(4) + "0") & vbTab
    
    verbruikBedrag = Val(vBibTekst(TABLE_PRODUCTS, "#e113 #")) * Val(Antwoord)
    aa = aa & Format(verbruikBedrag, "#,##0.00") & vbTab
    productieBedrag = verbruikBedrag * Val(txtVerbruik.Text)
    productiecumul = productiecumul + productieBedrag
    aa = aa & Format(productieBedrag, "#,##0.00")
    grdProduktDetail.AddItem aa, grdProduktDetail.Rows - 1
    GridText = Mid(GridText, InStr(GridText, vbCrLf) + 2)
Loop
Return

End Function


Private Sub Schoon()

LblProduktInfo = ""
CmdWegschrijven.Enabled = False
CmdProduktieFiche.Enabled = False
GrdProduktIsGewijzigd = False
CmdSProdukt.Caption = "Samengesteld produkt &Zoeken"
CmdSProdukt.Enabled = True

LblAankoop = Format(0, "#,##0.00")
LblVerkoop = Format(0, "#,##0.00")
LblAankoopNieuw = Format(0, "#,##0.00")
LblVerkoopNieuw = Format(0, "#,##0.00")

With grdProduktDetail
    .Rows = 2
    .Col = 0
    .Row = 1
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
    
    .Col = 6
    .Text = ""
    
    .Col = 7
    .Text = ""
End With

End Sub

Private Function LijstInstal() As Boolean

Dim aa As String
Dim Antwoord As String

LijstInstal = False
kpkProductOMS = vBibTekst(TABLE_PRODUCTS, "#v105 #")
kpkProductNR = vBibTekst(TABLE_PRODUCTS, "#v102 #")

bGet TABLE_VARIOUS, 1, vSet("11" + kpkProductNR, 20)
NietAanwezig = Ktrl
GrdProduktIsGewijzigd = False

TotaalAankoop = Val(vBibTekst(TABLE_PRODUCTS, "#e113 #"))
TotaalVerkoop = Val(vBibTekst(TABLE_PRODUCTS, "#e112 #"))
TotaalAankoopNieuw = 0
TotaalVerkoopNieuw = 0

If NietAanwezig Then
    AantalDeelProdukten = 0
Else
    Do
        RecordToVeld TABLE_VARIOUS
        GoSub VoegLijnERBij2
        bNext TABLE_VARIOUS
        If Ktrl Or vSet(KEY_BUF(TABLE_VARIOUS), 20) <> vSet("11" + kpkProductNR, 20) Then
            Exit Do
        End If
    Loop
End If

Exit Function

VoegLijnERBij2:
GridText = vBibTekst(TABLE_VARIOUS, "#v214 #")
Do While GridText <> ""
    bGet TABLE_PRODUCTS, 0, Mid(GridText, 1, InStr(GridText, vbTab) - 1)
    If Ktrl Then
        aa = Mid(GridText, 1, InStr(GridText, vbTab) - 1) & vbTab
        aa = aa & "Niet meer aanwezig !!" & vbTab
        aa = aa & "0" & vbTab
        aa = aa & "--" & vbTab
    Else
        RecordToVeld TABLE_PRODUCTS
        aa = vBibTekst(TABLE_PRODUCTS, "#v102 #") & vbTab
        aa = aa & vBibTekst(TABLE_PRODUCTS, "#v105 #") & vbTab
        aa = aa & Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#e113 #")), MASK_SY(2) + "00") & vbTab
        aa = aa & Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#e112 #")), MASK_SY(2) + "00") & vbTab
        aa = aa & Mid(fmarBoxText("004", "2", vBibTekst(TABLE_PRODUCTS, "#v106 #")), 4) & vbTab
    End If
       
    Antwoord = Mid(GridText, InStr(GridText, vbTab) + 1, InStr(GridText, vbCrLf) - InStr(GridText, vbTab) - 1)
    aa = aa & Dec$(Val(Antwoord), MASK_SY(4) + "0") & vbTab
    
    verbruikBedrag = Val(vBibTekst(TABLE_PRODUCTS, "#e113 #")) * Val(Antwoord)
    aa = aa & Format(verbruikBedrag, "#,##0.00") & vbTab
    productieBedrag = verbruikBedrag * Val(txtVerbruik.Text)
    productiecumul = productiecumul + productieBedrag
    aa = aa & Format(productieBedrag, "#,##0.00")
    GridText = Mid(GridText, InStr(GridText, vbCrLf) + 2)
    TotaalAankoopNieuw = TotaalAankoopNieuw + (Val(vBibTekst(TABLE_PRODUCTS, "#e113 #")) * Val(Antwoord))
    TotaalVerkoopNieuw = TotaalVerkoopNieuw + (Val(vBibTekst(TABLE_PRODUCTS, "#e112 #")) * Val(Antwoord))
Loop
Return

End Function

