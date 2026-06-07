VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form xDokument 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "xDokument"
   ClientHeight    =   1620
   ClientLeft      =   4830
   ClientTop       =   735
   ClientWidth     =   4470
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
   HelpContextID   =   30000
   Icon            =   "0XDOK.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1620
   ScaleWidth      =   4470
   Begin VB.CheckBox chkAfdrukInVenster 
      Caption         =   "Afdruk in venster"
      Height          =   255
      Left            =   60
      TabIndex        =   9
      Top             =   1260
      Width           =   2400
   End
   Begin VB.ComboBox cmbDokumentType 
      BackColor       =   &H00C0C0C0&
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
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   30
      Width           =   4305
   End
   Begin VB.CommandButton Knop 
      Caption         =   "&Editeren"
      Enabled         =   0   'False
      Height          =   345
      HelpContextID   =   30040
      Index           =   5
      Left            =   60
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   810
      Width           =   975
   End
   Begin VB.CommandButton Knop 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   360
      HelpContextID   =   30030
      Index           =   2
      Left            =   3900
      Picture         =   "0XDOK.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   390
      Width           =   480
   End
   Begin VB.CommandButton Knop 
      Caption         =   "&<<"
      Height          =   360
      Index           =   6
      Left            =   2820
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   390
      Width           =   495
   End
   Begin VB.CommandButton Knop 
      Caption         =   ">&>"
      Height          =   360
      Index           =   7
      Left            =   3360
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   390
      Width           =   495
   End
   Begin VB.CommandButton Knop 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "&Pictogram"
      Height          =   360
      Index           =   8
      Left            =   5700
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1485
   End
   Begin VB.CommandButton Knop 
      Caption         =   "&Andere Fiche"
      Height          =   360
      HelpContextID   =   30010
      Index           =   3
      Left            =   2820
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   780
      Width           =   1545
   End
   Begin VB.CommandButton Knop 
      Caption         =   "&Wegschrijven"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   345
      HelpContextID   =   30070
      Index           =   1
      Left            =   1110
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   810
      Width           =   1395
   End
   Begin MSMask.MaskEdBox TekstInfo 
      Height          =   315
      Index           =   0
      Left            =   90
      TabIndex        =   8
      Top             =   390
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      BackColor       =   16777215
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
End
Attribute VB_Name = "xDokument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit

Dim T                           As Integer
Dim Voorkeurquick(10, 10)       As String * 5

Dim psTekst(5)                  As String
Dim VeldTXT(17)                 As String

Dim aa                          As String
Dim dTas                        As Double
Dim IndexNR                     As Integer
Dim tMaxVeld                    As Integer


Private Sub PrintTotaal()
Dim T As Integer

For T = 0 To tMaxVeld
    VeldTXT(T) = ""
Next
VeldTXT(1) = "Totaal :"
VeldTXT(10) = Dec(dTas, MASK_2002)

If chkAfdrukInVenster Then
Else
    Printer.Print vbCrLf; FULL_LINE
End If

T = 0
aa = ""
Do While REPORT_TAB(T) <> 0
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
If chkAfdrukInVenster Then Xlog.X.AddItem aa, Xlog.X.Rows - 1

If chkAfdrukInVenster Then
    Xlog.X.Row = 1
    Xlog.X.Col = 0
        'Xlog.X.ColWidth(0) = 1335
        'Xlog.X.ColWidth(1) = 2025
        'Xlog.X.ColWidth(2) = 885
        'Xlog.X.ColWidth(3) = 480
        'Xlog.X.ColWidth(4) = 885
        'Xlog.X.ColWidth(5) = 525
        'Xlog.X.ColWidth(6) = 945
        'Xlog.X.ColWidth(7) = 945
        'Xlog.X.ColWidth(8) = 900
        'Xlog.X.ColWidth(9) = 1425
        Xlog.WindowState = 0

        Xlog.X.ColAlignment(2) = flexAlignRightTop
        Xlog.X.ColAlignment(4) = flexAlignRightTop
        Xlog.X.ColAlignment(6) = flexAlignRightTop
        Xlog.X.ColAlignment(7) = flexAlignRightTop
        Xlog.X.ColAlignment(8) = flexAlignRightTop
            
    Screen.MousePointer = vbNormal
    '---------------------------------------------------
    'TODO everywhere when choosen show print in a window
    Xlog.WindowState = vbNormal
    Xlog.statusForm = ""
    Xlog.Width = 11130
    '---------------------------------------------------
        
    Xlog.WijzigenLijn.Visible = False
    Xlog.Afsluiten.TabStop = False
    'Xlog.cbAfbeelding.Visible = False
    Xlog.cbAfbeelding.Visible = False
    XLogKey = ""
    Xlog.SSTab1.TabVisible(1) = False
    Xlog.Show 1
    Unload Xlog
End If

End Sub


Private Sub PrintVelden()
Dim T As Integer
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
If chkAfdrukInVenster Then Xlog.X.AddItem aa, Xlog.X.Rows - 1

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
    Printer.Print Tab(1); psTekst(2); Tab(108); "Pagina : " + Dec$((PAGE_COUNTER), "##########");

    Printer.Print Tab(108); "Datum  : "; psTekst(0); vbCrLf; vbCrLf;
    Printer.Print Tab(1); UCase$(psTekst(3));

Printer.Print vbCrLf$; FULL_LINE

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

PrtHandler3:
MsgBox "Kontroleer de printer."
Resume

End Sub


Private Sub cmbdokumentType_Click()

TekstInfo(0).text = ""
Knop(5).Enabled = False

End Sub

Private Sub FicheNaarRecord()

bGet TABLE_VARIOUS, 1, vSet(vBibTekst(TABLE_VARIOUS, "#" + JETTABLEUSE_INDEX(TABLE_VARIOUS, 1) + "#"), FLINDEX_LEN(TABLE_VARIOUS, 1))
If Ktrl = 0 Then
    Msg = "Gegevens bestaande fiche wijzigen.  Bent U zeker ?"
    KtrlBox = MsgBox(Msg, 292)
    If KtrlBox = 6 Then
        Select Case Left(cmbDokumentType.text, 2)
            Case "10"
                vBib TABLE_VARIOUS, (TekstInfo(0).text), "vs03"
                vBib TABLE_VARIOUS, Left(cmbDokumentType.text, 2) + vBibTekst(TABLE_VARIOUS, "#vs03 #"), "v005"
            Case "12"
                vBib TABLE_VARIOUS, (TekstInfo(0).text), "v152"
                vBib TABLE_VARIOUS, Left(cmbDokumentType.text, 2) + vBibTekst(TABLE_VARIOUS, "#v152 #"), "v005"
            Case "18"
                vBib TABLE_VARIOUS, (TekstInfo(0).text), "v087"
                vBib TABLE_VARIOUS, Left(cmbDokumentType.text, 2) + vBibTekst(TABLE_VARIOUS, "#v087 #"), "v005"
            Case "21"
                vBib TABLE_VARIOUS, (TekstInfo(0).text), "v216"
                vBib TABLE_VARIOUS, Left(cmbDokumentType.text, 2) + vBibTekst(TABLE_VARIOUS, "#v216 #"), "v005"
            Case "28"
                vBib TABLE_VARIOUS, (TekstInfo(0).text), "v231"
                vBib TABLE_VARIOUS, Left(cmbDokumentType.text, 2) + vBibTekst(TABLE_VARIOUS, "#v231 #"), "v005"
            Case Else
                MsgBox "stop"
        End Select
        bUpdate TABLE_VARIOUS, 1
    Else
        Exit Sub
    End If
Else
    bInsert TABLE_VARIOUS, 1
End If

End Sub

Private Sub Form_Activate()

TekstInfo(0).text = ""  'Right(FVT(TABLE_VARIOUS, 1), 18)

End Sub

Private Sub Form_Load()

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If

Top = 0
Left = 0

cmbDokumentType.AddItem "18: Investeringsfiches"
cmbDokumentType.AddItem "10: MuntKodes, dagkoersen"
cmbDokumentType.AddItem "12: Logboek Artikels"
cmbDokumentType.AddItem "21: Forfaitaire Btw cumulators"
cmbDokumentType.AddItem "28: Financiële Instellingen"
cmbDokumentType.ListIndex = 0

End Sub


Private Sub Knop_Click(Index As Integer)

Select Case Index
        Case 1  'Nieuw of Gewijzigd wegschrijven
            FicheNaarRecord
            Knop_Click 3
                
        Case 2  'Alle fiches afdrukken
            Select Case Left(cmbDokumentType.text, 2)
                Case "18"
                    If AfschrijvingenLijstOk Then
                        XLogKey = "SchrijfAF!"
                        Unload DiversePosten
                        DiversePosten.Show
                    End If
                Case Else
                    MsgBox "Nog geen afdrukdefinitie beschikbaar voor " + Mid(cmbDokumentType.text, 4)
            End Select
        
        Case 3  'Nieuwe Fiche
            NieuweFiche
            TekstInfo(0).SetFocus
        
        Case 5  'Fiche EDITEREN
            If TeleBibClick(Val(Left(cmbDokumentType.text, 2))) = False Then
                Knop(1).Enabled = False
            Else
                Knop(1).Enabled = True
                Knop(1).SetFocus
                Knop_Click 1
            End If
                
        Case 6  'Fiche Lager
            bPrev TABLE_VARIOUS
            If Ktrl = 9 Then
                bFirst TABLE_VARIOUS, 1
                If Ktrl Then
                    Beep
                    Knop(5).Enabled = False
                End If
            ElseIf Left(cmbDokumentType.text, 2) = Left(KEY_BUF(TABLE_VARIOUS), 2) Then
                INSERT_FLAG(TABLE_VARIOUS) = 0
                RecordNaarFiche
                Knop(5).Enabled = True
                Exit Sub
            End If
            If Ktrl Then
            Else
                bGetOrGreater TABLE_VARIOUS, 1, vSet(Left(cmbDokumentType.text, 2), 20)
                If Ktrl Or Left(KEY_BUF(TABLE_VARIOUS), 2) <> Left(cmbDokumentType.text, 2) Then
                Else
                    INSERT_FLAG(TABLE_VARIOUS) = 0
                    RecordNaarFiche
                    Knop(5).Enabled = True
                End If
            End If
      
        Case 7  'fiche hoger
            bNext TABLE_VARIOUS
            If Ktrl = 9 Then
                bLast TABLE_VARIOUS, 1
                If Ktrl Then
                    Beep
                    Knop(5).Enabled = False
                End If
            Else
                If Left(KEY_BUF(TABLE_VARIOUS), 2) <> Left(cmbDokumentType.text, 2) Then
                Else
                    INSERT_FLAG(TABLE_VARIOUS) = 0
                    RecordNaarFiche
                    Knop(5).Enabled = True
                End If
            End If

        Case 8 'verbergen
            WindowState = 1
End Select

End Sub

Private Sub Knop_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

If KeyCode = 46 Then
    RecordToVeld TABLE_VARIOUS
    If Left(cmbDokumentType, 2) = Left(FVT(TABLE_VARIOUS, 1), 2) And RTrim$(TekstInfo(0)) = RTrim$(Mid(FVT(TABLE_VARIOUS, 1), 3)) Then
        Msg = "Bestaande fiche verwijderen.  Bent U zeker ?"
        KtrlBox = MsgBox(Msg, 292)
        If KtrlBox = 6 Then
            bDelete TABLE_VARIOUS
            Knop_Click 3
        End If
    End If
End If

End Sub

Private Sub Knop_KeyPress(Index As Integer, KeyAscii As Integer)

If InStr("W?AAE<>PPP", UCase(Chr(KeyAscii))) = 0 Then
    Beep
    Exit Sub
ElseIf Knop(InStr("W?AAE<>PPP", UCase(Chr(KeyAscii)))).Enabled = False Then
    Beep
Else
    Knop_Click InStr("W?AAE<>PPP", UCase(Chr(KeyAscii)))
End If

End Sub

Private Sub NieuweFiche()

TekstInfo(0).text = ""
TLB_RECORD(TABLE_VARIOUS) = ""

INSERT_FLAG(TABLE_VARIOUS) = 1
Knop(1).Enabled = False     'wegschrijven uitschakelen
Knop(5).Enabled = False     'editeren uitschakelen
TekstInfo(0).Enabled = True

End Sub

Private Sub RecordNaarFiche()

TLB_RECORD(TABLE_VARIOUS) = ""
If Ktrl Then
    MsgBox "stop"
Else
    RecordToVeld TABLE_VARIOUS
End If
TekstInfo(0).text = Mid(vBibTekst(TABLE_VARIOUS, "#v005 #"), 3)
INSERT_FLAG(TABLE_VARIOUS) = 0

End Sub

Private Sub TekstInfo_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = 13 Then Knop(3).SetFocus  'SendKeys "{TAB}", True

End Sub

Private Sub TekstInfo_LostFocus(Index As Integer)

On Local Error GoTo LostFocusError

Select Case Index
    Case 0
        If TekstInfo(0).text = Space$(Len(TekstInfo(0).text)) Then Exit Sub
        bGet TABLE_VARIOUS, 1, Left(cmbDokumentType.text, 2) + vSet((TekstInfo(0).text), 18)
        If Ktrl = 0 Then
            RecordNaarFiche
            INSERT_FLAG(TABLE_VARIOUS) = 0
            Knop(1).Enabled = True
            Knop(1).Default = True
        Else
            Knop(1).Enabled = False
        End If
        If INSERT_FLAG(TABLE_VARIOUS) = 1 Then
            Select Case Left(cmbDokumentType.text, 2)
                Case "10"
                    vBib TABLE_VARIOUS, (TekstInfo(0).text), "vs03"
                    vBib TABLE_VARIOUS, Left(cmbDokumentType.text, 2) + vBibTekst(TABLE_VARIOUS, "#vs03 #"), "v005"
                Case "12"
                    vBib TABLE_VARIOUS, (TekstInfo(0).text), "v152"
                    vBib TABLE_VARIOUS, Left(cmbDokumentType.text, 2) + vBibTekst(TABLE_VARIOUS, "#v152 #"), "v005"
                Case "18"
                    vBib TABLE_VARIOUS, (TekstInfo(0).text), "v087"
                    vBib TABLE_VARIOUS, Left(cmbDokumentType.text, 2) + vBibTekst(TABLE_VARIOUS, "#v087 #"), "v005"
                Case "21"
                    vBib TABLE_VARIOUS, (TekstInfo(0).text), "v216"
                    vBib TABLE_VARIOUS, Left(cmbDokumentType.text, 2) + vBibTekst(TABLE_VARIOUS, "#v216 #"), "v005"
                Case "28"
                    vBib TABLE_VARIOUS, (TekstInfo(0).text), "v231"
                    vBib TABLE_VARIOUS, Left(cmbDokumentType.text, 2) + vBibTekst(TABLE_VARIOUS, "#v231 #"), "v005"
                    
                Case Else
                    MsgBox "stop"
            End Select
        End If
        Knop(5).Enabled = True
        Knop(5).SetFocus
End Select

LostFocusError:
Exit Sub
Resume

End Sub

Function AfschrijvingenLijstOk() As Boolean

Dim DeLijst As String
Dim Kontroleer As String
Dim omsLijn As String
Dim DatumJaarInvest As String

Dim dRest As Double
Dim ipct As Integer
Dim dbdrg As Double
Dim dRa As Double
Dim Das As Double

Dim bhKontrole As Boolean

Dim BedragBegin As Currency
Dim BedragRest As Currency
Dim RekKontrole As String * 7
Dim DummyDatum As String * 8

bhKontrole = False
If String99(READING, 63) = "1" Then
    Msg = "Afschrijvingsposten reeds gegenereerd voor dit boekjaar !" + vbCrLf + vbCrLf
    Msg = Msg + "Bijkomende posten kunnen uitsluitend via 'Diverse post'-optie" + vbCrLf
    Msg = Msg + "ingebracht worden !" + vbCrLf + vbCrLf
    Msg = Msg + "Wenst U boekhoudkundige kontrolelijst ?"
    KtrlBox = MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton2)
    If KtrlBox = vbYes Then
        bhKontrole = True
    Else
        AfschrijvingenLijstOk = False
        Exit Function
    End If
ElseIf String99(READING, 63) <> "0" Then
    MsgBox "Setup boekjaar en parameters bevat niet de juiste vlag geboekt of niet geboekt.  Kontroleer"
    AfschrijvingenLijstOk = False
    Exit Function
ElseIf String99(READING, 64) <> "1" Then
    MsgBox "Onlogische situatie.  Dit boekjaar bevat nog geen beginbalans ?  De beginbalans dient aanwezig te zijn.  Mogelijk bevindt U zich in het verkeerde boekjaar ?"
    AfschrijvingenLijstOk = False
    Exit Function
End If
   
dTas = 0
bGetOrGreater TABLE_VARIOUS, 1, vSet("18", 20)
If Ktrl Or Left(KEY_BUF(TABLE_VARIOUS), 2) <> "18" Then
    MsgBox "Er zijn geen investeringsfiches !"
    AfschrijvingenLijstOk = False
    Exit Function
Else
    Screen.MousePointer = vbHourglass
    xDokument.Enabled = False
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
        End If
    End If
    psTekst(0) = MIM_GLOBAL_DATE
    psTekst(2) = Mid(cmbDokumentType.text, 4) + " " + Mid(Mim.Caption, InStr(Mim.Caption, "["))
    InitVelden
    PrintTitel
    GoSub AfschrijvingsLijnErBij
    Do
        bNext TABLE_VARIOUS
        If Ktrl Or Left(KEY_BUF(TABLE_VARIOUS), 2) <> "18" Then
            Exit Do
        Else
            GoSub AfschrijvingsLijnErBij
        End If
    Loop
End If
PrintTotaal
Printer.EndDoc
If bhKontrole = True Then
    AfschrijvingenLijstOk = False
Else
    AfschrijvingenLijstOk = True
End If
xDokument.Enabled = True
Screen.MousePointer = vbNormal
Exit Function

AfschrijvingsLijnErBij:
omsLijn = ""
RecordToVeld TABLE_VARIOUS
If Len(vBibTekst(TABLE_VARIOUS, "#v083 #")) <> 8 Then
    omsLijn = "Datumformaat '" + vBibTekst(TABLE_VARIOUS, "#v083 #") + "' onjuist voor " + vBibTekst(TABLE_VARIOUS, "#v087 #")
    MsgBox omsLijn
    Return
End If
    
DatumJaarInvest = vBibTekst(TABLE_VARIOUS, "#v083 #")
DatumJaarInvest = Mid(DatumJaarInvest, 1, 2) + "/" + Mid(DatumJaarInvest, 3)
DatumJaarInvest = Mid(DatumJaarInvest, 1, 5) + "/" + Mid(DatumJaarInvest, 6)

If DATE_KEY(DatumJaarInvest) > Mid(BOOKYEAR_FROMTO, 9) Then
    Return
Else
    ipct = Val(vBibTekst(TABLE_VARIOUS, "#v082 #"))
    dbdrg = Val(vBibTekst(TABLE_VARIOUS, "#v084 #"))
    VeldTXT(4) = DatumJaarInvest
    VeldTXT(5) = Dec(dbdrg, MASK_2002)
    VeldTXT(7) = Dec((ipct), "#####")
    'investeringsrekening
    VeldTXT(3) = Trim$(vBibTekst(TABLE_VARIOUS, "#v019 #"))
    bGet TABLE_LEDGERACCOUNTS, 0, VeldTXT(3)
    If Ktrl Then
        VeldTXT(0) = "Niet (meer) aanwezig..."
    Else
        RecordToVeld TABLE_LEDGERACCOUNTS
        VeldTXT(0) = vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #")
    End If
    'afschrijvingsrekening
    VeldTXT(6) = Trim$(vBibTekst(TABLE_VARIOUS, "#v087 #"))
    bGet TABLE_LEDGERACCOUNTS, 0, VeldTXT(6)
    If Ktrl Then
        VeldTXT(1) = "Niet (meer) aanwezig..."
    Else
        RecordToVeld TABLE_LEDGERACCOUNTS
        VeldTXT(1) = vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #")
    End If
    'afschrijvingsKOSTrekening
    VeldTXT(9) = Trim$(vBibTekst(TABLE_VARIOUS, "#v088 #"))
    bGet TABLE_LEDGERACCOUNTS, 0, VeldTXT(9)
    If Ktrl Then
        VeldTXT(2) = "Niet (meer) aanwezig..."
    Else
        RecordToVeld TABLE_LEDGERACCOUNTS
        VeldTXT(2) = vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #")
    End If
    If bhKontrole Then
        GoSub bhSPECIAAL
    Else
        dRa = Val(vBibTekst(TABLE_VARIOUS, "#v085 #"))
        If Val(vBibTekst(TABLE_VARIOUS, "#v084 #")) = Val(vBibTekst(TABLE_VARIOUS, "#v085 #")) Then
            Das = 0
        Else
            'Bereken nieuwe afschrijving
            dRest = dbdrg - (dRa + Val(Format((dbdrg / ipct), "000000000")))

            If Abs(dRest) < 20 Then
                Das = Val(Format((dbdrg / ipct), "000000000")) + dRest
            Else
                Das = Val(Format((dbdrg / ipct), "000000000"))
            End If
        End If
        VeldTXT(8) = Dec(dRa, MASK_2002)
        VeldTXT(10) = Dec(Das, MASK_2002)
    End If
    dTas = dTas + Das
    PrintVelden
    Printer.Print vbCrLf;
End If
Return

bhSPECIAAL:
'Investeringsrekening
VeldTXT(3) = Trim$(vBibTekst(TABLE_VARIOUS, "#v019 #"))
'Afschrijvingsrekening
VeldTXT(6) = Trim$(vBibTekst(TABLE_VARIOUS, "#v087 #"))
'AfschrijvingsKOSTrekening
VeldTXT(9) = Trim$(vBibTekst(TABLE_VARIOUS, "#v088 #"))

'Reeds boekhoudkundig afgeschreven kontroleren
BedragBegin = 0
Das = 0
RekKontrole = VeldTXT(6)
GoSub OverLOOPJournaal
If Das <> 0 Then
    dRa = Val(vBibTekst(TABLE_VARIOUS, "#v085 #")) - Das
Else
    dRa = Val(vBibTekst(TABLE_VARIOUS, "#v085 #"))
End If
VeldTXT(8) = Dec(dRa, MASK_2002)
VeldTXT(10) = Dec(Das, MASK_2002)
Return

OverLOOPJournaal:
DummyDatum = ""
bGetOrGreater TABLE_JOURNAL, 0, RekKontrole + Left(BOOKYEAR_FROMTO, 8)
If Ktrl Or KEY_BUF(TABLE_JOURNAL) > RekKontrole + Right(BOOKYEAR_FROMTO, 8) Then
    Return
Else
    Do
        RecordToVeld TABLE_JOURNAL
        If Trim$(DummyDatum) = "" Then
            DummyDatum = vBibTekst(TABLE_JOURNAL, "#v066 #")
            BedragBegin = Val(vBibTekst(TABLE_JOURNAL, "#v068 #"))
        ElseIf DummyDatum = vBibTekst(TABLE_JOURNAL, "#v066 #") Then
            BedragBegin = BedragBegin + Val(vBibTekst(TABLE_JOURNAL, "#v068 #"))
        Else
            Das = Das + Val(vBibTekst(TABLE_JOURNAL, "#v068 #"))
        End If
        bNext TABLE_JOURNAL
        If Ktrl Or KEY_BUF(TABLE_JOURNAL) > RekKontrole + Right(BOOKYEAR_FROMTO, 8) Then
            Exit Do
        End If
    Loop
    If Das = 0 Then
        If Val(vBibTekst(TABLE_VARIOUS, "#v084 #")) = Val(vBibTekst(TABLE_VARIOUS, "#v085 #")) Then
            Das = 0
        Else
            Das = Abs(BedragBegin)
            BedragBegin = 0
        End If
    Else
        Das = Abs(Das)
    End If
End If
Return

End Function
Private Sub InitVelden()
Dim T As Integer

Select Case Left(cmbDokumentType.text, 2)
    Case "18"
        REPORT_FIELD(0) = "Investeringsrekening"
        REPORT_TAB(0) = 2

        REPORT_FIELD(1) = "Rekening Afschrijvingen"
        REPORT_TAB(1) = 43
        
        REPORT_FIELD(2) = "Afschrijvingskosten Rekening"
        REPORT_TAB(2) = 84
        
        REPORT_FIELD(3) = "Nummer"
        REPORT_TAB(3) = 2
        
        REPORT_FIELD(4) = "AankpDatum"
        REPORT_TAB(4) = 10

        REPORT_FIELD(5) = "Tot.Bedrag"
        REPORT_TAB(5) = 21
        
        REPORT_FIELD(6) = "Nummer"
        REPORT_TAB(6) = 43
        
        REPORT_FIELD(7) = "Jaren"
        REPORT_TAB(7) = 51
    
        REPORT_FIELD(8) = "Reeds afg."
        REPORT_TAB(8) = 57
        
        REPORT_FIELD(9) = "Nummer"
        REPORT_TAB(9) = 84
        
        REPORT_FIELD(10) = " Bedrag Af"
        REPORT_TAB(10) = 92
        tMaxVeld = 10
    Case Else
        MsgBox "Stop, nog niets voorzien."
End Select

If chkAfdrukInVenster Then
    Me.Hide
    Unload Xlog
    Xlog.Hide
    Xlog.Caption = cmbDokumentType.text
    Xlog.X.Cols = tMaxVeld + 1
    Xlog.X.Row = 0
    For T = 0 To tMaxVeld
        Xlog.X.Col = T
        Xlog.X.text = REPORT_FIELD(T)
    Next
    Me.Show
End If

End Sub


