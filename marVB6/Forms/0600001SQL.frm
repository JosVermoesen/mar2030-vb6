VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form HistoriekSQL 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6810
   ClientLeft      =   555
   ClientTop       =   2160
   ClientWidth     =   9045
   ForeColor       =   &H00000000&
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6810
   ScaleWidth      =   9045
   Begin VB.CommandButton cbKlembord 
      Caption         =   "Naar Klembord"
      Height          =   375
      Left            =   7560
      TabIndex        =   17
      Top             =   1320
      Width           =   1335
   End
   Begin VB.ComboBox cbv041 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   0
      Width           =   7335
   End
   Begin VB.CheckBox GansePeriode 
      Alignment       =   1  'Right Justify
      Caption         =   "&Boekjaar"
      Enabled         =   0   'False
      Height          =   255
      Left            =   7800
      TabIndex        =   15
      Top             =   5760
      Value           =   1  'Checked
      Width           =   1110
   End
   Begin MSFlexGridLib.MSFlexGrid grdJournaalDetail 
      Height          =   5865
      Left            =   0
      TabIndex        =   14
      Top             =   330
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   10345
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   16777088
      AllowUserResizing=   1
   End
   Begin VB.TextBox TxtLijnen 
      BackColor       =   &H00FFFFFF&
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
      Height          =   315
      Left            =   7560
      MaxLength       =   4
      TabIndex        =   9
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   6360
      Width           =   2535
   End
   Begin VB.CommandButton CmdLager 
      Caption         =   "&Lager"
      Height          =   255
      Left            =   7560
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   840
      Width           =   1275
   End
   Begin VB.CommandButton CmdHoger 
      Caption         =   "&Hoger"
      Height          =   315
      Left            =   7560
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   1275
   End
   Begin VB.TextBox Rekening 
      BackColor       =   &H00FFFFFF&
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
      Height          =   285
      Left            =   5280
      TabIndex        =   1
      ToolTipText     =   "Nummer algemene rekening.  Geef één of meer cijfers en druk CTRL om geïndexeerd te zoeken"
      Top             =   6360
      Width           =   1275
   End
   Begin VB.CommandButton Zoeken 
      Caption         =   "&Zoek"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6600
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6360
      Width           =   795
   End
   Begin VB.CommandButton Sluiten 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   315
      Left            =   7560
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Creditsaldo = bedrag in min !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   5
      Left            =   7560
      TabIndex        =   13
      Top             =   4320
      Width           =   1395
      WordWrap        =   -1  'True
   End
   Begin VB.Label LblSaldo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   12
      Top             =   4080
      Width           =   1395
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo huidige selektie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   4
      Left            =   7560
      TabIndex        =   11
      Top             =   3600
      Width           =   1395
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lijnen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   8340
      TabIndex        =   10
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tussenstop &Melden na:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   2
      Left            =   7560
      TabIndex        =   8
      Top             =   4920
      Width           =   1365
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "&Periode Van - Tot"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   6360
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "&Rekening"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   4320
      TabIndex        =   0
      Top             =   6360
      Width           =   855
   End
End
Attribute VB_Name = "HistoriekSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit

Dim RekeningNummer          As String * 7
Dim PositieStelsel          As Integer

Dim Van                     As String
Dim Tot                     As String

Dim dTotaalSaldo            As Currency

Dim rsJourHier As ADODB.Recordset

Private Sub cbKlembord_Click()

    On Error Resume Next
    
    If Me.grdJournaalDetail.Clip = "" Then
        MsgBox "Eerst selecteren a.u.b. !"
        Exit Sub
    Else
        Screen.MousePointer = vbHourglass
        Me.Refresh
        Err = 0
        Clipboard.Clear
        Clipboard.SetText Me.grdJournaalDetail.Clip
        Screen.MousePointer = vbNormal
    End If
    If Err Then
        Screen.MousePointer = vbNormal
        MsgBox "Kopiëren naar het klembord was onvolledig (afhankelijk van het gebruikte besturingssysteem & werkgeheugen). Verklein de selectie en probeer opnieuw.", vbExclamation
        Exit Sub
    End If
    
End Sub

Private Sub CmdHoger_Click()

bNext TABLE_LEDGERACCOUNTS
If Ktrl Then
Else
    Rekening.text = KEY_BUF(TABLE_LEDGERACCOUNTS)
    Zoeken_Click
End If

End Sub


Private Sub CmdLager_Click()

bPrev TABLE_LEDGERACCOUNTS
If Ktrl Then
Else
    Rekening.text = KEY_BUF(TABLE_LEDGERACCOUNTS)
    Zoeken_Click
End If

End Sub


Private Sub Form_Load()

    Dim T As Integer

    Top = 0
    Left = 0

    TxtLijnen = LaadTekst("HistoriekInScherm", "MaxLijnen")
    TekstLijn = DATE_TEXT(Left(BOOKYEAR_FROMTO, 8)) + " - " + DATE_TEXT(Right(BOOKYEAR_FROMTO, 8))

    Rekening.text = BasisB(TABLE_LEDGERACCOUNTS).TekstInfo
    HistoriekSQL.Caption = "Historiek (" + RTrim$(vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #")) + ")"

    With grdJournaalDetail
        .Cols = 7
        .Row = 0
    
        .Col = 0
        .text = "document"
        .Col = 1
        .text = "Datum Doc."
        .Col = 2
        .text = "Fin.stuk"
        .Col = 3
        .text = "Omschrijving"
        .Col = 4
        If bhEuro Then
            .text = "EUR"
        Else
            .text = "BEF"
        End If
        .Col = 5
        If bhEuro Then
            .text = "BEF"
        Else
            .text = "EUR"
        End If
        .Col = 6
        .text = "T.Rek."
    End With
    
    grdJournaalDetail.ColWidth(0) = 1140
    grdJournaalDetail.ColWidth(1) = 1020
    grdJournaalDetail.ColWidth(2) = 930
    grdJournaalDetail.ColWidth(3) = 1755
    grdJournaalDetail.ColWidth(4) = 870
    grdJournaalDetail.ColWidth(5) = 975
    grdJournaalDetail.ColWidth(6) = 735
    For COUNT_TO = 0 To grdJournaalDetail.Cols - 1
        grdJournaalDetail.ColAlignment(COUNT_TO) = flexAlignLeftCenter
    Next
    grdJournaalDetail.ColAlignment(4) = flexAlignRightCenter
    grdJournaalDetail.ColAlignment(5) = flexAlignRightCenter
    Zoeken_Click
    
End Sub



Private Sub Form_Unload(Cancel As Integer)

BeWaarTekst "HistoriekInScherm", "MaxLijnen", TxtLijnen

End Sub


Private Sub GansePeriode_Click()

If GansePeriode Then
    TekstLijn = DATE_TEXT(Left(BOOKYEAR_FROMTO, 8)) + " - " + DATE_TEXT(Right(BOOKYEAR_FROMTO, 8))
Else
    TekstLijn = DATE_TEXT(Left(PERIOD_FROMTO, 8)) + " - " + DATE_TEXT(Right(PERIOD_FROMTO, 8))
End If
Zoeken_Click

End Sub

Private Sub Rekening_GotFocus()

Rekening.SelLength = Len(Rekening.text)
SnelHelpPrint "Dubbelklikken of [Ctrl] voor geïndexeerd zoeken", BL_LOGGING

End Sub

Private Sub Rekening_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case 17
        aIndex = 0
        SharedFl = TABLE_LEDGERACCOUNTS
        GridText = Rekening.text
        SqlSearch.Show 1
        If Ktrl = 0 Then
            Rekening.text = vBibTekst(TABLE_LEDGERACCOUNTS, "#v019 #")
            Zoeken_Click
        End If
End Select

End Sub

Private Sub Rekening_LostFocus()

RekeningNummer = Rekening.text
bGet TABLE_LEDGERACCOUNTS, 0, RekeningNummer
If Ktrl Then
    HistoriekSQL.Caption = "Historiek"
Else
    RecordToVeld TABLE_LEDGERACCOUNTS
    HistoriekSQL.Caption = "Historiek (" + RTrim$(vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #")) + ")"
End If

End Sub

Private Sub sluiten_Click()

Unload Me

End Sub

Private Sub TekstLijn_LostFocus()

If DATE_INVALID(Right(TekstLijn, 10)) Then
    MsgBox "Respecteer : " + vbCrLf + vbCrLf + "DD/MM/EEJJ - DD/MM/EEJJ a.u.b. !"
    If GansePeriode Then
        TekstLijn = DATE_TEXT(Left(BOOKYEAR_FROMTO, 8)) + " - " + DATE_TEXT(Right(BOOKYEAR_FROMTO, 8))
    Else
        TekstLijn = DATE_TEXT(Left(PERIOD_FROMTO, 8)) + " - " + DATE_TEXT(Right(PERIOD_FROMTO, 8))
    End If
    TekstLijn.SetFocus
    Exit Sub
ElseIf Len(TekstLijn) <> 23 Then
    MsgBox "Respecteer : " + vbCrLf + vbCrLf + "DD/MM/EEJJ - DD/MM/EEJJ a.u.b. !"
    If GansePeriode Then
        TekstLijn = DATE_TEXT(Left(BOOKYEAR_FROMTO, 8)) + " - " + DATE_TEXT(Right(BOOKYEAR_FROMTO, 8))
    Else
        TekstLijn = DATE_TEXT(Left(PERIOD_FROMTO, 8)) + " - " + DATE_TEXT(Right(PERIOD_FROMTO, 8))
    End If
    TekstLijn.SetFocus
Else
    Zoeken_Click
End If

End Sub


Private Sub Zoeken_Click()
Dim aa          As String
Dim T           As Integer
Dim maxLijn     As Integer
Dim Zoekverder As Boolean

SnelHelpPrint "Bezig...", BL_LOGGING
Zoekverder = False
GridSchoon grdJournaalDetail
LblSaldo = ""

T = 0
dTotaalSaldo = 0

On Local Error Resume Next
Err = 0
maxLijn = TxtLijnen + 2
If Err Then
    TxtLijnen = 300
    maxLijn = TxtLijnen + 2
End If
RekeningNummer = Rekening.text
Van = RekeningNummer + Mid(TekstLijn, 7, 4) + Mid(TekstLijn, 4, 2) + Mid(TekstLijn, 1, 2)
Tot = RekeningNummer + Mid(TekstLijn, 20, 4) + Mid(TekstLijn, 17, 2) + Mid(TekstLijn, 14, 2)

GetIt:
bGet TABLE_LEDGERACCOUNTS, 0, RekeningNummer
If Ktrl Then
    HistoriekSQL.Caption = "Historiek"
Else
    RecordToVeld TABLE_LEDGERACCOUNTS
    HistoriekSQL.Caption = "Historiek (" + Trim$(rsMAR(TABLE_LEDGERACCOUNTS)("v020")) + ")"
    grdJournaalDetail.Visible = False
    Me.Refresh
End If

'recordset maken
Set rsJourHier = New ADODB.Recordset
On Error Resume Next
Err = 0
rsJourHier.CursorLocation = adUseClient

'Msg = "SELECT * FROM Journalen WHERE v041 like '%' AND v070 >='" & _
    Van & "' AND v070 <= '" & _
    Tot & "'"

Msg = "SELECT * FROM Journalen WHERE v070 >='" & _
    Van & "' AND v070 <= '" & _
    Tot & "' ORDER BY v066"
    
If BasisB(3).CheckBoxSorting.Value = vbChecked Then
    Msg = Msg & " DESC"
End If


rsJourHier.Open Msg, adntDB, adOpenForwardOnly, adLockReadOnly
If Err Then
    MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
    Screen.MousePointer = vbNormal
    Exit Sub
End If
'Stop
'If rsJourHier.RecordCount < 1 Then Exit Sub

'bFirst TABLE_JOURNAL, 0
'bGetOrGreater TABLE_JOURNAL, 0, Van
'If Ktrl Or KEY_BUF(TABLE_JOURNAL) > Tot Then
If rsJourHier.RecordCount < 1 Then
    aa = vbTab & vbTab & vbTab & vbTab & vbTab & vbTab
    grdJournaalDetail.AddItem aa, grdJournaalDetail.Rows - 1
    GoTo JournaalJump
Else
    'RecordToVeld TABLE_JOURNAL
    Rekening.text = rsJourHier("v019")
    Mid(Van, 1, 7) = Mid(rsJourHier("v070"), 1, 7)
    Mid(Tot, 1, 7) = Mid(rsJourHier("v070"), 1, 7)
    bGet TABLE_LEDGERACCOUNTS, 0, rsJourHier("v019")
    If Ktrl Then
        HistoriekSQL.Caption = "Historiek"
    Else
        RecordToVeld TABLE_LEDGERACCOUNTS
        HistoriekSQL.Caption = "Historiek (" + Trim$(rsMAR(TABLE_LEDGERACCOUNTS)("v020")) + ")"
    End If
End If
Screen.MousePointer = vbHourglass
Do While Not rsJourHier.EOF
    GoSub LijnErBij
    rsJourHier.MoveNext
Loop
rsJourHier.Close
Set rsJourHier = Nothing

JournaalJump:
grdJournaalDetail.Row = 1
grdJournaalDetail.Col = 0
Screen.MousePointer = vbNormal
LblSaldo = Format(dTotaalSaldo, "#,##0.00")
grdJournaalDetail.Visible = True
On Local Error Resume Next
grdJournaalDetail.SetFocus
Exit Sub

LijnErBij:
aa = rsJourHier("v033") & vbTab
aa = aa & DATE_TEXT(rsJourHier("v035")) & vbTab
aa = aa & rsJourHier("v038") & vbTab
aa = aa & rsJourHier("v067") & vbTab
aa = aa & Format(rsJourHier("dece068"), "#,##0.00") & vbTab
aa = aa & Format(EURO * rsJourHier("dece068"), "#,##0") & vbTab
dTotaalSaldo = dTotaalSaldo + rsJourHier("dece068").Value
aa = aa & rsJourHier("v069")

On Local Error Resume Next
grdJournaalDetail.AddItem aa, grdJournaalDetail.Rows - 1
If Err Then
    MsgBox Error
End If
If grdJournaalDetail.Rows >= maxLijn Then
    If Zoekverder Then
    Else
        Msg = "Meer dan " + Format(maxLijn - 2) + " journaallijnen.  Enkel de eerste " + Format(maxLijn - 2) + " tonen ?"
        KtrlBox = MsgBox(Msg, vbYesNo + vbDefaultButton1)
        If KtrlBox = vbYes Then
            bLast TABLE_JOURNAL, 0
        Else
            Zoekverder = True
        End If
    End If
End If
Return

End Sub
