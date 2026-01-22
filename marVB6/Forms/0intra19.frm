VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Intrastat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Intrastat"
   ClientHeight    =   3735
   ClientLeft      =   555
   ClientTop       =   1110
   ClientWidth     =   7890
   DrawStyle       =   1  'Dash
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
   ScaleHeight     =   3735
   ScaleWidth      =   7890
   Begin MSAdodcLib.Adodc InfoData 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   3405
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Intrastat"
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Eenheden 
      DataField       =   "Eenheden"
      DataSource      =   "InfoData"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   6270
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      Top             =   1560
      Width           =   1545
   End
   Begin VB.TextBox InfoTekst 
      DataField       =   "InfoTekst"
      DataSource      =   "InfoData"
      Height          =   975
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      Top             =   2400
      Width           =   7755
   End
   Begin VB.ComboBox KeuzeOpties 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   3480
   End
   Begin VB.ComboBox KeuzeOpties 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   540
      Width           =   2955
   End
   Begin VB.ComboBox KeuzeOpties 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   2
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   840
      Width           =   2970
   End
   Begin VB.ComboBox KeuzeOpties 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   3
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1140
      Width           =   4830
   End
   Begin VB.ComboBox KeuzeOpties 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   4
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1440
      Width           =   4830
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   3570
      TabIndex        =   1
      Top             =   240
      Width           =   765
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Goederenkode"
      DataSource      =   "InfoData"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   1380
      TabIndex        =   6
      Top             =   1740
      Width           =   1560
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   1380
      TabIndex        =   7
      Top             =   2040
      Width           =   1560
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   4650
      TabIndex        =   8
      Top             =   1770
      Width           =   1560
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   4650
      TabIndex        =   9
      Top             =   2070
      Width           =   1560
   End
   Begin VB.CommandButton Ok 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   6330
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   90
      Width           =   1440
   End
   Begin VB.CommandButton Annuleren 
      Cancel          =   -1  'True
      Caption         =   "Annuleren"
      Height          =   405
      Left            =   6330
      TabIndex        =   10
      Top             =   510
      Width           =   1440
   End
   Begin VB.Label Label1 
      Caption         =   "Aanvullende eenh."
      Height          =   240
      Index           =   11
      Left            =   6270
      TabIndex        =   26
      Top             =   1290
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "Nog toe te wijzen"
      Height          =   240
      Index           =   9
      Left            =   4560
      TabIndex        =   23
      Top             =   420
      Width           =   1560
   End
   Begin VB.Label LabelA6 
      Caption         =   "test"
      Height          =   240
      Left            =   90
      TabIndex        =   11
      Top             =   30
      Width           =   1620
   End
   Begin VB.Label Label1 
      Caption         =   "Aanvullende eenh."
      Height          =   240
      Index           =   1
      Left            =   3000
      TabIndex        =   12
      Top             =   1800
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "Statisch. waarde"
      Height          =   240
      Index           =   2
      Left            =   3000
      TabIndex        =   13
      Top             =   2070
      Width           =   1620
   End
   Begin VB.Label Label8 
      Caption         =   "test"
      Height          =   240
      Left            =   60
      TabIndex        =   14
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "VervoerWijze"
      Height          =   240
      Index           =   4
      Left            =   60
      TabIndex        =   15
      Top             =   540
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Netto (Kg)"
      Height          =   240
      Index           =   5
      Left            =   60
      TabIndex        =   16
      Top             =   2040
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "GoederenKode"
      Height          =   240
      Index           =   6
      Left            =   60
      TabIndex        =   17
      Top             =   1740
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Transaktie B"
      Height          =   240
      Index           =   7
      Left            =   60
      TabIndex        =   18
      Top             =   1440
      Width           =   1275
   End
   Begin VB.Label LabelB6 
      Caption         =   "test"
      Height          =   240
      Left            =   2520
      TabIndex        =   19
      Top             =   30
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Transaktie A"
      Height          =   240
      Index           =   10
      Left            =   60
      TabIndex        =   20
      Top             =   1140
      Width           =   1275
   End
   Begin VB.Label NogToeTeWijzen 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   4560
      TabIndex        =   21
      Top             =   690
      Width           =   1560
   End
End
Attribute VB_Name = "Intrastat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit

Dim sIInput(13) As String
Dim IntraType           As String * 2

Private Sub Annuleren_Click()

Msg = "Aangifte voor deze faktuur negeren !  Bent U zeker ?"
Ktrl = MsgBox(Msg, 292, "INSTRAT 19 overslaan")
If Ktrl = 6 Then
    Unload Me
End If

End Sub

Private Sub Eenheden_Change()

Select Case Eenheden.Text
    Case "-", ""
        TekstInfo(3).Enabled = False

    Case Else
        TekstInfo(3).Enabled = True
End Select
TekstInfo(3).Text = ""

End Sub


Private Sub Form_Initialize()

    InfoData.ConnectionString = ADOJET_PROVIDER & "Data Source=" + PROGRAM_LOCATION + "Default2022.mdb;Persist Security Info=False"
    InfoData.Refresh
    
End Sub

Private Sub Form_Load()
Dim T As Integer

On Error Resume Next
Err = 0

KeuzeOpties(0).AddItem "001: Frankrijk"
KeuzeOpties(0).AddItem "002: G.H. Luxemburg"
KeuzeOpties(0).AddItem "003: Nederland"
KeuzeOpties(0).AddItem "004: Duitsland"
KeuzeOpties(0).AddItem "005: Italië"
KeuzeOpties(0).AddItem "006: Verenigd Koninkrijk"
KeuzeOpties(0).AddItem "007: Ierland"
KeuzeOpties(0).AddItem "008: Denemarken"
KeuzeOpties(0).AddItem "009: Griekenland"
KeuzeOpties(0).AddItem "010: Portugal"
KeuzeOpties(0).AddItem "011: Spanje en Canarische Eilanden"
KeuzeOpties(0).AddItem "030: Zweden"
KeuzeOpties(0).AddItem "032: Finland"
KeuzeOpties(0).AddItem "038: Oostenrijk"
KeuzeOpties(0).AddItem "046: Malta"
KeuzeOpties(0).AddItem "053: Estland"
KeuzeOpties(0).AddItem "054: Letland"
KeuzeOpties(0).AddItem "055: Litouwen"
KeuzeOpties(0).AddItem "060: Polen"
KeuzeOpties(0).AddItem "061: Tsjechië"
KeuzeOpties(0).AddItem "063: Slowakije"
KeuzeOpties(0).AddItem "064: Hongarije"
KeuzeOpties(0).AddItem "091: Slovenië"
KeuzeOpties(0).AddItem "600: Cyprus"

KeuzeOpties(1).AddItem "1: Vervoer over zee"
KeuzeOpties(1).AddItem "2: Vervoer per spoor"
KeuzeOpties(1).AddItem "3: Wegvervoer"
KeuzeOpties(1).AddItem "4: Luchtvervoer"
KeuzeOpties(1).AddItem "5: Postzendingen"
KeuzeOpties(1).AddItem "7: Vaste transportinrichtingen"
KeuzeOpties(1).AddItem "8: Binnenwateren"
KeuzeOpties(1).AddItem "9: Eigen kracht"

KeuzeOpties(2).AddItem "X: Andere dan zee"
KeuzeOpties(2).AddItem "1: Zee via Antwerpen 2000-2070/9120-9130"
KeuzeOpties(2).AddItem "2: Zee via Gent 9000-9060"
KeuzeOpties(2).AddItem "3: Zee via Zeebrugge 8000-8380"
KeuzeOpties(2).AddItem "4: Zee via Oostende 8400"
KeuzeOpties(2).AddItem "5: Zee via Brussel 1020-1210"
KeuzeOpties(2).AddItem "6: Noorden Brussel 1800-1980/2830-2880"
KeuzeOpties(2).AddItem "7: Agglomeratie Luik 4000-4684"
KeuzeOpties(2).AddItem "8: Nieuwpoort 8620"
KeuzeOpties(2).AddItem "9: Overige"

KeuzeOpties(3).AddItem "1: Eigendomsoverdracht met financiële compensatie"
KeuzeOpties(3).AddItem "2: Retourzendingen"
KeuzeOpties(3).AddItem "3: Eigendomsoverdracht zonder fin. compensatie"
KeuzeOpties(3).AddItem "4: Loonveredeling, herstelling intergouvern. programma's"
KeuzeOpties(3).AddItem "5: Loonveredel., herstell. buiten intergouvern. prog."
KeuzeOpties(3).AddItem "7: Gecoördineerde defensieprojecten e.a. (vb. Airbus)"
KeuzeOpties(3).AddItem "8: Bouwmaterialen en uitrusting burgerlijke bouwkunde"
KeuzeOpties(3).AddItem "9: Andere, niet elders geregistreerde transacties"

IntraType = Format(aIndex)
Caption = Caption + " " + IntraType

Select Case IntraType
    Case "19"
        LabelA6.Caption = "LidStaat Herkomst"
        LabelB6.Caption = "Land van Oorsprong"
        Label8.Caption = "Plaats lossen"
    Case "29"
        LabelA6.Caption = "LidStaat bestemming"
        LabelB6.Caption = ""
        Label8.Caption = "Plaats laden"
        TekstInfo(0).Enabled = False
End Select
KeuzeOpties(0).Enabled = True
NogToeTeWijzen.Caption = Left(GridText, 9)
sIInput(6) = vBibTekst(Fl, "#v149 #")
T = KeuzeOpties(0).ListCount
Do While T >= 0
    KeuzeOpties(0).ListIndex = T - 1
    If Left(KeuzeOpties(0).Text, 3) = sIInput(6) Then
        Exit Do
    Else
        T = T - 1
    End If
Loop
KeuzeOpties(2).ListIndex = 0
KeuzeOpties(1).ListIndex = 2
KeuzeOpties(3).ListIndex = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

bClose TABLE_VARIOUS

End Sub


Private Sub KeuzeOpties_Click(Index As Integer)

Select Case Index
    Case 0
        If IntraType = "19" Then
            TekstInfo(0).Text = Left(KeuzeOpties(0).Text, 3)
        Else
            TekstInfo(0).Text = ""
        End If
    Case 1
        If Left(KeuzeOpties(Index).Text, 1) = "1" Then
            KeuzeOpties(2).Enabled = True
            KeuzeOpties(2).ListIndex = 1
        Else
            KeuzeOpties(2).Enabled = False
            KeuzeOpties(2).ListIndex = 0
        End If
    Case 3
        Do While KeuzeOpties(4).ListCount
            KeuzeOpties(4).RemoveItem 0
        Loop
        Select Case Left(KeuzeOpties(3).Text, 1)
            Case "1"
                KeuzeOpties(4).AddItem "1: Definitieve aankoop/verkoop"
                KeuzeOpties(4).AddItem "2: Zichtzending/op proef of via commissionair"
                KeuzeOpties(4).AddItem "3: Ruilhandel (compensatie in natura)"
                KeuzeOpties(4).AddItem "4: Persoonlijke aankopen door reiziger"
                KeuzeOpties(4).AddItem "5: Financiële leasing"
                KeuzeOpties(4).ListIndex = 0
            Case "2"
                KeuzeOpties(4).AddItem "1: Terugzending van goederen"
                KeuzeOpties(4).AddItem "2: Vervanging van teruggezonden goederen"
                KeuzeOpties(4).AddItem "3: Vervanging goederen die niet teruggezonden zijn"
                KeuzeOpties(4).ListIndex = 0
            Case "3"
                KeuzeOpties(4).AddItem "1: Door E.U. gefinancierde hulpprogramma's"
                KeuzeOpties(4).AddItem "2: Andere algemene regeringshulp"
                KeuzeOpties(4).AddItem "3: Andere hulp (particuliere, niet gouvern. organ.)"
                KeuzeOpties(4).ListIndex = 0
            Case "4", "5"
                KeuzeOpties(4).AddItem "1: Loonveredeling"
                KeuzeOpties(4).AddItem "2: Onderhoud en herstelling tegen betaling"
                KeuzeOpties(4).AddItem "3: Onderhoud en herstelling kosteloos"
                KeuzeOpties(4).ListIndex = 0
            Case Else
                KeuzeOpties(4).AddItem "0: zonder meer..."
                KeuzeOpties(4).ListIndex = 0
        End Select
End Select

End Sub

Private Sub NogToeTeWijzen_Change()

TekstInfo(4).Text = NogToeTeWijzen.Caption

End Sub

Private Sub Ok_Click()
Dim TempoTel As Integer

If Len(TekstInfo(1).Text) <> 8 Or Val(TekstInfo(2).Text) = 0 Then
    Beep
    Exit Sub
ElseIf TekstInfo(3).Text = "" And TekstInfo(3).Enabled = True Then
    Beep
    MsgBox "Aanvullende eenheden a.u.b."
Else
          Msg = "Informatielijn opslaan..." + vbCrLf
    Msg = Msg + "Bent U zeker ?"
    Ktrl = MsgBox(Msg, 292, "Intrastat 19")
    If Ktrl = 6 Then
        TLB_RECORD(TABLE_VARIOUS) = ""
        FVT(TABLE_VARIOUS, 1) = IntraType + Left(KeuzeOpties(0).Text, 3)
        vBib TABLE_VARIOUS, Left(KeuzeOpties(0).Text, 3), "v072"
        FVT(TABLE_VARIOUS, 1) = FVT(TABLE_VARIOUS, 1) + vSet((TekstInfo(0).Text), 3)
        vBib TABLE_VARIOUS, (TekstInfo(0).Text), "v073"
        FVT(TABLE_VARIOUS, 1) = FVT(TABLE_VARIOUS, 1) + Left(KeuzeOpties(1).Text, 1)
        vBib TABLE_VARIOUS, Left(KeuzeOpties(1).Text, 1), "v074"
        FVT(TABLE_VARIOUS, 1) = FVT(TABLE_VARIOUS, 1) + Left(KeuzeOpties(2).Text, 1)
        vBib TABLE_VARIOUS, Left(KeuzeOpties(2).Text, 1), "v075"
        FVT(TABLE_VARIOUS, 1) = FVT(TABLE_VARIOUS, 1) + Left(KeuzeOpties(3).Text, 1)
        vBib TABLE_VARIOUS, Left(KeuzeOpties(3).Text, 1), "v076"
        FVT(TABLE_VARIOUS, 1) = FVT(TABLE_VARIOUS, 1) + Left(KeuzeOpties(4).Text, 1)
        vBib TABLE_VARIOUS, Left(KeuzeOpties(4).Text, 1), "v077"
        FVT(TABLE_VARIOUS, 1) = FVT(TABLE_VARIOUS, 1) + TekstInfo(1).Text
        vBib TABLE_VARIOUS, (TekstInfo(1).Text), "v078"
        vBib TABLE_VARIOUS, FVT(TABLE_VARIOUS, 1), "v005"
        vBib TABLE_VARIOUS, Dec$(Val(TekstInfo(2).Text), MASK_SY(0)), "v079"
        vBib TABLE_VARIOUS, Dec$(Val(TekstInfo(3).Text), MASK_SY(0)), "v080"
        vBib TABLE_VARIOUS, Dec$(Val(TekstInfo(4).Text), MASK_SY(0)), "v081"
        vBib TABLE_VARIOUS, vBibTekst(TABLE_INVOICES, "#v033 #"), "v033"
        vBib TABLE_VARIOUS, vBibTekst(TABLE_INVOICES, "#v035 #"), "v035"

        bInsert TABLE_VARIOUS, 1
        If Ktrl Then
            Exit Sub
        End If
        NogToeTeWijzen.Caption = Dec$(Val(NogToeTeWijzen.Caption) - Val(TekstInfo(4).Text), MASK_SY(0))
        If Val(NogToeTeWijzen.Caption) = 0 Then
            Unload Me
        Else
            TekstInfo(2).Text = ""
            TekstInfo(3).Text = ""
        End If
    End If
End If

End Sub

Private Sub TekstInfo_LostFocus(Index As Integer)

On Local Error Resume Next

Select Case Index
    Case 0
        TekstInfo(Index).Text = Format(Val(TekstInfo(Index).Text), "000")
    Case 1
        If TekstInfo(1).Text = "" Then Exit Sub
        'InfoData.RecordSource = "SELECT * FROM Intrastat WHERE Goederenkode LIKE '" + TekstInfo(Index).Text + "*';"
        'SnelHelpPrint (InfoData.RecordSource), BL_LOGGING
        'InfoData.Refresh
        
    Case 4
        If Val(TekstInfo(Index).Text) > Val(NogToeTeWijzen.Caption) Or Val(TekstInfo(Index).Text) = 0 Then
            Beep
            TekstInfo(Index).Text = NogToeTeWijzen.Caption
        End If
End Select

End Sub



