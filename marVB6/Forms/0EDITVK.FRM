VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form WijzigenVerkoop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gegevens Invoegen / Wijzigen"
   ClientHeight    =   3870
   ClientLeft      =   4020
   ClientTop       =   420
   ClientWidth     =   5400
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
   ScaleHeight     =   3870
   ScaleWidth      =   5400
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbSortering 
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
      Left            =   3060
      Style           =   2  'Dropdown List
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   9
      Left            =   3960
      TabIndex        =   27
      Top             =   990
      Width           =   1275
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   8
      Left            =   1680
      TabIndex        =   26
      Top             =   3210
      Width           =   1335
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   7
      Left            =   300
      TabIndex        =   25
      Top             =   3210
      Width           =   1335
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   6
      Left            =   1500
      TabIndex        =   9
      Text            =   "1"
      Top             =   1290
      Width           =   1005
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   5
      Left            =   1500
      TabIndex        =   11
      Text            =   "0"
      Top             =   1590
      Width           =   675
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   1500
      TabIndex        =   7
      Text            =   "0"
      Top             =   990
      Width           =   1305
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   1500
      TabIndex        =   17
      Top             =   2220
      Width           =   1515
   End
   Begin VB.TextBox TekstInfo 
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
      Height          =   285
      Index           =   1
      Left            =   1500
      MaxLength       =   40
      TabIndex        =   4
      Top             =   690
      Width           =   3735
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   1500
      TabIndex        =   14
      Text            =   "1"
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   0
      Left            =   1500
      TabIndex        =   1
      Top             =   360
      Width           =   1515
   End
   Begin VB.CommandButton Ok 
      Caption         =   "Ok"
      Default         =   -1  'True
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
      Left            =   3960
      TabIndex        =   12
      Top             =   2760
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
      Height          =   315
      Left            =   4020
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1275
   End
   Begin VB.ComboBox Keuze 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1920
      Width           =   1440
   End
   Begin VB.ComboBox Keuze 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   1500
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CheckBox BTWType 
      Caption         =   "BTW &Filter"
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
      Left            =   2880
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1020
      Width           =   1050
   End
   Begin MSComctlLib.StatusBar LocalInfo 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   28
      Top             =   3585
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   503
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   12347
            MinWidth        =   12347
         EndProperty
      EndProperty
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
   Begin VB.Label lbMinStockInfo 
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ter Info: Minimum Stock overschreden"
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   3960
      TabIndex        =   31
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lbMilieu 
      Caption         =   "Product met heffing(en) !"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblSortering 
      Caption         =   "Ctrl Zoeken Standaard"
      Height          =   240
      Index           =   0
      Left            =   3180
      TabIndex        =   29
      Top             =   120
      Width           =   1995
   End
   Begin VB.Label ProduktLabel 
      Caption         =   "Produkt&Nr."
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Omschrijving"
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label VerkoopLabel 
      Caption         =   "&Verkoopprijs"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   1275
   End
   Begin VB.Label ProduktLabel 
      Caption         =   "&Verpakking"
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   1950
      Width           =   1065
   End
   Begin VB.Label VerkoopLabel 
      Caption         =   "&Aantal"
      Height          =   240
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   660
   End
   Begin VB.Label VerkoopLabel 
      Caption         =   "&Boekhouding"
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   2250
      Width           =   1230
   End
   Begin VB.Label VerkoopLabel 
      Caption         =   "B.T.W. kode"
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   2580
      Width           =   1155
   End
   Begin VB.Label VerkoopLabel 
      Caption         =   "&Korting"
      Height          =   240
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   1620
      Width           =   750
   End
   Begin VB.Label VerkoopLabel 
      Caption         =   "%"
      Height          =   240
      Index           =   7
      Left            =   2190
      TabIndex        =   21
      Top             =   1650
      Width           =   240
   End
   Begin VB.Label VerkoopLabel 
      Caption         =   "Totaal Ex. BTW"
      Height          =   240
      Index           =   5
      Left            =   120
      TabIndex        =   22
      Top             =   2940
      Width           =   1470
   End
   Begin VB.Label VerkoopLabel 
      Caption         =   "Totaal In. BTW"
      Height          =   240
      Index           =   6
      Left            =   1680
      TabIndex        =   23
      Top             =   2940
      Width           =   1380
   End
   Begin VB.Label VerkoopLabel 
      Caption         =   "Excl. BTW"
      Height          =   210
      Index           =   8
      Left            =   4260
      TabIndex        =   24
      Top             =   1320
      Width           =   975
   End
End
Attribute VB_Name = "WijzigenVerkoop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim DefaultKode             As String * 1
Dim LijnType                As String * 1
Dim OpbrengstDefault As String * 7

Dim SubEx                   As Double
Dim SubIn                   As Double
Dim dHg                     As Double
Dim nAnt                    As Single
Dim nVol                    As Single

Dim vkMaskAantal    As String
Dim adminNoVat As Boolean

Private Sub Annuleren_Click()

GridText = ""
WijzigenVerkoop.Hide

End Sub

Private Sub cmbSortering_Click()

BeWaarTekst Me.Name, "ProduktSortering", cmbSortering

End Sub

Private Sub Form_Load()
Dim T As Integer
Dim X As Integer
Dim Pipo As String

Taal = 2
If Trim(String99(READING, 20)) = "7" Then
    adminNoVat = True
Else
    adminNoVat = False
End If

For T = 0 To 9
    Keuze(0).AddItem fmarBoxText("004", Format(Taal), Format(T))
Next T
T = 65
Do While fmarBoxText("004", Format(Taal), Chr(T)) <> ""
    Keuze(0).AddItem fmarBoxText("004", Format(Taal), Chr(T))
    T = T + 1
Loop
Keuze(0).ListIndex = 0

'For T = 0 To 9
'    TaksKeuze(0).AddItem fmarBoxText("024", Format(Taal), Format(T))
'Next T
'Keuze(0).ListIndex = 0

'Pipo = SleutelDok(11)
OpbrengstDefault = XLogKey
'OpbrengstDefault = String99(READING, 25)

If adminNoVat Then
    DefaultKode = 6
Else
    DefaultKode = vBibTekst(TABLE_CUSTOMERS, "#v111 #")
    If DefaultKode = " " Then
        DefaultKode = String99(READING, 183)
    End If
End If


Dim PuntPos As Integer

vkMaskAantal = String99(READING, 300)
PuntPos = InStr(vkMaskAantal, ".")
If PuntPos Then
    vkMaskAantal = String(PuntPos - 1, "#") & Mid(vkMaskAantal, PuntPos)
End If
For T = 1 To 6
    If Len(fmarBoxText("002", "2", Format(T))) <> 0 Then
        Keuze(1).AddItem fmarBoxText("002", "2", Format(T))
    End If
Next T
If VAT_BOBTHEBUILDERS = True Then
    Keuze(1).Enabled = False
    Keuze(1).ListIndex = 0
ElseIf DefaultKode = "6" Then
    Keuze(1).ListIndex = Keuze(1).ListCount - 1
ElseIf Val(DefaultKode) > Keuze(1).ListCount - 1 Then
    MsgBox "Btw kode kan niet ! " + DefaultKode + vbCrLf + vbCrLf + "Kontroleer Setup boekjaar en parameters a.u.b. !"
    Keuze(1).ListIndex = 2
Else
    Keuze(1).ListIndex = Val(DefaultKode) - 1
End If

If GridText <> "" Then
    LijnType = Right(GridText, 1)
    Select Case LijnType
        Case "0", "1"
            Ok.Enabled = True
            TekstInfo(6).TabIndex = 0
            TekstInfo(1).MaxLength = 40
            TekstInfo(1).Text = RTrim$(Left(GridText, 40))
            GridText = Right(GridText, Len(GridText) - 41)
            TekstInfo(4).Text = Left(GridText, 11)
            GridText = Right(GridText, Len(GridText) - 12)
            TekstInfo(6).Text = Left(GridText, 7)
            GridText = Right(GridText, Len(GridText) - 8)
            TekstInfo(7).Text = Left(GridText, 12)
            GridText = Right(GridText, Len(GridText) - 13)
            TekstInfo(2).Text = Left(GridText, 6)
            GridText = Right(GridText, Len(GridText) - 7)
            For T = 0 To Keuze(0).ListCount - 1
                If Left(Keuze(0).List(T), 1) = Left(GridText, 1) Then
                    Keuze(0).ListIndex = T
                    Exit For
                End If
            Next
            GridText = Right(GridText, Len(GridText) - 2)
            TekstInfo(5).Text = Left(GridText, 3)
            GridText = Right(GridText, Len(GridText) - 4)
            If Val(Left(GridText, 1)) = 6 Then
                Keuze(1).ListIndex = Keuze(1).ListCount - 1
            Else
                Keuze(1).ListIndex = Val(Left(GridText, 1)) - 1
            End If
            GridText = Right(GridText, Len(GridText) - 2)
            TekstInfo(3).Text = Left(GridText, 7)
            GridText = Right(GridText, Len(GridText) - 8)
            TekstInfo(0).Text = Left(GridText, 13)
            GridText = Right(GridText, Len(GridText) - 14)
            GridText = ""
        Case "2"
            Ok.Enabled = True
            TekstInfo(1).MaxLength = 75
            TekstInfo(1).Text = RTrim$(Left(GridText, 75))
        Case Else
            MsgBox "Stop"
    End Select
Else
    TekstInfo(3).Text = OpbrengstDefault
    LijnType = Format(aIndex)
End If

Select Case LijnType
    Case "0"
        For T = 0 To FL_NUMBEROFINDEXEN(TABLE_PRODUCTS)
            cmbSortering.AddItem Format(T, "00") + ":" + FLINDEX_CAPTION(TABLE_PRODUCTS, T)
        Next
        On Error Resume Next
        cmbSortering = LaadTekst(Me.Name, "ProduktSortering")
        
        TekstInfo(1).Enabled = False
        'TekstInfo(4).Enabled = False
        
        MaakTotaal
    Case "1"
        TekstInfo(0).Text = "OMSCHRIJVING"
        TekstInfo(2).Text = Dec$(1, "###0.0")
        Keuze(0).ListIndex = 9
        
        ProduktLabel(0).Visible = False
        ProduktLabel(1).Visible = False
        TekstInfo(0).Visible = False
        TekstInfo(2).Visible = False
        Keuze(0).Visible = False
        MaakTotaal
        TekstInfo(1).TabIndex = 0
        
    Case "2"
        TekstInfo(1).MaxLength = 75
        'TekstInfo(1).Text = vSet((TekstInfo(1).Text), 75)
        Keuze(0).ListIndex = 0
        
        ProduktLabel(0).Visible = False
        ProduktLabel(1).Visible = False
        Keuze(0).Visible = False
        Keuze(1).Visible = False

        For T = 0 To 8
            TekstInfo(T).Visible = False
            VerkoopLabel(T).Visible = False
        Next
        TekstInfo(9).Visible = False
        BTWType.Visible = False
        TekstInfo(1).Visible = True
        Ok.Enabled = True
End Select
If Len(TekstInfo(6).Text) Then
Else
    TekstInfo(6).Text = "1"
End If

If XLogKassa = "" Then
    BTWType.Value = Val(String99(READING, 182))
Else
    BTWType.Value = 1
    TekstInfo(2).Enabled = False
    TekstInfo(3).Enabled = False
    'TekstInfo(4).Enabled = False
    TekstInfo(5).Enabled = False
    'TekstInfo(6).Enabled = False
    Keuze(0).Enabled = False
    Keuze(1).Enabled = False
    BTWType.Enabled = False
End If

End Sub

Private Sub Keuze_Click(Index As Integer)

Select Case Index
    Case 1
        MaakTotaal
End Select

End Sub

Private Sub MaakTotaal()

nVol = Val(TekstInfo(2).Text)
dHg = Val(TekstInfo(4).Text)
nAnt = Val(TekstInfo(6).Text)

SubEx = (dHg * nAnt) - ((dHg * nAnt) * Val(TekstInfo(5).Text) / 100)
SubIn = SubEx + (SubEx * Val(Mid(Keuze(1).Text, 4, 4)) / 100)
TekstInfo(7).Text = Dec$((SubEx), "#######0.000")
TekstInfo(8).Text = Dec$((SubIn), "#######0.000")
TekstInfo(9).Text = TekstInfo(4).Text

End Sub

Private Sub Ok_Click()

If TekstInfo(6) = "*" Then Exit Sub

Select Case LijnType
    Case "0", "1"
        TekstInfo(0).Text = vSet((TekstInfo(0).Text), 13) 'ProduktNummer
        TekstInfo(1).Text = vSet((TekstInfo(1).Text), 40) 'Omschrijving
        TekstInfo(2).Text = Dec$(Val(TekstInfo(2).Text), "###0.0") 'Verpakking
        TekstInfo(3).Text = vSet((TekstInfo(3).Text), 7) 'Boekhoudingrekening
        TekstInfo(4).Text = Dec$(Val(TekstInfo(4).Text), "######0.000") 'Verkoopprijs
        TekstInfo(5).Text = Dec$(Val(TekstInfo(5).Text), "##0") 'Korting
        TekstInfo(6).Text = Dec$(Val(TekstInfo(6).Text), vkMaskAantal) 'Aantal
        MaakTotaal
        
        'controle voor mileulijn!!
        If blMilieu = True Then
        
            MilieuGridText = ""
            
            Dim strMilieu() As String
            Dim telmilieu As Integer
            
            strMilieu = Split(vBibTekst(TABLE_PRODUCTS, "#v261 #"), ";")
            If LijnType = "1" Then
            Else
                For telmilieu = 0 To UBound(strMilieu)
                    bGet TABLE_PRODUCTS, 0, vSet(strMilieu(telmilieu), 13)
                    If Ktrl Then
                        MsgBox "subcode " & strMilieu(telmilieu) & " niet te vinden"
                        blMilieu = False
                    Else
                        RecordToVeld TABLE_PRODUCTS
                        MilieuGridText = MilieuGridText + vSet(vBibTekst(TABLE_PRODUCTS, "#v105 #"), 40) + Chr$(124) + Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#e112 #")), "######0.000") + Chr$(124) + TekstInfo(6) + Chr$(124) + Dec$(Val(TekstInfo(6)) * Val(vBibTekst(TABLE_PRODUCTS, "#e112 #")), "#######0.000") + Chr$(124) + Dec$(1, "###0.0") + Chr$(124)
                        MilieuGridText = MilieuGridText + vBibTekst(TABLE_PRODUCTS, "#v106 #") + Chr$(124) + Dec$(0, "##0") + Chr$(124) + vBibTekst(TABLE_PRODUCTS, "#v111 #") + Chr$(124) + vSet(vBibTekst(TABLE_PRODUCTS, "#v117 #"), 7) + Chr$(124) + vSet(vBibTekst(TABLE_PRODUCTS, "#v102 #"), 13) + Chr$(124) + "0" + vbCrLf
                    End If
                Next
            End If
        End If
        If TekstInfo(1).Text = Space$(Len(TekstInfo(1).Text)) Or TekstInfo(3).Text = Space$(Len(TekstInfo(3).Text)) Then
            TekstInfo(1).Text = ""
        Else
            GridText = TekstInfo(1).Text + Chr$(124) + TekstInfo(4).Text + Chr$(124) + TekstInfo(6).Text + Chr$(124) + TekstInfo(7).Text + Chr$(124) + TekstInfo(2).Text + Chr$(124)
            GridText = GridText + Left(Keuze(0).Text, 1) + Chr$(124) + TekstInfo(5).Text + Chr$(124) + Left(Keuze(1).Text, 1) + Chr$(124) + TekstInfo(3).Text + Chr$(124) + TekstInfo(0).Text + Chr$(124) + LijnType
            WijzigenVerkoop.Hide
        End If
        
    Case "2"
        TekstInfo(1).Text = vSet((TekstInfo(1).Text), 75)
        GridText = TekstInfo(1).Text + Chr$(124) + LijnType
        WijzigenVerkoop.Hide
End Select

End Sub

Private Function ProdukTekstInfo()
Dim HuidigeStock As Single
Dim MinimumStock As Single

If Trim(vBibTekst(TABLE_PRODUCTS, "#v261 #")) = "" Then
    Me.lbMilieu.Visible = False
Else
    Me.lbMilieu.Visible = True
End If
blMilieu = Me.lbMilieu.Visible

TekstInfo(0).Text = vBibTekst(TABLE_PRODUCTS, "#v102 #")
TekstInfo(1).Text = vBibTekst(TABLE_PRODUCTS, "#v105 #")
TekstInfo(5).Text = vBibTekst(TABLE_PRODUCTS, "#v300 #")

For T = 0 To Keuze(0).ListCount - 1
    If Left(Keuze(0).List(T), 1) = vBibTekst(TABLE_PRODUCTS, "#v106 #") Then
        Keuze(0).ListIndex = T
        Exit For
    End If
Next

If VAT_BOBTHEBUILDERS = True Then
    Keuze(1).ListIndex = 0
ElseIf Val(vBibTekst(TABLE_PRODUCTS, "#v111 #")) = 6 Then
    Keuze(1).ListIndex = 3
Else
    Keuze(1).ListIndex = Val(vBibTekst(TABLE_PRODUCTS, "#v111 #")) - 1
End If

nVol = Val(vBibTekst(TABLE_PRODUCTS, "#v107 #"))
TekstInfo(2).Text = Dec$((nVol), "###0.0")
If InStr(DIRECTSELL_STRING, "EUR") Then
    TekstInfo(4).Text = Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#e112 #")) * nVol, "######0.000")
Else
    TekstInfo(4).Text = Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#v112 #")) * nVol, "######0.000")
End If
TekstInfo(9).Text = TekstInfo(4).Text

dHg = Val(TekstInfo(4).Text)
If Val(TekstInfo(5).Text) <> 0 Then
Else
    TekstInfo(5).Text = Dec$(Val(vBibTekst(TABLE_CUSTOMERS, "#vs05 #")), "##0")
End If
TekstInfo(6).Text = Dec$(1, vkMaskAantal)
nAnt = 1

HuidigeStock = Val(vBibTekst(TABLE_PRODUCTS, "#v114 #")) + Val(vBibTekst(TABLE_PRODUCTS, "#v119 #")) - Val(vBibTekst(TABLE_PRODUCTS, "#v120 #"))
MinimumStock = Val(vBibTekst(TABLE_PRODUCTS, "#v115 #"))
If MinimumStock > HuidigeStock Then
    Me.lbMinStockInfo.Visible = True
Else
    Me.lbMinStockInfo.Visible = False
End If
LocalInfo.SimpleText = "Min.Stock: " + Format(MinimumStock) + " / Stock: " + Format(HuidigeStock) + " / Plaats: " + vBibTekst(TABLE_PRODUCTS, "#v109 #")


'TaksKeuze(0).ListIndex = Val(vBibTekst(TABLE_PRODUCTS, "#v168 #"))
'Select Case Val(vBibTekst(TABLE_PRODUCTS, "#v168 #"))
'    Case 1 To 9
'        TaksKeuze(0).Visible = True
'    Case Else
'        TaksKeuze(0).Visible = False
'End Select
MaakTotaal
'Ok.Enabled = True

End Function


Private Sub TekstInfo_DblClick(Index As Integer)

    TekstInfo_KeyDown Index, 17, 0

End Sub

Private Sub TekstInfo_GotFocus(Index As Integer)

TekstInfo(Index).SelStart = 0
TekstInfo(Index).SelLength = Len(TekstInfo(Index).Text)
Select Case Index
    Case 0, 3
        SnelHelpPrint "Dubbelklikken of [Ctrl] voor geïndexeerd zoeken", BL_LOGGING
        If TekstInfo(3).Text <> "" Then
            bGet TABLE_LEDGERACCOUNTS, 0, vSet((TekstInfo(3).Text), 7)
            If Ktrl Then
            Else
                RecordToVeld TABLE_LEDGERACCOUNTS
                SnelHelpPrint vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #"), BL_LOGGING
            End If
        End If
    Case 4
       Ok.Default = True
                    
    Case 6
        If Val(TekstInfo(Index).Text) = 0 Then
            TekstInfo(Index).Text = Dec$(1, "###.000")
        End If
End Select

Select Case Index
    Case 1
        TekstInfo(1).SelLength = 0
    Case Else
        TekstInfo(Index).SelLength = Len(TekstInfo(Index).Text)
End Select

End Sub

Private Sub TekstInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

If Index = 0 Then Ok.Enabled = False
Select Case KeyCode
    Case 13
        On Error Resume Next
        If Index = 4 Then Ok.Enabled = True:  Ok.SetFocus: Exit Sub
        If Index = 0 Then
            SendKeys "{TAB}", True
            Exit Sub
        End If
    Case 17
        Select Case Index
            Case 0
                Ok.Enabled = False
                SharedFl = TABLE_PRODUCTS
                aIndex = Val(Left(cmbSortering, 2))
                FVT(TABLE_PRODUCTS, 0) = String$(13, 0)
                GridText = TekstInfo(0).Text
                SqlSearch.Show 1
                If Ktrl Then
                    TekstInfo(0).Text = ""
                    TekstInfo(0).SetFocus: Exit Sub
                Else
                    RecordToVeld TABLE_PRODUCTS
                    ProdukTekstInfo
                    bGet TABLE_LEDGERACCOUNTS, 0, vSet(vBibTekst(TABLE_PRODUCTS, "#v117 #"), 7)
                    If Ktrl Then
                        TekstInfo(3).Text = ""
                        TekstInfo(0).SetFocus: Exit Sub
                    ElseIf Val(vBibTekst(TABLE_PRODUCTS, "#e113 #")) <= 0 Then
                        MsgBox "Aankoopprijs is onbekend.  Geef de inkoopprijs in van dit artikel a.u.b. via de produktfiche !"
                        TekstInfo(0).Text = ""
                        TekstInfo(0).SetFocus: Exit Sub
                    Else
                        RecordToVeld TABLE_LEDGERACCOUNTS
                        TekstInfo(3).Text = vBibTekst(TABLE_LEDGERACCOUNTS, "#v019 #")
                        SnelHelpPrint vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #"), BL_LOGGING
                        Ok.Enabled = True
                    End If
                End If
            Case 3
                SharedFl = TABLE_LEDGERACCOUNTS
                aIndex = 0
                FVT(TABLE_LEDGERACCOUNTS, 0) = String$(7, 0)
                GridText = TekstInfo(3).Text
                SqlSearch.Show 1
                If Ktrl Then
                    TekstInfo(1).Text = ""
                    FVT(TABLE_LEDGERACCOUNTS, 0) = ""
                    Ok.Enabled = False
                Else
                    RecordToVeld TABLE_LEDGERACCOUNTS
                    TekstInfo(3).Text = vBibTekst(TABLE_LEDGERACCOUNTS, "#v019 #")
                    Ok.Enabled = True
                End If
        End Select
End Select

End Sub

Private Sub TekstInfo_KeyPress(Index As Integer, KeyAscii As Integer)

Select Case Index
    Case 4
        TekstInfo(9).Text = ""
End Select

End Sub

Private Sub TekstInfo_LostFocus(Index As Integer)
Dim TempPCT As Single

On Local Error Resume Next

Select Case Index
    Case 0
        Ok.Enabled = True
        If Trim$(TekstInfo(0)) = "" Then
        Else
            bGet TABLE_PRODUCTS, 0, Trim$(TekstInfo(0).Text)
            If Ktrl Then
                TekstInfo(0).Text = ""
                TekstInfo(0).SetFocus: Exit Sub
            Else
                RecordToVeld TABLE_PRODUCTS
                ProdukTekstInfo
                bGet TABLE_LEDGERACCOUNTS, 0, vSet(vBibTekst(TABLE_PRODUCTS, "#v117 #"), 7)
                If Ktrl Then
                    MsgBox "Verkooprekening " & vSet(vBibTekst(TABLE_PRODUCTS, "#v117 #"), 7) & " bestaat niet (meer).  Eerst verbeteren a.u.b. in productfiche", vbExclamation
                    TekstInfo(0).SetFocus: Exit Sub
                Else
                    RecordToVeld TABLE_LEDGERACCOUNTS
                    TekstInfo(3).Text = vBibTekst(TABLE_LEDGERACCOUNTS, "#v019 #")
                    SnelHelpPrint vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #"), BL_LOGGING
                    Ok.Enabled = True
                End If
                TekstInfo(6).SetFocus
            End If
        End If
        
    Case 2
        nVol = Val(TekstInfo(2).Text)
        TekstInfo(2).Text = Dec$((nVol), "###0.0")
        If InStr(DIRECTSELL_STRING, "EUR") Then
            TekstInfo(4).Text = Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#e112 #")) * nVol, "######0.000")
        Else
            TekstInfo(4).Text = Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#v112 #")) * nVol, "######0.000")
        End If
        TekstInfo(9).Text = TekstInfo(4).Text

    Case 3
        bGet TABLE_LEDGERACCOUNTS, 0, vSet((TekstInfo(3).Text), 7)
        If Ktrl Then
            If FVT(TABLE_LEDGERACCOUNTS, 0) = String$(7, 0) Then
                Exit Sub
            Else
                TekstInfo(3).Text = OpbrengstDefault
                TekstInfo(3).SetFocus
            End If
        Else
            RecordToVeld TABLE_LEDGERACCOUNTS
            TekstInfo(3).Text = vBibTekst(TABLE_LEDGERACCOUNTS, "#v019 #")
            SnelHelpPrint vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #"), BL_LOGGING
        End If

    Case 4
        If TekstInfo(4).Text = TekstInfo(9).Text Then
        ElseIf BTWType.Value Then
            dHg = Val(TekstInfo(4).Text)
            TempPCT = Val(Mid(Keuze(1).Text, 4, 4))
            dHg = dHg * 100 / (100 + TempPCT)
            TekstInfo(4).Text = Dec$((dHg), "######0.000")
            MaakTotaal
            Ok.Enabled = True
        Else
            TekstInfo(4).Text = Dec$(Val(TekstInfo(4).Text), "######0.000")
            MaakTotaal
            Ok.Enabled = True
        End If
        
    Case 5
        If TekstInfo(0).Visible = True Then
            If (Val(TekstInfo(4).Text) / Val(TekstInfo(2).Text)) / (1 + (Val(TekstInfo(5).Text) / 100)) < Val(vBibTekst(TABLE_PRODUCTS, "#e113 #")) Then
                Msg = "Uw verkoopprijs wordt kleiner dan uw aankoopprijs !  Is dit de bedoeling ?"
                KtrlBox = MsgBox(Msg, 292)
                If KtrlBox = 6 Then
                Else
                    TekstInfo(5).Text = "0"
                End If
            End If
        End If
        TekstInfo(5).Text = Dec$(Val(TekstInfo(5).Text), "##0")

    Case 6
        If TekstInfo(6).Text = "*" Then
            If Keuze(0).Visible = True Then
                If Mid(WijzigenVerkoop.Keuze(0), 1, 1) = "5" Or Mid(WijzigenVerkoop.Keuze(0), 1, 1) = "3" Then
                    frmVanGucht.Show 1
                Else
                    TekstInfo(6).Text = Dec$(1, vkMaskAantal)
                End If
            Else
                TekstInfo(6).Text = Dec$(1, vkMaskAantal)
            End If
        End If
        nAnt = Val(TekstInfo(6).Text)
        TekstInfo(6).Text = Dec$((nAnt), vkMaskAantal)
        'If Ok.Enabled = True Then Ok.SetFocus
    End Select
MaakTotaal

End Sub

