VERSION 5.00
Begin VB.Form Wijzigen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gegevens Invoegen / Wijzigen"
   ClientHeight    =   2820
   ClientLeft      =   3285
   ClientTop       =   2220
   ClientWidth     =   10815
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
   ScaleHeight     =   2820
   ScaleWidth      =   10815
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton ButtonShowB2BDetails 
      Caption         =   "Lijndetails Toggle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4440
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.ListBox ListBoxLineItems 
      Columns         =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2310
      ItemData        =   "0EDIT.frx":0000
      Left            =   5640
      List            =   "0EDIT.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   240
      Width           =   4935
   End
   Begin VB.CommandButton cbFiche 
      Caption         =   "&Fiche"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox TekstInfo 
      Height          =   315
      Index           =   6
      Left            =   60
      TabIndex        =   13
      Top             =   2280
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.ComboBox cbEenheidsType 
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
      ItemData        =   "0EDIT.frx":0004
      Left            =   60
      List            =   "0EDIT.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1140
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.TextBox TekstInfo 
      Height          =   315
      Index           =   5
      Left            =   2940
      TabIndex        =   4
      Top             =   1140
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox TekstInfo 
      Enabled         =   0   'False
      Height          =   315
      Index           =   4
      Left            =   60
      TabIndex        =   3
      Top             =   780
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.TextBox TekstInfo 
      Height          =   315
      Index           =   3
      Left            =   2220
      TabIndex        =   1
      Top             =   420
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox TekstInfo 
      Height          =   315
      Index           =   2
      Left            =   2760
      TabIndex        =   9
      Top             =   2280
      Width           =   1635
   End
   Begin VB.TextBox TekstInfo 
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   60
      TabIndex        =   7
      Top             =   1920
      Width           =   4275
   End
   Begin VB.TextBox TekstInfo 
      Height          =   315
      Index           =   0
      Left            =   2820
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Ok 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Annuleren 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   360
      Left            =   4440
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1080
   End
   Begin VB.Label lPctFilter 
      Height          =   255
      Left            =   4560
      TabIndex        =   14
      Top             =   1980
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Te boeken"
      Height          =   240
      Index           =   1
      Left            =   4440
      TabIndex        =   12
      Top             =   2370
      Width           =   1200
   End
   Begin VB.Label Labelstock 
      Caption         =   "&Artikelnr."
      Height          =   240
      Index           =   0
      Left            =   1260
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "&Rekening"
      Height          =   240
      Index           =   0
      Left            =   1920
      TabIndex        =   5
      Top             =   1620
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "&Bedrag"
      Height          =   240
      Index           =   2
      Left            =   2040
      TabIndex        =   8
      Top             =   2400
      Width           =   675
   End
   Begin VB.Menu Opties 
      Caption         =   "Opties"
      Begin VB.Menu Filter 
         Caption         =   "Toggle Filter"
         Index           =   0
      End
   End
End
Attribute VB_Name = "Wijzigen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim CTRLFlag As Boolean

Dim VeldRekening As String * 7
Dim VeldNaam            As String * 40
Dim VeldBedrag          As String * 12
Dim VeldAantal          As String * 10
Dim VeldProdukt         As String * 13
Dim ProduktNaam         As String * 40
Dim iTabIndex           As Integer

Dim isB2BPeppol         As Boolean

Dim pctFilter           As String

Private Sub Annuleren_Click()

GridText = ""
Unload Me

End Sub

Private Sub CmdDommeKey_Click()

On Error Resume Next

'    If iTabIndex = 2 Then
'        Ok.SetFocus
'        Exit Sub
'    End If
'    Err = 0
'    On Local Error Resume Next
'    TekstInfo(iTabIndex + 1).SetFocus
'    If Err Then
'        iTabIndex = iTabIndex + 1
'        Err = 0
'        On Local Error Resume Next
'        TekstInfo(iTabIndex + 1).SetFocus
'        If Err Then
'            Ok.SetFocus
'        End If
'    End If

End Sub


Private Sub ButtonShowB2BDetails_Click()

    If Width = 5730 Then
        Width = 10900
    Else
        Width = 5730
    End If

End Sub

Private Sub cbEenheidsType_Click()

    On Error Resume Next
    TekstInfo(5).SetFocus
    
End Sub

Private Sub cbFiche_Click()

Dim dummykey As String

dummykey = TekstInfo(3).Text
bGet TABLE_PRODUCTS, 0, vSet((TekstInfo(3).Text), 13)
Unload frmProduktFiche
Load frmProduktFiche
frmProduktFiche.BorderStyle = 0
frmProduktFiche.TxtInfo(0).Text = TekstInfo(3).Text
frmProduktFiche.Show 1
TekstInfo(3).Text = Trim(dummykey)
TekstInfo(3).SetFocus
        
End Sub

Private Sub Filter_Click(Index As Integer)

    Me.Filter(0).Checked = Not Me.Filter(0).Checked
    If Me.Filter(0).Checked = True Then
        Me.TekstInfo(6).Visible = True
        Me.TekstInfo(6).TabIndex = 9
        If isB2BPeppol = False Then
            Me.TekstInfo(2).Enabled = False
        End If
        pctFilter = InputBox("Commissie Filter (vb. 25.75)", "Filter aan", "15")
    Else
        Me.TekstInfo(6).Visible = False
        If isB2BPeppol = False Then
            Me.TekstInfo(2).Enabled = True
        End If
        Me.TekstInfo(6).TabIndex = 9
        pctFilter = ""
    End If
    BeWaarTekst "Aankoopverrichting", "FilterToggle", Str$(Me.Filter(0).Checked)
    BeWaarTekst "Aankoopverrichting", "FilterPCT", pctFilter
    Me.lPctFilter.Caption = pctFilter
    
End Sub

Private Sub Form_Load()
Dim T As Integer
Dim X As Integer

If documentLinesDATA = "" Then
    isB2BPeppol = False
    Wijzigen.Width = 5730
    Wijzigen.Height = 3600
    TekstInfo(2).Enabled = True
    ButtonShowB2BDetails.Visible = False
Else
    isB2BPeppol = True
    TekstInfo(2).Enabled = False
    Wijzigen.Width = 5730
    Wijzigen.Height = 3600
    ButtonShowB2BDetails.Visible = True
    If DirekteAankoop.StockBeheer.Value = vbChecked Then
        ButtonShowB2BDetails_Click
    End If
    
       
    Dim counterHere As Integer
    Dim indexHere As Integer
    
    indexHere = DirekteAankoop.AankoopDetail.ListIndex
    For counterHere = 0 To 8
        ListBoxLineItems.AddItem documentLinesDATAArray(0, counterHere) & ": " & documentLinesDATAArray(indexHere + 1, counterHere)
    Next
    'AankoopDetail.ToolTipText = showString
End If

On Local Error Resume Next
Me.Filter(0).Checked = LaadTekst("Aankoopverrichting", "FilterToggle")
On Local Error Resume Next
pctFilter = LaadTekst("Aankoopverrichting", "FilterPCT")
If Me.Filter(0).Checked = True Then
    Me.TekstInfo(6).Visible = True
    Me.TekstInfo(6).TabIndex = 9
    If isB2BPeppol = False Then
        Me.TekstInfo(2).Enabled = False
    End If
Else
    Me.TekstInfo(6).Visible = False
    If isB2BPeppol = False Then
        Me.TekstInfo(2).Enabled = True
    End If
    Me.TekstInfo(6).TabIndex = 9
End If
Me.lPctFilter.Caption = pctFilter

If DirekteAankoop.StockBeheer.Value Then
    Labelstock(0).Visible = True
    cbEenheidsType.ListIndex = 0
    cbEenheidsType.Visible = True
    'Labelstock(1).Visible = True
    For T = 3 To 5
        TekstInfo(T).Visible = True
    Next
End If

T = 0
If GridText <> "" Then
    Do While GridText <> ""
        TekstInfo(T).Text = Left(GridText, InStr(GridText, Chr$(124)) - 1)
        GridText = Right(GridText, Len(GridText) - (Len(TekstInfo(T).Text) + 1))
        T = T + 1
    Loop
    
    TekstInfo(5).Text = Dec$(Val(TekstInfo(5).Text), "#####0.000")
    If TekstInfo(0).Text = Space$(Len(TekstInfo(0).Text)) Then
    Else
        bGet TABLE_LEDGERACCOUNTS, 0, vSet((TekstInfo(0).Text), 7)
        If Ktrl Then
            MsgBox "OnlogicaStop"
        Else
            RecordToVeld TABLE_LEDGERACCOUNTS
            TekstInfo(0).Text = vBibTekst(TABLE_LEDGERACCOUNTS, "#v019 #")
            TekstInfo(1).Text = vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #")
            If isB2BPeppol = False Then
                TekstInfo(2).TabIndex = 0
            End If
        End If
    End If
Else
    Exit Sub
End If

End Sub

Private Sub Ok_Click()
Dim HetBedrag As Currency

HetBedrag = Val(TekstInfo(2).Text)

If Val(TekstInfo(2).Text) = 0 Then
    Beep
ElseIf TekstInfo(1).Text = "" Then
    Beep
Else
    VeldRekening = TekstInfo(0).Text
    VeldNaam = TekstInfo(1).Text
    
    
        'ingave in EUR
        If bhEuro Then
            VeldBedrag = Dec$((HetBedrag), MASK_EURBH)
        Else
            'Boekhouding nog in BEF mét overgangstoggle !!
            'kan dus niet anders dan 1/ EURO zijn...
            'moet dus afgebleven worden
            VeldBedrag = Dec$((HetBedrag), MASK_EURBH)
        End If
    
        
        
    GridText = VeldRekening + Chr$(124) + VeldNaam + Chr$(124) + VeldBedrag + Chr$(124)
    If TekstInfo(3).Visible Then
        If TekstInfo(4).Text = "" Then
            Beep
        ElseIf Val(TekstInfo(5).Text) = 0 Then
            Beep
        Else
            VeldProdukt = TekstInfo(3).Text
            VeldAantal = Dec$(Val(TekstInfo(5).Text), "#####0.000")
            ProduktNaam = TekstInfo(4).Text
            GridText = GridText + VeldProdukt + Chr$(124) + ProduktNaam + Chr$(124) + VeldAantal + Chr$(124)
            Unload Me
        End If
    Else
        Unload Me
    End If
End If

End Sub

Private Sub TekstInfo_Change(Index As Integer)

Dim bedragA As Double
Dim bedragB As Double

Select Case Index
    Case 4
        Me.cbFiche.Visible = True
    Case 6
        bedragA = Val(TekstInfo(6).Text)
        bedragB = bedragA - (bedragA * pctFilter / 100)
        If isB2BPeppol = False Then
            TekstInfo(2).Text = Str$(bedragB)
        End If
        
End Select

End Sub

Private Sub TekstInfo_GotFocus(Index As Integer)

TekstInfo(Index).SelLength = Len(TekstInfo(Index).Text)
iTabIndex = Index
If Index = 0 Then Ok.Default = False: SnelHelpPrint "[Ctrl] voor geïndexeerd zoeken", BL_LOGGING
CTRLFlag = False

End Sub


Private Sub TekstInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

If KeyCode = 17 Then CTRLFlag = True
Select Case KeyCode
    Case 9, 13
        
        If Index = 0 Or Index = 5 Then
            If isB2BPeppol = False Then
                If TekstInfo(2).Enabled = True Then
                    TekstInfo(2).SetFocus
                Else
                    TekstInfo(6).SetFocus
                End If
            End If
        End If
            
        If Index = 2 Or Index = 6 Then Ok_Click
        If Index = 3 Then TekstInfo(5).SetFocus
                
            
    Case 17
        Select Case Index
            Case 0
                SharedFl = TABLE_LEDGERACCOUNTS
                aIndex = 0
                GridText = TekstInfo(0).Text
                SqlSearch.Show 1
                If Ktrl Then
                    TekstInfo(1).Text = ""
                    Ok.Enabled = False
                Else
                    RecordToVeld TABLE_LEDGERACCOUNTS
                    TekstInfo(0).Text = vBibTekst(TABLE_LEDGERACCOUNTS, "#v019 #")
                    TekstInfo(1).Text = vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #")
                    Ok.Enabled = True
                    If isB2BPeppol = False Then
                        TekstInfo(2).SetFocus
                    End If
                End If
            Case 3
                SharedFl = TABLE_PRODUCTS
                aIndex = 0
                GridText = TekstInfo(3).Text
                SqlSearch.Show 1
                If Ktrl Then
                    TekstInfo(4).Text = ""
                    Ok.Enabled = False
                Else
                    RecordToVeld TABLE_PRODUCTS
                    TekstInfo(3).Text = vBibTekst(TABLE_PRODUCTS, "#v102 #")
                    TekstInfo(4).Text = vBibTekst(TABLE_PRODUCTS, "#v105 #")
                    bGet TABLE_LEDGERACCOUNTS, 0, vSet(vBibTekst(TABLE_PRODUCTS, "#v116 #"), 7)
                    If Ktrl Then
                        TekstInfo(1).Text = ""
                        Ok.Enabled = False
                    Else
                        RecordToVeld TABLE_LEDGERACCOUNTS
                        TekstInfo(0).Text = vBibTekst(TABLE_LEDGERACCOUNTS, "#v019 #")
                        TekstInfo(1).Text = vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #")
                        Ok.Enabled = True
                    End If
                End If
        End Select
End Select

End Sub

Private Sub TekstInfo_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim RekenHier As Variant

If Index = 2 Then
    Select Case KeyCode
        Case 106, 107, 109, 111
            If Len(TekstInfo(Index)) = 1 Then
                Exit Sub
            End If
            If isB2BPeppol = False Then
                RekenHier = EenLijnRekenen(TekstInfo(2).Text)
                If RekenHier Then
                    TekstInfo(2).Text = Str(Round(RekenHier, 2))
                End If
            End If
    End Select
End If

End Sub

Private Sub TekstInfo_LostFocus(Index As Integer)

Select Case Index
    Case 0
        If CTRLFlag = True Then Exit Sub
        bGet TABLE_LEDGERACCOUNTS, 0, vSet((TekstInfo(0).Text), 7)
        If Ktrl Then
            TekstInfo(0).Text = ""
            TekstInfo(1).Text = ""
        Else
            RecordToVeld TABLE_LEDGERACCOUNTS
            TekstInfo(0).Text = vBibTekst(TABLE_LEDGERACCOUNTS, "#v019 #")
            TekstInfo(1).Text = vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #")
        End If
    Case 3
        If CTRLFlag = True Then Exit Sub
        bGet TABLE_PRODUCTS, 0, vSet((TekstInfo(3).Text), 13)
        If Ktrl Then
            KtrlBox = MsgBox("Code " & TekstInfo(3).Text & " bestaat niet." & vbCr & "Nieuw produkt aanmaken", vbQuestion + vbYesNo + vbDefaultButton2)
            If KtrlBox = vbNo Then
                'Annuleren_Click
                Exit Sub
            Else
                Unload frmProduktFiche
                Load frmProduktFiche
                frmProduktFiche.BorderStyle = 0
                frmProduktFiche.TxtInfo(0).Text = TekstInfo(3).Text
                frmProduktFiche.Show 1
                TekstInfo(3).SetFocus
            End If
        Else
            RecordToVeld TABLE_PRODUCTS
            TekstInfo(3).Text = vBibTekst(TABLE_PRODUCTS, "#v102 #")
            TekstInfo(4).Text = vBibTekst(TABLE_PRODUCTS, "#v105 #")
            bGet TABLE_LEDGERACCOUNTS, 0, vSet(vBibTekst(TABLE_PRODUCTS, "#v116 #"), 7)
            If Ktrl Then
                TekstInfo(1).Text = ""
                Ok.Enabled = False
            Else
                RecordToVeld TABLE_LEDGERACCOUNTS
                TekstInfo(0).Text = vBibTekst(TABLE_LEDGERACCOUNTS, "#v019 #")
                TekstInfo(1).Text = vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #")
                Ok.Enabled = True
            End If
        End If
    Case 5
        Dim tmpAantal As Single
        
        On Error Resume Next
        If Left(cbEenheidsType.Text, 1) = "1" Then
            tmpAantal = TekstInfo(5).Text
        ElseIf Left(cbEenheidsType.Text, 1) = "2" Then
            tmpAantal = Val(TekstInfo(5).Text) * Val(vBibTekst(TABLE_PRODUCTS, "#v107 #"))
            TekstInfo(5).Text = Trim$(Dec((tmpAantal), "####.000"))
            cbEenheidsType.ListIndex = 0
        Else
            MsgBox "stop"
        End If
        
        If isB2BPeppol = False Then
            TekstInfo(2).Text = Dec$(Round(Val(TekstInfo(5).Text) * Val(vBibTekst(TABLE_PRODUCTS, "#e113 #")), 2), MASK_EURBH)
            XDoEvents = DoEvents
            TekstInfo(2).SetFocus
        End If
End Select
CTRLFlag = False

End Sub

