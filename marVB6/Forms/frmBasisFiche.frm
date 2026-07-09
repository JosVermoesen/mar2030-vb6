VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form BasicTable 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "BasisFiche"
   ClientHeight    =   2730
   ClientLeft      =   1005
   ClientTop       =   1125
   ClientWidth     =   4770
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
   Icon            =   "frmBasisFiche.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2730
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.CheckBox CheckBoxSorting 
      Alignment       =   1  'Right Justify
      Caption         =   "Dalend Sorteren"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1320
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.ComboBox ComboBoxTemplate 
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
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2280
      Width           =   4395
   End
   Begin VB.CommandButton ButtonNext 
      Caption         =   "&Hoger"
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
      Left            =   960
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Volgende"
      Top             =   1380
      Width           =   975
   End
   Begin VB.CommandButton ButtonPrevious 
      Caption         =   "&Lager"
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
      Left            =   960
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Vorige"
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton ButtonFirst 
      Caption         =   "&Top"
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
      Left            =   180
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Eerste fiche in rij"
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton ButtonLast 
      Caption         =   "B&odem"
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
      Left            =   180
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Laatste fiche in rij"
      Top             =   1380
      Width           =   735
   End
   Begin VB.CommandButton ButtonRemove 
      Caption         =   "Verwijderen"
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
      Left            =   180
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Aktieve fiche verwijderen"
      Top             =   1800
      Width           =   1725
   End
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
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   540
      Width           =   3195
   End
   Begin VB.CommandButton ButtonEdit 
      Caption         =   "B&ewerken"
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
      HelpContextID   =   30040
      Left            =   180
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Aktieve fiche wijzigen en/of in detail bekijken"
      Top             =   120
      Width           =   1125
   End
   Begin VB.CommandButton ButtonSearchOn 
      Caption         =   "&Zoeken op..."
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
      HelpContextID   =   30030
      Left            =   180
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "GeSELECTeerd zoeken (ANSI-92 SQL)"
      Top             =   540
      Width           =   1125
   End
   Begin VB.CommandButton ButtonMinimize 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "&Minimaliseren"
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
      Left            =   2040
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Venster minimaliseren"
      Top             =   1800
      Width           =   2535
   End
   Begin VB.CommandButton ButtonJournal 
      Caption         =   "&Relaties"
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
      HelpContextID   =   30060
      Left            =   2040
      TabIndex        =   5
      ToolTipText     =   "Dokumenten en journaaldetail"
      Top             =   960
      Width           =   2535
   End
   Begin MSMask.MaskEdBox TekstInfo 
      Height          =   315
      Left            =   1380
      TabIndex        =   1
      ToolTipText     =   "Breng hier fiche UNIEKE ID-kode in !"
      Top             =   120
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      BackColor       =   16777215
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
End
Attribute VB_Name = "BasicTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit

Dim T                                   As Integer
Dim Voorkeurquick(10, 10)               As String * 5

Private Sub FicheNaarRecord(Fl As Integer)

    bGet Fl, 0, vSet(TekstInfo.text, FLINDEX_LEN(Fl, 0))
    If Ktrl = 0 Then
        bUpdate Fl, 0
    Else
        bInsert Fl, 0
    End If
    'Knop_Click 3 was for new Record

End Sub


Private Sub ButtonEdit_Click()

    Dim TeZoeken As String

    Fl = Val(Me.Tag)
    Mim.InfoData.Visible = False

    'EDIT
    If PeppolFlag And Fl = TABLE_SUPPLIERS Then
        vBib TABLE_SUPPLIERS, TekstInfo.text, "A110"   'Uniek Codenummer
    Else
        TeZoeken = Trim$(TekstInfo)
        If TeZoeken = "" Then Beep: Exit Sub
        bGet Fl, 0, TekstInfo
        If Ktrl = 0 Then
            INSERT_FLAG(Fl) = 0
            RecordNaarFiche Fl
            ButtonEdit.Enabled = True
        Else
            NieuweFiche Fl
            TekstInfo = TeZoeken
        End If
        If Fl = TABLE_LEDGERACCOUNTS Then DbKontrole (TekstInfo.text), TABLE_LEDGERACCOUNTS
            
        If INSERT_FLAG(Fl) = 1 Then
            Select Case Fl
                Case TABLE_CUSTOMERS, TABLE_SUPPLIERS
                    vBib Fl, (TekstInfo.text), "A110"   'Klant/Levnummer
                Case TABLE_LEDGERACCOUNTS
                    vBib Fl, (TekstInfo.text), "v019"   'Rekeningnummer
                Case Else
                    MsgBox "Stop"
            End Select
        End If
    End If
            
    If TeleBibClick(Fl) = False Then
    Else
        FicheNaarRecord Fl
    End If
    Mim.InfoData.Visible = False
    INSERT_FLAG(Fl) = 0
    RecordNaarFiche Fl
    ButtonEdit.Enabled = True
        
End Sub

Private Sub ButtonFirst_Click()

    Dim TeZoeken As String

    Fl = Val(Me.Tag)
    Mim.InfoData.Visible = False
        
    'Move First
    bFirst Fl, 0
    If Ktrl Then
        Beep
        ButtonEdit.Enabled = False
    Else
        INSERT_FLAG(Fl) = 0
        RecordNaarFiche Fl
        ButtonEdit.Enabled = True
    End If
    
End Sub

Private Sub ButtonJournal_Click()

    Dim TeZoeken As String

    Fl = Val(Me.Tag)
    Mim.InfoData.Visible = False
        
    'Ledger or balance buyer and seller
    Select Case Fl
        Case TABLE_CUSTOMERS To TABLE_SUPPLIERS
            'BalansKontrole Fl
            BalansKontroleWithRecordSet Fl
                    
        Case TABLE_LEDGERACCOUNTS
            HistoriekSQL.Show
        Case Else
            MsgBox "boekhoudkontrole (nog) niet aanwezig"
    End Select
    INSERT_FLAG(Fl) = 0
    RecordNaarFiche Fl
    ButtonEdit.Enabled = True
        
End Sub

Private Sub ButtonLast_Click()

    Dim TeZoeken As String

    Fl = Val(Me.Tag)
    Mim.InfoData.Visible = False

    'Move Last
    bLast Fl, 0
    If Ktrl Then
        Beep
        ButtonEdit.Enabled = False
    Else
        INSERT_FLAG(Fl) = 0
        RecordNaarFiche Fl
        ButtonEdit.Enabled = True
    End If
    
End Sub

Private Sub ButtonMinimize_Click()

    WindowState = 1

End Sub

Private Sub ButtonNext_Click()

    Dim TeZoeken As String

    Fl = Val(Me.Tag)
    Mim.InfoData.Visible = False
    
    'Move Next
    bNext Fl
    If Ktrl = 9 Then
        bLast Fl, 0
        If Ktrl Then
            Beep
            ButtonEdit.Enabled = False
        End If
    End If
    If Ktrl Then
    Else
        INSERT_FLAG(Fl) = 0
        RecordNaarFiche Fl
        ButtonEdit.Enabled = True
    End If

End Sub

Private Sub ButtonPrevious_Click()

    Dim TeZoeken As String

    Fl = Val(Me.Tag)
    Mim.InfoData.Visible = False
        
    'Move Previous
    bPrev Fl
    If Ktrl = 9 Then
        bFirst Fl, 0
        If Ktrl Then
            Beep
            ButtonEdit.Enabled = False
        End If
    End If
    If Ktrl Then
    Else
        INSERT_FLAG(Fl) = 0
        RecordNaarFiche Fl
        ButtonEdit.Enabled = True
    End If
    
End Sub

Private Sub ButtonRemove_Click()

    If INSERT_FLAG(Fl) = 0 Then
        Msg = "Bestaande fiche " + Caption + " verwijderen.  Bent U zeker ?"
        KtrlBox = MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton2, TekstInfo.text)
        If KtrlBox = vbYes Then
            bDelete Fl
            'Knop_Click 3 was for new Record
        End If
    End If
    
End Sub

Private Sub ButtonSearchOn_Click()
    Dim TeZoeken As String

    Fl = Val(Me.Tag)
    Select Case Fl
        Case TABLE_CONTRACTS
            Venster.Show 1
        Case Else
            SharedFl = Fl
            aIndex = Val(Left(cmbSortering.text, 2))
            GridText = TekstInfo.text
            SqlSearch.Show 1
    End Select
    If Ktrl = 0 Then
        TekstInfo.text = vBibTekst(Fl, "#" + JETTABLEUSE_INDEX(Fl, 0) + "#")
        INSERT_FLAG(Fl) = 0
        RecordNaarFiche Fl
        ButtonEdit.Enabled = True
    Else
        ButtonEdit.Enabled = False
        TekstInfo.text = ""
        INSERT_FLAG(Fl) = 1
    End If
        
End Sub



Private Sub cmbSortering_GotFocus()

    ButtonSearchOn.Default = True

End Sub



Private Sub NieuweFiche(Fl As Integer)

    Dim xTMP As Integer

    If PeppolFlag Then
    Else
        xTMP = daoBlankoRecord(Fl)

        TekstInfo.text = ""
        INSERT_FLAG(Fl) = 1
        ButtonEdit.Enabled = False
        ButtonEdit.Default = True
        TekstInfo.Enabled = True

        On Error Resume Next
        TekstInfo.SetFocus
    End If

End Sub

Private Sub RecordNaarFiche(Fl)

    On Error Resume Next

    TLB_RECORD(Fl) = ""
    If Ktrl Then
    Else
        RecordToVeld Fl
    End If
    TekstInfo.text = vBibTekst(Fl, "#" + JETTABLEUSE_INDEX(Fl, 0) + "#")
    XDoEvents = DoEvents

    Msg = ""
    For T = 0 To 1
        Msg = Msg + RTrim$(FVT(Fl, T)) + " "
    Next
    SnelHelpPrint Msg, BL_LOGGING
    INSERT_FLAG(Fl) = 0

End Sub


Private Sub Form_Activate()

On Local Error Resume Next
    Fl = Val(Me.Tag)

    If LOCATION_COMPANYDATA = "" Then
        Exit Sub
    Else
        If cmbSortering.ListCount Then
        Else
            For T = 0 To FL_NUMBEROFINDEXEN(Fl)
                cmbSortering.AddItem Format(T, "00") + ":" + FLINDEX_CAPTION(Fl, T) + " (" + Trim$(JETTABLEUSE_INDEX(Fl, T)) + ")"
            Next
            Err = 0
            On Error Resume Next
            cmbSortering = LaadTekst(Me.Name, "ProduktSortering")
            If Err Then cmbSortering.ListIndex = 1
        End If
        NieuweFiche Fl
        If ComboBoxTemplate.ListCount Then
        Else
            Select Case Fl
                Case 1
                    ComboBoxTemplate.AddItem "001.DEF: Alle velden klantfiche"
                
                Case 2
                    ComboBoxTemplate.AddItem "002.DEF: Alle velden leveranciersfiche"
                
                Case 3
                    ComboBoxTemplate.AddItem "003.DEF: Alle velden rekeningfiche"
            End Select
            ComboBoxTemplate.ListIndex = 0
            ComboBoxTemplate.Enabled = False
        End If
    End If

End Sub

Private Sub TekstInfo_Change()

    If TekstInfo.text = "" Then
        INSERT_FLAG(Fl) = 1
        ButtonEdit.Enabled = False
    Else
        ButtonEdit.Enabled = True
    End If

End Sub

Private Sub TekstInfo_GotFocus()

    SnelHelpPrint TekstInfo.ToolTipText & ", " & "[Ctrl] voor geSELECTeerd zoeken !", BL_LOGGING

    ButtonEdit.Default = True
    TekstInfo.SelLength = Len(TekstInfo.text)

End Sub

Private Sub TekstInfo_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case 17
        aIndex = Val(Left(cmbSortering.text, 2))
        SharedFl = Fl
        GridText = TekstInfo.text
        SqlSearch.Show 1
        If Ktrl = 0 Then
            INSERT_FLAG(Fl) = 0
            RecordNaarFiche Fl
            ButtonEdit.Enabled = True
        End If
End Select

End Sub

