VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form BasisFiche 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "BasisFiche"
   ClientHeight    =   1740
   ClientLeft      =   1005
   ClientTop       =   1125
   ClientWidth     =   4230
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
   ScaleHeight     =   1740
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.CommandButton Knop 
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
      Index           =   7
      Left            =   1920
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Volgende"
      Top             =   900
      Width           =   735
   End
   Begin VB.CommandButton Knop 
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
      Index           =   6
      Left            =   1920
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Vorige"
      Top             =   540
      Width           =   735
   End
   Begin VB.CommandButton Knop 
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
      Index           =   0
      Left            =   1140
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Eerste fiche in rij"
      Top             =   540
      Width           =   735
   End
   Begin VB.CommandButton Knop 
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
      Index           =   9
      Left            =   1140
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Laatste fiche in rij"
      Top             =   900
      Width           =   735
   End
   Begin VB.CommandButton Knop 
      Caption         =   "&Verwijderen"
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
      Index           =   10
      Left            =   -60
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Aktieve fiche verwijderen"
      Top             =   900
      Width           =   1125
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
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1260
      Width           =   2955
   End
   Begin VB.CommandButton Knop 
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
      Index           =   5
      Left            =   -60
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Aktieve fiche wijzigen en/of in detail bekijken"
      Top             =   120
      Width           =   1125
   End
   Begin VB.CommandButton Knop 
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
      Index           =   2
      Left            =   -60
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "GeSELECTeerd zoeken (ANSI-92 SQL)"
      Top             =   1260
      Width           =   1125
   End
   Begin VB.CommandButton Knop 
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
      Index           =   8
      Left            =   3000
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Venster minimaliseren"
      Top             =   900
      Width           =   1095
   End
   Begin VB.CommandButton Knop 
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
      Index           =   4
      Left            =   3000
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Dokumenten en journaaldetail"
      Top             =   540
      Width           =   1095
   End
   Begin MSMask.MaskEdBox TekstInfo 
      Height          =   315
      Index           =   0
      Left            =   1140
      TabIndex        =   1
      ToolTipText     =   "Breng hier fiche UNIEKE ID-kode in !"
      Top             =   120
      Width           =   2955
      _ExtentX        =   5212
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
Attribute VB_Name = "BasisFiche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit

Dim T                                   As Integer
Dim Voorkeurquick(10, 10)               As String * 5


Private Sub FicheNaarRecord(Fl As Integer)

bGet Fl, 0, vSet(TekstInfo(0).text, FLINDEX_LEN(Fl, 0))
If Ktrl = 0 Then
    bUpdate Fl, 0
Else
    bInsert Fl, 0
End If
Knop_Click 3

End Sub


Private Sub cmbSortering_GotFocus()

    Knop(2).Default = True

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
End If

End Sub

Private Sub Knop_Click(Index As Integer)
Dim TeZoeken As String

Fl = Val(Me.Tag)
Mim.InfoData.Visible = False
Select Case Index
        Case 0  'Eerste record
            bFirst Fl, 0
            If Ktrl Then
                Beep
                Knop(5).Enabled = False
            Else
                INSERT_FLAG(Fl) = 0
                RecordNaarFiche Fl
                Knop(5).Enabled = True
            End If
                
        Case 2  'GeSELECTeerd Zoeken
        Select Case Fl
            Case TABLE_CONTRACTS
                Venster.Show 1
            Case Else
                SharedFl = Fl
                aIndex = Val(Left(cmbSortering.text, 2))
                GridText = TekstInfo(0).text
                SqlSearch.Show 1
        End Select
        If Ktrl = 0 Then
            TekstInfo(0).text = vBibTekst(Fl, "#" + JETTABLEUSE_INDEX(Fl, 0) + "#")
            INSERT_FLAG(Fl) = 0
            RecordNaarFiche Fl
            Knop(5).Enabled = True
        Else
            Knop(5).Enabled = False
            TekstInfo(0).text = ""
            INSERT_FLAG(Fl) = 1
        End If
        
        Case 3  'Nieuwe Fiche
            'NieuweFiche Fl
            
        Case 4  'Boekhouding
            Select Case Fl
                Case TABLE_CUSTOMERS To TABLE_SUPPLIERS
                    'BalansKontrole Fl
                    'Stop
                    BalansKontroleWithRecordSet Fl
                    
                Case TABLE_LEDGERACCOUNTS
                    HistoriekSQL.Show
                Case Else
                    MsgBox "boekhoudkontrole (nog) niet aanwezig"
            End Select
            INSERT_FLAG(Fl) = 0
            RecordNaarFiche Fl
            Knop(5).Enabled = True
        
        Case 5  'Fiche EDITEREN
            If PeppolFlag Then
                vBib TABLE_SUPPLIERS, Me.TekstInfo(0).text, "A110"   'Uniek Codenummer
            Else
                TeZoeken = Trim$(TekstInfo(0))
                If TeZoeken = "" Then Beep: Exit Sub
                    bGet Fl, 0, TekstInfo(0)
                    If Ktrl = 0 Then
                        INSERT_FLAG(Fl) = 0
                        RecordNaarFiche Fl
                        Knop(5).Enabled = True
                    Else
                        NieuweFiche Fl
                        TekstInfo(0) = TeZoeken
                    End If
                    If Fl = TABLE_LEDGERACCOUNTS Then DbKontrole (TekstInfo(0).text), TABLE_LEDGERACCOUNTS
            
                    If INSERT_FLAG(Fl) = 1 Then
                    Select Case Fl
                        Case TABLE_CUSTOMERS, TABLE_SUPPLIERS
                            vBib Fl, (TekstInfo(0).text), "A110"   'Klant/Levnummer
                        Case TABLE_LEDGERACCOUNTS
                            vBib Fl, (TekstInfo(0).text), "v019"   'Rekeningnummer
                        Case Else
                            MsgBox "Stop"
                    End Select
                End If
            End If
            
            If TeleBibClick(Fl) = False Then
                'Knop(1).Enabled = False
            Else
                'Knop(1).Enabled = True
                'Knop(1).SetFocus
                FicheNaarRecord Fl
            End If
            Mim.InfoData.Visible = False
            'Knop_Click 3
            INSERT_FLAG(Fl) = 0
            RecordNaarFiche Fl
            Knop(5).Enabled = True
                            
        Case 6  'Fiche Lager
            bPrev Fl
            If Ktrl = 9 Then
                bFirst Fl, 0
                If Ktrl Then
                    Beep
                    Knop(5).Enabled = False
                End If
            End If
            If Ktrl Then
            Else
                INSERT_FLAG(Fl) = 0
                RecordNaarFiche Fl
                Knop(5).Enabled = True
            End If
      
        Case 7  'fiche hoger
            bNext Fl
            If Ktrl = 9 Then
                bLast Fl, 0
                If Ktrl Then
                    Beep
                    Knop(5).Enabled = False
                End If
            End If
            If Ktrl Then
            Else
                INSERT_FLAG(Fl) = 0
                RecordNaarFiche Fl
                Knop(5).Enabled = True
            End If

        Case 8 'verbergen
            WindowState = 1

        Case 9
            bLast Fl, 0
            If Ktrl Then
                Beep
                Knop(5).Enabled = False
            Else
                INSERT_FLAG(Fl) = 0
                RecordNaarFiche Fl
                Knop(5).Enabled = True
            End If
        Case 10
            If INSERT_FLAG(Fl) = 0 Then
                msg = "Bestaande fiche " + Caption + " verwijderen.  Bent U zeker ?"
                KtrlBox = MsgBox(msg, vbYesNo + vbQuestion + vbDefaultButton2, TekstInfo(0).text)
                If KtrlBox = vbYes Then
                    bDelete Fl
                    Knop_Click 3
                End If
            End If
        
        Case Else
End Select

End Sub


Private Sub NieuweFiche(Fl As Integer)

Dim xTMP As Integer

If PeppolFlag Then
Else
    xTMP = daoBlankoRecord(Fl)

    TekstInfo(0).text = ""
    INSERT_FLAG(Fl) = 1
    Knop(5).Enabled = False
    Knop(5).Default = True
    TekstInfo(0).Enabled = True

    On Error Resume Next
    TekstInfo(0).SetFocus
End If

End Sub

Private Sub RecordNaarFiche(Fl)

On Error Resume Next

TLB_RECORD(Fl) = ""
If Ktrl Then
    'MsgBox "stop"
Else
    RecordToVeld Fl
End If
TekstInfo(0).text = vBibTekst(Fl, "#" + JETTABLEUSE_INDEX(Fl, 0) + "#")
XDoEvents = DoEvents

msg = ""
For T = 0 To 1
    msg = msg + RTrim$(FVT(Fl, T)) + " "
Next
SnelHelpPrint msg, BL_LOGGING
INSERT_FLAG(Fl) = 0

End Sub

Private Sub TekstInfo_Change(Index As Integer)

If TekstInfo(0).text = "" Then
    INSERT_FLAG(Fl) = 1
    Knop(5).Enabled = False
Else
    Knop(5).Enabled = True
End If

End Sub

Private Sub TekstInfo_GotFocus(Index As Integer)

SnelHelpPrint TekstInfo(0).ToolTipText & ", " & "[Ctrl] voor geSELECTeerd zoeken !", BL_LOGGING

Knop(5).Default = True
TekstInfo(0).SelLength = Len(TekstInfo(0).text)

End Sub

Private Sub TekstInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case 17
        aIndex = Val(Left(cmbSortering.text, 2))
        SharedFl = Fl
        GridText = TekstInfo(0).text
        SqlSearch.Show 1
        If Ktrl = 0 Then
            INSERT_FLAG(Fl) = 0
            RecordNaarFiche Fl
            Knop(5).Enabled = True
        End If
End Select

End Sub

