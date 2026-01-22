VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form BasisFiche 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BasisFiche"
   ClientHeight    =   1260
   ClientLeft      =   1005
   ClientTop       =   1170
   ClientWidth     =   5685
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
   Icon            =   "00100mk.frx":0000
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1260
   ScaleWidth      =   5685
   WindowState     =   1  'Minimized
   Begin VB.CheckBox cmdGrafischVoorstellen 
      Alignment       =   1  'Right Justify
      Caption         =   "&Grafische voorstelling"
      Height          =   255
      Left            =   3120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Toggle voor knop VERWIJDEREN"
      Top             =   240
      Width           =   2445
   End
   Begin VB.CheckBox VerwijderenKan 
      Alignment       =   1  'Right Justify
      Caption         =   "Verwijderen &mogelijk"
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Toggle voor knop VERWIJDEREN"
      Top             =   30
      Width           =   2445
   End
   Begin VB.CommandButton Knop 
      Caption         =   "&Verwijderen"
      Enabled         =   0   'False
      Height          =   315
      Index           =   10
      Left            =   3330
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Aktieve fiche verwijderen"
      Top             =   900
      Width           =   1125
   End
   Begin VB.CommandButton Knop 
      Caption         =   "B&odem"
      Height          =   315
      Index           =   9
      Left            =   30
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Laatste fiche in rij"
      Top             =   900
      Width           =   825
   End
   Begin VB.CommandButton Knop 
      Caption         =   "&Top"
      Height          =   315
      Index           =   0
      Left            =   30
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Eerste fiche in rij"
      Top             =   540
      Width           =   825
   End
   Begin VB.CommandButton Knop 
      Caption         =   "&Editeren"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   315
      HelpContextID   =   30040
      Index           =   5
      Left            =   3330
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Aktieve fiche wijzigen en/of in detail bekijken"
      Top             =   540
      Width           =   1125
   End
   Begin VB.CommandButton Knop 
      Caption         =   "&Zoeken"
      Height          =   315
      HelpContextID   =   30030
      Index           =   2
      Left            =   2430
      TabIndex        =   1
      ToolTipText     =   "Geïndexeerd zoeken"
      Top             =   540
      Width           =   855
   End
   Begin VB.CommandButton Knop 
      Caption         =   "&Lager"
      Height          =   315
      Index           =   6
      Left            =   900
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Vorige"
      Top             =   540
      Width           =   705
   End
   Begin VB.CommandButton Knop 
      Caption         =   "&Hoger"
      Height          =   315
      Index           =   7
      Left            =   1650
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Volgende"
      Top             =   540
      Width           =   735
   End
   Begin VB.CommandButton Knop 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "&Pictogram"
      Height          =   300
      Index           =   8
      Left            =   4500
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Venster minimaliseren"
      Top             =   540
      Width           =   1125
   End
   Begin VB.CommandButton Knop 
      Caption         =   "&Nieuw"
      Height          =   315
      HelpContextID   =   30010
      Index           =   3
      Left            =   2430
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Nieuwe fiche maken of via kode direkt activeren"
      Top             =   900
      Width           =   855
   End
   Begin VB.CommandButton Knop 
      Caption         =   "&Relaties"
      Height          =   315
      HelpContextID   =   30060
      Index           =   4
      Left            =   900
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Dokumenten en journaaldetail"
      Top             =   900
      Width           =   1485
   End
   Begin MSMask.MaskEdBox TekstInfo 
      Height          =   420
      Index           =   0
      Left            =   30
      TabIndex        =   7
      ToolTipText     =   "Breng hier fiche ID-kode in na eerst de knop NIEUW te hebben gedrukt !"
      Top             =   30
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   741
      _Version        =   393216
      ClipMode        =   1
      BackColor       =   16777215
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
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

Dim FotoBestand                         As String
Dim VideoBestand                        As String


Private Sub FicheNaarRecord(Fl As Integer)

bGet Fl, 0, vSet(TekstInfo(0).Text, FlIndexLen(Fl, 0))
If Ktrl = 0 Then
    bUpdate Fl, 0
Else
    bInsert Fl, 0
End If
Knop_Click 3

End Sub

Private Sub cmdGrafischVoorstellen_Click()

SnelHelpPrint "Optie in ontwikkelfase.  Laatste bijwerking dd. 15/01/1999"
If BasisB(Fl).cmdGrafischVoorstellen.Value = False Then
    BasisB(Fl).Height = 1635
Else
    BasisB(Fl).Height = 4700
End If

End Sub

Private Sub Form_Activate()

On Local Error Resume Next

Fl = Val(Me.Tag)
If BedrijfsLokatie = "" Then
    Exit Sub
Else
    NieuweFiche Fl
End If

End Sub

Private Sub Knop_Click(Index As Integer)

Mim.InfoData.Visible = False
Select Case Index
        Case 0  'Eerste record
            bFirst Fl, 0
            If Ktrl Then
                Beep
                Knop(5).Enabled = False
            Else
                InsertFlag(Fl) = 0
                RecordNaarFiche Fl
                Knop(5).Enabled = True
            End If

        'Case 1  'Nieuw of Gewijzigd wegschrijven
        '    FicheNaarRecord Fl
            
                
        Case 2  'Geïndexeerd Zoeken
        Select Case Fl
            Case FlPolis
                Venster.Show 1
            Case Else
                SharedFl = Fl
                aIndex = 1
                GridText = TekstInfo(0).Text
                AlfaSearch.Show 1
        End Select
        If Ktrl = 0 Then
            TekstInfo(0).Text = vBibTekst(Fl, "#" + FlIndexIs(Fl, 0) + "#")
            InsertFlag(Fl) = 0
            RecordNaarFiche Fl
            Knop(5).Enabled = True
        Else
            Knop(5).Enabled = False
            TekstInfo(0).Text = ""
            InsertFlag(Fl) = 1
        End If
        
        Case 3  'Nieuwe Fiche
            NieuweFiche Fl
            Knop(5).Default = True
            Knop(5).Enabled = False
            TekstInfo(0).SetFocus
        
        Case 4  'Boekhouding
            Select Case Fl
                Case FlKlant To FlLeverancier
                    BalansKontrole Fl
                Case FlRekening
                    Historiek.Show
                Case Else
                    MsgBox "boekhoudkontrole (nog) niet aanwezig"
            End Select
            InsertFlag(Fl) = 0
            RecordNaarFiche Fl
            Knop(5).Enabled = True
        
        Case 5  'Fiche EDITEREN
            bGet Fl, 0, vSet((TekstInfo(0).Text), FlIndexLen(Fl, 0))
            If Ktrl = 0 Then
                InsertFlag(Fl) = 0
                RecordNaarFiche Fl
                Knop(5).Enabled = True
            Else
                'Knop(1).Enabled = False
            End If
            
            If Fl = FlRekening Then DbKontrole (TekstInfo(0).Text), FlRekening
            
            If InsertFlag(Fl) = 1 Then
                Select Case Fl
                    Case FlKlant, FlLeverancier
                        vBib Fl, (TekstInfo(0).Text), "A110"   'Klant/Levnummer
                    Case FlRekening
                        vBib Fl, (TekstInfo(0).Text), "v019"   'Rekeningnummer
                    Case Else
                        MsgBox "Stop"
                End Select
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
            InsertFlag(Fl) = 0
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
                InsertFlag(Fl) = 0
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
                InsertFlag(Fl) = 0
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
                InsertFlag(Fl) = 0
                RecordNaarFiche Fl
                Knop(5).Enabled = True
            End If
        Case 10
            If InsertFlag(Fl) = 0 Then
                Msg = "Bestaande fiche " + Caption + " verwijderen.  Bent U zeker ?"
                KtrlBox = MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton2, TekstInfo(0).Text)
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

TekstInfo(0).Text = ""
TLBRecord(Fl) = ""

Select Case Fl
    Case FlKlant, FlLeverancier
        vBib Fl, "2", "A10C"    'Taalkode
        vBib Fl, "002", "v149"  'Landnummer  ISO kode
        vBib Fl, "B  ", "A109"  'Landkode Postkantoor
        vBib Fl, "BE", "v150"   'Landkode    ISO kode
        vBib Fl, "BEF", "vs03"  'Munteenheid ISO kode
        vBib Fl, "1", "vs07"    'exemplaren dokumenten
    Case FlRekening
        vBib Fl, "O", "v032"    'Budgetcode
End Select
InsertFlag(Fl) = 1
'Knop(1).Enabled = False     'wegschrijven uitschakelen
TekstInfo(0).Enabled = True

End Sub

Private Sub RecordNaarFiche(Fl)

On Error Resume Next

TLBRecord(Fl) = ""
If Ktrl Then
    'MsgBox "stop"
Else
    RecordToVeld Fl
End If
TekstInfo(0).Text = vBibTekst(Fl, "#" + FlIndexIs(Fl, 0) + "#")
XDoEvents = DoEvents
'Msg = ""
'T = 0
'Do While Voorkeurquick(Fl, T) <> Space$(5)
'    Msg = Msg + RTrim$(vBibTekst(Fl, "#" + Voorkeurquick(Fl, T) + "#")) + ", "
'    T = T + 1
'Loop
'SnelHelpPrint Msg

InsertFlag(Fl) = 0
End Sub

Private Sub TekstInfo_Change(Index As Integer)

If TekstInfo(0).Text = "" Then
    InsertFlag(Fl) = 1
    Knop(5).Enabled = False
Else
    Knop(5).Enabled = True
End If

End Sub

Private Sub TekstInfo_GotFocus(Index As Integer)

BasisB(Fl).WindowState = 0
TekstInfo(0).SelLength = Len(TekstInfo(0).Text)
SnelHelpPrint "[Ctrl] + minstens 1 teken voor geïndexeerd zoeken !"

End Sub

Private Sub TekstInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case 17
        If Len(TekstInfo(0).Text) < 1 Then
            SnelHelpPrint "Minstens 1 tekens voor geïndexeerd zoeken a.u.b. !"
            Beep
            Exit Sub
        End If
        aIndex = 0
        SharedFl = Fl
        GridText = TekstInfo(0).Text
        AlfaSearch.Show 1
        If Ktrl = 0 Then
            InsertFlag(Fl) = 0
            RecordNaarFiche Fl
            Knop(5).Enabled = True
        End If
End Select

End Sub



Private Sub VerwijderenKan_Click()

If VerwijderenKan.Value = 0 Then
    Knop(10).Enabled = False
Else
    Knop(10).Enabled = True
End If

End Sub


