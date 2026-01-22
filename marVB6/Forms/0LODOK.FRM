VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form LayOutDokument 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AfdrukTester Briefwisseling"
   ClientHeight    =   1470
   ClientLeft      =   210
   ClientTop       =   645
   ClientWidth     =   9240
   ForeColor       =   &H80000008&
   HelpContextID   =   23000
   Icon            =   "0LODOK.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   98
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   616
   Begin VB.ComboBox cmbFiguurnamen 
      Height          =   315
      Left            =   6960
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   960
      Width           =   2235
   End
   Begin VB.CommandButton cmdFiguurEnter 
      Caption         =   "Opnemen"
      Default         =   -1  'True
      Height          =   375
      Left            =   8100
      TabIndex        =   32
      Top             =   540
      Width           =   1035
   End
   Begin VB.TextBox txtFiguurY 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6960
      TabIndex        =   30
      Top             =   660
      Width           =   855
   End
   Begin VB.TextBox FiguurIndex 
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
      Left            =   6960
      TabIndex        =   26
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox txtFiguurX 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6960
      TabIndex        =   24
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox LijnBeginX 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4140
      TabIndex        =   22
      Top             =   300
      Width           =   795
   End
   Begin VB.TextBox LijnBeginY 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5340
      TabIndex        =   21
      Top             =   300
      Width           =   795
   End
   Begin VB.TextBox LijnTotX 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4140
      TabIndex        =   20
      Top             =   600
      Width           =   795
   End
   Begin VB.TextBox LijnTotY 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5340
      TabIndex        =   19
      Top             =   600
      Width           =   795
   End
   Begin VB.TextBox VsoftGebied 
      Height          =   315
      Index           =   0
      Left            =   60
      TabIndex        =   18
      ToolTipText     =   "Van (in twips) bvb. 4000"
      Top             =   900
      Width           =   585
   End
   Begin VB.TextBox VsoftGebied 
      Height          =   315
      Index           =   1
      Left            =   660
      TabIndex        =   17
      ToolTipText     =   "Tot (in twips) bvb. 14000"
      Top             =   900
      Width           =   645
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   315
      Left            =   3660
      TabIndex        =   15
      Top             =   0
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      BuddyControl    =   "Lokatie"
      BuddyDispid     =   196628
      OrigLeft        =   76
      OrigTop         =   42
      OrigRight       =   92
      OrigBottom      =   63
      Max             =   99
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton BoxEnter 
      Caption         =   "Opnemen"
      Height          =   375
      Left            =   4140
      TabIndex        =   14
      Top             =   960
      Width           =   1035
   End
   Begin VB.CheckBox BoxMaken 
      Alignment       =   1  'Right Justify
      Caption         =   "Box &maken"
      Height          =   375
      Left            =   5220
      TabIndex        =   9
      Top             =   960
      Width           =   915
   End
   Begin VB.CommandButton AutoOpBouw 
      Caption         =   "&AutoOpbouw"
      Height          =   345
      Left            =   2760
      TabIndex        =   13
      Top             =   1020
      Width           =   1185
   End
   Begin VB.CommandButton Command2 
      Caption         =   "A&lles schoon"
      Height          =   345
      Left            =   2760
      TabIndex        =   12
      Top             =   660
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Schoon&scherm"
      Height          =   345
      Left            =   1500
      TabIndex        =   11
      Top             =   660
      Width           =   1185
   End
   Begin VB.CommandButton EnterReturn 
      Caption         =   "Lijn O&pnemen"
      Height          =   345
      Left            =   1500
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1020
      Width           =   1185
   End
   Begin VB.TextBox LijnIndex 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5400
      TabIndex        =   8
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox Lokatie 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3240
      TabIndex        =   6
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   510
      TabIndex        =   3
      Top             =   0
      Width           =   675
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1200
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   300
      Width           =   2715
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   510
      TabIndex        =   5
      Top             =   300
      Width           =   675
   End
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   315
      Left            =   5940
      TabIndex        =   16
      Top             =   0
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      BuddyControl    =   "LijnIndex"
      BuddyDispid     =   196627
      OrigLeft        =   76
      OrigTop         =   42
      OrigRight       =   92
      OrigBottom      =   63
      Max             =   9
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDown3 
      Height          =   315
      Left            =   7440
      TabIndex        =   27
      Top             =   0
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      BuddyControl    =   "FiguurIndex"
      BuddyDispid     =   196613
      OrigLeft        =   76
      OrigTop         =   42
      OrigRight       =   92
      OrigBottom      =   63
      Max             =   99
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.Image imgFiguur 
      Height          =   495
      Left            =   4020
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Naam"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   6360
      TabIndex        =   31
      Top             =   1020
      Width           =   525
   End
   Begin VB.Label Label1 
      Caption         =   "&yPos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   6360
      TabIndex        =   29
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "&xPos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   6360
      TabIndex        =   28
      Top             =   420
      Width           =   525
   End
   Begin VB.Label Label1 
      Caption         =   "&Figuur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   6300
      TabIndex        =   25
      Top             =   60
      Width           =   525
   End
   Begin VB.Label Label1 
      Caption         =   "&Kader"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   4140
      TabIndex        =   23
      Top             =   60
      Width           =   525
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      Index           =   1
      X1              =   416
      X2              =   416
      Y1              =   8
      Y2              =   86
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      Index           =   0
      X1              =   268
      X2              =   268
      Y1              =   8
      Y2              =   86
   End
   Begin VB.Label Label1 
      Caption         =   "&Vsoft Gebied"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   660
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "&Tekstlijn"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   1320
      TabIndex        =   0
      Top             =   60
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "&yPos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   4
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "&xPos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   525
   End
   Begin VB.Menu Bestanden 
      Caption         =   "&Bestand"
      Begin VB.Menu Bestand 
         Caption         =   "&Openen"
         Index           =   1
      End
      Begin VB.Menu Bestand 
         Caption         =   "O&pslaan"
         Index           =   2
      End
      Begin VB.Menu Bestand 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu Bestand 
         Caption         =   "Af&drukken"
         Index           =   4
      End
      Begin VB.Menu Bestand 
         Caption         =   "&Naar Afbeeldingen"
         Index           =   5
      End
      Begin VB.Menu Bestand 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu Bestand 
         Caption         =   "&Afsluiten"
         Index           =   7
      End
      Begin VB.Menu Bestand 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu Bestand 
         Caption         =   "Via kladblok"
         Index           =   9
         Begin VB.Menu ViaKB 
            Caption         =   "Tekst- en Lijnobjecten"
            Index           =   0
         End
         Begin VB.Menu ViaKB 
            Caption         =   "Grafische Objecten"
            Index           =   1
         End
      End
   End
   Begin VB.Menu opmaak 
      Caption         =   "&Opmaak"
      Begin VB.Menu Teken 
         Caption         =   "&Teken"
      End
   End
   Begin VB.Menu TaalKeuze 
      Caption         =   "&Taal"
      Begin VB.Menu mTaal 
         Caption         =   "&Frans"
         Index           =   1
      End
      Begin VB.Menu mTaal 
         Caption         =   "&Nederlands"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu mTaal 
         Caption         =   "&Engels"
         Index           =   3
      End
      Begin VB.Menu mTaal 
         Caption         =   "&Duits"
         Index           =   4
      End
   End
   Begin VB.Menu Document 
      Caption         =   "Document"
      Begin VB.Menu DokType 
         Caption         =   "Factuur/Creditnota"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu DokType 
         Caption         =   "LeveringsBon"
         Index           =   1
      End
      Begin VB.Menu DokType 
         Caption         =   "BestelBon"
         Index           =   2
      End
      Begin VB.Menu DokType 
         Caption         =   "Offerte"
         Index           =   3
      End
      Begin VB.Menu DokType 
         Caption         =   "Briefwisseling"
         Index           =   4
      End
      Begin VB.Menu DokType 
         Caption         =   "Rekeninguitttreksel"
         Index           =   5
      End
   End
   Begin VB.Menu Figuren 
      Caption         =   "Figuren"
      Begin VB.Menu FigurenDef 
         Caption         =   "Definiëren"
      End
   End
End
Attribute VB_Name = "LayOutDokument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim KopVoet As Integer
Dim TaalKode As Integer
Dim DocuType As Integer
Dim Diepgang As Integer

Dim psTekst(50) As String              'TekstString zelf
Dim psX(50) As Single                 'Xpositie
Dim psY(50) As Single                 'Ypositie
Dim psFontSize(50) As Single
Dim psFontName(50) As String
Dim psFontBold(50)         As Integer
Dim psFontItalic(50) As Integer
Dim psFontUnderLine(50) As Integer
Dim psColor(50) As Long
Dim MaxPslokatie As Integer                 'aantal STANDAARD af te drukken

Dim BoxTeller As Integer
Dim Box(10) As Integer
Dim BeginX(10) As Integer
Dim BeginY(10) As Integer
Dim TotX(10) As Integer
Dim TotY(10) As Integer

Dim PsLokatie As Integer

Dim FiguurTeller As Integer
Dim FiguurX(4) As Integer
Dim FiguurY(4) As Integer
Dim FiguurName(4) As String

Private Sub Afdrukken()

On Error GoTo Fout

Mim.Teken.FLAGS = cdlPDPrintSetup
MousePointer = vbHourglass
Mim.Teken.PrinterDefault = True
Mim.Teken.CancelError = True
Mim.Teken.ShowPrinter
MousePointer = vbHourglass

PrintUserDef

Printer.CurrentY = Val(VsoftGebied(0).Text)
Printer.CurrentX = 1000
Printer.Print "Begin gereserveerde ruimte voor R&VSOFT"

Printer.CurrentY = Val(VsoftGebied(1).Text)
Printer.CurrentX = 1000
Printer.Print "Einde gereserveerde ruimte voor R&VSOFT"
Printer.NewPage
Printer.EndDoc
MousePointer = vbNormal

Fout:
    Exit Sub
    Resume
End Sub

Private Sub AfDrukVerVersen()
Dim Tabulatie As Integer

On Error Resume Next

psTekst(PsLokatie) = Text1(0).Text
psX(PsLokatie) = Afbeeldingen.Picture2.CurrentX
psY(PsLokatie) = Afbeeldingen.Picture2.CurrentY + Diepgang
psFontSize(PsLokatie) = Afbeeldingen.Picture2.FontSize
psFontName(PsLokatie) = Afbeeldingen.Picture2.FontName
psFontBold(PsLokatie) = Afbeeldingen.Picture2.FontBold
psFontItalic(PsLokatie) = Afbeeldingen.Picture2.FontItalic
psFontUnderLine(PsLokatie) = Afbeeldingen.Picture2.FontUnderline
psColor(PsLokatie) = Afbeeldingen.Picture2.ForeColor

Afbeeldingen.Picture2.Print Text1(0).Text;
If psX(PsLokatie + 1) + psY(PsLokatie + 1) = 0 Then
    psX(PsLokatie + 1) = Afbeeldingen.Picture2.CurrentX
    psY(PsLokatie + 1) = Afbeeldingen.Picture2.CurrentY + Diepgang
End If
PsLokatie = PsLokatie + 1
FicheVerVersen

End Sub

Private Sub AutoOpBouw_Click()

Dim Telhier As Integer

Screen.MousePointer = vbHourglass
DoEvents
Afbeeldingen.Picture2.Cls
PsLokatie = 0
FicheVerVersen
Do While PsLokatie < MaxPslokatie
    AfDrukVerVersen
Loop

PsLokatie = 0
FicheVerVersen
For Telhier = 0 To UBound(Box)
    BoxTeller = Telhier
    PrintBox
Next

For Telhier = 0 To UBound(FiguurX)
    FiguurTeller = Telhier
    FiguurAfdrukTest 0
Next
Afbeeldingen.Show
Afbeeldingen.SetFocus
Screen.MousePointer = vbNormal

End Sub

Private Sub Bestand_Click(Index As Integer)

Select Case Index
    Case 1
        PapierCls
        Inladen Format(KopVoet) + Format(TaalKode) + Format(DocuType)
        PsLokatie = 0
        FicheVerVersen
        BoxTeller = 0: BoxVeldVerVersen
        FiguurTeller = 0: FiguurVerversen
        AutoOpBouw_Click
    Case 2
        Msg$ = "Wegschrijven.  Bent U zeker ?"
        Ktrl = MsgBox(Msg$, 292)
        If Ktrl = 6 Then
            Wegschrijven Format(KopVoet) + Format(TaalKode) + Format(DocuType)
        End If
    Case 4
        Afdrukken
    Case 5
        Afbeeldingen.SetFocus
    Case 7
        Unload Me
    Case Else
End Select

End Sub

Private Sub BoxEnter_Click()

BoxIndexVerVersen
PrintBox
BoxTeller = BoxTeller + 1
If BoxTeller > 10 Then
    BoxTeller = 10
Else
    BoxVeldVerVersen
End If

End Sub

Private Sub BoxIndexVerVersen()

BeginX(BoxTeller) = Val(LijnBeginX.Text)
BeginY(BoxTeller) = Val(LijnBeginY.Text)
TotX(BoxTeller) = Val(LijnTotX.Text)
TotY(BoxTeller) = Val(LijnTotY.Text)
Box(BoxTeller) = BoxMaken.Value

End Sub

Private Sub FiguurIndexVerVersen()

On Error Resume Next
FiguurX(FiguurTeller) = Val(txtFiguurX.Text)
FiguurY(FiguurTeller) = Val(txtFiguurY.Text)
FiguurName(FiguurTeller) = cmbFiguurnamen.Text

End Sub

Sub FiguurVerversen()

Dim NietMeerAanwezig As Boolean

FiguurIndex.Text = Str$(FiguurTeller)
txtFiguurX.Text = Str$(FiguurX(FiguurTeller))
txtFiguurY.Text = Str$(FiguurY(FiguurTeller))

NietMeerAanwezig = True
If FiguurName(FiguurTeller) = "" Then
    cmbFiguurnamen.ListIndex = 0
    NietMeerAanwezig = False
Else
    For T = 0 To cmbFiguurnamen.ListCount - 1
        If cmbFiguurnamen.List(T) = FiguurName(FiguurTeller) Then
            cmbFiguurnamen.ListIndex = T
            NietMeerAanwezig = False
            Exit For
        End If
    Next
End If
If NietMeerAanwezig Then
    MsgBox FiguurName(FiguurTeller) & " niet meer beschikbaar in keuzelijst.  Wordt hierna '-- blanco --' geplaatst", vbInformation
    FiguurName(FiguurTeller) = ""
    cmbFiguurnamen.ListIndex = 0
End If

End Sub

Private Sub BoxMaken_Click()

Box(BoxTeller) = BoxMaken.Value

End Sub

Private Sub BoxVeldVerVersen()

LijnIndex.Text = Str$(BoxTeller)
LijnBeginX.Text = Str$(BeginX(BoxTeller))
LijnBeginY.Text = Str$(BeginY(BoxTeller))
LijnTotX.Text = Str$(TotX(BoxTeller))
LijnTotY.Text = Str$(TotY(BoxTeller))
BoxMaken.Value = Box(BoxTeller)

End Sub

Private Sub cmdFiguurEnter_Click()

FiguurIndexVerVersen

'hier eventueel tekenen
FiguurAfdrukTest (0)


FiguurTeller = FiguurTeller + 1
If FiguurTeller > 4 Then
    FiguurTeller = 10
Else
    FiguurVerversen
End If

End Sub

Private Sub Command1_Click()

Afbeeldingen.Picture2.Cls

End Sub

Private Sub Command2_Click()

PapierCls

PsLokatie = 0
FicheVerVersen
BoxVeldVerVersen

End Sub

Private Sub DokType_Click(Index As Integer)

For T = 0 To 5
    DokType(T).Checked = False
Next

DokType(Index).Checked = True
DocuType = Index

End Sub

Private Sub EnterReturn_Click()

AfDrukVerVersen

End Sub

Private Sub FicheVerVersen()

On Error GoTo ErIsIetsFOUT

LOCATION_.Text = Str$(PsLokatie)
Text1(0).Text = psTekst(PsLokatie)
Text1(2).Text = Str$(psX(PsLokatie))
Text1(3).Text = Str$(psY(PsLokatie))
If psTekst(PsLokatie) <> "" Then
    Afbeeldingen.Picture2.CurrentX = psX(PsLokatie)
    Afbeeldingen.Picture2.CurrentY = psY(PsLokatie) - Diepgang
    Afbeeldingen.Picture2.FontSize = psFontSize(PsLokatie)
    Afbeeldingen.Picture2.FontName = psFontName(PsLokatie)
    Afbeeldingen.Picture2.FontBold = psFontBold(PsLokatie)
    Afbeeldingen.Picture2.FontItalic = psFontItalic(PsLokatie)
    Afbeeldingen.Picture2.FontUnderline = psFontUnderLine(PsLokatie)
    Afbeeldingen.Picture2.ForeColor = psColor(PsLokatie)
End If
Exit Sub

ErIsIetsFOUT:
MsgBox "Mogelijk werd een FONT verwijderd.  Wijzig via kladblok a.u.b." & vbCrLf & vbCrLf & Err.Description

End Sub

Private Sub FigurenDef_Click()

    frmFiguren.Show 1
    Dim Bestanden As String
    cmbFiguurnamen.Clear
    cmbFiguurnamen.AddItem "-- blanco --"
    Bestanden = Dir(LOCATION_COMPANYDATA & "*.mfd")
    Do While Bestanden <> ""
        cmbFiguurnamen.AddItem Left(Bestanden, Len(Bestanden) - 4)
        Bestanden = Dir
    Loop
    cmbFiguurnamen.ListIndex = 0

End Sub

Sub FiguurAfdrukTest(Dest As Integer)

Dim FigBestandsnaam As String
Dim Hoogte As Variant
Dim Breedte As Variant

If ScrLeesTekstBestand(Msg, LOCATION_COMPANYDATA & FiguurName(FiguurTeller) & ".mfd") Then
        FigBestandsnaam = Left(Msg, InStr(Msg, vbTab) - 1)
        imgFiguur.Picture = LoadPicture(FigBestandsnaam)
        Msg = Mid(Msg, InStr(Msg, vbTab) + 1)
        Hoogte = Left(Msg, InStr(Msg, vbTab) - 1)
        Msg = Mid(Msg, InStr(Msg, vbTab) + 1)
        Breedte = Msg
        If Dest = 1 Then
            Printer.PaintPicture imgFiguur.Picture, FiguurX(FiguurTeller), FiguurY(FiguurTeller), Breedte, Hoogte
        Else
            Afbeeldingen.Picture2.PaintPicture imgFiguur.Picture, FiguurX(FiguurTeller), FiguurY(FiguurTeller), Breedte, Hoogte
        End If
Else
    'MsgBox "afdrukprobleemke figuurobject: " & FiguurName
End If
    
End Sub

Private Sub FiguurIndex_Change()

FiguurTeller = Val(FiguurIndex.Text)
If FiguurTeller > 4 Then
    FiguurTeller = 4
End If
FiguurVerversen

End Sub

Private Sub Form_Load()

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If

Afbeeldingen.Show
Mim.WindowState = 1

    Afbeeldingen.Width = 9510
    Afbeeldingen.Top = 2070
    Afbeeldingen.Left = 60
    Afbeeldingen.Height = 4245
    
    Afbeeldingen.Picture2.Width = 95
    Afbeeldingen.Picture2.Height = 70
    Afbeeldingen.Picture2.AutoRedraw = True
    Afbeeldingen.Picture2.AutoSize = True
    
    Afbeeldingen.Picture2.Cls
    Afbeeldingen.Height = Afbeeldingen.Height + 100
PapierCls
KopVoet = 1
TaalKode = 2
VsoftGebied(0).Text = 2000
VsoftGebied(1).Text = 5000

Dim Bestanden As String
cmbFiguurnamen.Clear
cmbFiguurnamen.AddItem "-- blanco --"
Bestanden = Dir(LOCATION_COMPANYDATA & "*.mfd")
Do While Bestanden <> ""
    cmbFiguurnamen.AddItem Left(Bestanden, Len(Bestanden) - 4)
    Bestanden = Dir
Loop
cmbFiguurnamen.ListIndex = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

Close
Unload Afbeeldingen
Mim.WindowState = 0

End Sub

Private Sub Inladen(TypeEnTaal As String)
Dim FlFree As Integer
Dim Teller As Integer
 
On Error GoTo ErrorLoad

If Dir$(LOCATION_COMPANYDATA + "DDEF" + TypeEnTaal + ".Txt") = "" Then
    Beep
    Exit Sub
Else
    FlFree = FreeFile
    Open LOCATION_COMPANYDATA + "DDEF" + TypeEnTaal + ".Txt" For Input As FlFree
    Input #FlFree, Diepgang
    Input #FlFree, A$
    VsoftGebied(0).Text = A$
    Input #FlFree, A$
    VsoftGebied(1).Text = A$
    For Teller = 0 To 10
        Input #FlFree, BeginX(Teller), BeginY(Teller), TotX(Teller), TotY(Teller), Box(Teller)
    Next
    Teller = 0
    While Not EOF(FlFree)
        Input #FlFree, psTekst(Teller), psX(Teller), psY(Teller), psFontSize(Teller), psFontName(Teller), psFontBold(Teller), psFontItalic(Teller), psFontUnderLine(Teller), psColor(Teller)
        Teller = Teller + 1
    Wend
    MaxPslokatie = Teller
    Close FlFree
    If Dir$(LOCATION_COMPANYDATA + "DDEF" + TypeEnTaal + "G.Txt") = "" Then
    Else
        FlFree = FreeFile
        Open LOCATION_COMPANYDATA + "DDEF" + TypeEnTaal + "G.Txt" For Input As FlFree
        Teller = 0
        While Not EOF(FlFree)
            Input #FlFree, FiguurX(Teller), FiguurY(Teller), FiguurName(Teller)
            Teller = Teller + 1
        Wend
        Close FlFree
    End If
End If

Exit Sub

ErrorLoad:
Exit Sub
Resume

End Sub

Private Sub LijnIndex_Change()

BoxTeller = Val(LijnIndex.Text)
If BoxTeller > 10 Then
    MaxPslokatie = 10
End If
BoxVeldVerVersen

End Sub

Private Sub LOCATION__Change()

PsLokatie = Val(LOCATION_.Text)
If MaxPslokatie < PsLokatie Then
    MaxPslokatie = PsLokatie
End If
FicheVerVersen

End Sub

Private Sub mTaal_Click(Index As Integer)

For T = 1 To 4
    mTaal(T).Checked = False
Next

mTaal(Index).Checked = True
TaalKode = Index

End Sub

Private Sub PapierCls()
Dim T As Integer

For T = 0 To MaxPslokatie
    psX(T) = 0
    psY(T) = 0
    psFontSize(T) = 0
    psFontName(T) = ""
    psFontBold(T) = 0
    psFontItalic(T) = 0
    psFontUnderLine(T) = 0
    psColor(T) = 0
    psTekst(T) = ""
Next
For T = 0 To 10
    BeginX(T) = 0
    BeginY(T) = 0
    TotX(T) = 0
    TotY(T) = 0
    Box(T) = 0
Next
For T = 0 To 4
    FiguurX(T) = 0
    FiguurY(T) = 0
    FiguurName(T) = ""
Next
MaxPslokatie = 0

End Sub

Private Sub PrintUserDef()
Dim T         As Integer
Dim Tabulatie As Integer

    PsLokatie = 0
    
    Do While PsLokatie <= MaxPslokatie
        If psTekst(PsLokatie) <> "" Then
            Printer.FontName = psFontName(PsLokatie)
            Printer.FontItalic = psFontItalic(PsLokatie)
            Printer.FontSize = psFontSize(PsLokatie)
            Printer.FontBold = psFontBold(PsLokatie)
            Printer.FontUnderline = psFontUnderLine(PsLokatie)
            Printer.ForeColor = psColor(PsLokatie)
        End If
        Printer.CurrentX = psX(PsLokatie)
        Printer.CurrentY = psY(PsLokatie)
        Printer.Print psTekst(PsLokatie);
        PsLokatie = PsLokatie + 1
    Loop
    For T = 0 To 10
        If BeginX(T) = 0 Then
        ElseIf Box(T) Then
            Printer.Line (BeginX(T), BeginY(T))-(TotX(T), TotY(T)), , B
        Else
            Printer.Line (BeginX(T), BeginY(T))-(TotX(T), TotY(T))
        End If
    Next

End Sub

Private Sub Teken_Click()

On Local Error GoTo FontError

Mim.Teken.DialogTitle = Text1(0).Text
Mim.Teken.CancelError = True
Mim.Teken.FLAGS = cdlCFPrinterFonts + cdlCFEffects
'set initial values for the dialog
    Mim.Teken.FontSize = Afbeeldingen.Picture2.FontSize
    Mim.Teken.FontName = Afbeeldingen.Picture2.FontName
    Mim.Teken.FontBold = Afbeeldingen.Picture2.FontBold
    Mim.Teken.FontItalic = Afbeeldingen.Picture2.FontItalic
    Mim.Teken.FontUnderline = Afbeeldingen.Picture2.FontUnderline
    Mim.Teken.Color = Afbeeldingen.Picture2.ForeColor
'display Choose Font dialog
Mim.Teken.ShowFont
    'change text font according to options selected
    Afbeeldingen.Picture2.FontName = Mim.Teken.FontName
    Afbeeldingen.Picture2.FontSize = Mim.Teken.FontSize
    Afbeeldingen.Picture2.FontBold = Mim.Teken.FontBold
    Afbeeldingen.Picture2.FontItalic = Mim.Teken.FontItalic
    Afbeeldingen.Picture2.FontUnderline = Mim.Teken.FontUnderline
    Afbeeldingen.Picture2.ForeColor = Mim.Teken.Color
EnterReturn_Click

FontError:
    Exit Sub
    Resume

End Sub

Private Sub Text1_Change(Index As Integer)

On Error GoTo HandlerText

Select Case Index
    Case 2
        Afbeeldingen.Picture2.CurrentX = Val(Text1(2).Text)
    Case 3
        Afbeeldingen.Picture2.CurrentY = Val(Text1(3).Text)
    Case 4
        PsLokatie = Val(Text1(4).Text)
    Case Else
End Select
Exit Sub

HandlerText:
    Beep
    Exit Sub

End Sub

Private Sub Text1_GotFocus(Index As Integer)

Select Case Index
    Case 0
        EnterReturn.Default = True
End Select

End Sub

Private Sub Text1_LostFocus(Index As Integer)

Select Case Index
    Case 0
        EnterReturn.Default = False
End Select

End Sub

Private Sub Wegschrijven(TypeEnTaal As String)
Dim FlFree As Integer
Dim Teller As Integer

FlFree = FreeFile
Open LOCATION_COMPANYDATA + "DDEF" + TypeEnTaal + ".Txt" For Output As FlFree

Write #FlFree, Diepgang
Write #FlFree, VsoftGebied(0).Text
Write #FlFree, VsoftGebied(1).Text
For Teller = 0 To 10
    Write #FlFree, BeginX(Teller), BeginY(Teller), TotX(Teller), TotY(Teller), Box(Teller)
Next
For Teller = 0 To MaxPslokatie
    If psTekst(Teller) = "" Then
        Exit For
    Else
        Write #FlFree, psTekst(Teller), psX(Teller), psY(Teller), psFontSize(Teller), psFontName(Teller), psFontBold(Teller), psFontItalic(Teller), psFontUnderLine(Teller), psColor(Teller)
    End If
Next
Close FlFree

If Dir$(LOCATION_COMPANYDATA + "DDEF" + TypeEnTaal + "G.Txt") <> "" Then
    Kill LOCATION_COMPANYDATA + "DDEF" + TypeEnTaal + "G.Txt"
End If
For Teller = 0 To 4
    If FiguurName(Teller) = "" Or FiguurName(Teller) = "-- blanco --" Then
    Else
        FlFree = FreeFile
        Open LOCATION_COMPANYDATA + "DDEF" + TypeEnTaal + "G.Txt" For Append As FlFree
        Write #FlFree, FiguurX(Teller), FiguurY(Teller), FiguurName(Teller)
        Close FlFree
    End If
Next

End Sub

Function PrintBox()

If Box(BoxTeller) Then
    Afbeeldingen.Picture2.Line (BeginX(BoxTeller), BeginY(BoxTeller) - Diepgang)-(TotX(BoxTeller), TotY(BoxTeller) - Diepgang), , B
Else
    Afbeeldingen.Picture2.Line (BeginX(BoxTeller), BeginY(BoxTeller) - Diepgang)-(TotX(BoxTeller), TotY(BoxTeller) - Diepgang)
End If

End Function

Private Sub ViaKB_Click(Index As Integer)

Dim HetBESTAND As String

Select Case Index
    Case 0
        HetBESTAND = LOCATION_COMPANYDATA + "DDEF" + Format(KopVoet) + Format(TaalKode) + Format(DocuType) + ".Txt"
        If Dir(HetBESTAND) = "" Then
            MsgBox HetBESTAND & " is niet aanwezig", vbInformation
        Else
            Shell "notepad.exe " + HetBESTAND, vbNormalFocus
        End If
    Case 1
        HetBESTAND = LOCATION_COMPANYDATA + "DDEF" + Format(KopVoet) + Format(TaalKode) + Format(DocuType) + "G.Txt"
        If Dir(HetBESTAND) = "" Then
            MsgBox HetBESTAND & " is niet aanwezig", vbInformation
        Else
            Shell "notepad.exe " + HetBESTAND, vbNormalFocus
        End If
End Select

End Sub
