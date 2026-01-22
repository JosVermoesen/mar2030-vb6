VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Wissel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EURO Wisselfunctie (c) 2002 Roelandt & Vermoesen Bvba"
   ClientHeight    =   4935
   ClientLeft      =   5610
   ClientTop       =   3030
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6675
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOvernemen 
      Caption         =   "Overnemen"
      Height          =   315
      Left            =   5520
      TabIndex        =   51
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   3960
      TabIndex        =   47
      Top             =   660
      Width           =   2655
      Begin VB.OptionButton oMunt 
         Alignment       =   1  'Right Justify
         Caption         =   "BEF"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   49
         Top             =   180
         Width           =   615
      End
      Begin VB.OptionButton oMunt 
         Caption         =   "EUR"
         Height          =   255
         Index           =   1
         Left            =   1860
         TabIndex        =   48
         Top             =   180
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ontvangen"
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
         Index           =   4
         Left            =   780
         TabIndex        =   50
         Top             =   180
         Width           =   1095
      End
   End
   Begin VB.OptionButton obTerug 
      Caption         =   "Option1"
      Height          =   195
      Index           =   1
      Left            =   6420
      TabIndex        =   46
      Top             =   1860
      Width           =   195
   End
   Begin VB.OptionButton obTerug 
      Caption         =   "Option1"
      Height          =   195
      Index           =   0
      Left            =   6420
      TabIndex        =   45
      Top             =   1260
      Value           =   -1  'True
      Width           =   195
   End
   Begin VB.TextBox txtTeruggave 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   4380
      TabIndex        =   44
      Text            =   "0"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton txtMunt 
      Caption         =   "0"
      Height          =   255
      Index           =   7
      Left            =   3780
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   1500
      Width           =   495
   End
   Begin VB.CommandButton txtMunt 
      Caption         =   "0"
      Height          =   255
      Index           =   6
      Left            =   3240
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   1500
      Width           =   495
   End
   Begin VB.CommandButton txtMunt 
      Caption         =   "0"
      Height          =   255
      Index           =   5
      Left            =   2700
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   1500
      Width           =   495
   End
   Begin VB.CommandButton txtMunt 
      Caption         =   "0"
      Height          =   255
      Index           =   4
      Left            =   2160
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   1500
      Width           =   495
   End
   Begin VB.CommandButton txtMunt 
      Caption         =   "0"
      Height          =   255
      Index           =   3
      Left            =   1620
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   1500
      Width           =   495
   End
   Begin VB.CommandButton txtMunt 
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   1500
      Width           =   495
   End
   Begin VB.CommandButton txtMunt 
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   540
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   1500
      Width           =   495
   End
   Begin VB.CommandButton txtMunt 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1500
      Width           =   495
   End
   Begin VB.TextBox txtTeruggave 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   5520
      TabIndex        =   34
      Text            =   "0"
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtBedragBEF 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   4380
      TabIndex        =   32
      Text            =   "0"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton txtBiljet 
      Caption         =   "0"
      Height          =   255
      Index           =   6
      Left            =   3240
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   300
      Width           =   495
   End
   Begin VB.CommandButton txtBiljet 
      Caption         =   "0"
      Height          =   255
      Index           =   5
      Left            =   2700
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   300
      Width           =   495
   End
   Begin VB.CommandButton txtBiljet 
      Caption         =   "0"
      Height          =   255
      Index           =   4
      Left            =   2160
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   300
      Width           =   495
   End
   Begin VB.CommandButton txtBiljet 
      Caption         =   "0"
      Height          =   255
      Index           =   3
      Left            =   1620
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   300
      Width           =   495
   End
   Begin VB.CommandButton txtBiljet 
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   300
      Width           =   495
   End
   Begin VB.CommandButton txtBiljet 
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   540
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   300
      Width           =   495
   End
   Begin VB.CommandButton txtBiljet 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   300
      Width           =   495
   End
   Begin VB.TextBox txtTeBetalen 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   5520
      TabIndex        =   24
      Text            =   "0"
      Top             =   300
      Width           =   1095
   End
   Begin VB.TextBox txtTeBetalen 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   4380
      TabIndex        =   21
      Text            =   "0"
      Top             =   300
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3015
      Left            =   60
      TabIndex        =   20
      Top             =   1860
      Width           =   5355
      Begin VB.Image Image1 
         Height          =   2655
         Left            =   120
         Top             =   240
         Width           =   5115
      End
   End
   Begin VB.CommandButton cbSluiten 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   315
      Left            =   5520
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4500
      Width           =   1095
   End
   Begin VB.CommandButton cbKontroleer 
      Caption         =   "Kontroleer"
      Default         =   -1  'True
      Height          =   315
      Left            =   5520
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox txtBedrag 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5520
      TabIndex        =   0
      Text            =   "0"
      Top             =   1200
      Width           =   855
   End
   Begin MSComctlLib.ImageList ilMunten 
      Left            =   1860
      Top             =   2460
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   63
      ImageHeight     =   63
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMuntEnBiljetWissel.frx":0000
            Key             =   "1CENTR"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMuntEnBiljetWissel.frx":0680
            Key             =   "2CENTR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMuntEnBiljetWissel.frx":0E6C
            Key             =   "5CENTR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMuntEnBiljetWissel.frx":18D0
            Key             =   "10CENTR"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMuntEnBiljetWissel.frx":2218
            Key             =   "20CENTR"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMuntEnBiljetWissel.frx":2D60
            Key             =   "50CENTR"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMuntEnBiljetWissel.frx":39FC
            Key             =   "1EUROR"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMuntEnBiljetWissel.frx":430C
            Key             =   "2EUROR"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilBiljetten 
      Left            =   480
      Top             =   2700
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   261
      ImageHeight     =   134
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMuntEnBiljetWissel.frx":4D70
            Key             =   "5b"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMuntEnBiljetWissel.frx":6614
            Key             =   "10b"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMuntEnBiljetWissel.frx":83F0
            Key             =   "20b"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMuntEnBiljetWissel.frx":A448
            Key             =   "50b"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMuntEnBiljetWissel.frx":C628
            Key             =   "100b"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMuntEnBiljetWissel.frx":EA5C
            Key             =   "200b"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMuntEnBiljetWissel.frx":11214
            Key             =   "500b"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Teruggave"
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
      Index           =   6
      Left            =   5520
      TabIndex        =   35
      Top             =   1500
      Width           =   1095
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Totaal"
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
      Index           =   5
      Left            =   3840
      TabIndex        =   33
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EUR"
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
      Index           =   3
      Left            =   6060
      TabIndex        =   23
      Top             =   0
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BEF"
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
      Index           =   2
      Left            =   4980
      TabIndex        =   22
      Top             =   0
      Width           =   495
   End
   Begin VB.Label lblMunt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.02"
      Height          =   255
      Index           =   6
      Left            =   3240
      TabIndex        =   19
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblMunt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.05"
      Height          =   255
      Index           =   5
      Left            =   2700
      TabIndex        =   18
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblMunt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.10"
      Height          =   255
      Index           =   4
      Left            =   2160
      TabIndex        =   17
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblMunt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.20"
      Height          =   255
      Index           =   3
      Left            =   1620
      TabIndex        =   16
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblMunt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.50"
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   15
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblMunt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   540
      TabIndex        =   14
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblMunt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblMunt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.01"
      Height          =   255
      Index           =   7
      Left            =   3780
      TabIndex        =   12
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Munten"
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
      Index           =   1
      Left            =   0
      TabIndex        =   11
      Top             =   900
      Width           =   3735
   End
   Begin VB.Label lblBiljet 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      Height          =   255
      Index           =   6
      Left            =   3240
      TabIndex        =   10
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblBiljet 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
      Height          =   255
      Index           =   5
      Left            =   2700
      TabIndex        =   9
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblBiljet 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "20"
      Height          =   255
      Index           =   4
      Left            =   2160
      TabIndex        =   8
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblBiljet 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "50"
      Height          =   255
      Index           =   3
      Left            =   1620
      TabIndex        =   7
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblBiljet 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   6
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblBiljet 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "200"
      Height          =   255
      Index           =   1
      Left            =   540
      TabIndex        =   5
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblBiljet 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "500"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Biljetten"
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
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3735
   End
   Begin VB.Menu Instellingen 
      Caption         =   "Instellingen"
      Begin VB.Menu Init 
         Caption         =   "Initialistatie EUR Geldlade"
         Index           =   0
      End
      Begin VB.Menu Init 
         Caption         =   "Initialisatie BEF Geldlade"
         Index           =   1
      End
   End
End
Attribute VB_Name = "Wissel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const EurMunten = 8
Const EurBiljetten = 7

Dim Biljet(EurBiljetten - 1)
Dim Munt(EurMunten - 1)
Dim isBerekend As Boolean
Dim TelTot As Integer
Dim Bedrag As Currency


Private Sub cbKontroleer_Click()

    Dim KtrlAantal As Single
    Dim TrimString As String
                       
    KtrlAantal = KiesBedrag
    KtrlBox = Schoon
    'Hoeveel BILJETTEN ?
    For TelTot = 0 To UBound(Biljet)
        On Error Resume Next
        KtrlAantal = Bedrag \ (Biljet(TelTot) * 100)
        If Err Then MsgBox "Plezante(n)", vbExclamation: isBerekend = True: Exit Sub
        If KtrlAantal <> 0 And KtrlAantal * Biljet(TelTot) * 100 <= Bedrag Then
            Bedrag = Bedrag - (KtrlAantal * Biljet(TelTot) * 100)
            txtBiljet(TelTot).Caption = KtrlAantal
            lblBiljet(TelTot).BackColor = &H80FFFF
        Else
            txtBiljet(TelTot).Caption = 0
        End If
    Next
        
    'Hoeveel MUNTEN ?
    For TelTot = 0 To UBound(Munt)
        KtrlAantal = Bedrag \ (Munt(TelTot) * 100)
        If KtrlAantal <> 0 And KtrlAantal * Munt(TelTot) * 100 <= Bedrag Then
            Bedrag = Bedrag - (KtrlAantal * Munt(TelTot) * 100)
            txtMunt(TelTot).Caption = KtrlAantal
            lblMunt(TelTot).BackColor = &H80FFFF
        Else
            txtMunt(TelTot).Caption = 0
        End If
    Next
    If Bedrag <> 0 Then
        MsgBox Bedrag, vbExclamation, "Rest!?!?!?"
    End If
    isBerekend = True
        
End Sub


Private Sub cbSluiten_Click()

    KasBetalingBEF = 0
    KasBetalingEUR = 0
    KasTerugEUR = 0
    Unload Me
    
End Sub


Private Sub cmdOvernemen_Click()

    If KasBetalingBEF <> 0 Then
        KasBetalingEUR = 0
    Else
        If txtBedrag.Enabled Then
            KasBetalingEUR = Val(txtBedrag.Text)
        Else
            KasBetalingEUR = CDbl(txtBedrag.Text)
        End If
    End If
    KasTerugEUR = CDbl(txtTeruggave(0).Text)
    Unload Me
    
End Sub

Private Sub Form_Load()

    Biljet(0) = 500
    Biljet(1) = 200
    Biljet(2) = 100
    Biljet(3) = 50
    Biljet(4) = 20
    Biljet(5) = 10
    Biljet(6) = 5

    Munt(0) = 2
    Munt(1) = 1
    Munt(2) = 0.5
    Munt(3) = 0.2
    Munt(4) = 0.1
    Munt(5) = 0.05
    Munt(6) = 0.02
    Munt(7) = 0.01
    
    If KasTicketTotaal Then
        oMunt(1).Visible = False
        txtBedrag.Text = Format(KasTicketTotaal, "#,##0.00")
        txtTeBetalen(1).Text = txtBedrag.Text
        txtTeBetalen(0).Text = Format(Round(KasTicketTotaal * Euro), "#,##0.00")
        txtBedragBEF.Text = txtTeBetalen(0).Text
        txtBedrag.Text = Str$(KasTicketTotaal)
        oMunt(1).Visible = True
        cbKontroleer_Click
    End If
    KasBetalingBEF = 0
    'KasTotaal (enkel in EUR)
    
    'KasBetalingBEF
    'KasBetalingEUR
    'KasTerugEUR
    
    'KasTotaalBEF
    'KasTotaalEUR
    isBerekend = False

End Sub


Private Sub Init_Click(Index As Integer)

    Select Case Index
        Case 0 'EUR
            MsgBox "Stop"
        
        Case 1 'BEF
            MsgBox "Stop"
            
    End Select
    
End Sub


Private Sub lblBiljet_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Image1.Picture = ilBiljetten.ListImages(EurBiljetten - Index).Picture
    If Frame1.Caption = txtBiljet(Index).Caption & " biljetten" Then
    Else
        Frame1.Caption = txtBiljet(Index).Caption & " biljetten"
    End If
    
End Sub


Private Sub lblMunt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Image1.Picture = ilMunten.ListImages(EurMunten - Index).Picture
    If Frame1.Caption = txtMunt(Index).Caption & " stukken" Then
    Else
        Frame1.Caption = txtMunt(Index).Caption & " stukken"
    End If
    
End Sub


Private Sub obTerug_Click(Index As Integer)

    cbKontroleer_Click

End Sub


Private Sub oMunt_Click(Index As Integer)

    If Index = 0 Then
        txtBedragBEF.Text = Str$(txtBedragBEF.Text)
        txtBedrag.Text = Format(Round(txtBedrag.Text, 2), "#,##0.00")
        txtBedrag.Enabled = False
        txtBedragBEF.Enabled = True
        txtBedragBEF.SelLength = Len(txtBedragBEF.Text)
        txtBedragBEF.SetFocus
    Else
        txtBedrag.Text = Str$(txtBedrag.Text)
        txtBedragBEF.Text = Format(Round(txtBedragBEF), "#,##0.00")
        txtBedrag.Enabled = True
        txtBedragBEF.Enabled = False
        txtBedrag.SelLength = Len(txtBedrag.Text)
        txtBedrag.SetFocus
    End If
    cbKontroleer_Click
    
End Sub


Private Sub txtBedrag_Change()
    
    If oMunt(1).Visible = False Then Exit Sub
    If oMunt(0).Value = True Then
        Exit Sub
    End If
    If isBerekend = True Then
        For TelTot = 0 To UBound(Biljet)
            txtBiljet(TelTot).Caption = 0
            lblBiljet(TelTot).BackColor = &H80000005
        Next
        For TelTot = 0 To UBound(Munt)
            txtMunt(TelTot).Caption = 0
            lblMunt(TelTot).BackColor = &H80000005
        Next
        isBerekend = False
    End If
    Image1.Picture = Nothing
    If txtBedrag.Text = "" Then
        txtBedrag = 0
    End If
    txtBedragBEF.Text = Format(Round(Val(txtBedrag.Text) * Euro), "#,##0.00")
    txtTeruggave(0).Text = Format(-CDbl(txtTeBetalen(1).Text) + Val(txtBedrag.Text), "#,##0.00")
    
End Sub


Private Sub txtBedragBEF_Change()

    If oMunt(1).Visible = False Then Exit Sub
    If oMunt(1).Value = True Then
        KasBetalingBEF = 0
        Exit Sub
    Else
        KasBetalingBEF = Val(txtBedragBEF.Text)
    End If
    If txtBedragBEF.Text = "" Then
        txtBedragBEF.Text = 0
    End If
    txtBedrag.Text = Format(Round(Val(txtBedragBEF.Text) / Euro, 2), "#,##0.00")
    txtTeruggave(0).Text = Format(-CDbl(txtTeBetalen(1).Text) + CDbl(txtBedrag.Text), "#,##0.00")
    
End Sub


Private Sub txtBiljet_GotFocus(Index As Integer)

    Image1.Picture = ilBiljetten.ListImages(EurBiljetten - Index).Picture
    If Frame1.Caption = txtBiljet(Index).Caption & " biljetten" Then
    Else
        Frame1.Caption = txtBiljet(Index).Caption & " biljetten"
    End If
    
End Sub


Private Sub txtBiljet_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Image1.Picture = ilBiljetten.ListImages(EurBiljetten - Index).Picture
    If Frame1.Caption = txtBiljet(Index).Caption & " biljetten" Then
    Else
        Frame1.Caption = txtBiljet(Index).Caption & " biljetten"
    End If
    
End Sub


Private Sub txtMunt_GotFocus(Index As Integer)

    Image1.Picture = ilMunten.ListImages(EurMunten - Index).Picture
    If Frame1.Caption = txtMunt(Index).Caption & " stukken" Then
    Else
        Frame1.Caption = txtMunt(Index).Caption & " stukken"
    End If
    
End Sub


Private Sub txtMunt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Image1.Picture = ilMunten.ListImages(EurMunten - Index).Picture
    If Frame1.Caption = txtMunt(Index).Caption & " stukken" Then
    Else
        Frame1.Caption = txtMunt(Index).Caption & " stukken"
    End If

End Sub


Private Sub txtTeruggave_Change(Index As Integer)

    If Index = 0 Then
        txtTeruggave(1).Text = Format(Round(CDbl(txtTeruggave(0).Text) * Euro), "#,##0.00")
    End If

End Sub


Function KiesBedrag() As Boolean

    Dim TrimString As String
    
    If obTerug(0).Value = True Then
        If oMunt(1).Value = True Then
            'niks
            TrimString = Trim$(Str$(Round(Val(txtBedrag.Text), 2)))
            Bedrag = Val(TrimString) * 100
        Else
            't is in franken !
            Bedrag = CDbl(txtBedrag.Text) * 100
        End If
    Else
        Bedrag = CDbl(txtTeruggave(0).Text) * 100
    End If
    
End Function


Function Schoon()

    For TelTot = 0 To UBound(Biljet)
        txtBiljet(TelTot).Caption = 0
        lblBiljet(TelTot).BackColor = &H80000005
    Next
    For TelTot = 0 To UBound(Munt)
        txtMunt(TelTot).Caption = 0
        lblMunt(TelTot).BackColor = &H80000005
    Next
    isBerekend = False

End Function
