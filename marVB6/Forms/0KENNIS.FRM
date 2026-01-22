VERSION 5.00
Begin VB.Form FrmKennisBeheer 
   Caption         =   "Kennisbeheer"
   ClientHeight    =   6120
   ClientLeft      =   1635
   ClientTop       =   1605
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6120
   ScaleWidth      =   9825
   Begin VB.TextBox txtComm 
      Height          =   315
      Index           =   3
      Left            =   1320
      TabIndex        =   9
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox txtComm 
      Height          =   315
      Index           =   4
      Left            =   3360
      TabIndex        =   10
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox txtComm 
      Height          =   315
      Index           =   5
      Left            =   5520
      TabIndex        =   11
      Top             =   3000
      Width           =   4155
   End
   Begin VB.TextBox txtComm 
      Height          =   315
      Index           =   2
      Left            =   5520
      TabIndex        =   7
      Top             =   2640
      Width           =   4155
   End
   Begin VB.TextBox txtComm 
      Height          =   315
      Index           =   1
      Left            =   3360
      TabIndex        =   6
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox txtComm 
      Height          =   315
      Index           =   0
      Left            =   1320
      TabIndex        =   5
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CheckBox chkTakGroep 
      Caption         =   "Alle info binnen takgroep"
      Height          =   240
      Left            =   120
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5670
      Width           =   2505
   End
   Begin VB.CheckBox chkAlleMijen 
      Caption         =   "Alle maatschappijen kontroleren voor deze tak"
      Height          =   240
      Left            =   120
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5430
      Width           =   3915
   End
   Begin VB.CommandButton CmdBewaar 
      Caption         =   "&Bewaren"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6960
      TabIndex        =   15
      Top             =   5460
      Width           =   1335
   End
   Begin VB.TextBox TxtTakInfo 
      Height          =   1635
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   3720
      Width           =   9615
   End
   Begin VB.TextBox TxtMijInfo 
      Height          =   1995
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   600
      Width           =   9675
   End
   Begin VB.ComboBox CmbKeuzeInfo 
      Height          =   315
      Index           =   1
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   3360
      Width           =   8775
   End
   Begin VB.ComboBox CmbKeuzeInfo 
      Height          =   315
      Index           =   0
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   5535
   End
   Begin VB.CommandButton CmdSluiten 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   375
      Left            =   8280
      TabIndex        =   16
      Top             =   5460
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "&Schade Tel :"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Fax:"
      Height          =   255
      Left            =   3000
      TabIndex        =   22
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Email:"
      Height          =   255
      Left            =   5040
      TabIndex        =   21
      Top             =   3000
      Width           =   555
   End
   Begin VB.Label Label5 
      Caption         =   "Email:"
      Height          =   255
      Left            =   5040
      TabIndex        =   20
      Top             =   2640
      Width           =   555
   End
   Begin VB.Label Label4 
      Caption         =   "Fax:"
      Height          =   255
      Left            =   3000
      TabIndex        =   19
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "&Productie Tel :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "&Info deze Maatschappij/Tak"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2235
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   6660
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label2 
      Caption         =   "&Tak 1998"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "&Maatschappij"
      Height          =   195
      Index           =   0
      Left            =   4200
      TabIndex        =   0
      Top             =   0
      Width           =   1035
   End
End
Attribute VB_Name = "FrmKennisBeheer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReedsAanwezig As Boolean

Sub InstallTak()
Dim Pipo As String
Dim Keuze As Integer
Dim Aantallijnen As Integer

Pipo = ZoekEnPlaats(CmbKeuzeInfo(1), "NTKB2897", 0, Keuze, vBibTekst(TABLE_CONTRACTS, "#v223 #"))
CmbKeuzeInfo(1).ListIndex = Keuze

End Sub


Private Sub CmbKeuzeInfo_Click(Index As Integer)
Dim TeZoekenString As String

bGet TABLE_VARIOUS, 1, "27" + Left(CmbKeuzeInfo(1), 3) + Left(CmbKeuzeInfo(0), 4)
If Ktrl Then
    ReedsAanwezig = False
    TxtMijInfo = "- Nihil -"
    txtComm(0).Text = "" 'Tel. Productie
    txtComm(1).Text = "" 'Fax Productie
    txtComm(2).Text = "" 'email Productie
    txtComm(3).Text = "" 'Tel. Schade
    txtComm(4).Text = "" 'Fax Schade
    txtComm(5).Text = "" 'email Schade
Else
    RecordToVeld TABLE_VARIOUS
    TxtMijInfo = vBibTekst(TABLE_VARIOUS, "#v132 #")
    txtComm(0).Text = vBibTekst(TABLE_VARIOUS, "A10A")
    txtComm(1).Text = vBibTekst(TABLE_VARIOUS, "vs02")
    txtComm(2).Text = vBibTekst(TABLE_VARIOUS, "v224")
    
    txtComm(3).Text = vBibTekst(TABLE_VARIOUS, "s10A")
    txtComm(4).Text = vBibTekst(TABLE_VARIOUS, "ss02")
    txtComm(5).Text = vBibTekst(TABLE_VARIOUS, "s224")
            
    ReedsAanwezig = True
End If
cmdBewaar.Enabled = False

End Sub


Private Sub CmdBewaar_Click()

If ReedsAanwezig Then
    'eerst toelating vragen tot overschrijven
    Msg = "Bestaande tekst overschrijven.  Bent U zeker ?"
    KtrlBox = MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton2)
    If KtrlBox = vbYes Then
        vBib TABLE_VARIOUS, TxtMijInfo, "v132"
        vBib TABLE_VARIOUS, txtComm(0).Text, "A10A"
        vBib TABLE_VARIOUS, txtComm(1).Text, "vs02"
        vBib TABLE_VARIOUS, txtComm(2).Text, "v224"
        
        vBib TABLE_VARIOUS, txtComm(3).Text, "s10A"
        vBib TABLE_VARIOUS, txtComm(4).Text, "ss02"
        vBib TABLE_VARIOUS, txtComm(5).Text, "s224"
        
        vBib TABLE_VARIOUS, "27" + vBibTekst(TABLE_VARIOUS, "#v223 #") + vBibTekst(TABLE_VARIOUS, "#A010 #"), "v005"
        bUpdate TABLE_VARIOUS, 1
    End If
Else
    TLB_RECORD(TABLE_VARIOUS) = ""
    vBib TABLE_VARIOUS, Left(CmbKeuzeInfo(0), 4), "A010"
    vBib TABLE_VARIOUS, Left(CmbKeuzeInfo(1), 3), "v223"
    vBib TABLE_VARIOUS, TxtMijInfo, "v132"
    vBib TABLE_VARIOUS, txtComm(0).Text, "A10A"
    vBib TABLE_VARIOUS, txtComm(1).Text, "vs02"
    vBib TABLE_VARIOUS, txtComm(2).Text, "v224"
    
    vBib TABLE_VARIOUS, txtComm(3).Text, "s10A"
    vBib TABLE_VARIOUS, txtComm(4).Text, "ss02"
    vBib TABLE_VARIOUS, txtComm(5).Text, "s224"
       
    vBib TABLE_VARIOUS, "27" + vBibTekst(TABLE_VARIOUS, "#v223 #") + vBibTekst(TABLE_VARIOUS, "#A010 #"), "v005"
    bInsert TABLE_VARIOUS, 1
End If
CmbKeuzeInfo_Click 0

End Sub

Private Sub cmdSluiten_Click()

Unload Me

End Sub


Private Sub Form_Load()

    InstallMij
    InstallTak

End Sub

Private Sub InstallMij()
Dim T As Integer
Dim MijIs As Integer

MijIs = 0
CmbKeuzeInfo(0).Clear
T = -1
bGetOrGreater TABLE_SUPPLIERS, 0, vSet("CO", 12)
If Ktrl Then
    Beep
    Exit Sub
Else
    RecordToVeld TABLE_SUPPLIERS
End If
Screen.MousePointer = vbHourglass
If Left(KEY_BUF(TABLE_SUPPLIERS), 2) <> "CO" Then
    Beep
Else
    Do
        T = T + 1
        If Mid(vBibTekst(TABLE_SUPPLIERS, "#A110 #"), 3, 4) = vBibTekst(TABLE_CONTRACTS, "#A010 #") Then
            MijIs = T
        End If
        CmbKeuzeInfo(0).AddItem Mid(vBibTekst(TABLE_SUPPLIERS, "#A110 #"), 3, 4) + ": " + vBibTekst(TABLE_SUPPLIERS, "#A100 #") + "/" + vSet(vBibTekst(TABLE_SUPPLIERS, "#A400 #"), 8)
        bNext TABLE_SUPPLIERS
        If Ktrl Or Left(KEY_BUF(TABLE_SUPPLIERS), 2) <> "CO" Then
            Exit Do
        Else
            RecordToVeld TABLE_SUPPLIERS
        End If
    Loop
End If
If CmbKeuzeInfo(0).ListCount Then
    CmbKeuzeInfo(0).ListIndex = MijIs
End If
Screen.MousePointer = vbNormal

End Sub


Private Sub txtComm_KeyPress(Index As Integer, KeyAscii As Integer)

If cmdBewaar.Enabled = False Then cmdBewaar.Enabled = True

End Sub

Private Sub TxtMijInfo_Change()

If cmdBewaar.Enabled = False Then cmdBewaar.Enabled = True

End Sub


