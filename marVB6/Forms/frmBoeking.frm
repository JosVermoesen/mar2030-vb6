VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBoeking 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BoekingsDetail"
   ClientHeight    =   3450
   ClientLeft      =   2010
   ClientTop       =   2790
   ClientWidth     =   8280
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   8280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNegeren 
      Cancel          =   -1  'True
      Caption         =   "Boeking Terugzetten"
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdBoeken 
      Caption         =   "&Boeking laten doorgaan"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   1935
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   5530
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Journaal"
      TabPicture(0)   =   "frmBoeking.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "mshfBoekLijst"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Document"
      TabPicture(1)   =   "frmBoeking.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshfBoekLijst 
         Height          =   2535
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   4471
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         AllowUserResizing=   1
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
End
Attribute VB_Name = "frmBoeking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBoeken_Click()

Unload Me

End Sub

Private Sub cmdNegeren_Click()

DKTRL_CUMUL = 99
Unload Me

End Sub

Private Sub Form_Load()

With mshfBoekLijst
    .TextMatrix(0, 0) = "Rekening"
    .TextMatrix(0, 1) = "Boekingsomschrijving"
    .ColWidth(1) = 2775
    .TextMatrix(0, 2) = "EUR Debet"
    .TextMatrix(0, 3) = "EUR Credit"
    .TextMatrix(0, 4) = "BEF Debet"
    .TextMatrix(0, 5) = "BEF Credit"
End With

End Sub

Private Sub SSTab1_GotFocus()

If DKTRL_EUR <> 0 Or DKTRL_BEF <> 0 Then
    If mshfBoekLijst.TextMatrix(mshfBoekLijst.Rows - 1, 0) = "" Then
    
        Dim TempoRij
        TempoRij = mshfBoekLijst.Rows - 1
        'mshfBoekLijst.TextMatrix(TempoRij, 0) = "Ter info"
        mshfBoekLijst.TextMatrix(TempoRij, 1) = "Ter info BEF Rekenenverschil"
        If DKTRL_BEF Then
            If DKTRL_BEF < 0 Then
                mshfBoekLijst.TextMatrix(TempoRij, 4) = Format(-DKTRL_BEF, "#,##0.00")
            Else
                mshfBoekLijst.TextMatrix(TempoRij, 5) = Format(DKTRL_BEF, "#,##0.00")
            End If
    
        Else
            If DKTRL_EUR < 0 Then
                mshfBoekLijst.TextMatrix(TempoRij, 2) = Format(-DKTRL_EUR, "#,##0.00")
            Else
                mshfBoekLijst.TextMatrix(TempoRij, 3) = Format(DKTRL_EUR, "#,##0.00")
            End If
        End If
    End If
    
End If

End Sub
