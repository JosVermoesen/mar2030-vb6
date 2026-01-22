VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmDataOmgeving 
   Caption         =   "DataOmgevingTest"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frmDataOmgeving.frx":0000
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   5318
      _Version        =   393216
      Cols            =   3
      WordWrap        =   -1  'True
      DataMember      =   "KlantenLopendePolissen"
      BandDisplay     =   1
      RowSizingMode   =   1
      _NumberOfBands  =   2
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0)._NumMapCols=   2
      _Band(0)._MapCol(0)._Name=   "Document"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(1)._Name=   "Klant"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(1).BandIndent=   1
      _Band(1).Cols   =   2
      _Band(1).GridLinesBand=   1
      _Band(1).TextStyleBand=   0
      _Band(1).TextStyleHeader=   3
      _Band(1).ColHeader=   1
   End
End
Attribute VB_Name = "frmDataOmgeving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

MSHFlexGrid1.CollapseAll

End Sub

Private Sub Form_Resize()

 MSHFlexGrid1.Top = Me.ScaleTop
 MSHFlexGrid1.Left = Me.ScaleLeft
 MSHFlexGrid1.Height = Me.ScaleHeight
 MSHFlexGrid1.Width = Me.ScaleWidth
 
End Sub

