VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   1380
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   873
      _Version        =   393216
      Silent          =   -1  'True
      DeviceType      =   "Scanner"
      FileName        =   "PipoDeKloo"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

MMControl1.DeviceType = "Scanner"
MMControl1.Command = "0pen"

Stop

End Sub
