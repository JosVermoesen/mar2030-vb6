VERSION 5.00
Begin VB.Form frmDokHistoriek 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Geactiveerde Documenten"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   2550
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstDokHistoriek 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmDokHistoriek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()

    LaadInstellingen Me
 
End Sub


Private Sub Form_Resize()

    lstDokHistoriek.Top = Me.ScaleTop
    lstDokHistoriek.Left = Me.ScaleLeft
    lstDokHistoriek.Height = Me.ScaleHeight
    lstDokHistoriek.Width = Me.ScaleWidth

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim X As Boolean

X = InstellingenBewaard(Me)

End Sub
