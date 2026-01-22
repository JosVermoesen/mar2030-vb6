VERSION 5.00
Begin VB.Form frmKopiePlak 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Gekopiëerd om te plakken"
   ClientHeight    =   1710
   ClientLeft      =   570
   ClientTop       =   6195
   ClientWidth     =   8190
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstKopiePlak 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmKopiePlak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()

    LaadInstellingen Me
 
End Sub


Private Sub Form_Resize()

    lstKopiePlak.Top = Me.ScaleTop
    lstKopiePlak.Left = Me.ScaleLeft
    lstKopiePlak.Height = Me.ScaleHeight
    lstKopiePlak.Width = Me.ScaleWidth

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim x As Boolean

x = InstellingenBewaard(Me)

End Sub
