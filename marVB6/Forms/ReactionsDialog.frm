VERSION 5.00
Begin VB.Form FormReactionsDialog 
   Caption         =   "Reacties"
   ClientHeight    =   3615
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   6405
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   6405
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TextBoxReactions 
      Height          =   2895
      Left            =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "ReactionsDialog.frx":0000
      Top             =   100
      Width           =   6200
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   3100
      Width           =   1215
   End
End
Attribute VB_Name = "FormReactionsDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()

    Unload Me
    
End Sub


Private Sub Form_Resize()
    
    On Local Error Resume Next
    Me.TextBoxReactions.Width = Me.Width - 300
    Me.TextBoxReactions.Height = Me.Height - 1200
    Me.CancelButton.Top = Me.Height - 1000

End Sub

