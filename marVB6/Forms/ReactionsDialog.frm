VERSION 5.00
Begin VB.Form FormReactionsDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reacties"
   ClientHeight    =   3615
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TextBoxReactions 
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "ReactionsDialog.frx":0000
      Top             =   120
      Width           =   6135
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   3120
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


