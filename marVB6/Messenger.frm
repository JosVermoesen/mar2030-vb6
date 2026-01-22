VERSION 5.00
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
   Begin VB.CommandButton mnuMyContacts 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1380
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton btnOpenInbox 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public WithEvents MsgrUIA As MessengerAPI.Messenger
Attribute MsgrUIA.VB_VarHelpID = -1

Private Sub btnOpenInbox_Click()

    Set MsgrUIA = New MessengerAPI.Messenger
    On Error Resume Next
    MsgrUIA.OpenInbox
    ErrorTrap ("IMessenger.OpenInbox")  'Error handling routine
End Sub


Private Sub mnuMyContacts_Click()
    'On Error Resume Next
    Set MsgrUIA = New MessengerAPI.Messenger
    Set MsgrContacts = Nothing
    Set MsgrContacts = MsgrUIA.MyContacts
    ' ("IMessenger.MyContacts") 'Error handling routine
    Stop
    
End Sub

