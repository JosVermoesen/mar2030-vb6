VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   6570
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5055
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   8916
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"RTBPicture.frx":0000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5160
      TabIndex        =   0
      Top             =   5280
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) As Long
Private Const WM_PASTE = &H302

Private Sub Command1_Click()

    InsertPictureInRichTextBox RichTextBox1, LoadPicture("c:\mijn documenten\karte.gif")
    
End Sub


Sub InsertPictureInRichTextBox(RTB As RichTextBox, Picture As StdPicture)
    ' copy into the clipboard
    ' Copy the picture into the clipboard.
    Clipboard.Clear
    Clipboard.SetData Picture
    ' paste into the RichTextBox control
    SendMessage RTB.hwnd, WM_PASTE, 0, 0
End Sub






