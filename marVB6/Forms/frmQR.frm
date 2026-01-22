VERSION 5.00
Begin VB.Form QR 
   Caption         =   "Form1"
   ClientHeight    =   8745
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   9525
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   9525
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   6000
      ScaleHeight     =   915
      ScaleWidth      =   1155
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cbPictureData 
      Caption         =   "PictureData"
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cbSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   6000
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Square"
      Height          =   192
      Left            =   6000
      TabIndex        =   1
      Top             =   252
      Width           =   936
   End
   Begin VB.TextBox Text1 
      Height          =   2610
      Left            =   252
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmQR.frx":0000
      Top             =   168
      Width           =   5610
   End
   Begin VB.Image Image1 
      Height          =   3090
      Left            =   240
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   3810
   End
End
Attribute VB_Name = "QR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' QR Code generator library (VB6/VBA)
'
' Copyright (c) Project Nayuki. (MIT License)
' https://www.nayuki.io/page/qr-code-generator-library
'
' Copyright (c) wqweto@gmail.com (MIT License)
'
'=========================================================================
Option Explicit
DefObj A-Z

Private Sub cbPictureData_Click()

    'Set Picture1.Picture = QRCodegenBarcode(Text1.Text)
    Clipboard.Clear
    Const CF_BITMAP = 2   ' Define bitmap format.
    Clipboard.SetData QRCodegenBarcode(Text1.Text)
    Picture1.Picture = Clipboard.GetData()
    'Picture1.Picture = Image1.Picture
    
    
    
    SavePicture Picture1, "anothertry.bmp"
    

End Sub

Private Sub cbSave_Click()

    Const CF_BITMAP = 2   ' Define bitmap format.
    SavePicture Image1.Picture, "filename.bmp"


End Sub

Private Sub Form_Click()
   ' Declare variables.
   Dim CX As Integer
   Dim CY As Integer
   Dim Limit As Integer
   Dim Radius   As Integer
   Dim Msg As String
   
   ScaleMode = vbPixels   ' Set scale to pixels.
   AutoRedraw = True ' Turn on AutoRedraw.
   Width = Height   ' Change width to match height.
   CX = ScaleWidth / 2   ' Set X position.
   CY = ScaleHeight / 2   ' Set Y position.
   'Limit = CX   ' Limit size of circles.
   'For Radius = 0 To Limit   ' Set radius.
   '   Circle (CX, CY), Radius, RGB(Rnd * 255, Rnd * 255, Rnd * 255)
   '   DoEvents   ' Yield for other processing.
   'Next Radius
   Me.Picture = Image1.Picture
        
   'Msg = "Choose OK to save the graphics from this form "
   'Msg = Msg & "to a bitmap file."
   'MsgBox Msg
   SavePicture Image, "TEST.BMP"   ' Save picture to file.
   
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 67 And Shift = vbCtrlMask Then
        Clipboard.Clear
        Const CF_BITMAP = 2   ' Define bitmap format.
        
        Clipboard.SetData Image1.Picture
                
    End If
End Sub

Private Sub Form_Load()

    Text1_Change
    
End Sub

Private Sub Form_Resize()
    Dim lWidth          As Long
    Dim lHeight         As Long
    
    If WindowState <> vbMinimized Then
        lWidth = ScaleWidth - Image1.Left - Image1.Left
        lHeight = ScaleHeight - Image1.Top - Image1.Left
        If lWidth > lHeight Then
            lWidth = lHeight
        End If
        Image1.Width = lWidth
        Image1.Height = lWidth
    End If
End Sub

Private Sub Text1_Change()
    Set Image1.Picture = QRCodegenBarcode(Text1.Text, SquareModules:=(Check1.Value = vbChecked))
End Sub

Private Sub Check1_Click()
    Text1_Change
End Sub
