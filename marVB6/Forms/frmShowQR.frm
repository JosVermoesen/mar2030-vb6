VERSION 5.00
Begin VB.Form ShowAndSaveQR 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "ShowAndSaveQR"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   Icon            =   "frmShowQR.frx":0000
   LinkTopic       =   "FormQR"
   ScaleHeight     =   4335
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2610
      Left            =   6240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmShowQR.frx":000C
      Top             =   120
      Visible         =   0   'False
      Width           =   5610
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Square"
      Height          =   192
      Left            =   12000
      TabIndex        =   1
      Top             =   120
      Width           =   936
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   6840
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   3000
   End
End
Attribute VB_Name = "ShowAndSaveQR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
   
   ' Declare variables.
   'Dim CX As Integer
   'Dim CY As Integer
   'Dim Limit As Integer
   'Dim Radius   As Integer
   'Dim Msg As String
   
   'ScaleMode = vbPixels   ' Set scale to pixels.
   'AutoRedraw = True ' Turn on AutoRedraw.
   'Width = Height   ' Change width to match height.
   'CX = ScaleWidth / 2   ' Set X position.
   'CY = ScaleHeight / 2   ' Set Y position.
   'Limit = CX   ' Limit size of circles.
   'For Radius = 0 To Limit   ' Set radius.
   '   Circle (CX, CY), Radius, RGB(Rnd * 255, Rnd * 255, Rnd * 255)
   '   DoEvents   ' Yield for other processing.
   'Next Radius
   'Me.Picture = Image1.Picture
        
   'Msg = "Choose OK to save the graphics from this form "
   'Msg = Msg & "to a bitmap file."
   'MsgBox Msg
   'SavePicture Image, "QR.BMP"   ' Save picture to file.
   'Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

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
      
   Dim myFile
      
   myFile = Dir(LOCATION_COMPANYDATA + "BMP-qr\" + FileNameQR + ".bmp")
   
   If myFile = "" Then
   Else
        Kill LOCATION_COMPANYDATA + "BMP-qr\" + myFile
        DoEvents
   End If
      
   SavePicture Image, LOCATION_COMPANYDATA + "BMP-qr\" + FileNameQR + ".bmp"  ' Save picture to file.
       
End Sub

Private Sub Text1_Change()

    Set Image1.Picture = QRCodegenBarcode(Text1.Text, SquareModules:=(Check1.Value = vbChecked))
            
End Sub
