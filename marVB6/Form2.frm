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
   Begin VB.CommandButton cmdStartScanning 
      Caption         =   "&Start Scanner"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   1395
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStartScanning_Click()

Dim marScanner As ImgScan

Set marScanner = New ImgScan

marScanner.FileType = TIFF
Stop
If marScanner.ScannerAvailable Then
    marScanner.OpenScanner
    marScanner.StartScan
    If marScanner.StatusCode Then
        MsgBox "scan is mislukt.  Probeer eventueel opnieuw.", vbExclamation
    End If
End If
    
End Sub

