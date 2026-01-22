VERSION 5.00
Object = "{A5EDEDF4-2BBC-45F3-822B-E60C278A1A79}#11.0#0"; "MDIVWCTL.DLL"
Begin VB.Form OfficeDocView 
   Caption         =   "OfficeDocView"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin MODICtl.MiDocView MiDocView1 
      Height          =   2895
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      _cx             =   5106
      _cy             =   5106
      ActionState     =   0
      DocViewMode     =   0
      FitMode         =   0
      FileName        =   ""
   End
End
Attribute VB_Name = "OfficeDocView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    TestLayout
    
End Sub

Sub TestLayout()
  
  Dim miDoc As MODICtl.Document
  Dim miLayout As MODICtl.Layout
  Dim strLayoutInfo As String
    
  ' Load an existing TIFF file.
  Set miDoc = New MODICtl.Document
  miDoc.Create "C:\mvb.mdi"
  
  ' Perform OCR.
  miDoc.Images(0).OCR
  
  ' Retrieve and summarize the OCR results.
  Set miLayout = miDoc.Images(0).Layout
  strLayoutInfo = _
   "Language: " & miLayout.Language & vbCrLf & _
    "Number of characters: " & miLayout.NumChars & vbCrLf & _
    "Number of fonts: " & miLayout.NumFonts & vbCrLf & _
    "Number of words: " & miLayout.NumWords & vbCrLf & _
    "Beginning of text: " & Left(miLayout.Text, 50) & vbCrLf & _
    "First word of text: " & miLayout.Words(0).Text
  MsgBox strLayoutInfo, vbInformation + vbOKOnly, _
    "Layout Information"
  
  Set miLayout = Nothing
  Set miDoc = Nothing

End Sub


Private Sub Form_Load()

    MiDocView1.filename = "C:\mvb.mdi"
    
End Sub
