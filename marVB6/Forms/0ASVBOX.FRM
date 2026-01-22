VERSION 5.00
Begin VB.Form KeuzeVSF 
   Caption         =   "KeuzeLijst"
   ClientHeight    =   2055
   ClientLeft      =   5505
   ClientTop       =   4440
   ClientWidth     =   4065
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2055
   ScaleWidth      =   4065
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox NTBoxLijst 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4035
   End
End
Attribute VB_Name = "KeuzeVSF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit


Private Sub Form_Activate()
Dim X As String
Dim ZoekTekst As String
Dim Aantallijnen As Integer
Dim Keuze As Integer
Dim OptieTxt As String

OptieTxt = GridText
NTBoxLijst.Clear
ZoekTekst = "NTKB2"
If aIndex >= 1000 Then
    ZoekTekst = ZoekTekst + Format(aIndex - 1000, "000")
Else
    ZoekTekst = ZoekTekst + "9" + Format(aIndex, "00")
End If
Keuze = -1
X = ZoekEnPlaats(KeuzeVSF.NTBoxLijst, ZoekTekst, Aantallijnen, Keuze, GridText)

If Keuze = -1 Then
Else
    NTBoxLijst.ListIndex = Keuze
End If

End Sub

Private Sub Form_Resize()

On Local Error Resume Next
NTBoxLijst.Width = KeuzeVSF.Width - 100
NTBoxLijst.Height = KeuzeVSF.Height - 390

End Sub

Private Sub NTBoxLijst_DblClick()

NTBoxLijst_KeyPress 13

End Sub

Private Sub NTBoxLijst_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case 17, 27
        NTBoxLijst_KeyPress 27
End Select

End Sub

Private Sub NTBoxLijst_KeyPress(KeyAscii As Integer)

On Local Error Resume Next

Select Case KeyAscii
    Case 13
        GridText = NTBoxLijst.Text
        Unload Me
    Case 27
        Unload Me
    Case Else
End Select

End Sub

