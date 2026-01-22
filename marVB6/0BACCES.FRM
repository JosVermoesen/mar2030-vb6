VERSION 5.00
Begin VB.Form WelkeModus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Btrieve/MS Access instellingen"
   ClientHeight    =   1185
   ClientLeft      =   2760
   ClientTop       =   2415
   ClientWidth     =   4095
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1185
   ScaleWidth      =   4095
   Begin VB.CommandButton CmdSluiten 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   345
      Left            =   1020
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   720
      Width           =   1965
   End
   Begin VB.OptionButton OptModus 
      Caption         =   "&Ms-Access gebruiken"
      Height          =   195
      Index           =   1
      Left            =   750
      TabIndex        =   1
      Top             =   360
      Value           =   -1  'True
      Width           =   2685
   End
   Begin VB.OptionButton OptModus 
      Caption         =   "&Btrieve gebruiken"
      Height          =   195
      Index           =   0
      Left            =   750
      TabIndex        =   0
      Top             =   90
      Width           =   2685
   End
End
Attribute VB_Name = "WelkeModus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSluiten_Click()

If OptModus(0) Then
    BAModus = 0
ElseIf OptModus(1) Then
    BAModus = 1
Else
    Stop
End If
Unload Me

End Sub







