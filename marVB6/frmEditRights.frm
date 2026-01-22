VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditRights 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5760
   Icon            =   "frmEditRights.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAction 
      Cancel          =   -1  'True
      Caption         =   "&Annuleren"
      Height          =   300
      Index           =   1
      Left            =   4440
      TabIndex        =   5
      Top             =   540
      Width           =   1215
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Bijwerken"
      Default         =   -1  'True
      Height          =   300
      Index           =   0
      Left            =   4440
      TabIndex        =   4
      Top             =   180
      Width           =   1215
   End
   Begin VB.TextBox txtRights 
      Height          =   300
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   540
      Width           =   360
   End
   Begin VB.TextBox txtRights 
      BackColor       =   &H80000004&
      Height          =   300
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   180
      Width           =   2880
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
   Begin VB.Label lblRights 
      Caption         =   "Niveau"
      Height          =   240
      Index           =   1
      Left            =   600
      TabIndex        =   3
      Top             =   600
      Width           =   720
   End
   Begin VB.Label lblRights 
      Caption         =   "Object"
      Height          =   300
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Width           =   1200
   End
End
Attribute VB_Name = "frmEditRights"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' shared vars
Public UserID As String
Public Object As String
Public Level As Integer
Public UpdateRec As Boolean
Public Action As FormAction

Private Sub cmdAction_Click(Index As Integer)
    '
    Select Case Index
        Case 0 ' update
            Level = txtRights(1).Text
            Object = txtRights(0).Text
            UpdateRec = True
        Case 1 ' cancel
            UpdateRec = False
    End Select
    '
    Me.Hide
    '
End Sub

Private Sub Form_Activate()

    Select Case Action
        Case faAdd
            Me.Caption = "Nieuwe rechten toevoegen"
        Case faEdit
            Me.Caption = "Rechten voor [" & UserID & "]"
        Case faDelete
            Me.Caption = "Bestaande gebruiksrecht verwijderen"
    End Select
    '
    txtRights(0).Text = Object
    If Action <> faAdd Then
        txtRights(0).Locked = True
        txtRights(0).BackColor = vbYellow
    Else
        txtRights(0).Locked = False
        txtRights(0).BackColor = vbWhite
    End If
    '
    If Action <> faDelete Then
        txtRights(1).Text = Level
        txtRights(1).SetFocus
    Else
        txtRights(1).Visible = False
    End If
    '
    UpdateRec = False
    '
End Sub
