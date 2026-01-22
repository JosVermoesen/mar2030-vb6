VERSION 5.00
Object = "{3D588416-8F31-4516-9E89-12B4D80C9CFC}#1.0#0"; "eIDInteropUserControlLibrary1.tlb"
Begin VB.Form frmLoginEID 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gebruiker Aanmelding"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   4995
   Icon            =   "frmLoginEID.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin eIDInteropUserControlLibrary1Ctl.InteropUserControl InteropUserControl1 
      Height          =   2895
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   4815
      Object.Visible         =   "True"
      Enabled         =   "True"
      ForegroundColor =   "-2147483630"
      BackgroundColor =   "-2147483633"
      BackColor       =   "Control"
      ForeColor       =   "ControlText"
      Location        =   "8, 96"
      Name            =   "InteropUserControl"
      Size            =   "321, 193"
      Object.TabIndex        =   "0"
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "eID"
      Height          =   300
      Index           =   2
      Left            =   3000
      TabIndex        =   8
      Top             =   600
      Width           =   720
   End
   Begin VB.CommandButton cmdAction 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   285
      Index           =   1
      Left            =   3840
      TabIndex        =   4
      Top             =   1005
      Width           =   1080
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "marLogin"
      Default         =   -1  'True
      Height          =   300
      Index           =   0
      Left            =   3840
      TabIndex        =   3
      Top             =   600
      Width           =   1080
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   960
      Width           =   2760
   End
   Begin VB.TextBox txtUserID 
      Height          =   300
      Left            =   960
      TabIndex        =   1
      Text            =   "DEFAULT"
      Top             =   600
      Width           =   1920
   End
   Begin VB.Label lblPassword 
      Caption         =   "Paswoord"
      Height          =   300
      Left            =   0
      TabIndex        =   6
      Top             =   1020
      Width           =   840
   End
   Begin VB.Label lbluserID 
      Caption         =   "Gebruiker"
      Height          =   240
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   840
   End
   Begin VB.Label lblAppTitle 
      Alignment       =   2  'Center
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   5055
   End
End
Attribute VB_Name = "frmLoginEID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status As Boolean
Public Title As String
Public UserID As String
Public Password As String
Public NationalNumber As String


Private Sub cmdAction_Click(Index As Integer)
    
    Dim dummystring As String
    Dim dummy As String
    Dim strReturn() As String
    Dim returnlong As Long
    
    Select Case Index
        Case 0: Status = True
        Case 1: Status = False
        Case 2
            On Local Error Resume Next
            dummy$ = Me.InteropUserControl1.clickon_eid()
            Me.InteropUserControl1.Refresh
            If dummy$ <> "" Then
                strReturn = Split(dummy$, ";")
                Me.txtUserID = strReturn(0)
                Me.txtPassword = strReturn(3)
                Status = True
                returnlong = waituntilspring
            End If
            If Err Then
                MsgBox Err.Description
            End If
    End Select
    '
    Me.Hide
    '
End Sub

Private Function waituntilspring() As Long

    Dim telhier As Long
    
    For telhier = 0 To 2000000
        DoEvents
    Next
    waituntilspring = telhier
    
End Function

Private Sub Form_Activate()
    '
    lblAppTitle = Title
    If UserID <> "" Then
        txtUserID = UserID
    End If
    txtPassword = Password
    Status = False
    On Local Error Resume Next
    Me.SetFocus
    
End Sub

