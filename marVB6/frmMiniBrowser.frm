VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmMiniBrowser 
   ClientHeight    =   4395
   ClientLeft      =   3060
   ClientTop       =   3345
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboAddress 
      Height          =   315
      Left            =   480
      TabIndex        =   2
      Text            =   "¯¯END!"
      Top             =   900
      Visible         =   0   'False
      Width           =   3795
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   3795
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5400
      ExtentX         =   9525
      ExtentY         =   6694
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   2700
      Top             =   1500
   End
   Begin VB.CommandButton cmdSluiten 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   375
      Left            =   660
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "frmMiniBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public StartingAddress As String
Dim mbDontNavigateNow As Boolean

Private Sub cmdSluiten_Click()

    Msg = "Internet ExplorerVenster Sluiten.  Bent U zeker"
    If MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
        Unload Me
    End If
    
End Sub

Private Sub cmdSluiten_GotFocus()

    On Error Resume Next
    brwWebBrowser.SetFocus
    
End Sub

Private Sub Form_Load()
    On Error Resume Next
    'Me.Show
    'tbToolBar.Refresh
    'Form_Resize

    'cboAddress.Move 50, lblAddress.Top + lblAddress.Height + 15

    If Len(StartingAddress) > 0 Then
        cboAddress.Text = StartingAddress
        cboAddress.AddItem cboAddress.Text
        'try to navigate to the starting address
        timTimer.Enabled = True
        brwWebBrowser.Navigate StartingAddress
    End If

End Sub



Private Sub brwWebBrowser_DownloadComplete()
    On Error Resume Next
    'Me.Caption = brwWebBrowser.LocationName
    brwWebBrowser.SetFocus
    
End Sub

Private Sub brwWebBrowser_NavigateComplete(ByVal URL As String)
    Dim i As Integer
    Dim bFound As Boolean
    'Me.Caption = brwWebBrowser.LocationName
       
    For i = 0 To cboAddress.ListCount - 1
        If cboAddress.List(i) = brwWebBrowser.LocationURL Then
            bFound = True
            Exit For
        End If
    Next i
    mbDontNavigateNow = True
    If bFound Then
        cboAddress.RemoveItem i
    End If
    cboAddress.AddItem brwWebBrowser.LocationURL, 0
    cboAddress.ListIndex = 0
    mbDontNavigateNow = False
End Sub

Private Sub cboAddress_Click()
    If mbDontNavigateNow Then Exit Sub
    timTimer.Enabled = True
    brwWebBrowser.Navigate cboAddress.Text
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cboAddress_Click
    End If
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    cboAddress.Width = Me.ScaleWidth - 100
    brwWebBrowser.Width = Me.ScaleWidth - 100
    brwWebBrowser.Height = Me.ScaleHeight - 500
    
End Sub

Private Sub timTimer_Timer()
    If brwWebBrowser.Busy = False Then
        timTimer.Enabled = False
        'Me.Caption = brwWebBrowser.LocationName
    Else
        'Me.Caption = "Bezig..."
    End If
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As Button)
    On Error Resume Next
     
    timTimer.Enabled = True
     
    Select Case Button.Key
        Case "Vorige"
            brwWebBrowser.GoBack
        Case "Volgende"
            brwWebBrowser.GoForward
        Case "Vernieuwen"
            brwWebBrowser.Refresh
        Case "StartPagina"
            brwWebBrowser.GoHome
        Case "Stop"
            timTimer.Enabled = False
            brwWebBrowser.Stop
            'Me.Caption = brwWebBrowser.LocationName
        Case Else
            MsgBox "stop"
    End Select

End Sub

