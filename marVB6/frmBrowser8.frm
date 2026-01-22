VERSION 5.00
Object = "{8E10A278-6945-4466-8DF7-0AED5BDDEFDB}#1.0#0"; "BrowserInteropUserControlLibrary1.tlb"
Begin VB.Form frmBrowser8 
   Caption         =   "Explorer"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5010
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   3630
   ScaleWidth      =   5010
   Begin BrowserInteropUserControlLibrary1Ctl.InteropUserControl interopBrowser 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Object.Visible         =   "True"
      Enabled         =   "True"
      ForegroundColor =   "-2147483630"
      BackgroundColor =   "-2147483633"
      BackColor       =   "Control"
      ForeColor       =   "ControlText"
      Location        =   "8, 8"
      Name            =   "InteropUserControl"
      Size            =   "313, 217"
      Object.TabIndex        =   "0"
   End
End
Attribute VB_Name = "frmBrowser8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()

    Me.interopBrowser.Width = Me.Width
    Me.interopBrowser.Height = Me.Height

End Sub
