VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0E400164-0685-434B-B5FF-CE76FE37CF14}#6.10#0"; "vpectrl61.ocx"
Begin VB.Form PreviewVPE 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   4680
      TabIndex        =   1
      Top             =   630
      Width           =   4680
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   1111
      ButtonWidth     =   1429
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Zoom In"
            Key             =   "cmdZoomIn"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Zoom Out"
            Key             =   "cmdZoomout"
         EndProperty
      EndProperty
   End
   Begin VPECTRLLib.VPE VPE 
      Height          =   1155
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3495
      _Version        =   393226
      Appearance      =   0
      Caption         =   "VPE Preview"
      Enabled         =   -1  'True
      PageWidth       =   21
      PageHeight      =   29.7
      SwapFileName    =   ""
      RulersMeasure   =   0
      ExternalWindow  =   -1  'True
      GridMode        =   0
      tbGrid          =   0   'False
      Rulers          =   -1  'True
      ToolBar         =   -1  'True
      tbClose         =   -1  'True
      tbScale         =   -1  'True
      StatusBar       =   -1  'True
      tbOpen          =   -1  'True
      tbSave          =   -1  'True
      tbPrint         =   -1  'True
      tbHelp          =   -1  'True
      tbAbout         =   -1  'True
      tbNavigate      =   -1  'True
      GridVisible     =   0   'False
      PageFormat      =   -1
      PaperView       =   -1  'True
      PageScroller    =   -1  'True
      StatusSegment   =   -1  'True
      EnableURLs      =   -1  'True
      PageScrollerTracking=   -1  'True
      PreviewCtrl     =   1
      DocFileReadOnly =   0   'False
      EnableHelpRouting=   0   'False
      EnablePrintSetupDialog=   -1  'True
      tbMail          =   -1  'True
      PageOrientation =   1
      PictureCacheSize=   65536
   End
End
Attribute VB_Name = "PreviewVPE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ZOOM_STEP As Integer = 25
Private CurrentZoom As Integer

Private Sub Form_Load()

    ' Initial zoom
    CurrentZoom = 100

    ' Resize preview container
    ResizePreviewContainer

    ' Initialize VPE
    VPE.Device = "VPE-Preview"
    Set VPE.PreviewControl = picPreview.hwnd
    VPE.PreviewZoom = CurrentZoom

    ' Generate the document
    Call GenerateReport

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    ResizePreviewContainer
End Sub

Private Sub ResizePreviewContainer()
    If Me.WindowState <> vbMinimized Then
        picPreview.Move 0, 500, Me.ScaleWidth, Me.ScaleHeight - 500
    End If
End Sub

Private Sub cmdZoomIn_Click()
    CurrentZoom = CurrentZoom + ZOOM_STEP
    If CurrentZoom > 400 Then CurrentZoom = 400
    VPE.PreviewZoom = CurrentZoom
End Sub

Private Sub cmdZoomOut_Click()
    CurrentZoom = CurrentZoom - ZOOM_STEP
    If CurrentZoom < 25 Then CurrentZoom = 25
    VPE.PreviewZoom = CurrentZoom
End Sub

Private Sub GenerateReport()

    ' Example VPE output
    VPE.StartDoc

    VPE.FontName = "Arial"
    VPE.FontSize = 14
    VPE.DrawText 1000, 1000, "VPE Embedded Preview Example"

    VPE.FontSize = 10
    VPE.DrawText 1000, 1500, "This preview is inside a VB6 PictureBox."

    VPE.EndDoc

End Sub

