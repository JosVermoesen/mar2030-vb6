VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Afbeeldingen 
   Caption         =   "Multimedia"
   ClientHeight    =   1785
   ClientLeft      =   1350
   ClientTop       =   1560
   ClientWidth     =   3585
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
   Icon            =   "frmMMedia.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7.438
   ScaleMode       =   4  'Character
   ScaleWidth      =   29.875
   Begin MCI.MMControl mmspeler 
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   855
      Left            =   3300
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   870
      Width           =   3555
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   2955
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   2955
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   3285
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   3285
   End
   Begin VB.Menu Bestand 
      Caption         =   "&Bestand"
   End
   Begin VB.Menu MimFocus 
      Caption         =   "mar&Focus"
   End
   Begin VB.Menu Media 
      Caption         =   "&Media"
      Begin VB.Menu MediaMenu 
         Caption         =   "&Video voor Windows"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu MediaMenu 
         Caption         =   "&Geluid"
         Index           =   1
      End
      Begin VB.Menu MediaMenu 
         Caption         =   "&Midi Sequencer"
         Index           =   2
      End
      Begin VB.Menu MediaMenu 
         Caption         =   "MMM&ovie"
         Index           =   3
      End
      Begin VB.Menu MediaMenu 
         Caption         =   "&CD Muziek"
         Index           =   4
      End
      Begin VB.Menu MediaMenu 
         Caption         =   "Scanner"
         Index           =   5
      End
   End
End
Attribute VB_Name = "Afbeeldingen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MediaFilter As String
Dim MediaType As String

Dim MMTABLEDEF_ONT(200)         As String
Dim MMNu As Integer

Private Sub Bestand_Click()

On Local Error GoTo ErrorJump

mmspeler.Command = "Close"
mmspeler.Wait = True
MMNu = 0

If GridText <> "" Then
    MMNu = 1
    MMTABLEDEF_ONT(1) = GridText
    MMTABLEDEF_ONT(2) = ""
    GoTo GeenBestand
End If

' Set the "Open File..." dialog to display files.
If MediaFilter = "" Then GoTo GeenBestand
    Mim.Teken.filename = ""
    'Mim.Teken.FLAGS = cdlOFNExplorer Or cdlOFNAllowMultiselect
    Mim.Teken.Filter = MediaFilter
    Mim.Teken.CancelError = True
    Mim.Teken.ShowOpen
    MMTABLEDEF_ONT(0) = Mim.Teken.filename + " "
    
    If InStr(InStr(MMTABLEDEF_ONT(0), " ") + 1, MMTABLEDEF_ONT(0), " ") = 0 Then
        MMNu = 1
        MMTABLEDEF_ONT(1) = MMTABLEDEF_ONT(0)
    Else
        Dim InhoudsOpgave As String
        InhoudsOpgave = Left(MMTABLEDEF_ONT(0), InStr(MMTABLEDEF_ONT(0), " ") - 1)
        MMTABLEDEF_ONT(0) = Right(MMTABLEDEF_ONT(0), Len(MMTABLEDEF_ONT(0)) - InStr(MMTABLEDEF_ONT(0), " "))
        Do While MMTABLEDEF_ONT(0) <> ""
            MMTABLEDEF_ONT(MMNu + 1) = InhoudsOpgave + "\" + Left(MMTABLEDEF_ONT(0), InStr(MMTABLEDEF_ONT(0), " ") - 1)
            MMTABLEDEF_ONT(0) = Right(MMTABLEDEF_ONT(0), Len(MMTABLEDEF_ONT(0)) - InStr(MMTABLEDEF_ONT(0), " "))
            MMNu = MMNu + 1
        Loop
    End If
    
    Dim T As Integer
    For T = MMNu + 1 To 200
        MMTABLEDEF_ONT(T) = ""
    Next
    MMNu = 1

GeenBestand:
    mmspeler.DeviceType = MediaType
    mmspeler.filename = MMTABLEDEF_ONT(MMNu)
    Caption = mmspeler.filename
    mmspeler.UpdateInterval = 500
    mmspeler.Command = "Open"
    HScroll1.Enabled = True
    HScroll1.Min = 0
    HScroll1.Max = mmspeler.Length / 5
    mmspeler.Command = "Play"
    Picture2.Cls
    Picture2.Print "R&Vsoft Multi-Media speler"; vbCrLf; "Modus : "; MediaType
Exit Sub

ErrorJump:
'Resume Next
Exit Sub

End Sub

Private Sub Form_Load()

MediaFilter = "Video for windows (*.avi)|*.avi"
MediaType = "AVIVideo"

End Sub

Private Sub Form_Resize()

On Local Error GoTo ResizeError

   ' Initialize location of both pictures.
   If WindowState = 1 Then
   Else
    Picture1.Move 0, 0, ScaleWidth - VScroll1.Width, ScaleHeight - HScroll1.Height - mmspeler.Height
    Picture2.Move 0, 0
   End If

   ' Position the horizontal scroll bar.
   HScroll1.Top = Picture1.Height
   HScroll1.Left = 0
   HScroll1.Width = Picture1.Width + VScroll1.Width

   ' Position the vertical scroll bar.
   VScroll1.Top = 0
   VScroll1.Left = Picture1.Width

   VScroll1.Height = Picture1.Height

   ' Set the Max value for the scroll bars.
   HScroll1.Max = Picture2.Width - Picture1.Width
   VScroll1.Max = Picture2.Height - Picture1.Height

   ' Determine if child picture will fill up screen.
   ' If so, then there is no need to use scroll bars.

   VScroll1.Enabled = (Picture1.Height < Picture2.Height)
   HScroll1.Enabled = (Picture1.Width < Picture2.Width)
   mmspeler.Top = HScroll1.Top + HScroll1.Height
   mmspeler.Width = Picture1.Width + VScroll1.Width + 1

ResizeError:
Resume Next
Exit Sub

End Sub

Private Sub HScroll1_Change()

On Local Error Resume Next
  ' Picture2.Left is set to the negative of the value because
  ' as you scroll the scroll bar to the right, the display
  ' should move to the Left, showing more of the right

  ' of the display, and vice-versa when scrolling to the
  ' left.

   Picture2.Left = -HScroll1.Value

End Sub

Private Sub MediaMenu_Click(Index As Integer)
Dim T As Integer

Select Case Index
    Case 0
        MediaFilter = "Video for windows (*.avi)|*.avi"
        MediaType = "AVIVideo"
    Case 1
        MediaFilter = "Geluid (*.wav)|*.wav"
        MediaType = "WaveAudio"
    Case 2
        MediaFilter = "Midi Sequencer (*.mid;*.rmi)|*.mid;*.rmi"
        MediaType = "Sequencer"
    Case 3
        MediaFilter = "MMMovie speler (*.mmm)|*.mmm"
        MediaType = "MMMovie"
    Case 4
        MediaFilter = ""
        MediaType = "CDAudio"
    Case 5
        MediaFilter = ""
        MediaType = "Scanner"
        
End Select
For T = 0 To 5
    MediaMenu(T).Checked = False
Next
MediaMenu(Index).Checked = True
Bestand_Click

End Sub

Private Sub MimFocus_Click()

On Local Error Resume Next

Mim.SetFocus

End Sub

Private Sub MMSpeler_Done(NotifyCode As Integer)

GridText = ""
If MMTABLEDEF_ONT(MMNu + 1) <> "" Then
    mmspeler.Command = "Close"
    MMNu = MMNu + 1
    mmspeler.filename = MMTABLEDEF_ONT(MMNu)
    Caption = mmspeler.filename
    mmspeler.UpdateInterval = 500
    mmspeler.Command = "Open"
    HScroll1.Enabled = True
    HScroll1.Min = 0
    HScroll1.Max = mmspeler.Length / 5
    mmspeler.Command = "Play"
End If

End Sub

Private Sub MMSpeler_StatusUpdate()

On Local Error GoTo ScrollBarError

HScroll1.Value = (mmspeler.Position / 5)
Exit Sub

ScrollBarError:
HScroll1.Enabled = True
If MediaType = "CDAudio" Or MediaType = "Scanner" Then
Else
    HScroll1.Max = mmspeler.Length / 5
End If
Resume Next

End Sub

Private Sub Picture2_Click()

Form_Resize

End Sub



Private Sub VScroll1_Change()
  ' Picture2.Top is set to the negative of the value because
  ' as you scroll the scroll bar down, the display
  ' should move up, showing more of the bottom
  ' of the display, and vice-versa when scrolling up.

   Picture2.Top = -VScroll1.Value

End Sub

