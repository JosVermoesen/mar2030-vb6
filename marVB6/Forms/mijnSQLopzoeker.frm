VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form mijnSQLOpzoeker 
   Caption         =   "OLEDB SQL Opvraagscherm"
   ClientHeight    =   2730
   ClientLeft      =   480
   ClientTop       =   3810
   ClientWidth     =   6285
   HelpContextID   =   11500
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfSQL 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   4048
      _Version        =   393216
      FixedCols       =   0
      WordWrap        =   -1  'True
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtSQL 
      Height          =   735
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   660
      Visible         =   0   'False
      Width           =   6105
   End
   Begin VB.CommandButton cmdSluiten 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   300
      Left            =   4140
      TabIndex        =   2
      Top             =   1200
      Width           =   1440
   End
   Begin VB.CommandButton cmdSQL 
      Caption         =   "SQL SELECT"
      Default         =   -1  'True
      Height          =   300
      Left            =   1260
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1500
      Width           =   1230
   End
   Begin MSComctlLib.ImageList imlTreePics 
      Left            =   2640
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mijnSQLopzoeker.frx":0000
            Key             =   "Tabel"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mijnSQLopzoeker.frx":0112
            Key             =   "Opzoeking"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mijnSQLopzoeker.frx":0224
            Key             =   "Index"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mijnSQLopzoeker.frx":0336
            Key             =   "Eigenschap"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mijnSQLopzoeker.frx":0448
            Key             =   "Verbonden"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mijnSQLopzoeker.frx":055A
            Key             =   "Veld"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblRecordCount 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   1425
   End
End
Attribute VB_Name = "mijnSQLOpzoeker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim grdColWidth(20) As Integer

Private datPrimaryRS As ADODB.Recordset

Private Sub cmdSluiten_Click()

Unload Me

End Sub

Private Sub cmdSQL_Click()

On Error Resume Next
Screen.MousePointer = vbHourglass
' create a ADO recordset using the bedrijfs collection
    Set datPrimaryRS = New ADODB.Recordset
    datPrimaryRS.CursorLocation = adUseClient
    On Error Resume Next
    datPrimaryRS.Open txtSQL.Text, adntDB 'is toch optioneel: , adOpenForwardOnly, adLockReadOnly
    If Err Then
        MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
        Set msfSQL.DataSource = Nothing
        msfSQL.Refresh
    Else
        Set msfSQL.DataSource = datPrimaryRS
    End If
    lblRecordCount = datPrimaryRS.RecordCount
    
For COUNT_TO = 0 To msfSQL.Cols - 1
    If grdColWidth(COUNT_TO) = 0 Then
        Exit For
    Else
        msfSQL.ColWidth(COUNT_TO) = grdColWidth(COUNT_TO)
    End If
    msfSQL.ColAlignment(COUNT_TO) = flexAlignLeftTop
Next
Screen.MousePointer = vbNormal

End Sub


Private Sub Form_Resize()

    lblRecordCount.Top = Me.ScaleTop
    lblRecordCount.Left = Me.ScaleLeft
    
    msfSQL.Top = Me.ScaleTop + lblRecordCount.Height
    msfSQL.Left = Me.ScaleLeft
    msfSQL.Height = Me.ScaleHeight - lblRecordCount.Height
    msfSQL.Width = Me.ScaleWidth

End Sub


Private Sub Form_Unload(Cancel As Integer)

Set datPrimaryRS = Nothing

End Sub

Private Sub msfSQL_DblClick()

If msfSQL.Top = 5 Then
    'vergroten
    Me.WindowState = vbMaximized
    msfSQL.Top = 0
    msfSQL.Left = 0
    msfSQL.Height = Me.Height - 700
    msfSQL.Width = Me.Width - 100
Else
    'terug normaal
    Me.WindowState = vbNormal
    msfSQL.Top = 5
    msfSQL.Left = 0
    msfSQL.Height = 2295
    msfSQL.Width = 7995
End If

End Sub


Private Sub txtSQL_GotFocus()

cmdSQL.Default = True

End Sub

Private Sub txtSQL_LostFocus()

cmdSQL.Default = False

End Sub


