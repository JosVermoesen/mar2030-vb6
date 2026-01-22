VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form mijnSQLEditor 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "OLEDB SQL Opvraag- en Opbouwgenerator"
   ClientHeight    =   1125
   ClientLeft      =   195
   ClientTop       =   435
   ClientWidth     =   9330
   HelpContextID   =   11500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSQL 
      Height          =   735
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   11445
   End
   Begin MSDataGridLib.DataGrid msfSQL 
      Height          =   795
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   1402
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2067
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2067
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlTreePics 
      Left            =   0
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
            Picture         =   "mijnSQLEditor.frx":0000
            Key             =   "Tabel"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mijnSQLEditor.frx":0112
            Key             =   "Opzoeking"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mijnSQLEditor.frx":0224
            Key             =   "Index"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mijnSQLEditor.frx":0336
            Key             =   "Eigenschap"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mijnSQLEditor.frx":0448
            Key             =   "Verbonden"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mijnSQLEditor.frx":055A
            Key             =   "Veld"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSluiten 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   300
      Left            =   4140
      TabIndex        =   1
      Top             =   120
      Width           =   1440
   End
   Begin VB.CommandButton cmdSQL 
      Caption         =   "SQL SELECT"
      Default         =   -1  'True
      Height          =   300
      Left            =   2100
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   240
      Width           =   1230
   End
End
Attribute VB_Name = "mijnSQLEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim datPrimaryRS As ADODB.Recordset
Dim rvLockMode As Integer

Private Sub cmdSluiten_Click()

    Unload Me

End Sub

Private Sub cmdSQL_Click()

    On Error Resume Next
    Screen.MousePointer = vbHourglass
    ' create a ADO recordset using the bedrijfs collection
    Set datPrimaryRS = New ADODB.Recordset
    datPrimaryRS.CursorLocation = adUseClient
     
    If msfSQL.AllowUpdate = True Then
        rvLockMode = adLockOptimistic
    Else
        rvLockMode = adLockReadOnly
    End If
    datPrimaryRS.Open txtSQL.Text, adntDB, adOpenDynamic, rvLockMode
    If Err Then
        MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
        Set msfSQL.DataSource = Nothing
        msfSQL.Refresh
    Else
        Set msfSQL.DataSource = datPrimaryRS
    End If
    Screen.MousePointer = vbNormal

End Sub


Private Sub Form_Resize()

    msfSQL.Top = Me.ScaleTop
    msfSQL.Left = Me.ScaleLeft
    msfSQL.Height = Me.ScaleHeight
    msfSQL.Width = Me.ScaleWidth

End Sub

Private Sub Form_Unload(Cancel As Integer)

Set datPrimaryRS = Nothing

End Sub

