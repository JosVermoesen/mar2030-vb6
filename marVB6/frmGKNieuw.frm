VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form GroeneKaartNieuw 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "R&Vsoft 2001 SCANBEHEER"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   7080
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   1560
      TabIndex        =   0
      Top             =   540
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Invoer"
      TabPicture(0)   =   "frmGKNieuw.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Voorbeeld"
      TabPicture(1)   =   "frmGKNieuw.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "imgFiguur"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Uitgebreid"
      TabPicture(2)   =   "frmGKNieuw.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "cmdScan"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdScanOpties"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdScanNieuweAfbeelding"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdBestand"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdBewaar"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "tbGK(2)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "tbGK(1)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "tbGK(0)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "cmdAfdruk"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "cmdTop"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "cmdVorige"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "cmdVolgende"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "cmdBodem"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).ControlCount=   13
      Begin VB.CommandButton cmdBodem 
         Caption         =   "Bodem"
         Height          =   495
         Left            =   3300
         TabIndex        =   13
         Top             =   1980
         Width           =   615
      End
      Begin VB.CommandButton cmdVolgende 
         Caption         =   ">"
         Height          =   495
         Left            =   3000
         TabIndex        =   12
         Top             =   1980
         Width           =   255
      End
      Begin VB.CommandButton cmdVorige 
         Caption         =   "<"
         Height          =   495
         Left            =   2700
         TabIndex        =   11
         Top             =   1980
         Width           =   255
      End
      Begin VB.CommandButton cmdTop 
         Caption         =   "Top"
         Height          =   495
         Left            =   2160
         TabIndex        =   10
         Top             =   1980
         Width           =   495
      End
      Begin VB.CommandButton cmdAfdruk 
         Caption         =   "Afdrukken"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   5040
         Width           =   1935
      End
      Begin VB.TextBox tbGK 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Tag             =   "tbSleutel"
         ToolTipText     =   "Code Verzekeraar"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox tbGK 
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Tag             =   "tbOmschrijving"
         ToolTipText     =   "Naam Verzekeraar"
         Top             =   3000
         Width           =   3675
      End
      Begin VB.TextBox tbGK 
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Tag             =   "tbDocRef"
         ToolTipText     =   "Referte document Verzekeraar"
         Top             =   3360
         Width           =   3675
      End
      Begin VB.CommandButton cmdBewaar 
         Caption         =   "Bewaren"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2100
         Width           =   975
      End
      Begin VB.CommandButton cmdBestand 
         Caption         =   "Bestand"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2100
         Width           =   855
      End
      Begin VB.CommandButton cmdScanNieuweAfbeelding 
         Caption         =   "Afbeelding Scannen"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   1500
         Width           =   1935
      End
      Begin VB.CommandButton cmdScanOpties 
         Caption         =   "Scan Opties"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton cmdScan 
         Caption         =   "Software Scan Starten"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   420
         Width           =   1935
      End
      Begin VB.Image imgFiguur 
         Height          =   5055
         Left            =   -74820
         Stretch         =   -1  'True
         Top             =   420
         Width           =   3675
      End
   End
End
Attribute VB_Name = "GroeneKaartNieuw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cnn As ADODB.Connection
Dim rs  As ADODB.Recordset
Dim BestandsNaamFiguur As String

Private Sub cmdAfdruk_Click()

  PrintAnywhere imgFiguur, Printer
  
End Sub


Private Sub cmdBestand_Click()
   On Error GoTo CancelError
    Mim.Teken.filename = ""
    Mim.Teken.CancelError = True
    Mim.Teken.Filter = _
        "Alle Figuurbestanden|*.bmp;*.dib;*.gif;*.jpg;*.wmf;*.emf;*.ico;*.cur" & _
        "|JPEG bestanden (*.jpg)|*.jpg" & _
        "|GIF bestanden (*.gif)|*.gif" & _
        "|BITMAP bestanden (*.bmp;*.dib)|*.bmp;*.dib" & _
        "|META bestanden (*.wmf, *.emf)|*.wmf;*.emf" & _
        "|CURSOR bestanden (*.ico, *.cur)|*.ico;*.cur"
                
    Mim.Teken.ShowOpen
    imgFiguur.Picture = LoadPicture(Mim.Teken.filename)
    cmdBewaar.Enabled = True
    cmdBewaar.SetFocus
    BestandsNaamFiguur = Mim.Teken.filename
    Exit Sub
    
CancelError:
Screen.MousePointer = vbNormal
End Sub

Private Sub CmdBewaar_Click()

    If rs.State Then
    Else
        fnOpenDatabase
    End If
    
    'BewaarAfbeelding
    rs.AddNew
    rs("tbDatum") = Now
    rs("tbOmschrijving") = "Een testafbeelding " & Now
    DoEvents
    FileToBlob rs("tbobject"), BestandsNaamFiguur
    rs.Update
    cmdBestand.SetFocus
    cmdBewaar.Enabled = False

End Sub

Private Sub cmdBodem_Click()

    cmdBewaar.Enabled = False
    
    On Error GoTo ErrorMOVELAST
    If rs.State Then
    Else
        fnOpenDatabase
    End If
    rs.MoveLast
    BlobToFile rs("tbobject"), App.path & "\marscan.jpg"
    On Error Resume Next
    imgFiguur.Picture = LoadPicture(App.path & "\marscan.jpg")
    If Err Then MsgBox Error
    
ErrorMOVELAST:
        
End Sub

Private Sub cmdScan_Click()
   
    cmdBewaar.Enabled = False

    If rs.State Then
    Else
        fnOpenDatabase
    End If
    
    Screen.MousePointer = vbHourglass
    'ImgScan1.ShowSetupBeforeScan = True
    'ImgScan1.ShowSelectScanner
        
    'If ImgScan1.StatusCode = 0 Then
        'scanner werd gekozen
    '    ImgScan1.OpenScanner
    '    If ImgScan1.StatusCode = 0 Then
    '        'scanner kan geopend worden
    '        'TIFF, JPG_File, BMP_Bitmap, AWD_MicrosoftFax
    '        '  1 ,       6 ,         3 ,                2
    '
    '        ImgScan1.FileType = JPG_File
    '        ImgScan1.Image = App.path & "\marscan.jpg"
    '        ImgScan1.PageOption = PromptToCreateNewFile
    '
    '        ImgScan1.StartScan
    '        If ImgScan1.StatusCode = 0 Then
    '            'scan met succes
    '            'opslaan en afbeelden
    '            imgFiguur.Picture = LoadPicture(ImgScan1.Image)
    '            'BewaarAfbeelding
    '            rs.AddNew
    '            rs("tbDatum") = Now
    '            rs("tbOmschrijving") = "Een testafbeelding " & Now
    '            DoEvents
    '            FileToBlob rs("tbobject"), ImgScan1.Image
    '            rs.Update
    '        End If
    '    End If
    '    ImgScan1.CloseScanner
    'End If
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub cmdScanNieuweAfbeelding_Click()

    cmdBewaar.Enabled = False
    
    'ImgScan1.ShowScanNew
    
End Sub

Private Sub cmdScanOpties_Click()

    cmdBewaar.Enabled = False

    'ImgScan1.ShowScanPreferences
    
End Sub

Private Sub cmdTop_Click()

    cmdBewaar.Enabled = False
    
    On Local Error GoTo ErrorMOVEFIRST
    If rs.State Then
    Else
        fnOpenDatabase
    End If
    rs.MoveFirst
    BlobToFile rs("tbobject"), App.path & "\marscan.jpg"
    On Error Resume Next
    imgFiguur.Picture = LoadPicture(App.path & "\marscan.jpg")
    If Err Then MsgBox Error
    
ErrorMOVEFIRST:

End Sub

Private Sub cmdVolgende_Click()

    cmdBewaar.Enabled = False
        
    On Error GoTo ErrorMOVENEXT
    If rs.State Then
    Else
        cmdBodem_Click
        Exit Sub
    End If
    rs.MoveNext
    If rs.EOF Then
        cmdBodem_Click
        Exit Sub
    Else
        Screen.MousePointer = vbHourglass
        BlobToFile rs("tbobject"), App.path & "\marscan.jpg"
        On Error Resume Next
        imgFiguur.Picture = LoadPicture(App.path & "\marscan.jpg")
        If Err Then MsgBox Error
        Screen.MousePointer = vbNormal
    End If
    
ErrorMOVENEXT:

End Sub

Private Sub cmdVorige_Click()

    cmdBewaar.Enabled = False
        
    On Local Error GoTo ErrorBOF
    If rs.State Then
    Else
        cmdTop_Click
        Exit Sub
    End If
    rs.MovePrevious
    If rs.BOF Then
        cmdTop_Click
        Exit Sub
    Else
        Screen.MousePointer = vbHourglass
        BlobToFile rs("tbobject"), App.path & "\marscan.jpg"
        On Error Resume Next
        imgFiguur.Picture = LoadPicture(App.path & "\marscan.jpg")
        
        If Err Then MsgBox Error
        Screen.MousePointer = vbNormal
    End If
    
ErrorBOF:
    
End Sub


Function fnOpenDatabase()

If Dir(BedrijfsLokatie & "scan.mdb") = "" Then
    If Not CopyFile(App.path, BedrijfsLokatie, "scan.mdb") Then
        MsgBox "Scan.mdb is noodzakelijk.", vbExclamation
        Unload Me
        Exit Function
    End If
End If

    cnn.Open adoJetProvider & _
        "Data Source=" & BedrijfsLokatie & _
        "scan.mdb;" & _
        "Persist Security Info=False"
    rs.CursorLocation = adUseClient
    rs.Open "NL_GroeneKaart", cnn, adOpenDynamic, adLockOptimistic

End Function


' Copy a BLOB field's contents to a binary file.
Function BlobToFile(fld As ADODB.Field, filename As String, _
    Optional ChunkSize As Long = 8192)
    Dim fnum As Integer, bytesLeft As Long, bytes As Long
    Dim tmp() As Byte
    
    ' Raise an error if the field doesn't support GetChunk.
    If (fld.Attributes And adFldLong) = 0 Then
        Err.Raise 1001, , "Field doesn't support the GetChunk method."
    End If
    ' Open the file;, delete it firstoverwrite it if necessary.' Delete the
    ' file if it exists already, then create a new one.
    If Dir$(filename) <> "" Then Kill filename
    
    fnum = FreeFile
    Open filename For Binary As fnum
    ' Read the field's contents, and write it the data to the file.
    bytesLeft = fld.ActualSize
    
    Do While bytesLeft
        bytes = bytesLeft
        If bytes > ChunkSize Then bytes = ChunkSize
        tmp = fld.GetChunk(bytes)
        Put #fnum, , tmp
        bytesLeft = bytesLeft - bytes
    Loop
    Close #fnum
    
End Function


' Copy a file's contents into a BLOB field.
Function FileToBlob(fld As ADODB.Field, filename As String, _
    Optional ChunkSize As Long = 8192)
    Dim fnum As Integer, bytesLeft As Long, bytes As Long
    Dim tmp() As Byte
    
    ' Raise an error if the field doesn't support GetChunk.
    If (fld.Attributes And adFldLong) = 0 Then
        Err.Raise 1001, , "Field doesn't support the GetChunk method."
    End If
    ' Open the file; raise an error if the file doesn't exist.
    If Dir$(filename) = "" Then Err.Raise 53, , "File not found"
    
    fnum = FreeFile
    Open filename For Binary As fnum
    ' Read the file in chunks, and append data to the field.
    bytesLeft = LOF(fnum)
    Do While bytesLeft
        bytes = bytesLeft
        If bytes > ChunkSize Then bytes = ChunkSize
        ReDim tmp(1 To bytes) As Byte
        Get #1, , tmp
        fld.AppendChunk tmp
        bytesLeft = bytesLeft - bytes
    Loop
    
    Close #fnum
    
End Function


Sub PrintAnywhere(Src As Object, Dest As Object)
    
    Err = 0
    On Error Resume Next
    If Dest Is Printer Then
        'Dest.PaintPicture Src.Picture, 0, 0, Src.Width, Src.Height
        Printer.PaintPicture Src.Picture, 0, 0, Printer.Width - 500, Printer.Height - 800
        
        Printer.EndDoc
    Else
        Dest.PaintPicture Src.Picture, Dest.Width / 2, Dest.Height / 2
    End If
    If Err Then MsgBox Error, vbInformation
   
End Sub

Sub ResizePictureBoxToImage(pic As PictureBox, twipWd _
  As Integer, twipHt As Integer)
 ' This code assumes that all units are in twips.  If
 ' not, you must convert it to twips before calling
 ' this routine.  This also assumes that the image
 ' was blt'ed to 0,0.
 Dim BorderHt As Integer, BorderWd As Integer
 BorderWd = pic.Width - pic.ScaleWidth
 BorderHt = pic.Height - pic.ScaleHeight
 pic.Move pic.Left, pic.Top, twipWd + BorderWd, _
   twipHt + BorderHt
End Sub



