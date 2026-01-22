VERSION 5.00
Begin VB.Form ScanBeheer 
   Caption         =   "R&Vsoft 2003 SCANBEHEER"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   8445
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBewaar 
      Caption         =   "Bewaren"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdBestand 
      Caption         =   "Bestand"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdAfdruk 
      Caption         =   "Afdrukken"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Software Scan Starten"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdScanOpties 
      Caption         =   "Scan Opties"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton cmdScanNieuweAfbeelding 
      Caption         =   "Afbeelding Scannen"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton cmdTop 
      Caption         =   "Top"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton cmdBodem 
      Caption         =   "Bodem"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton cmdVorige 
      Caption         =   "<"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton cmdVolgende 
      Caption         =   ">"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   1440
      Width           =   255
   End
   Begin VB.Image imgFiguur 
      Height          =   7935
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "ScanBeheer"
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
    
    On Local Error GoTo errorEOF
    
    If rs.State Then
    Else
        fnOpenDatabase
    End If
    rs.MoveLast
    BlobToFile rs("tbobject"), App.path & "\marscan.jpg"
    On Error Resume Next
    imgFiguur.Picture = LoadPicture(App.path & "\marscan.jpg")
    If Err Then MsgBox Error
        
errorEOF:
    
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
    
    On Local Error GoTo ErrorBOF
    
    If rs.State Then
    Else
        fnOpenDatabase
    End If
    rs.MoveFirst
    BlobToFile rs("tbobject"), App.path & "\marscan.jpg"
    On Error Resume Next
    imgFiguur.Picture = LoadPicture(App.path & "\marscan.jpg")
    If Err Then MsgBox Error
    
ErrorBOF:
    
End Sub

Private Sub cmdVolgende_Click()

    cmdBewaar.Enabled = False
    
    On Local Error GoTo ErrorMOVENEXT
    
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
        
    On Local Error GoTo ErrorPREVIOUS
    
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
    
ErrorPREVIOUS:
End Sub



Private Sub Form_Load()

    Set cnn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    rs.Close
    cnn.Close
    Set rs = Nothing
    Set cnn = Nothing
        
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
    rs.Open "Afbeeldingen", cnn, adOpenDynamic, adLockOptimistic

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
    If Dest Is Printer Then
        'Dest.PaintPicture Src.Picture, 0, 0, Src.Width, Src.Height
        Printer.PaintPicture Src.Picture, 0, 0, Printer.Width - 500, Printer.Height - 800
        
        Printer.EndDoc
    Else
        Dest.PaintPicture Src.Picture, Dest.Width / 2, Dest.Height / 2
    End If
   
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



