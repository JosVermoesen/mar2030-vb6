Attribute VB_Name = "modScripting"
Option Explicit

Private Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long

Function GenerateVpeDoc(location As String, docName As String) As Boolean

    GenerateVpeDoc = False

    Dim FlFree As Integer
    Dim Teller As Integer
    Dim pdfCmd As String

    Dim Ktrl As Double
    
    Dim YPOS As Double
    Dim xPos As Double
    Dim YPOS2 As Double
    Dim xPos2 As Double
    Dim pdfFontSize As Double
    Dim pdfFontName As String
    Dim pdfFontBold As Double
    Dim pdfFontItalic As Double
    Dim pdfFontUnderLine As Double
    Dim pdfColor As Double
    Dim pdfAlign As Double
    Dim pdfBoxPen As Double
        
    Dim textstring As String
    Dim texttmp As String
    Dim filename$
      
    Dim COUNT_TO As Integer
                            
VPE_ADD:
If Dir(location & docName) = "" Then
    Beep
    MsgBox location & docName & vbCrLf & "kan niet gevonden worden.", vbCritical
    MsgBox "marIntegraal probeert een standaardversie te laden. Vergeet niet te wijzigen met uw bedrijfsinfo." & vbCrLf & "De map: " & location & " wordt ter info geopend", vbExclamation
    
    If Not CopyFile(App.path & "\vpeSjbs\", location, docName) Then
        MsgBox App.path & "\vpeSjbs\" & docName & vbCrLf & "kan niet gevonden worden.", vbCritical
        MsgBox "Installeer marIntegraal opnieuw.", vbInformation
        Exit Function
    Else
        Dim r As Long
        r = ShellExecute(0, "open", location, 0, 0, 1)
        GoTo VPE_ADD
    End If
Else
    Mim.Report.PageBreak
    FlFree = FreeFile
    Open location & docName For Input As FlFree
    pdfCmd = ""
    
    Do While Not EOF(FlFree)
        Line Input #FlFree, pdfCmd
        If Left(pdfCmd, 1) = "'" Then
        Else
            Select Case Trim(UCase(pdfCmd))
                Case "CMD-WRITE"
                    GoSub VPEWRITE
                
                Case "CMD-WRITEBOX"
                    GoSub VPEWRITEBOX
                
                Case "CMD-PRINT"
                    GoSub VPEPRINT
            
                Case "CMD-PICTURE"
                    GoSub VPEPICTURE
            
                Case Else
                    MsgBox pdfCmd & " niet voorzien in deze function", vbCritical
            End Select
        End If
    Loop
    Close FlFree
    GenerateVpeDoc = True
End If

Exit Function

VPEPICTURE:
Input #FlFree, xPos, YPOS, xPos2, YPOS2, pdfBoxPen, filename$
Mim.Report.PictureBestFit = True
Mim.Report.PenSize = pdfBoxPen
If Left(filename$, 4) = "[BL]" Then
    filename$ = LOCATION_COMPANYDATA & Mid(filename$, 5)
ElseIf Left(filename$, 4) = "[PL]" Then
    filename$ = PROGRAM_LOCATION & Mid(filename$, 5)
End If
Ktrl = Mim.Report.Picture(xPos, YPOS, xPos2, YPOS2, filename$)
Return

VPEWRITE:
Input #FlFree, xPos, YPOS, xPos2, YPOS2, pdfFontSize, pdfFontName, pdfColor, pdfAlign, pdfFontBold, pdfFontItalic, pdfFontUnderLine
Line Input #FlFree, textstring
Do
    Line Input #FlFree, texttmp
    If texttmp = "CMD-ENDWRITE" Then
        Exit Do
    Else
        textstring = textstring & vbCrLf & texttmp
    End If
Loop
Mim.Report.FontName = pdfFontName
Mim.Report.FontSize = pdfFontSize
Mim.Report.TextColor = pdfColor
Mim.Report.TextBold = pdfFontBold
Mim.Report.TextItalic = pdfFontItalic
Mim.Report.TextUnderline = pdfFontUnderLine

Mim.Report.TextAlignment = pdfAlign
Ktrl = Mim.Report.Write(xPos, YPOS, xPos2, YPOS2, textstring)
Return
           
VPEWRITEBOX:
Input #FlFree, xPos, YPOS, xPos2, YPOS2, pdfFontSize, pdfFontName, pdfColor, pdfAlign, pdfFontBold, pdfFontItalic, pdfFontUnderLine
Line Input #FlFree, textstring
Do
    Line Input #FlFree, texttmp
    If texttmp = "CMD-ENDWRITE" Then
        Exit Do
    Else
        textstring = textstring & vbCrLf & texttmp
    End If
Loop
Mim.Report.FontName = pdfFontName
Mim.Report.FontSize = pdfFontSize
Mim.Report.TextColor = pdfColor
Mim.Report.TextBold = pdfFontBold
Mim.Report.TextItalic = pdfFontItalic
Mim.Report.TextUnderline = pdfFontUnderLine

Mim.Report.TextAlignment = pdfAlign
Ktrl = Mim.Report.WriteBox(xPos, YPOS, xPos2, YPOS2, textstring)
Return
           
VPEPRINT:
Input #FlFree, xPos, YPOS, pdfFontSize, pdfFontName, pdfColor, pdfFontBold, pdfFontItalic, pdfFontUnderLine
Line Input #FlFree, textstring
Mim.Report.FontName = pdfFontName
Mim.Report.FontSize = pdfFontSize
Mim.Report.TextColor = pdfColor
Mim.Report.TextBold = pdfFontBold
Mim.Report.TextItalic = pdfFontItalic
Mim.Report.TextUnderline = pdfFontUnderLine
Ktrl = Mim.Report.VPEPRINT(xPos, YPOS, textstring)
Return
      
End Function


Function ScrMaakTekstBestand(TekstZelf As String, Bestandsnaam As String) As Boolean

    On Error Resume Next
    Kill Bestandsnaam
    
    'On Error Resume Next
    
    'Dim A As TextStream
    Screen.MousePointer = vbHourglass
    'Set A = fs.CreateTextFile(Bestandsnaam, True)
    ''A.WriteLine (TekstZelf)
    'A.Write (TekstZelf)
    'A.Close
    'If Err Then
    '    ScrMaakTekstBestand = False
    'Else
    '    ScrMaakTekstBestand = True
    'End If
    'A = Nothing
    Call MarWriteUtf8File(Bestandsnaam, TekstZelf)
    Screen.MousePointer = vbNormal
    ScrMaakTekstBestand = True
    
End Function

Function ScrLeesTekstBestand(TekstZelf As String, Bestandsnaam As String) As Boolean
    
    On Error Resume Next
    'Const ForReading = 1
    'Dim A As TextStream
    
    'Set A = fs.OpenTextFile(Bestandsnaam, ForReading, False, TristateFalse)
    'If Err Then ScrLeesTekstBestand = False: Exit Function
    
    'TekstZelf = A.ReadLine
    TekstZelf = MarReadUtf8File(Bestandsnaam)
    ScrLeesTekstBestand = True
        
End Function

Function ScrLeesBestandAlleTekst(TekstZelf As String, Bestandsnaam As String) As Boolean
    
    On Error Resume Next
    'Const ForReading = 1
    'Dim A As TextStream
    
    'Set A = fs.OpenTextFile(Bestandsnaam, ForReading)
    'If Err Then ScrLeesBestandAlleTekst = False: Exit Function
    
    TekstZelf = MarReadUtf8File(Bestandsnaam)
    ScrLeesBestandAlleTekst = True
        
End Function

Function ScrAlleDrives() As String
    
    'Dim d As Drive
    'Dim dc, s, n
    
    'Set dc = fs.Drives
    'For Each d In dc
    '    s = s & d.DriveLetter & " - "
    '    On Error Resume Next
    '    Err = 0
    '    If d.DriveType = Remote Then
    '        n = d.ShareName
    '        If Err Then n = "(geen media aanwezig)"
    '    Else
    '        n = d.VolumeName
    '        If Err Then n = "(geen media aanwezig)"
    '    End If
    '    s = s & n & vbCrLf
    'Next
    
    'Set d = Nothing
    'ScrAlleDrives = s
    
'Dim fso As FileSystemObject
'Dim CDDrive As Drive

'Set fso = New FileSystemObject
'Set CDDrive = fso.GetDrive("D:")
'If CDDrive.IsReady Then
'     MsgBox CDDrive.VolumeName
'Else
'     MsgBox "Please enter a CD."
'End If

'Set CDDrive = Nothing
'Set fso = Nothing
    
End Function


Function scrFolderBestaat(Folder As String) As Boolean

    scrFolderBestaat = fs.FolderExists(Folder)
        
End Function

Function scrExcelStartDemo()

    'Stop
    'Dim xlApp As Object   ' Declare variable to hold the reference.
   
    'Set xlApp = CreateObject("excel.application")
    ' You may have to set Visible property to True
    '    if you want to see the application.
    'xlApp.Visible = True
    ' Use xlApp to access Microsoft Excel's
    ' other objects.
    'Set xlApp = Nothing   ' the application, then release the reference.
    
    ' Declare an object variable to hold the object
    ' reference. Dim as Object causes late binding.
    'Dim ExcelSheet As Object
    'Set ExcelSheet = CreateObject("Excel.Sheet")

    ' Make Excel visible through the Application object.
    'ExcelSheet.Application.Visible = True
    ' Place some text in the first cell of the sheet.
    'ExcelSheet.Cells(1, 1).Value = "Dit is kolom A, rij 1"
    ' Save the sheet to C:\test.doc directory.
    'ExcelSheet.SaveAs "C:\TEST.DOC"
    ' Close Excel with the Quit method on the Application object.
    'ExcelSheet.Application.Quit
    ' Release the object variable.
    'Set ExcelSheet = Nothing

'Dim xlApp As Excel.Application
'Dim xlBook As Excel.Workbook
'Dim xlSheet As Excel.WorkSheet
'Set xlApp = CreateObject("Excel.Application")
'Set xlBook = xlApp.Workbooks.Add'
'Set xlSheet = xlBook.Worksheets(1)

End Function


Function LineCalculating(StartWith As Variant) As Variant

Dim StartWithHere As Variant
Dim Resulting As Variant
StartWithHere = StartWith

JumpLijnRekenen:
StartWith = InputBox("Rekenen met '" & Trim(StartWithHere) & "'" & vbCrLf & "(voer in CLR de startgegevens te verwijderen)", "1-Lijn Rekenen")
If InStr(UCase(StartWith), "CLR") Then Stop: StartWithHere = "": GoTo JumpLijnRekenen
On Error Resume Next
Resulting = Mim.ScriptControl1.Eval(StartWithHere & StartWith)
If Err Then
    Select Case Err
        Case 11
            MsgBox "Deling door nul", vbExclamation
        Case 1002
            MsgBox "Schrijffout of ongeldige notering", vbExclamation

        Case Else
            MsgBox Error
    End Select
ElseIf Resulting <> "" Then
    LineCalculating = Resulting
End If

End Function

Public Function IsValidEmail(ByVal sEmail As String) As Boolean
    Dim oReg As Object
    Set oReg = CreateObject("VBScript.RegExp")

    With oReg
        .Pattern = "^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
        .IgnoreCase = True
        .Global = False
    End With

    IsValidEmail = oReg.Test(Trim$(sEmail))
End Function

