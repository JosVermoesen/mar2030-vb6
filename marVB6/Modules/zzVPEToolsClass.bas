Attribute VB_Name = "thisVPEToolsClass"
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

Sub pdfPrintUserDef(TypeEnTaal As String, pdfOVSStrook As Double)
    
    Dim FlFree As Integer
    Dim Teller As Integer
    Dim pdfCmd As String
    Dim fileName As String
    
    Dim Ktrl As Double
    
    Dim adresYpos As Double
    Dim adresXpos As Double
    Dim adresYpos2 As Double
    Dim adresXpos2 As Double
    Dim adresBox As Boolean
    
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
      
    Dim COUNT_TO As Integer
                
If Dir(LOCATION_COMPANYDATA + "vpeSjbs\pdfDDEF" + TypeEnTaal + ".Txt") = "" Then
    Exit Sub
Else
    With Mim.Report
        .nTopMargin = 1
        .nLeftMargin = 0.5
        .nRightMargin = 20.8
        .nBottomMargin = 29.8
    End With
    FlFree = FreeFile
    Open LOCATION_COMPANYDATA + "vpeSjbs\pdfDDEF" + TypeEnTaal + ".Txt" For Input As FlFree
    
    'Dim pdfCmd As String
    
    Do While Not EOF(FlFree)
        Line Input #FlFree, pdfCmd
        If Left(pdfCmd, 1) = "'" Then
        Else
            Select Case Trim(UCase(pdfCmd))
                Case "CMD-VSOFTSPACE"
                    GoSub CMDVSOFTSPACE
                
                Case "CMD-ADRESSPACE"
                    GoSub CMDADRESSPACE
                
                Case "CMD-WRITE"
                    GoSub CMDWRITE
                
                Case "CMD-WRITEBOX"
                    GoSub CMDWRITEBOX
                
                Case "CMD-PRINT"
                    GoSub CMDPRINT
                
                Case "CMD-PICTURE"
                    GoSub CMDPICTURE
              Case Else
                    MsgBox pdfCmd & " nog niet voorzien", vbCritical
            End Select
        End If
    Loop
    Close FlFree
End If
Exit Sub

CMDPICTURE:
Input #FlFree, xPos, YPOS, xPos2, YPOS2, pdfBoxPen, fileName
Mim.Report.PenSize = pdfBoxPen
Mim.Report.PictureBestFit = True
If Left(fileName$, 4) = "[BL]" Then
    fileName$ = LOCATION_COMPANYDATA & Mid(fileName$, 5)
ElseIf Left(fileName$, 4) = "[PL]" Then
    fileName$ = PROGRAM_LOCATION & Mid(fileName$, 5)
End If
'2.5, 0.5, 19, 2.5
Ktrl = Mim.Report.Picture(xPos, YPOS, xPos2, YPOS2, fileName)
'.VpePrint(xPos, YPOS, textstring)
Return

CMDVSOFTSPACE:
Input #FlFree, pdfVsoftVanaf, pdfVsoftTot
Return

CMDADRESSPACE:
Input #FlFree, pdfadresXpos, pdfadresYpos, pdfadresXpos2, pdfadresYpos2 ', adresBox
'Ktrl = Mim.Report.WriteBox(adresXpos, adresYpos, adresXpos2, adresYpos2, "")
Return

CMDWRITE:
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
           
CMDWRITEBOX:
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
           
CMDPRINT:
Input #FlFree, xPos, YPOS, pdfFontSize, pdfFontName, pdfColor, pdfFontBold, pdfFontItalic, pdfFontUnderLine
Line Input #FlFree, textstring
Mim.Report.FontName = pdfFontName
Mim.Report.FontSize = pdfFontSize
Mim.Report.TextColor = pdfColor
Mim.Report.TextBold = pdfFontBold
Mim.Report.TextItalic = pdfFontItalic
Mim.Report.TextUnderline = pdfFontUnderLine
If YPOS > pdfVsoftTot Then
    Ktrl = Mim.Report.VPEPRINT(xPos, YPOS - pdfOVSStrook, textstring)
Else
    Ktrl = Mim.Report.VPEPRINT(xPos, YPOS, textstring)
End If
Return

End Sub



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
    Dim fileName$
      
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
Input #FlFree, xPos, YPOS, xPos2, YPOS2, pdfBoxPen, fileName$
Mim.Report.PictureBestFit = True
Mim.Report.PenSize = pdfBoxPen
If Left(fileName$, 4) = "[BL]" Then
    fileName$ = LOCATION_COMPANYDATA & Mid(fileName$, 5)
ElseIf Left(fileName$, 4) = "[PL]" Then
    fileName$ = PROGRAM_LOCATION & Mid(fileName$, 5)
End If
Ktrl = Mim.Report.Picture(xPos, YPOS, xPos2, YPOS2, fileName$)
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



