Attribute VB_Name = "thisTB2Class"
Option Explicit


Function tb2Indent(MAPIString As String) As String

    Dim Xarray() As String
    Dim strComment As String

    Xarray = Split(MAPIString, "'")

    Dim strIndent As String
    Dim telIndent As Integer
    Dim tabIndent As Integer
    Dim telTabIndent As Integer
            
    Screen.MousePointer = vbHourglass
    strIndent = ""
    telTabIndent = 0
    For telIndent = 0 To UBound(Xarray) - 1
        'Maak eerst het aantal tabs klaar voor de volgende stringreeks
        For telTabIndent = 1 To tabIndent
            strIndent = strIndent & vbTab
        Next
        
        'TELEBIB2 extra info eventueel toe te voegen...
        If Len(Xarray(telIndent)) = 0 Then
            'Stop
        Else
            strComment = TB2commentaar(Xarray(telIndent))
            strIndent = strIndent & Xarray(telIndent) & strComment & vbCrLf
            'Aanpassen van de tabIndent voor de volgende tekenreeks
            Select Case Left(Xarray(telIndent), 3)
                Case "XGH", "XEH", "XRH"
                    'Teller verhogen
                    tabIndent = tabIndent + 1
                    
                Case "XGT", "XET", "XRT"
                    'niks doen, zie hieronder !
                                    
            End Select
                
            'Aanpassen van de tabIndent voor afsluiting tekenreeks
            Select Case Left(Xarray(telIndent + 1), 3)
                Case "XGT", "XET", "XRT"
                    tabIndent = tabIndent - 1
            End Select
    
        End If
    Next
    
    If IsXP Then
        strIndent = Replace$(strIndent, vbTab, Space$(4))
    End If
    
    Screen.MousePointer = vbNormal
    tb2Indent = strIndent

End Function

Function headComLists(strDE As String, strQualifiant As String, strOOD As String) As String

    Dim rsTB2Qualifiers As ADODB.Recordset
    
    Set rsTB2Qualifiers = New ADODB.Recordset
    rsTB2Qualifiers.CursorLocation = adUseClient
    Msg = "SELECT * From A_DE_QUALIFIANT WHERE DE = '" & strDE & "' AND Qualifiant = '" & strQualifiant & "'"
    rsTB2Qualifiers.Open Msg, adTBIB, adOpenDynamic, adLockOptimistic, adCmdText
    If rsTB2Qualifiers.RecordCount <> 1 Then
        headComLists = " - "
        Exit Function
    Else
        headComLists = rsTB2Qualifiers("Lbc-2").Value
    End If
    'Tot slot nog controleren of gebruiksdatum niet overschreden (meestal indien problemen)..
    If rsTB2Qualifiers("Datdel").Value <> "" Then
        'ok blijkbaar niet meer in gebruik !
        strOOD = rsTB2Qualifiers("Datdel").Value
    End If
    rsTB2Qualifiers.Close
            
End Function


Function subComLists(strCode As String, strValeur As String) As String

    Dim rsTB2Lists As ADODB.Recordset
    
    Set rsTB2Lists = New ADODB.Recordset
    rsTB2Lists.CursorLocation = adUseClient
    Msg = "SELECT * From VALEUR WHERE Code = '" & strCode & "' AND Valeur = '" & strValeur & "'"
    rsTB2Lists.Open Msg, adTBIB, adOpenDynamic, adLockOptimistic, adCmdText
    If rsTB2Lists.RecordCount <> 1 Then
        subComLists = ""
    Else
        subComLists = rsTB2Lists("Lbl-2").Value
    End If
    rsTB2Lists.Close
            
End Function


Function TB2commentaar(TB2code As String) As String

    TB2commentaar = ""
    
    Dim Yarray() As String
    Dim Zarray() As String
        Yarray = Split(TB2code, "+")
                
    Dim headCommentaar As String
    Dim subCommentaar As String
    Dim outOfDateSince As String
    
        headCommentaar = ""
        subCommentaar = ""
        
    Select Case Yarray(0)
        Case "GIS"
            If UBound(Yarray) > 1 Then
                subCommentaar = subComLists(Yarray(1), Yarray(2))
            End If
            'GIS= Process indicator= X021
            headCommentaar = headComLists("X021", Yarray(1), outOfDateSince)
            
        Case "IPD"
            If UBound(Yarray) > 1 Then
                Zarray = Split(Yarray(2), ":")
                subCommentaar = subComLists(Yarray(1), Zarray(0)) ' Yarray(2))
            End If
            'IPD: Segment Product=X060, dus...
            headCommentaar = headComLists("X060", Yarray(1), outOfDateSince)
            
        Case "ATT"
            If UBound(Yarray) > 1 Then
                subCommentaar = subComLists(Yarray(1), Yarray(2))
            End If
            'ATT: Attribute=X010, dus...
            headCommentaar = headComLists("X010", Yarray(1), outOfDateSince)
    
        Case "RFF"
            'RFF: Reference qualifier=X050 én subReference Details=X032, dus...
            Zarray = Split(Yarray(1), ":")
            If UBound(Zarray) > 1 Then
                subCommentaar = headComLists("X032", Zarray(2), outOfDateSince)
            End If
            headCommentaar = headComLists("X050", Zarray(0), outOfDateSince)
    
        Case "DTM"
            'DTM: Date/Time qualifier=X016 én format qualifier X018, dus...
            Zarray = Split(Yarray(1), ":")
            If UBound(Zarray) = 2 Then
                subCommentaar = "=" & headComLists("X018", Zarray(2), outOfDateSince)
            End If
            headCommentaar = headComLists("X016", Zarray(0), outOfDateSince)
        
        Case "QRS"
            'QRS: Declaration qualifier=X045 én eventuele response code, dus...
            Zarray = Split(Yarray(1), ":")
            If UBound(Yarray) = 2 Then
                subCommentaar = "=" & headComLists("X046", Yarray(2), outOfDateSince)
            End If
            headCommentaar = headComLists("X045", Zarray(0), outOfDateSince)
        
        Case "BIN"
            'BIN: Boolean indicator=X069 én eventueel boolean indicator value X070, dus...
            If UBound(Yarray) = 2 Then
                subCommentaar = headComLists("X070", Yarray(2), outOfDateSince)
            End If
            headCommentaar = headComLists("X069", Yarray(1), outOfDateSince)
        
        Case "QTY"
            'QTY: Quantity qualifier=X047, cijfer zelf,
            '     aantaldecimalen en gecodeerde maateenheid X049 dus...
            Zarray = Split(Yarray(1), ":")
            If UBound(Zarray) > 2 Then
                subCommentaar = headComLists("X049", Zarray(3), outOfDateSince)
            End If
            headCommentaar = headComLists("X047", Zarray(0), outOfDateSince)
              
        Case "COM"
            'COM:
            Zarray = Split(Yarray(1), ":")
            headCommentaar = headComLists("X013", Zarray(0), outOfDateSince)
            
        Case "ICD"
            'ICD: Waarborgen=X058
            Zarray = Split(Yarray(1), ":")
            headCommentaar = headComLists("X058", Zarray(0), outOfDateSince)
                                
        Case "MOA"
            'MOA: Monetairy Amount=X028, dus...
            Zarray = Split(Yarray(1), ":")
            If UBound(Zarray) > 1 Then
                subCommentaar = headComLists("X031", Zarray(2), outOfDateSince)
            End If
            headCommentaar = headComLists("X028", Zarray(0), outOfDateSince)
                
        Case "PCD"
            'Percentaga qualifier= X038, cijfer zelf, aantal decimalen
            Zarray = Split(Yarray(1), ":")
            headCommentaar = headComLists("X038", Zarray(0), outOfDateSince)
            
        Case "PTY"
            'PTY: Party Identification, party qualifier=X043, dus...
            headCommentaar = headComLists("X043", Yarray(1), outOfDateSince)
        
        Case "NME"
            'NME: Name qualifier=X033, dus...
            headCommentaar = headComLists("X033", Yarray(1), outOfDateSince)
        
        Case "ADR"
            'ADR: Adres qualifier=X001, dus...
            headCommentaar = headComLists("X001", Yarray(1), outOfDateSince)
        
        Case "DOC"
            'DOC: Document=X015, dus...
            headCommentaar = headComLists("X015", Yarray(1), outOfDateSince)
        
        Case "PFN"
            'PFN: Beroep
            headCommentaar = headComLists("X040", Yarray(1), outOfDateSince)
        
        Case "PER"
            'PER: Period Qualifier=X072, dus...
            headCommentaar = headComLists("X072", Yarray(1), outOfDateSince)
        
        Case "ROD"
            'ROD: Risico Object=X052
            headCommentaar = headComLists("X052", Yarray(1), outOfDateSince)
            
        Case "IFD"
            'Stop
            
            
        Case Else
            Exit Function
    End Select
    TB2commentaar = vbTab & "(" & headCommentaar & " " & subCommentaar & Trim(" " & outOfDateSince) & ")"
            
End Function

Function rodCheck(ByRef MAPIString As String) As String

        Dim rodArray() As String
        Dim telRod As Integer
        Dim rodTmp As String
        Dim rodHier As String

        rodArray = Split(MAPIString, "ROD+")

        For telRod = 1 To UBound(rodArray)
            rodHier = headComLists("X052", Left(rodArray(telRod), 3), "")
            If InStr(rodTmp, rodHier) Then
            Else
                If telRod = UBound(rodArray) Then
                    rodTmp = rodTmp + rodHier
                Else
                    rodTmp = rodTmp + rodHier + "; "
                End If
            End If
        Next
        If rodTmp <> "" Then
            rodCheck = "Verzekerd risico: " + rodTmp
        Else
            rodCheck = ""
        End If

    End Function

    Function icdCheck(ByRef MAPIString As String) As String

        Dim icdArray() As String
        Dim telIcd As Integer
        Dim icdTmp As String
        Dim icdHier As String
        Dim icdDouble As String

        icdArray = Split(MAPIString, "ICD+")

        For telIcd = 1 To UBound(icdArray)
            icdHier = headComLists("X058", Left(icdArray(telIcd), 3), "")
            If InStr(icdTmp, icdHier) Then
            Else
                If telIcd = UBound(icdArray) Then
                    icdTmp = icdTmp + icdHier
                Else
                    icdTmp = icdTmp + icdHier + "; "
                    'icdTmp = icdTmp + vbCrLf
                End If
            End If
        Next
        If icdTmp <> "" Then
            icdCheck = "Waarborgen: " + icdTmp
        Else
            icdCheck = ""
        End If

    End Function

Function ftxCheck(ByRef MAPIString As String) As String

        Dim ftxArray() As String
        Dim telFtx As Integer
        Dim ftxTmp As String
        Dim ftxHier As String

        ftxArray = Split(MAPIString, "FTX+018+")

        For telFtx = 1 To UBound(ftxArray)
            ftxTmp = ftxTmp + " " + Left(ftxArray(telFtx), InStr(ftxArray(telFtx), "'") - 1)
        Next

        'checken voor :
        Do While InStr(ftxTmp, "+08")
            ftxTmp = Replace(ftxTmp, "+08", "")
        Loop
        
        Do While InStr(ftxTmp, ":")
            ftxTmp = Replace(ftxTmp, ":", "")
        Loop
        
        'Do While InStr(ftxTmp, ". ")
        '    ftxTmp = Replace(ftxTmp, ". ", vbCrLf)
        'Loop
        
        If ftxTmp <> "" Then
            ftxCheck = "Bericht van de verzekeraar:" + vbCrLf + ftxTmp
        Else
            ftxCheck = ""
        End If

    End Function
