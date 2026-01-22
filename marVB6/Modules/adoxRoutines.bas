Attribute VB_Name = "modAdxRoutines"
Option Explicit

Function adxMaakDatabase(dbNaam As String, dbPath As String) As Boolean
    
    'adxMaakDatabase = False
    'Dim cat As New ADOX.Catalog
    'On Error Resume Next
    'cat.Create adojetprovider & _
    '            "Data Source=" + dbPath + "\" + dbNaam + ".mdb"
    'If Err Then
    '    MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
    'Else
    '    adxMaakDatabase = True
    'End If
    'Set cat.ActiveConnection = Nothing
        
End Function

Function adxKolom(tbNaam As String, clNaam As String, clType As Long, clLengte As Long) As Boolean

    Dim cat     As New ADOX.Catalog
    Dim clCol   As New ADOX.Column
                  
' Waarde van clType:
    ' adDate voor DATUM/TIJD
    ' adInteger voor INTEGER
    ' adCurrency  'voor geldwaarden (4 cijfers na de comma)
    ' adSingle voor SINGLE
    ' adDouble voor DUBBELE NAUWKEURIGHEID
    ' adBoolean voor BOOLEAN
    ' adVarWChar voor TEKST. Parameters Size indien vaste lengte.
    ' adLongVarWChar voor MEMO
    
    adxKolom = False
    Err = 0
    On Error Resume Next
    If InStr(adntDB.Properties("DBMS Name"), "SQL Server") Then
        ' De catalogus openen voor SQL-SERVER
        cat.ActiveConnection = SQLConnect
    Else
        ' De catalogus openen voor JET4
        cat.ActiveConnection = jetConnect
    End If
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        Set cat.ActiveConnection = Nothing
        Exit Function
    Else
        With clCol
            .Name = clNaam
            .ParentCatalog = cat
            .Type = clType
            If clLengte <> 0 Then
                .DefinedSize = clLengte
            End If
            .Properties("Nullable") = True
        End With
        cat.Tables(tbNaam).Columns.Append clCol
        Set cat = Nothing
        If Err Then
            MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
            MsgBox "Aanmaak Kolom " + clNaam + " zonder succes.", vbInformation
        Else
            adxKolom = True
        End If
    End If
    Set cat.ActiveConnection = Nothing

End Function

Function adxMaakTabel(tbNaam As String) As Boolean

    Dim cat     As New ADOX.Catalog
    Dim tbl     As New ADOX.Table
                           
    adxMaakTabel = False
    Err = 0
    On Error Resume Next
    If InStr(adntDB.Properties("DBMS Name"), "SQL Server") Then
        ' De catalogus openen voor SQL-SERVER
        cat.ActiveConnection = SQLConnect
    Else
        ' De catalogus openen voor JET4
        cat.ActiveConnection = jetConnect
    End If
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        Set cat.ActiveConnection = Nothing
        Exit Function
    Else
        With tbl
            .Name = tbNaam
            Set .ParentCatalog = cat
            .Columns.Append "ID", adInteger, adBigInt
                                                          
            ' De kolom volgnummer moet automatisch groter worden
            .Columns("ID").Properties("AutoIncrement") = True
        End With
        cat.Tables.Append tbl
        Set cat = Nothing
        If Err Then
            MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
        Else
            MsgBox "Aanmaak tabel " + tbNaam + " met succes.", vbInformation
            adxMaakTabel = True
        End If
    End If
    Set cat.ActiveConnection = Nothing
        
End Function


Function TabelKontrole() As Boolean

   Dim cnnHier As New ADODB.Connection
   Dim catHier As New ADOX.Catalog

   TabelKontrole = False
   cnnHier.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source= " & LOCATION_COMPANYDATA & _
      "marnt.mdv;"
   Set catHier.ActiveConnection = cnnHier
   
   For COUNT_TO = 0 To catHier.Tables.Count - 1
        Select Case catHier.Tables(COUNT_TO).Name
            Case "1900" To "2004"
                catHier.Tables(COUNT_TO).Name = "jr" & catHier.Tables(COUNT_TO).Name
        End Select
   Next
   catHier.Tables.Refresh
   cnnHier.Close
   Set catHier = Nothing
   Set cnnHier = Nothing
   
End Function

