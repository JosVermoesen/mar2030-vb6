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

