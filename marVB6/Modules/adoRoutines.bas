Attribute VB_Name = "modAdoRoutines"
Option Explicit

Public Function SQLPopUp(OpzoekReeks As String, tbNaam As String, vldNaam As String, idKode As String) As Boolean

Dim Zoekstring As String
Dim SQLstring As String
Dim SQLPopString As String
Dim strSQL As String

Dim FormHier As Object

Dim rs  As ADODB.Recordset

SQLPopUp = False
Do While OpzoekReeks <> ""
    If InStr(OpzoekReeks, ";") = 0 Then
        Zoekstring = OpzoekReeks
        OpzoekReeks = ""
    Else
        Zoekstring = Left(OpzoekReeks, InStr(OpzoekReeks, ";") - 1)
        OpzoekReeks = Mid(OpzoekReeks, InStr(OpzoekReeks, ";") + 1)
    End If
    bGet TABLE_VARIOUS, 1, "29" + Zoekstring
    If Ktrl Then
    Else
        RecordToVeld TABLE_VARIOUS
        If InStr(vBibTekst(TABLE_VARIOUS, "#v132 #"), "[Colwidth]") Then
            SQLstring = Left(vBibTekst(TABLE_VARIOUS, "#v132 #"), InStr(vBibTekst(TABLE_VARIOUS, "#v132 #"), "[Colwidth]") - 1)
        Else
            SQLstring = vBibTekst(TABLE_VARIOUS, "#v132 #")
        End If
        SQLPopString = Mid(SQLstring, InStr(SQLstring, "FROM ") + 5)
        If SQLPopString = "" Then
            MsgBox "Ongeldige syntax in : " + SQLstring, vbCritical
            Exit Function
        End If
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
         
        Err = 0
        On Error Resume Next
        rs.Open "SELECT * FROM " + SQLPopString + " WHERE " + Left(tbNaam, 1) + vldNaam + " Like '" + idKode + "%'", adntDB, adOpenStatic, adLockOptimistic
        If Err Then
            'Dus eerst verbindingsveld nog invoegen
            MsgBox "Verbindingsveld bestaat nog niet.  Voeg eerst een veld toe: " + Left(tbNaam, 1) + vldNaam
        ElseIf rs.RecordCount = 0 Then
            'deze partij bestaat nog niet, dus invoegen
            strSQL = "INSERT INTO " & SQLPopString & " (" & Left(tbNaam, 1) & vldNaam & ") VALUES ('" & idKode & "');"
            adntDB.Execute strSQL, , adCmdText
        End If
    
        'Msg = Mid(vBibTekst(TABLE_VARIOUS, "#v132 #"), InStr$(vBibTekst(TABLE_VARIOUS, "#v132 #"), "[Colwidth]") + 10)
        'If Msg = "" Then
        '    grdColWidth(0) = 0
        'Else
        'COUNT_TO = 0
        'Do While Msg <> ""
        '    If InStr$(Msg, vbTab) <> 0 Then
        '        grdColWidth(COUNT_TO) = Val(Left(Msg, InStr$(Msg, vbTab) - 1))
        '        Msg = Mid(Msg, InStr$(Msg, vbTab) + 1)
        '        COUNT_TO = COUNT_TO + 1
        '    Else
        '        Exit Do
        '    End If
        'Loop
        'grdColWidth(COUNT_TO) = 0
        Set FormHier = New mijnSQLEditor
        'If optBewerkbaar.Value = True Then
            FormHier!msfSQL.AllowUpdate = True
            FormHier!msfSQL.AllowAddNew = False
            FormHier!msfSQL.AllowDelete = False
        'Else
        '    FormHier!msfSQL.AllowUpdate = False
        '    FormHier!msfSQL.AllowAddNew = False
        '    FormHier!msfSQL.AllowDelete = False
        'End If
        FormHier.txtSQL = SQLstring + " WHERE " + Left(tbNaam, 1) + vldNaam + " Like '" + idKode + "%'"
        FormHier.Caption = SQLstring
        FormHier.Show

        FormHier.SetFocus
        SendKeys "{ENTER}"
    End If
Loop
'DirekteVerkoop.SetFocus
SQLPopUp = True

End Function

Public Function OpenSchemaString(tbType As String) As Variant

   Dim rstSchema As ADODB.Recordset
   Dim strCnn As String
   Dim retString As Variant
   
   retString = ""
   
   Err = 0
   On Error Resume Next
   Set rstSchema = adntDB.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))
   If Err Then MsgBox Error: Exit Function
   Do Until rstSchema.EOF
        retString = retString & rstSchema!TABLE_NAME & vbCr
        rstSchema.MoveNext
   Loop
   rstSchema.Close
   Set rstSchema = Nothing
   OpenSchemaString = retString
        
End Function


Public Sub ToonIndexen(tbNaam As String, obObject As Object)

   Dim rstSchema As ADODB.Recordset
       
   obObject.Clear
  
   Set rstSchema = adntDB.OpenSchema(adSchemaIndexes)
  
   Do Until rstSchema.EOF
    If UCase(tbNaam) = UCase(rstSchema!TABLE_NAME) Then
        obObject.AddItem "+" & rstSchema!COLUMN_NAME & "; " & rstSchema!INDEX_NAME
        'for t=0 to rstschema.Fields.Count-1:print t;" ";rstschema.Fields(t).Name;" = ";rstschema.Fields(t).Value:next
    End If
      rstSchema.MoveNext
   Loop
   rstSchema.Close
   
End Sub

Function vBT(TLBR As Variant, TBS As String) As String

    Dim tbsHier As String * 7
        
    tbsHier = "#     #"
    Mid(tbsHier, 2) = TBS
    Err = 0
    On Error Resume Next
    If TLBR = "" Then
        vBT = ""
    Else
        vBT = Mid(TLBR, InStr(TLBR, tbsHier) + 7, InStr(InStr(TLBR, tbsHier) + 7, TLBR, "#") - (InStr(TLBR, tbsHier) + 7))
    End If
    If Err Then vBT = ""

End Function

