Attribute VB_Name = "modAdoRoutines"
Option Explicit

Function ADOBIB_TEXT(adoField As ADODB.Field, TBS As String) As String

Err = 0
On Error Resume Next
If adoField = "" Then
    ADOBIB_TEXT = ""
Else
    ADOBIB_TEXT = Mid(adoField, InStr(adoField, TBS) + 7, InStr(InStr(adoField, TBS) + 7, adoField, "#") - (InStr(adoField, TBS) + 7))
End If
If Err Then ADOBIB_TEXT = ""

End Function

Function ADO_GET(iTabel As Integer, iIndex As Integer, sZoals As String, sZoek As Variant) As Boolean

Dim MsgHier As String

On Error Resume Next
If rsMAR(iTabel).State = adStateClosed Then
    Ktrl = bOpen(iTabel)
End If
ADO_GET = False
Err = 0
If InStr(SQLConnect, "SQLOLEDB") Then
    rsMAR(iTabel).Close
    If iTabel = TABLE_COUNTERS Then
        MsgHier = "SELECT * FROM " & "jr" & bstNaam(iTabel) & _
            " WHERE " & JETTABLEUSE_INDEX(iTabel, iIndex) & _
            " " & sZoals & " " & _
            " '" & sZoek & "'"
    Else
        MsgHier = "SELECT * FROM " & bstNaam(iTabel) & _
            " WHERE " & JETTABLEUSE_INDEX(iTabel, iIndex) & _
            " " & sZoals & " " & _
            " '" & sZoek & "'"
    End If
    'SnelHelpPrint MsgHier, BL_LOGGING
    rsMAR(iTabel).Open MsgHier, adntDB, adOpenForwardOnly, adLockOptimistic
    If Err Then
        MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
        Exit Function
    ElseIf rsMAR(iTabel).RecordCount Then
        ADO_GET = True
    End If
Else
    rsMAR(iTabel).Index = FLINDEX_CAPTION(iTabel, iIndex)
    If sZoals = "=" Then
        rsMAR(iTabel).Seek sZoek
    ElseIf sZoals = ">=" Then
        rsMAR(iTabel).Seek sZoek, adSeekAfterEQ
    Else
        MsgBox sZoals & " nog niet beschikbaar"
    End If
    If rsMAR(iTabel).EOF Then
    Else
        ADO_GET = True
    End If
End If

End Function

Function RV(adoRecord As ADODB.Recordset, TBS As String) As Variant

If IsNull(adoRecord(TBS)) Or IsEmpty(adoRecord(TBS)) Then
    RV = ""
Else
    RV = adoRecord(TBS)
End If

End Function


Function ADOJOURNAL_OK() As Boolean

ADOJOURNAL_OK = False
If Val(RV(rsJournaal, "v068")) = 0 Then
    SnelHelpPrint "BoekBedrag is 0", BL_LOGGING
Else
    DKTRL_CUMUL = DKTRL_CUMUL + Val(RV(rsJournaal, "v068"))
    DKTRL_BEF = DKTRL_BEF + Round(Val(RV(rsJournaal, "v068")) * EURO, 0)
    DKTRL_EUR = DKTRL_EUR + Round(Val(RV(rsJournaal, "v068")), 2)
    On Error Resume Next
    rsJournaal("dece068") = Val(RV(rsJournaal, "v068"))
    Err = 0
    On Error Resume Next
    
    Dim Pipo As Variant
    Pipo = rsJournaal("v019") & vbTab & rsJournaal("v067") & vbTab
    If Val(rsJournaal("v068")) < 0 Then
        Pipo = Pipo & "" & vbTab & Format(-Val(rsJournaal("v068")), "#,##0.00")
        Pipo = Pipo & vbTab & "" & vbTab & Format(Round(-Val(rsJournaal("v068")) * EURO), "#,##0.00")
    Else
        Pipo = Pipo & Format(Val(rsJournaal("v068")), "#,##0.00") & vbTab & ""
        Pipo = Pipo & vbTab & Format(Round(Val(rsJournaal("v068")) * EURO), "#,##0.00") & vbTab & ""
    End If
    frmBoeking.mshfBoekLijst.AddItem Pipo, frmBoeking.mshfBoekLijst.Rows - 1
End If

rsJournaal("v070") = vSet(rsJournaal("v019"), 7) + rsJournaal("v066")
bGet TABLE_LEDGERACCOUNTS, 0, rsJournaal("v019")
If Ktrl Then
    MsgBox "Rekening " + rsJournaal("v019") + " niet te vinden." + vbCrLf + "Eerst SETUPrekening inbrengen a.u.b. !"
    DKTRL_CUMUL = DKTRL_CUMUL + 99
    Exit Function
ElseIf ACTIVE_BOOKYEAR Then
    RecordToVeld TABLE_LEDGERACCOUNTS
    vBib TABLE_LEDGERACCOUNTS, Str$(Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#e023 #")) + Val(RV(rsJournaal, "v068"))), "e023"
    rsMAR(TABLE_LEDGERACCOUNTS)("dece023") = rsMAR(TABLE_LEDGERACCOUNTS)("dece023") + RV(rsJournaal, "dece068")
Else
    RecordToVeld TABLE_LEDGERACCOUNTS
    vBib TABLE_LEDGERACCOUNTS, Str$(Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#e022 #")) + Val(RV(rsJournaal, "v068"))), "e022"
    rsMAR(TABLE_LEDGERACCOUNTS)("dece022") = rsMAR(TABLE_LEDGERACCOUNTS)("dece022") + RV(rsJournaal, "dece068")
End If
rsMAR(TABLE_LEDGERACCOUNTS)("dnnsync") = False
bUpdate TABLE_LEDGERACCOUNTS, 0
Err = 0
On Error Resume Next
rsJournaal.Update
If Err Then
    MsgBox Error
Else
    ADOJOURNAL_OK = True
End If

End Function


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

