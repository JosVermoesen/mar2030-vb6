Attribute VB_Name = "thisAdoToolsClass"
DefInt A-Z
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


