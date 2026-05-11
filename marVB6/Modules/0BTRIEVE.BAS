Attribute VB_Name = "modDataBaseRoutines"
DefInt A-Z
Option Explicit

Dim X   As Integer

Sub SchrijfLog(BtrieveAktie As String, FlNummer As Integer, IndexNummer As Integer, IndexSleutel As String)
Dim FlLog As Integer
Dim RecordLijn As String

RecordLijn = BtrieveAktie + Format(IndexNummer)
Select Case FlNummer
    Case TABLE_VARIOUS
        RecordLijn = RecordLijn + "ALLERLEI"
    Case TABLE_CUSTOMERS
        RecordLijn = RecordLijn + "KLANTEN "
    Case TABLE_SUPPLIERS
        RecordLijn = RecordLijn + "LEVERANC"
    Case TABLE_LEDGERACCOUNTS
        RecordLijn = RecordLijn + "REKENING"
    Case TABLE_PRODUCTS
        RecordLijn = RecordLijn + "PRODUKT "
    Case TABLE_INVOICES
        RecordLijn = RecordLijn + "dokument"
    Case TABLE_JOURNAL
        RecordLijn = RecordLijn + "JOURNAAL"
    Case TABLE_CONTRACTS
        RecordLijn = RecordLijn + "POLISSEN"
    Case TABLE_COUNTERS
        RecordLijn = RecordLijn + "TELLERS "
End Select

Select Case BtrieveAktie
    Case "INSERT", "UPDATE"
        RecordLijn = RecordLijn + TLB_RECORD(FlNummer)
    Case Else
        RecordLijn = RecordLijn + IndexSleutel
End Select

FlLog = FreeFile
Open PROGRAM_LOCATION + "NTIMPORT.LOG" For Append Lock Write As FlLog
    Print #FlLog, RecordLijn; vbCrLf;
Close FlLog

End Sub


