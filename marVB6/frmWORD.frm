VERSION 5.00
Begin VB.Form frmWORD 
   Caption         =   "WordTest"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   9945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1215
   End
End
Attribute VB_Name = "frmWORD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim wApp As Word.Application

    
Private Sub Command1_Click()

    Dim TT As Integer
    Dim volNaam As String
    
    Set wApp = New Word.Application
    volNaam = BedrijfsLokatie & "sjb\sBrief.doc"
    If Dir(volNaam) = "" Then
        Set wApp = Nothing
        Exit Sub
    Else
        wApp.Documents.Open (volNaam)
    End If
        
    Dim Book() As Variant
    ReDim Book(12, 2)
    Book(1, 1) = "Dhr."
    Book(2, 1) = "Van Den Plas"
    Book(3, 1) = "Armand"
    Book(4, 1) = "Mw."
    Book(5, 1) = "Regenbui"
    Book(6, 1) = "Inge"
    Book(7, 1) = "Populierenstraat"
    Book(8, 1) = 39
    Book(9, 1) = "bus 8"
    Book(10, 1) = "B"
    Book(11, 1) = 9310
    Book(12, 1) = "Moorsel"
        
    Book(1, 2) = "Aan1"
    Book(2, 2) = "Naam1"
    Book(3, 2) = "Voornaam1"
    Book(4, 2) = "Aan2"
    Book(5, 2) = "Naam2"
    Book(6, 2) = "Voornaam2"
    Book(7, 2) = "Straat"
    Book(8, 2) = "Nr"
    Book(9, 2) = "Bus"
    Book(10, 2) = "LND"
    Book(11, 2) = "PC"
    Book(12, 2) = "Plaats"
        
    With wApp.Documents(1)
        For TT = 1 To 12
            .Bookmarks(Book(TT, 2)).Select
            .Bookmarks(Book(TT, 2)).Range.InsertAfter (Book(TT, 1))
        Next
    End With
    wApp.Documents(1).SaveAs volNaam & Format(Now, "yyyymmddhhmmss")
    SnelHelpPrint wApp.Documents(1).Name, blLogging
    wApp.Documents(1).PrintOut
    wApp.Documents(1).Close SaveChanges:=False
    Set wApp = Nothing
    'Me.OLE1.SourceDoc
    
    
End Sub

