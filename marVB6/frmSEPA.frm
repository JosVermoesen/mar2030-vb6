VERSION 5.00
Object = "{E0348C7E-8303-439F-A0A3-3F4D8F38DDAC}#1.0#0"; "vsftIUCL.tlb"
Begin VB.Form SEPA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SEPA Converter"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4455
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4455
   StartUpPosition =   1  'CenterOwner
   Begin vsftIUCLCtl.SEPAUserControl SEPAUserControl1 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Object.Visible         =   "True"
      Enabled         =   "True"
      ForegroundColor =   "-2147483630"
      BackgroundColor =   "-2147483633"
      strBB           =   ""
      Location        =   "0, 8"
      Name            =   "InteropUserControl"
      Size            =   "297, 241"
      Object.TabIndex        =   "0"
   End
End
Attribute VB_Name = "SEPA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
