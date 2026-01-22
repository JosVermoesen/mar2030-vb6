VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbMAPI_Folders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *****************************************************************************
' *  Virtual-COM library: vbMAPI  http://www.everythingaccess.com/vbmapi.htm  *
' *  v1.16.3.0 DEMO              built Nov  8 2023                            *
' *---------------------------------------------------------------------------*
' *  This is a demonstration version of the library for TESTING purposes only.*
' *****************************************************************************

Option Explicit

Property Get Count() As Long: End Property
Attribute Count.VB_UserMemId = -101
Property Get Item(ByVal IndexOrName As Variant) As vbMAPI_Folder: End Property
Attribute Item.VB_UserMemId = 0
Property Get Session() As vbMAPI_Session: End Property
Attribute Session.VB_UserMemId = -102
Function Add(ByVal FolderName As String) As vbMAPI_Folder: End Function
Attribute Add.VB_UserMemId = -103
