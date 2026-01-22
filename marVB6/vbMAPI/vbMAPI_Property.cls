VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbMAPI_Property"
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

Property Get PropTag() As Long: End Property
Attribute PropTag.VB_UserMemId = -101
Property Get PropTagSymbol() As String: End Property
Attribute PropTagSymbol.VB_UserMemId = -102
Property Get Value() As Variant: End Property
Attribute Value.VB_UserMemId = 0
Property Let Value(ByVal Value As Variant): End Property
Attribute Value.VB_UserMemId = 0
Property Get IsCached() As Boolean: End Property
Attribute IsCached.VB_UserMemId = -103
Property Get Session() As vbMAPI_Session: End Property
Attribute Session.VB_UserMemId = -104
