VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private dReeks() As Double
Private mlPositie As Long
Private mlUBound As Long
Private mdRunningTotal As Double
Private mbCompleet As Boolean

Private Sub Class_Initialize()
    mlPositie = 0
    mlUBound = 0
    mdRunningTotal = 0
    mbCompleet = False
End Sub

Public Property Let aantalElementen(lAantalElementen As Long)
    mlUBound = lAantalElementen - 1
    ReDim dArray(mlUBound)
End Property

Public Sub AddItem(dItem As Double)
    mdRunningTotal = _
        (mdRunningTotal - dArray(mlPositie)) + dItem
    dArray(mlPositie) = dItem
    If mlPositie = mlUBound Then
        mlPositie = 0
        mbCompleet = True
    Else
        mlPositie = mlPositie + 1
    End If
End Sub

Public Property Get GeMiddelde() As Double
    GeMiddelde = mdRunningTotal / (mlUBound + 1)
End Property

Public Property Get Volledig() As Boolean
    Volledig = mbCompleet
End Property


