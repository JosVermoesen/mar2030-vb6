Attribute VB_Name = "MultiMediaRoutines"
Sub BeginPlaySound(ByVal ResourceId As Integer)
Dim Ret As Variant
    
On Local Error Resume Next
Err = 0

#If Win32 Then
    ' Important: The returned string is converted to Unicode
    SoundBuffer = StrConv(LoadResData(ResourceId, "VSOFT_SOUND"), vbUnicode)
#Else
    SoundBuffer = LoadResData(ResourceId, "VSOFT_SOUND")
#End If

If Err Then
    Exit Sub
Else
    Ret = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
    ' Important: This function is neccessary for playing sound asynchronously
    DoEvents
End If

End Sub

Sub EndPlaySound()
    Dim Ret As Variant
    Ret = sndPlaySound(0&, 0&)
End Sub



Sub PrintAnywhere(Src As Object, Dest As Object)

On Error Resume Next
Dest.PaintPicture Src.Picture, Dest.Width / 2, Dest.Height / 2
If Dest Is Printer Then
    Printer.EndDoc
End If

End Sub


