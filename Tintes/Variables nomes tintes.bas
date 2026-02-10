Attribute VB_Name = "Module3"
Global nummaq As Byte
Type POINTAPI
        X As Long
        Y As Long
End Type

Function isloaded(vnomform As String) As Boolean
  Dim f
  For Each f In Forms
   If f.Name = vnomform Then
         isloaded = True
   End If
  Next
End Function

