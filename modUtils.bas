Attribute VB_Name = "modUtils"
Option Explicit

Function LoadPictureBin(sID As String, Target As Object)
Target.Picture = LoadPicture(PATH & "\" & sID)
End Function

Function StripPath(sPath As String) As String
Dim bak As String
bak = StrReverse(sPath)

Dim pos As Integer
pos = InStr(1, bak, "/")
If pos = 0 Then pos = InStr(1, bak, "\")

StripPath = Mid(sPath, 1, Len(sPath) - pos)
End Function
