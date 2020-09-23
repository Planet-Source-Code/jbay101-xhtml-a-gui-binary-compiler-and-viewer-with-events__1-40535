Attribute VB_Name = "modControl"
Option Explicit

Function CreateControl(sClassName As String, sName As String) As clsObject
Dim tmp As clsObject
On Error GoTo err_handle:
Set tmp = New clsObject


Set tmp.VBC = frmHidden.Controls.Add(sClassName, sName, frmHidden.picClientArea)
Set CreateControl = tmp

Exit Function
err_handle:
Debug.Print Timer & " > " & "Class: " & sClassName & " could not be created!"
End Function
