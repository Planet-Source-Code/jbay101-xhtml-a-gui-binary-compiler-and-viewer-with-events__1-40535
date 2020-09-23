Attribute VB_Name = "modProperty"
Option Explicit
Dim Controls(1000) As clsObject
Dim lastcontrol As Long


Function ProcessProperty(sClassName As String, MyProperty() As tProperty)
Dim i As Long
Dim sName As String
On Error Resume Next
sName = "ctrl" & Round(Timer * Rnd(10), 0)

Set Controls(lastcontrol) = CreateControl(sClassName, sName)

For i = 0 To UBound(MyProperty)
    If IsPropertyAlign(MyProperty(i).sName) Then
        ConvertAlign MyProperty(i).sName, CStr(MyProperty(i).sValue), Controls(lastcontrol).VBC
        GoTo NextProperty
    End If
    
    If IsPropertyReallyEvent(MyProperty(i).sName) Then
        Controls(lastcontrol).AddEventTag sName & "." & GetEventName(MyProperty(i).sName), MyProperty(i).sValue
        CallByName Controls(lastcontrol).VBC.object, "HasHand", VbLet, CVar(True)
        GoTo NextProperty
    End If
    
    If IsPropertyGeneric(MyProperty(i).sName) Then
        If IsPropertyMeasurement(MyProperty(i).sName) Then
            CallByName Controls(lastcontrol).VBC, MyProperty(i).sName, VbLet, CVar(FormatMeasurement(LCase(MyProperty(i).sName), MyProperty(i).sValue))
        Else
            CallByName Controls(lastcontrol).VBC, MyProperty(i).sName, VbLet, CVar(MyProperty(i).sValue)
        End If
        
    Else
        CallByName Controls(lastcontrol).VBC.object, MyProperty(i).sName, VbLet, CVar(MyProperty(i).sValue)
    End If

NextProperty:
Next i

lastcontrol = lastcontrol + 1
End Function

Function FreeObjects()
Dim i As Long
Dim b As Long
b = lastcontrol
lastcontrol = 0
On Error Resume Next

For i = 0 To b
    
    Unload Controls(i).VBC.object
    Controls(i).VBC.Visible = False
    Set Controls(i).VBC = Nothing
    Set Controls(i) = Nothing
Next i

End Function
Function VBC_EventProc(sLink As String)
MsgBox sLink
End Function

Function SetPageTitle(sNewTitle As String)
frmHidden.Caption = sNewTitle
End Function

Private Function IsPropertyGeneric(sProp As String) As Boolean
IsPropertyGeneric = True
If LCase(sProp) = "left" Then Exit Function
If LCase(sProp) = "top" Then Exit Function
If LCase(sProp) = "height" Then Exit Function
If LCase(sProp) = "width" Then Exit Function
If LCase(sProp) = "visible" Then Exit Function

IsPropertyGeneric = False

End Function

Private Function IsPropertyMeasurement(sProp As String) As Boolean
IsPropertyMeasurement = True
If LCase(sProp) = "left" Then Exit Function
If LCase(sProp) = "top" Then Exit Function
If LCase(sProp) = "height" Then Exit Function
If LCase(sProp) = "width" Then Exit Function

IsPropertyMeasurement = False

End Function

Private Function IsPropertyReallyEvent(sProp As String) As Boolean
If LCase(Left(sProp, 6)) = "event " Then
    IsPropertyReallyEvent = True
End If
End Function

Private Function GetEventName(sProp As String) As String
GetEventName = UCase(Mid(sProp, 7, Len(sProp) - 6))
End Function

Function FormatMeasurement(vName As String, vValue As Variant) As Variant
Dim vCache  As Variant

If InStr(1, vValue, "%") <> 0 Then
    vCache = Replace(vValue, "%", "") 'cheap hack!
    
    If (vName = "width") Or (vName = "left") Then vCache = vCache / 100 * frmHidden.picClientArea.ScaleWidth
    If (vName = "height") Or (vName = "top") Then vCache = vCache / 100 * frmHidden.picClientArea.ScaleHeight
    FormatMeasurement = vCache
ElseIf InStr(1, vValue, "px") <> 0 Then
    vCache = Replace(vValue, "px", "") 'cheap hack!
    If (vName = "width") Or (vName = "left") Then vCache = frmHidden.picClientArea.ScaleX(vCache, 3, frmHidden.picClientArea.ScaleMode)
    If (vName = "height") Or (vName = "top") Then vCache = frmHidden.picClientArea.ScaleY(vCache, 3, frmHidden.picClientArea.ScaleMode)

    
    FormatMeasurement = vCache
Else
    FormatMeasurement = vValue
End If

End Function

Function IsPropertyAlign(sName As String) As Boolean
IsPropertyAlign = False
If InStr(1, LCase(sName), "align") <> 0 Then IsPropertyAlign = True
End Function

Function ConvertAlign(sName As String, sValue As String, object As VBControlExtender)
Select Case LCase(sName)
Case "halign"
    Select Case LCase(sValue)
    Case "centered"
        object.Left = frmHidden.picClientArea.ScaleWidth / 2 - object.Width / 2
    Case "left"
        object.Left = 0
    Case "right"
        object.Left = frmHidden.picClientArea.ScaleWidth - object.Width
    End Select
Case "valign"
    Select Case LCase(sValue)
    Case "centered"
        object.Top = frmHidden.picClientArea.ScaleHeight / 2 - object.Height / 2
    Case "top"
        object.Top = 0
    Case "bottom"
        object.Top = frmHidden.picClientArea.ScaleHeight - object.Height
    End Select
    
End Select
End Function
