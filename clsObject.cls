VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'this is an object to store our control. It is the only way we can
' use events
Option Explicit

Public WithEvents VBC As VBControlExtender
Attribute VBC.VB_VarHelpID = -1

Private Type et
    sEvent As String
    sTag As String
End Type

Dim event_tag(100) As et
Dim lastet As Long


Private Sub VBC_ObjectEvent(Info As EventInfo)
    If GetEventTag(VBC.Name & "." & UCase(Info.Name)) <> "" Then
        VBC_EventProc GetEventTag(VBC.Name & "." & UCase(Info.Name))
    End If


End Sub

Function AddEventTag(sEvent As String, ByVal sTag As String)
event_tag(lastet).sEvent = sEvent
event_tag(lastet).sTag = sTag
End Function

Function GetEventTag(sEvent As String) As String
Dim i As Integer
For i = 0 To lastet
If event_tag(i).sEvent = sEvent Then GetEventTag = event_tag(i).sTag
Next i
End Function
