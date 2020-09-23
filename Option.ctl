VERSION 5.00
Begin VB.UserControl OptionButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   MouseIcon       =   "Option.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox States 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   1
      Left            =   1365
      Picture         =   "Option.ctx":0152
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   7
      Top             =   3045
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox States 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   0
      Left            =   1365
      Picture         =   "Option.ctx":02AB
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   2835
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox States 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   3
      Left            =   1575
      Picture         =   "Option.ctx":0415
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   3045
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox States 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   2
      Left            =   1575
      Picture         =   "Option.ctx":0577
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   2835
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox States 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   5
      Left            =   1785
      Picture         =   "Option.ctx":07B5
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   3045
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox States 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   4
      Left            =   1785
      Picture         =   "Option.ctx":090E
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   2835
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox States 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   7
      Left            =   1995
      Picture         =   "Option.ctx":0A78
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   3045
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox States 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   6
      Left            =   1995
      Picture         =   "Option.ctx":0B12
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   2835
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "OptionButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal HWND As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Enum AllStates
    CheckedDown = 0
    UncheckedDown = 1
    CheckedHot = 2
    UncheckedHot = 3
    CheckedIdle = 4
    UncheckedIdle = 5
    CheckedDisabled = 6
    UncheckedDisabled = 7
End Enum

Private Const LabelMargin As Integer = 15
Private Const DisabledForeColor As Long = 9740965

Private Const SRCCOPY = &HCC0020
Private PropEnabled As Boolean
Private PropCaption As String
Private PropForeColor As Long
Private PropValue As Boolean

Private MouseOver As Boolean
Private MouseDown As Boolean

Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Click()

Private Sub Redraw()
Cls

If PropEnabled = False Then
    If PropValue = True Then
        DrawAState CheckedDisabled
            Else
                DrawAState UncheckedDisabled
    End If
    
    GoTo DrawCaption
End If

If MouseDown = True Then
    If PropValue = True Then
        DrawAState CheckedDown
            Else
                DrawAState UncheckedDown
    End If
        Else
            If MouseOver = True Then
                If PropValue = True Then
                    DrawAState CheckedHot
                        Else
                            DrawAState UncheckedHot
                End If
                    Else
                        If PropValue = True Then
                            DrawAState CheckedIdle
                                Else
                                    DrawAState UncheckedIdle
                        End If
            End If
End If
        
DrawCaption:
If PropCaption = "" Then Exit Sub

With UserControl
    If PropEnabled = True Then
        UserControl.ForeColor = PropForeColor
            Else
                UserControl.ForeColor = DisabledForeColor
    End If
    
    'Draws the caption.
        SCaption = Split(PropCaption, " ")
        
        For X = 0 To UBound(SCaption)
            'See how much text can fit on one line before I add a line break.
            
            If TextWidth(EndCaption & SCaption(X)) > .ScaleWidth - 3 Then
                If EndCaption <> "" Then EndCaption = Left(EndCaption, Len(EndCaption) - 1)
                EndCaption = EndCaption & vbCrLf
            End If
            
            EndCaption = EndCaption & SCaption(X) & " "
        Next
        
        EndCaption = Left(EndCaption, Len(EndCaption) - 1)
        SCaption = Split(EndCaption, vbCrLf)
        
        .CurrentY = (.ScaleHeight / 2) - (TextHeight(EndCaption) / 2) - 1

        For X = 0 To UBound(SCaption)
            'Now draw each new line in the middle of the control.
            
            .CurrentX = LabelMargin
            Print SCaption(X)
        Next
End With
End Sub

Private Sub DrawAState(TheState As AllStates)
BitBlt UserControl.hDC, 0, (UserControl.ScaleHeight / 2) - (States(0).Height / 2), States(0).Width, States(0).Height, States(TheState).hDC, 0, 0, SRCCOPY
End Sub


Private Sub UserControl_InitProperties()
Caption = Ambient.DisplayName
Enabled = True
Value = False
Set Font = UserControl.Parent.Font
ForeColor = vbBlack
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseDown = True
Redraw
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim DrawIt As Boolean

SetCapture UserControl.HWND

If X < 0 Or X > UserControl.ScaleWidth Or Y < 0 Or Y > UserControl.ScaleHeight Then
    ReleaseCapture
    MouseOver = False
    MouseDown = False
    Redraw
        Else
            If MouseOver = False Then DrawIt = True
            MouseOver = True
            If DrawIt = True Then Redraw
End If

RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseDown = False

If PropEnabled = True Then
    UncheckOthers
              
    PropValue = True
    RaiseEvent Click
End If

Redraw
End Sub

Private Sub UncheckOthers()
UserControl.ParentControls.ParentControlsType = vbExtender

For X = 0 To UserControl.ParentControls.Count - 1
    If TypeOf UserControl.ParentControls(X) Is OptionButtonC Then
        If UserControl.ParentControls(X).Name <> UserControl.Ambient.DisplayName Then
            UserControl.ParentControls(X).Value = False
        End If
    End If
Next
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
Enabled = PropBag.ReadProperty("Enabled", True)
Value = PropBag.ReadProperty("Value", False)
Set Font = PropBag.ReadProperty("Font", UserControl.Parent.Font)
ForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
End Sub

Private Sub UserControl_Resize()
Redraw
End Sub
Public Property Let ForeColor(NewValue As OLE_COLOR)
PropForeColor = NewValue
Redraw
End Property
Public Property Get ForeColor() As OLE_COLOR
ForeColor = PropForeColor
End Property
Public Property Let Caption(NewValue As String)
PropCaption = NewValue
Redraw
End Property
Public Property Get Caption() As String
Caption = PropCaption
End Property

Public Property Let Enabled(NewValue As Boolean)
PropEnabled = NewValue
Redraw
End Property
Public Property Get Enabled() As Boolean
Enabled = PropEnabled
End Property

Public Property Let Value(NewValue As Boolean)
If NewValue = True Then UncheckOthers
PropValue = NewValue
Redraw
End Property
Public Property Get Value() As Boolean
Value = PropValue
End Property

Public Property Set Font(NewValue As StdFont)
Set UserControl.Font = NewValue
Redraw
End Property
Public Property Get Font() As StdFont
Set Font = UserControl.Font
End Property

Public Property Let BackColor(NewValue As OLE_COLOR)
UserControl.BackColor = NewValue
Redraw
End Property
Public Property Get BackColor() As OLE_COLOR
BackColor = UserControl.BackColor
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "Caption", PropCaption, Ambient.DisplayName
PropBag.WriteProperty "Enabled", PropEnabled, True
PropBag.WriteProperty "Value", PropValue, False
PropBag.WriteProperty "Font", UserControl.Font, UserControl.Parent.Font
PropBag.WriteProperty "ForeColor", PropForeColor, vbBlack
End Sub
Public Property Let HasHand(ByVal vNewValue As Boolean)
If vNewValue = True Then
    UserControl.MousePointer = 99
End If

End Property

