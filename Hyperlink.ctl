VERSION 5.00
Begin VB.UserControl Hyperlink 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   1740
      Top             =   2910
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3030
      Top             =   1980
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      MouseIcon       =   "Hyperlink.ctx":0000
      TabIndex        =   0
      Top             =   -15
      Width           =   3855
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Hyperlink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type

Event Click()
Public Property Get Caption() As String
Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal vNewValue As String)
Label1.Caption = vNewValue
End Property

Public Property Get FontName() As String
FontName = Label1.FontName
End Property

Public Property Let FontName(ByVal vNewValue As String)
Label1.FontName = vNewValue
End Property

Public Property Get FontSize() As Variant
FontSize = Label1.FontSize
End Property

Public Property Let FontSize(ByVal vNewValue As Variant)
Label1.FontSize = vNewValue
End Property

Public Property Get FontBold() As Variant
FontBold = Label1.FontBold
End Property

Public Property Let FontBold(ByVal vNewValue As Variant)
Label1.Fontbond = vNewValue
End Property

Public Property Get FontItalic() As Variant
FontItalic = Label1.FontItalic
End Property

Public Property Let FontItalic(ByVal vNewValue As Variant)
FontItalic = Label1.FontItalic
End Property

Private Sub Label1_Click()
RaiseEvent Click
End Sub

Private Sub Timer1_Timer()
If UserControl.Height <> Label1.Height Then
    UserControl.Height = Label1.Height
    Timer1.Enabled = False
End If

If UserControl.Width <> Label1.Width Then
    UserControl.Width = Label1.Width
    Timer1.Enabled = False
End If

End Sub

Private Sub UserControl_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_InitProperties()

Label1.Font = UserControl.Parent.picClientArea.Font

Label1.BackColor = UserControl.Parent.picClientArea.BackColor
UserControl.BackColor = UserControl.Parent.picClientArea.BackColor


End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Label1.Caption = PropBag.ReadProperty("Caption", "[]")
Set Label1.Font = PropBag.ReadProperty("Font", UserControl.Parent.picClientArea.Font)
UserControl.BackColor = PropBag.ReadProperty("BackColor", UserControl.Parent.picClientArea.BackColor)
Label1.BackColor = PropBag.ReadProperty("BackColor", UserControl.Parent.picClientArea.BackColor)

UserControl.ForeColor = PropBag.ReadProperty("forecolor", UserControl.Parent.picClientArea.ForeColor)
Label1.ForeColor = PropBag.ReadProperty("forecolor", UserControl.Parent.picClientArea.ForeColor)
Label1.AutoSize = PropBag.ReadProperty("AutoSize", True)
End Sub

Private Sub UserControl_Resize()

Label1.Move 0, 0, UserControl.Width, UserControl.Height
If Label1.AutoSize = True Then
    Label1.AutoSize = True
End If


End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "Caption", Label1.Caption, "[no caption set]"
PropBag.WriteProperty "Font", Label1.Font, Parent.Font
PropBag.WriteProperty "AutoSize", Label1.AutoSize, True
End Sub

Public Property Get BackColor() As Variant
BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal vNewValue As Variant)
UserControl.BackColor = vNewValue
Label1.BackColor = vNewValue
End Property

Public Property Get ForeColor() As Variant
ForeColor = UserControl.BackColor
End Property

Public Property Let ForeColor(ByVal vNewValue As Variant)
UserControl.ForeColor = vNewValue
Label1.ForeColor = vNewValue
End Property


Public Property Let HasHand(ByVal vNewValue As Boolean)
If vNewValue = True Then
    Label1.MousePointer = 99
    UserControl.MousePointer = 99
End If

End Property

Public Property Let AutoSize(ByVal vNewValue As Boolean)
Label1.AutoSize = vNewValue
End Property

Public Property Get AutoSize() As Boolean
AutoSize = Label1.AutoSize
End Property

Private Sub Timer2_Timer()
Dim X As POINTAPI
GetCursorPos X
If WindowFromPoint(X.X, X.Y) <> UserControl.HWND Then
    If Label1.FontUnderline = True Then
        Label1.FontUnderline = False
        Label1.ForeColor = vbBlack
    End If
Else
    If Label1.FontUnderline = False Then
        Label1.FontUnderline = True
        Label1.ForeColor = vbBlue
    End If
End If

End Sub

