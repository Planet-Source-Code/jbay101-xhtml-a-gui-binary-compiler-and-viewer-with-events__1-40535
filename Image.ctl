VERSION 5.00
Begin VB.UserControl Image 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3300
   MouseIcon       =   "Image.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   220
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2385
      Top             =   2520
   End
   Begin VB.Image Image1 
      Height          =   915
      Left            =   -15
      Stretch         =   -1  'True
      Top             =   15
      Width           =   1635
   End
End
Attribute VB_Name = "Image"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim auto As Boolean
Event Click()

Private Sub Timer1_Timer()
If UserControl.Height <> Image1.Height Then
    'Image1.Height = UserControl.Height
    'UserControl.Height = Image1.Height
    Timer1.Enabled = False
End If

If UserControl.Width <> Image1.Width Then
    'UserControl.Width = Image1.Width
    'Image1.Width = UserControl.Width
    
    Timer1.Enabled = False
End If

UserControl.Refresh
End Sub

Private Sub UserControl_Click()
RaiseEvent Click

End Sub

Private Sub Image1_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_InitProperties()
UserControl.BackColor = UserControl.Parent.picClientArea.BackColor
End Sub

Private Sub UserControl_Resize()
Image1.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight

End Sub


Public Property Let HasHand(ByVal vNewValue As Boolean)
If vNewValue = True Then
    UserControl.MousePointer = 99
End If

End Property


Public Property Let AutoSize(ByVal vNewValue As Boolean)
If vNewValue = True Then
Image1.Stretch = False
Image1.Move 0, 0, Image1.Picture.Width, Image1.Picture.Height
UserControl.Width = Image1.Width
UserControl.Height = Image1.Height
End If
auto = vNewValue
End Property

Public Property Let Picture(ByVal sID As String)
LoadPictureBin sID, Image1
If auto = True Then
    UserControl.Width = Image1.Width
    UserControl.Height = Image1.Height
End If
End Property

Public Property Let BorderStyle(ByVal Value As Integer)
UserControl.BorderStyle = Value
End Property

