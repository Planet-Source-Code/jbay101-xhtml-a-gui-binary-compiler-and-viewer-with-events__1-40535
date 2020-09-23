VERSION 5.00
Begin VB.UserControl CommandButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   MouseIcon       =   "Command.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   900
      Top             =   2280
   End
End
Attribute VB_Name = "CommandButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Declare Function SetCapture Lib "user32" (ByVal HWND As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Type SepRGB
    Red As Single
    Green As Single
    Blue As Single
End Type

Private Const DisabledForeColor As Long = 9740965

'Idle Colors
    Private Const BorderColorLines As Long = 7549440
    
    Private Const FirstBottomLine As Long = 15199215
    Private Const SecondBottomLine As Long = 14082023
    Private Const ThirdBottomLine As Long = 13030358
    
    Private Const FirstCornerPixel As Long = 8672545
    Private Const SecondCornerPixel As Long = 11376251
    Private Const ThirdCornerPixel As Long = 10845522
    Private Const FourthCornerPixel As Long = 14602182
    
    Private Const FromColorFade As Long = 16250871
    Private Const ToColorFade As Long = 15199215

'Disabled Colors
    Private Const BorderColorLinesX As Long = 12437454
    
    Private Const FirstBottomLineX As Long = 15726583
    Private Const SecondBottomLineX As Long = 15726583
    Private Const ThirdBottomLineX As Long = 15726583
    
    Private Const FirstCornerPixelX As Long = 12437454
    Private Const SecondCornerPixelX As Long = 12437454
    Private Const ThirdCornerPixelX As Long = 12437454
    Private Const FourthCornerPixelX As Long = 12437454
    
    Private Const FromColorFadeX As Long = 15726583
    Private Const ToColorFadeX As Long = 15726583
    
'Down colors.
    Private Const BorderColorLinesD As Long = 7549440
    
    Private Const FirstBottomLineD As Long = 15199215
    Private Const SecondBottomLineD As Long = 14082023
    Private Const ThirdBottomLineD As Long = 15725559
    
    Private Const FirstCornerPixelD As Long = 8672545
    Private Const SecondCornerPixelD As Long = 11376251
    Private Const ThirdCornerPixelD As Long = 10845522
    Private Const FourthCornerPixelD As Long = 14602182
    
    Private Const FromColorFadeD As Long = 14607335
    Private Const ToColorFadeD As Long = 14607335
    
'Has focus colors
    Private Const BorderColorLinesF As Long = 7549440
    
    Private Const FirstTopLineF As Long = 16771022
    Private Const SecondTopLineF As Long = 16242621
    
    Private Const FirstBottomLineF As Long = 15199215
    Private Const SecondBottomLineF As Long = 15183500
    Private Const ThirdBottomLineF As Long = 15696491
    
    Private Const FirstCornerPixelF As Long = 8672545
    Private Const SecondCornerPixelF As Long = 11376251
    Private Const ThirdCornerPixelF As Long = 10845522
    Private Const FourthCornerPixelF As Long = 14602182
    
    Private Const FromColorFadeF As Long = 16250871
    Private Const ToColorFadeF As Long = 15199215
    
    Private Const SideFromColorFadeF As Long = 16241597
    Private Const SideToColorFadeF As Long = 15183500

'HOT Colors
    Private Const BorderColorLinesH As Long = 7549440
    
    Private Const FirstTopLineH As Long = 13562879
    Private Const SecondTopLineH As Long = 9231359
    
    Private Const FirstBottomLineH As Long = 15199215
    Private Const SecondBottomLineH As Long = 3257087
    Private Const ThirdBottomLineH As Long = 38630
    
    Private Const FirstCornerPixelH As Long = 8672545
    Private Const SecondCornerPixelH As Long = 11376251
    Private Const ThirdCornerPixelH As Long = 10845522
    Private Const FourthCornerPixelH As Long = 14602182
    
    Private Const FromColorFadeH As Long = 16250871
    Private Const ToColorFadeH As Long = 15199215
    
    Private Const SideFromColorFadeH As Long = 10280929
    Private Const SideToColorFadeH As Long = 3192575
    
Private PropCaption As String
Private HasFocus As Boolean
Private MouseOver As Boolean
Private MouseDown As Boolean
Private PropEnabled As Boolean
Private PropForeColor As Long

Event Click()
Event DblClick()
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

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

Public Property Set Font(NewValue As StdFont)
Set UserControl.Font = NewValue
Redraw
End Property
Public Property Get Font() As StdFont
Set Font = UserControl.Font
End Property

Public Property Let Enabled(NewValue As Boolean)
PropEnabled = NewValue
Redraw
End Property
Public Property Get Enabled() As Boolean
Enabled = PropEnabled
End Property
Private Function CreateFade(FromColor As Long, ToColor As Long, FadeLength As Long) As Collection
Dim Increment As SepRGB
Dim ToColorRGB As SepRGB
Dim FromColorRGB As SepRGB
Dim SubRGB As SepRGB
Dim Final As SepRGB
Dim Results As New Collection

If FromColor = ToColor Then
    For X = 1 To FadeLength
        Results.Add FromColor
    Next
    GoTo ThatsIt
End If

ToColorRGB = GetRGB(ToColor)
FromColorRGB = GetRGB(FromColor)

With SubRGB
    .Red = Abs(ToColorRGB.Red - FromColorRGB.Red)
    .Green = Abs(ToColorRGB.Green - FromColorRGB.Green)
    .Blue = Abs(ToColorRGB.Blue - FromColorRGB.Blue)
End With

With Increment
    .Red = SubRGB.Red / FadeLength
    .Green = SubRGB.Green / FadeLength
    .Blue = SubRGB.Blue / FadeLength
End With

With Final
    .Red = FromColorRGB.Red
    .Green = FromColorRGB.Green
    .Blue = FromColorRGB.Blue

For X = 1 To FadeLength
        Results.Add RGB(.Red, .Green, .Blue)
        
        If .Red <> ToColorRGB.Red Then If .Red > ToColorRGB.Red Then .Red = .Red - Increment.Red Else .Red = .Red + Increment.Red
        If .Green <> ToColorRGB.Green Then If .Green > ToColorRGB.Green Then .Green = .Green - Increment.Green Else .Green = .Green + Increment.Green
        If .Blue <> ToColorRGB.Blue Then If .Blue > ToColorRGB.Blue Then .Blue = .Blue - Increment.Blue Else .Blue = .Blue + Increment.Blue
    Next
End With

ThatsIt:
Set CreateFade = Results
End Function
Private Function GetRGB(ByVal LongValue As Long) As SepRGB
LongValue = Abs(LongValue)
GetRGB.Red = LongValue And 255
GetRGB.Green = (LongValue \ 256) And 255
GetRGB.Blue = (LongValue \ 65536) And 255
End Function
Private Sub Redraw()
Cls

If PropEnabled = False Then
    DrawDisabled
    GoTo DrawCaption
End If

If MouseDown = True Then
    If MouseOver = True Then
        DrawDown
            Else
                GoTo DoOthers
    End If
        Else
DoOthers:
            If MouseOver = True Then
                DrawHot
                    Else
                        If HasFocus = False Then
                            DrawIdle
                                Else
                                    DrawFocus
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
        
        .CurrentY = (.ScaleHeight / 2) - (TextHeight(EndCaption) / 2)

        For X = 0 To UBound(SCaption)
            'Now draw each new line in the middle of the control.
            
            .CurrentX = (.ScaleWidth / 2) - (TextWidth(SCaption(X)) / 2)
            Print SCaption(X)
        Next
End With
End Sub

Private Sub DrawFocus()
Dim Gradient As New Collection
Dim SCaption() As String
Dim EndCaption As String

With UserControl
    'Draws border lines (not corners)
        Line (3, 0)-(.ScaleWidth - 3, 0), BorderColorLinesF
        Line (0, 3)-(0, .ScaleHeight - 3), BorderColorLinesF
        Line (3, .ScaleHeight - 1)-(.ScaleWidth - 3, .ScaleHeight - 1), BorderColorLinesF
        Line (.ScaleWidth - 1, 3)-(.ScaleWidth - 1, .ScaleHeight - 3), BorderColorLinesF
    
    'Draws the fade at the bottom.
        Line (1, .ScaleHeight - 4)-(.ScaleWidth - 1, .ScaleHeight - 4), FirstBottomLineF
        Line (2, .ScaleHeight - 3)-(.ScaleWidth - 2, .ScaleHeight - 3), SecondBottomLineF
        Line (3, ScaleHeight - 2)-(.ScaleWidth - 3, .ScaleHeight - 2), ThirdBottomLineF
        
    'Draws the background gradient.
        Set Gradient = CreateFade(FromColorFadeF, ToColorFadeF, .ScaleHeight - 5)
        
        For X = 1 To Gradient.Count
            Select Case X
            Case 1
                Line (3, X + 1)-(.ScaleWidth - 4, X + 1), FirstTopLineF
            Case 2
                Line (2, X + 1)-(.ScaleWidth - 3, X + 1), SecondTopLineF
            Case Else
                Line (1, X + 1)-(.ScaleWidth - 2, X + 1), Gradient(X)
            End Select
        Next
    
    'Draws side gradients
        Set Gradient = CreateFade(SideFromColorFadeF, SideToColorFadeF, .ScaleHeight - 7)
        
        For X = 1 To Gradient.Count
            PSet (1, X + 3), Gradient(X)
            PSet (2, X + 3), Gradient(X)
            
            PSet (.ScaleWidth - 2, X + 3), Gradient(X)
            PSet (.ScaleWidth - 3, X + 3), Gradient(X)
        Next
        
    'Draws corners
    'First set of pixels
        'Upper Left Corner
        PSet (2, 0), FirstCornerPixelF
        PSet (0, 2), FirstCornerPixelF
        
        'Bottom left corner
        PSet (0, .ScaleHeight - 3), FirstCornerPixelF
        PSet (2, .ScaleHeight - 1), FirstCornerPixelF
        
        'Top right corner
        PSet (.ScaleWidth - 1, 2), FirstCornerPixelF
        PSet (.ScaleWidth - 3, 0), FirstCornerPixelF
        
        'Bottom right corner
        PSet (.ScaleWidth - 3, .ScaleHeight - 1), FirstCornerPixelF
        PSet (.ScaleWidth - 1, .ScaleHeight - 3), FirstCornerPixelF
        
    'Second set of pixels.
        'Upper Left Corner
        PSet (1, 0), SecondCornerPixelF
        PSet (0, 1), SecondCornerPixelF
        
        'Bottom left corner
        PSet (0, .ScaleHeight - 2), SecondCornerPixelF
        PSet (1, .ScaleHeight - 1), SecondCornerPixelF
        
        'Top right corner
        PSet (.ScaleWidth - 1, 1), SecondCornerPixelF
        PSet (.ScaleWidth - 2, 0), SecondCornerPixelF
        
        'Bottom right corner
        PSet (.ScaleWidth - 2, .ScaleHeight - 1), SecondCornerPixelF
        PSet (.ScaleWidth - 1, .ScaleHeight - 2), SecondCornerPixelF
    
    'Third pixel.
        PSet (1, 1), ThirdCornerPixelF
        PSet (1, .ScaleHeight - 2), ThirdCornerPixelF
        PSet (.ScaleWidth - 2, 1), ThirdCornerPixelF
        PSet (.ScaleWidth - 2, .ScaleHeight - 2), ThirdCornerPixelF
    
    'Fourth set of pixels.
        'Upper left corner.
            PSet (2, 1), FourthCornerPixelF
            PSet (1, 2), FourthCornerPixelF
        
        'Bottom left corner.
            PSet (1, .ScaleHeight - 3), FourthCornerPixelF
            PSet (2, .ScaleHeight - 2), FourthCornerPixelF
        
        'Bottom right corner.
            PSet (.ScaleWidth - 3, .ScaleHeight - 2), FourthCornerPixelF
            PSet (.ScaleWidth - 2, .ScaleHeight - 3), FourthCornerPixelF
        
        'Top right corner.
            PSet (.ScaleWidth - 3, 1), FourthCornerPixelF
            PSet (.ScaleWidth - 2, 2), FourthCornerPixelF
End With
End Sub
Private Sub DrawDown()
Dim Gradient As New Collection
Dim SCaption() As String
Dim EndCaption As String

With UserControl
    'Draws border lines (not corners)
        Line (3, 0)-(.ScaleWidth - 3, 0), BorderColorLinesD
        Line (0, 3)-(0, .ScaleHeight - 3), BorderColorLinesD
        Line (3, .ScaleHeight - 1)-(.ScaleWidth - 3, .ScaleHeight - 1), BorderColorLinesD
        Line (.ScaleWidth - 1, 3)-(.ScaleWidth - 1, .ScaleHeight - 3), BorderColorLinesD
    
    'Draws the fade at the bottom.
        Line (1, .ScaleHeight - 4)-(.ScaleWidth - 1, .ScaleHeight - 4), FirstBottomLineD
        Line (2, .ScaleHeight - 3)-(.ScaleWidth - 2, .ScaleHeight - 3), SecondBottomLineD
        Line (3, ScaleHeight - 2)-(.ScaleWidth - 3, .ScaleHeight - 2), ThirdBottomLineD
        
    'Draws the background gradient.
        Set Gradient = CreateFade(FromColorFadeD, ToColorFadeD, .ScaleHeight - 5)
        
        For X = 1 To Gradient.Count
            Select Case X
            Case 1
                Line (3, X + 1)-(.ScaleWidth - 4, X + 1), Gradient(X)
            Case 2
                Line (2, X + 1)-(.ScaleWidth - 3, X + 1), Gradient(X)
            Case Else
                Line (1, X + 1)-(.ScaleWidth - 2, X + 1), Gradient(X)
            End Select
        Next
        
    'Draws corners
    'First set of pixels
        'Upper Left Corner
        PSet (2, 0), FirstCornerPixelD
        PSet (0, 2), FirstCornerPixelD
        
        'Bottom left corner
        PSet (0, .ScaleHeight - 3), FirstCornerPixelD
        PSet (2, .ScaleHeight - 1), FirstCornerPixelD
        
        'Top right corner
        PSet (.ScaleWidth - 1, 2), FirstCornerPixelD
        PSet (.ScaleWidth - 3, 0), FirstCornerPixelD
        
        'Bottom right corner
        PSet (.ScaleWidth - 3, .ScaleHeight - 1), FirstCornerPixelD
        PSet (.ScaleWidth - 1, .ScaleHeight - 3), FirstCornerPixelD
        
    'Second set of pixels.
        'Upper Left Corner
        PSet (1, 0), SecondCornerPixelD
        PSet (0, 1), SecondCornerPixelD
        
        'Bottom left corner
        PSet (0, .ScaleHeight - 2), SecondCornerPixelD
        PSet (1, .ScaleHeight - 1), SecondCornerPixelD
        
        'Top right corner
        PSet (.ScaleWidth - 1, 1), SecondCornerPixelD
        PSet (.ScaleWidth - 2, 0), SecondCornerPixelD
        
        'Bottom right corner
        PSet (.ScaleWidth - 2, .ScaleHeight - 1), SecondCornerPixelD
        PSet (.ScaleWidth - 1, .ScaleHeight - 2), SecondCornerPixelD
    
    'Third pixel.
        PSet (1, 1), ThirdCornerPixelD
        PSet (1, .ScaleHeight - 2), ThirdCornerPixelD
        PSet (.ScaleWidth - 2, 1), ThirdCornerPixelD
        PSet (.ScaleWidth - 2, .ScaleHeight - 2), ThirdCornerPixelD
    
    'Fourth set of pixels.
        'Upper left corner.
            PSet (2, 1), FourthCornerPixelD
            PSet (1, 2), FourthCornerPixelD
        
        'Bottom left corner.
            PSet (1, .ScaleHeight - 3), FourthCornerPixelD
            PSet (2, .ScaleHeight - 2), FourthCornerPixelD
        
        'Bottom right corner.
            PSet (.ScaleWidth - 3, .ScaleHeight - 2), FourthCornerPixelD
            PSet (.ScaleWidth - 2, .ScaleHeight - 3), FourthCornerPixelD
        
        'Top right corner.
            PSet (.ScaleWidth - 3, 1), FourthCornerPixelD
            PSet (.ScaleWidth - 2, 2), FourthCornerPixelD
End With
End Sub
Private Sub DrawDisabled()
Dim Gradient As New Collection
Dim SCaption() As String
Dim EndCaption As String

With UserControl
    'Draws border lines (not corners)
        Line (3, 0)-(.ScaleWidth - 3, 0), BorderColorLinesX
        Line (0, 3)-(0, .ScaleHeight - 3), BorderColorLinesX
        Line (3, .ScaleHeight - 1)-(.ScaleWidth - 3, .ScaleHeight - 1), BorderColorLinesX
        Line (.ScaleWidth - 1, 3)-(.ScaleWidth - 1, .ScaleHeight - 3), BorderColorLinesX
    
    'Draws the fade at the bottom.
        Line (1, .ScaleHeight - 4)-(.ScaleWidth - 1, .ScaleHeight - 4), FirstBottomLineX
        Line (2, .ScaleHeight - 3)-(.ScaleWidth - 2, .ScaleHeight - 3), SecondBottomLineX
        Line (3, ScaleHeight - 2)-(.ScaleWidth - 3, .ScaleHeight - 2), ThirdBottomLineX
        
    'Draws the background gradient.
        Set Gradient = CreateFade(FromColorFadeX, ToColorFadeX, .ScaleHeight - 5)
        
        For X = 1 To Gradient.Count
            Select Case X
            Case 1
                Line (3, X + 1)-(.ScaleWidth - 4, X + 1), Gradient(X)
            Case 2
                Line (2, X + 1)-(.ScaleWidth - 3, X + 1), Gradient(X)
            Case Else
                Line (1, X + 1)-(.ScaleWidth - 2, X + 1), Gradient(X)
            End Select
        Next
        
    'Draws corners
    'First set of pixels
        'Upper Left Corner
        PSet (2, 0), FirstCornerPixelX
        PSet (0, 2), FirstCornerPixelX
        
        'Bottom left corner
        PSet (0, .ScaleHeight - 3), FirstCornerPixelX
        PSet (2, .ScaleHeight - 1), FirstCornerPixelX
        
        'Top right corner
        PSet (.ScaleWidth - 1, 2), FirstCornerPixelX
        PSet (.ScaleWidth - 3, 0), FirstCornerPixelX
        
        'Bottom right corner
        PSet (.ScaleWidth - 3, .ScaleHeight - 1), FirstCornerPixelX
        PSet (.ScaleWidth - 1, .ScaleHeight - 3), FirstCornerPixelX
        
    'Second set of pixels.
        'Upper Left Corner
        PSet (1, 0), SecondCornerPixelX
        PSet (0, 1), SecondCornerPixelX
        
        'Bottom left corner
        PSet (0, .ScaleHeight - 2), SecondCornerPixelX
        PSet (1, .ScaleHeight - 1), SecondCornerPixelX
        
        'Top right corner
        PSet (.ScaleWidth - 1, 1), SecondCornerPixelX
        PSet (.ScaleWidth - 2, 0), SecondCornerPixelX
        
        'Bottom right corner
        PSet (.ScaleWidth - 2, .ScaleHeight - 1), SecondCornerPixelX
        PSet (.ScaleWidth - 1, .ScaleHeight - 2), SecondCornerPixelX
    
    'Third pixel.
        PSet (1, 1), ThirdCornerPixelX
        PSet (1, .ScaleHeight - 2), ThirdCornerPixelX
        PSet (.ScaleWidth - 2, 1), ThirdCornerPixelX
        PSet (.ScaleWidth - 2, .ScaleHeight - 2), ThirdCornerPixelX
    
    'Fourth set of pixels.
        'Upper left corner.
            PSet (2, 1), FourthCornerPixelX
            PSet (1, 2), FourthCornerPixelX
        
        'Bottom left corner.
            PSet (1, .ScaleHeight - 3), FourthCornerPixelX
            PSet (2, .ScaleHeight - 2), FourthCornerPixelX
        
        'Bottom right corner.
            PSet (.ScaleWidth - 3, .ScaleHeight - 2), FourthCornerPixelX
            PSet (.ScaleWidth - 2, .ScaleHeight - 3), FourthCornerPixelX
        
        'Top right corner.
            PSet (.ScaleWidth - 3, 1), FourthCornerPixelX
            PSet (.ScaleWidth - 2, 2), FourthCornerPixelX
End With
End Sub
Private Sub DrawIdle()
Dim Gradient As New Collection
Dim SCaption() As String
Dim EndCaption As String

With UserControl
    'Draws border lines (not corners)
        Line (3, 0)-(.ScaleWidth - 3, 0), BorderColorLines
        Line (0, 3)-(0, .ScaleHeight - 3), BorderColorLines
        Line (3, .ScaleHeight - 1)-(.ScaleWidth - 3, .ScaleHeight - 1), BorderColorLines
        Line (.ScaleWidth - 1, 3)-(.ScaleWidth - 1, .ScaleHeight - 3), BorderColorLines
    
    'Draws the fade at the bottom.
        Line (1, .ScaleHeight - 4)-(.ScaleWidth - 1, .ScaleHeight - 4), FirstBottomLine
        Line (2, .ScaleHeight - 3)-(.ScaleWidth - 2, .ScaleHeight - 3), SecondBottomLine
        Line (3, ScaleHeight - 2)-(.ScaleWidth - 3, .ScaleHeight - 2), ThirdBottomLine
        
    'Draws the background gradient.
        Set Gradient = CreateFade(FromColorFade, ToColorFade, .ScaleHeight - 5)
        
        For X = 1 To Gradient.Count
            Select Case X
            Case 1
                Line (3, X + 1)-(.ScaleWidth - 4, X + 1), Gradient(X)
            Case 2
                Line (2, X + 1)-(.ScaleWidth - 3, X + 1), Gradient(X)
            Case Else
                Line (1, X + 1)-(.ScaleWidth - 2, X + 1), Gradient(X)
            End Select
        Next
        
    'Draws corners
    'First set of pixels
        'Upper Left Corner
        PSet (2, 0), FirstCornerPixel
        PSet (0, 2), FirstCornerPixel
        
        'Bottom left corner
        PSet (0, .ScaleHeight - 3), FirstCornerPixel
        PSet (2, .ScaleHeight - 1), FirstCornerPixel
        
        'Top right corner
        PSet (.ScaleWidth - 1, 2), FirstCornerPixel
        PSet (.ScaleWidth - 3, 0), FirstCornerPixel
        
        'Bottom right corner
        PSet (.ScaleWidth - 3, .ScaleHeight - 1), FirstCornerPixel
        PSet (.ScaleWidth - 1, .ScaleHeight - 3), FirstCornerPixel
        
    'Second set of pixels.
        'Upper Left Corner
        PSet (1, 0), SecondCornerPixel
        PSet (0, 1), SecondCornerPixel
        
        'Bottom left corner
        PSet (0, .ScaleHeight - 2), SecondCornerPixel
        PSet (1, .ScaleHeight - 1), SecondCornerPixel
        
        'Top right corner
        PSet (.ScaleWidth - 1, 1), SecondCornerPixel
        PSet (.ScaleWidth - 2, 0), SecondCornerPixel
        
        'Bottom right corner
        PSet (.ScaleWidth - 2, .ScaleHeight - 1), SecondCornerPixel
        PSet (.ScaleWidth - 1, .ScaleHeight - 2), SecondCornerPixel
    
    'Third pixel.
        PSet (1, 1), ThirdCornerPixel
        PSet (1, .ScaleHeight - 2), ThirdCornerPixel
        PSet (.ScaleWidth - 2, 1), ThirdCornerPixel
        PSet (.ScaleWidth - 2, .ScaleHeight - 2), ThirdCornerPixel
    
    'Fourth set of pixels.
        'Upper left corner.
            PSet (2, 1), FourthCornerPixel
            PSet (1, 2), FourthCornerPixel
        
        'Bottom left corner.
            PSet (1, .ScaleHeight - 3), FourthCornerPixel
            PSet (2, .ScaleHeight - 2), FourthCornerPixel
        
        'Bottom right corner.
            PSet (.ScaleWidth - 3, .ScaleHeight - 2), FourthCornerPixel
            PSet (.ScaleWidth - 2, .ScaleHeight - 3), FourthCornerPixel
        
        'Top right corner.
            PSet (.ScaleWidth - 3, 1), FourthCornerPixel
            PSet (.ScaleWidth - 2, 2), FourthCornerPixel
End With
End Sub

Private Sub DrawHot()
Dim Gradient As New Collection
Dim SCaption() As String
Dim EndCaption As String

With UserControl
    'Draws border lines (not corners)
        Line (3, 0)-(.ScaleWidth - 3, 0), BorderColorLinesH
        Line (0, 3)-(0, .ScaleHeight - 3), BorderColorLinesH
        Line (3, .ScaleHeight - 1)-(.ScaleWidth - 3, .ScaleHeight - 1), BorderColorLinesH
        Line (.ScaleWidth - 1, 3)-(.ScaleWidth - 1, .ScaleHeight - 3), BorderColorLinesH
    
    'Draws the fade at the bottom.
        Line (1, .ScaleHeight - 4)-(.ScaleWidth - 1, .ScaleHeight - 4), FirstBottomLineH
        Line (2, .ScaleHeight - 3)-(.ScaleWidth - 2, .ScaleHeight - 3), SecondBottomLineH
        Line (3, ScaleHeight - 2)-(.ScaleWidth - 3, .ScaleHeight - 2), ThirdBottomLineH
        
    'Draws the background gradient.
        Set Gradient = CreateFade(FromColorFadeH, ToColorFadeH, .ScaleHeight - 5)
        
        For X = 1 To Gradient.Count
            Select Case X
            Case 1
                Line (3, X + 1)-(.ScaleWidth - 4, X + 1), FirstTopLineH
            Case 2
                Line (2, X + 1)-(.ScaleWidth - 3, X + 1), SecondTopLineH
            Case Else
                Line (1, X + 1)-(.ScaleWidth - 2, X + 1), Gradient(X)
            End Select
        Next

    'Draws side gradients
        Set Gradient = CreateFade(SideFromColorFadeH, SideToColorFadeH, .ScaleHeight - 7)
        
        For X = 1 To Gradient.Count
            PSet (1, X + 3), Gradient(X)
            PSet (2, X + 3), Gradient(X)
            
            PSet (.ScaleWidth - 2, X + 3), Gradient(X)
            PSet (.ScaleWidth - 3, X + 3), Gradient(X)
        Next
        
    'Draws corners
    'First set of pixels
        'Upper Left Corner
        PSet (2, 0), FirstCornerPixelH
        PSet (0, 2), FirstCornerPixelH
        
        'Bottom left corner
        PSet (0, .ScaleHeight - 3), FirstCornerPixelH
        PSet (2, .ScaleHeight - 1), FirstCornerPixelH
        
        'Top right corner
        PSet (.ScaleWidth - 1, 2), FirstCornerPixelH
        PSet (.ScaleWidth - 3, 0), FirstCornerPixelH
        
        'Bottom right corner
        PSet (.ScaleWidth - 3, .ScaleHeight - 1), FirstCornerPixelH
        PSet (.ScaleWidth - 1, .ScaleHeight - 3), FirstCornerPixelH
        
    'Second set of pixels.
        'Upper Left Corner
        PSet (1, 0), SecondCornerPixelH
        PSet (0, 1), SecondCornerPixelH
        
        'Bottom left corner
        PSet (0, .ScaleHeight - 2), SecondCornerPixelH
        PSet (1, .ScaleHeight - 1), SecondCornerPixelH
        
        'Top right corner
        PSet (.ScaleWidth - 1, 1), SecondCornerPixelH
        PSet (.ScaleWidth - 2, 0), SecondCornerPixelH
        
        'Bottom right corner
        PSet (.ScaleWidth - 2, .ScaleHeight - 1), SecondCornerPixelH
        PSet (.ScaleWidth - 1, .ScaleHeight - 2), SecondCornerPixelH
    
    'Third pixel.
        PSet (1, 1), ThirdCornerPixelH
        PSet (1, .ScaleHeight - 2), ThirdCornerPixelH
        PSet (.ScaleWidth - 2, 1), ThirdCornerPixelH
        PSet (.ScaleWidth - 2, .ScaleHeight - 2), ThirdCornerPixelH
    
    'Fourth set of pixels.
        'Upper left corner.
            PSet (2, 1), FourthCornerPixelH
            PSet (1, 2), FourthCornerPixelH
        
        'Bottom left corner.
            PSet (1, .ScaleHeight - 3), FourthCornerPixelH
            PSet (2, .ScaleHeight - 2), FourthCornerPixelH
        
        'Bottom right corner.
            PSet (.ScaleWidth - 3, .ScaleHeight - 2), FourthCornerPixelH
            PSet (.ScaleWidth - 2, .ScaleHeight - 3), FourthCornerPixelH
        
        'Top right corner.
            PSet (.ScaleWidth - 3, 1), FourthCornerPixelH
            PSet (.ScaleWidth - 2, 2), FourthCornerPixelH
End With
End Sub

Private Sub Timer1_Timer()
Dim X As POINTAPI
GetCursorPos X
If WindowFromPoint(X.X, X.Y) <> UserControl.HWND Then
    MouseOver = False
Else
    MouseOver = True
End If

End Sub

Private Sub UserControl_DblClick()
RaiseEvent DblClick
End Sub

Private Sub UserControl_EnterFocus()
HasFocus = True
Redraw
End Sub

Private Sub UserControl_ExitFocus()
HasFocus = False
Redraw
End Sub


Private Sub UserControl_InitProperties()
Caption = Ambient.DisplayName
Set Font = UserControl.Parent.Font
Enabled = True
ForeColor = vbBlack
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseDown = True
Redraw
RaiseEvent MouseDown(Button, Shift, X, Y)
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
If PropEnabled = True Then RaiseEvent Click
MouseDown = False
Redraw
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
Set Font = PropBag.ReadProperty("Font", UserControl.Parent.Font)
Enabled = PropBag.ReadProperty("Enabled", True)
ForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
End Sub

Private Sub UserControl_Resize()
Redraw
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "Caption", PropCaption, Ambient.DisplayName
PropBag.WriteProperty "Font", UserControl.Font, UserControl.Parent.Font
PropBag.WriteProperty "Enabled", PropEnabled, True
PropBag.WriteProperty "ForeColor", PropForeColor, vbBlack
End Sub

Public Property Let HasHand(ByVal vNewValue As Boolean)
If vNewValue = True Then
    UserControl.MousePointer = 99
End If

End Property

