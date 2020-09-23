VERSION 5.00
Begin VB.Form frmHidden 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   Caption         =   "Form1"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8310
   Icon            =   "frmHidden.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   469
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   554
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picClientArea 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6780
      Left            =   15
      ScaleHeight     =   448
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   313
      TabIndex        =   0
      Top             =   -45
      Width           =   4755
   End
End
Attribute VB_Name = "frmHidden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Initialize()
picClientArea.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight

End Sub

Private Sub Form_Resize()
picClientArea.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
FreeObjects
ReadBinaryFile Me.Tag
End Sub

Private Sub Form_Unload(Cancel As Integer)
FreeObjects
End
End Sub
