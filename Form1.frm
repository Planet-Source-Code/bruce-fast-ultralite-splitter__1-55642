VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton Splitter1 
      BackColor       =   &H00FFFF00&
      Height          =   3255
      Left            =   1440
      MousePointer    =   9  'Size W E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   75
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Code by Bruce Fast, submitted to the public domain.

Private Capturing As Boolean
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Sub Form_Resize()
    Splitter1.Height = ScaleHeight
    DoMove
End Sub

Private Sub Splitter1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim a1&

    If Button = 1 Then 'The mouse is down
        If Capturing = False Then
            Splitter1.ZOrder
            SetCapture Splitter1.hwnd
            Capturing = True
        End If
        With Splitter1
            a1 = .Left + X
            If MoveOk(a1) Then
                .Left = a1
            End If
        End With
    End If
    
    
End Sub

Private Sub Splitter1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Capturing Then
        ReleaseCapture
        Capturing = False
        DoMove
    End If
    
End Sub

'******************************************************************************
' Modify below to make the splitter responsive
'******************************************************************************

Private Sub DoMove()    'Put in actions to be taken when the splitter is moved
Dim Right&  'The right hand edge of the splitter

    Right = Splitter1.Left + Splitter1.Width
    
    'Center the two buttons horizontally in their 'paine'
    Command1.Left = (Splitter1.Left / 2) - (Command1.Width / 2)
    Command2.Left = Right + ((ScaleWidth - Right) / 2) - (Command2.Width / 2)
End Sub


Private Function MoveOk(X&) As Boolean  'Put in any limiters you desire
    MoveOk = False
    If X > Command1.Width And X < ScaleWidth - Command2.Width Then
        MoveOk = True
    End If
End Function
