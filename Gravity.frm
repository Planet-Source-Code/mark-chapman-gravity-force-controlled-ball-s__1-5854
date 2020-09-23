VERSION 5.00
Begin VB.Form Gravity 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bouncy Thing"
   ClientHeight    =   5040
   ClientLeft      =   3885
   ClientTop       =   1410
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   7050
   Begin VB.CommandButton Command2 
      Caption         =   "Decrease"
      Height          =   405
      Left            =   5265
      TabIndex        =   1
      Top             =   585
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Increase"
      Height          =   405
      Left            =   5265
      TabIndex        =   0
      Top             =   195
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400000&
      Height          =   5085
      Left            =   0
      ScaleHeight     =   5025
      ScaleWidth      =   5025
      TabIndex        =   10
      Top             =   0
      Width           =   5085
      Begin VB.Shape Ball 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   210
         Index           =   0
         Left            =   195
         Shape           =   3  'Circle
         Top             =   195
         Width           =   210
      End
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Reset"
      Height          =   210
      Left            =   5265
      TabIndex        =   2
      Top             =   1170
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Shake"
      Height          =   405
      Left            =   5265
      TabIndex        =   7
      Top             =   3900
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Add Ball"
      Height          =   405
      Left            =   5265
      TabIndex        =   3
      Top             =   1950
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Force <----------"
      Height          =   405
      Left            =   5265
      TabIndex        =   6
      Top             =   3315
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Force ---------->"
      Height          =   405
      Left            =   5265
      TabIndex        =   5
      Top             =   2925
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Reset Balls"
      Height          =   405
      Left            =   5265
      TabIndex        =   4
      Top             =   2340
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   405
      Left            =   5265
      TabIndex        =   8
      Top             =   4485
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Left            =   195
      Top             =   3315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "Gravity Control"
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   5265
      TabIndex        =   9
      Top             =   1560
      Width           =   1575
   End
End
Attribute VB_Name = "Gravity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private colBalls As New Collection
Public mlngGravity As Long
Private Amount As Integer


Private Sub Command1_Click() 'increase gravity
mlngGravity = mlngGravity + 5
End Sub

Private Sub Command2_Click() 'gravity decrease
mlngGravity = mlngGravity - 2
If mlngGravity <= 0 Then
    mlngGravity = 2
    MsgBox "No gravity is unexeceptable on this planet"
End If
End Sub

Private Sub Command3_Click() 'exit program
Unload Me
End Sub

Private Sub Command4_Click()
Dim a As Integer
    
For a = 1 To Ball.UBound
    Unload Ball(a) ' removes shape
Next a
If colBalls.Count > 2 Then
For a = colBalls.Count To 2 Step -1
    colBalls.Remove (a) 'removes from collection instance of class ball
Next a
End If
Amount = 0
Ball(0).Left = 195
Ball(0).Top = 195
colBalls(1).Left = 195
colBalls(1).Top = 195

End Sub

Private Sub Command5_Click() 'increase force on all balls to the right
Dim a As Integer
For a = 0 To Ball.UBound
    colBalls(a + 1).HSpeed = colBalls(a + 1).HSpeed + 5
Next a
End Sub

Private Sub Command6_Click() 'increase force on all balls to the left
Dim a As Integer
For a = 0 To Ball.UBound
    colBalls(a + 1).HSpeed = colBalls(a + 1).HSpeed - 5
Next a
End Sub

Private Sub Command7_Click() ' shake
Dim a As Integer
For a = 0 To Ball.UBound
colBalls(a + 1).VSpeed = colBalls(a + 1).VSpeed + 300 + Rnd(500)
colBalls(a + 1).HSpeed = colBalls(a + 1).HSpeed + 100 + Rnd(300)
Next a
End Sub

Private Sub Command8_Click() 'Add upto 6 balls
Dim objBall As ClassBall
Randomize
If Amount < 9 Then
    Load Ball(Ball.UBound + 1)
    Ball(Ball.UBound).Visible = True
    Set objBall = New ClassBall
    colBalls.Add objBall
    Ball(Ball.UBound).Left = colBalls(Ball.UBound).Left
    Ball(Ball.UBound).Top = colBalls(Ball.UBound).Top
End If
Amount = Amount + 1

End Sub

Private Sub Command9_Click()
mlngGravity = 10
End Sub

Private Sub Form_Load()
Dim objBall As ClassBall

Timer1.Interval = 20
mlngGravity = 10
Amount = 0

Set objBall = New ClassBall
colBalls.Add objBall
Command8.Enabled = True

End Sub

Private Sub Form_Paint()
'Line (0, 0)-(5000, 0), &HFFFF& ' paints edge of area balls within
'Line (0, 0)-(0, 5000), &HFFFF&
'Line (0, 5000)-(5000, 5000), &HFFFF&
'Line (5000, 0)-(5000, 5000), &HFFFF&
End Sub

Private Sub Timer1_Timer()
Dim a As Integer
Dim b As Integer
Dim temp As Long
Dim Bool As Boolean
Dim c As Integer

If (Ball(Ball.UBound).Left > 195 + Ball(Ball.UBound).Width) Or (Ball(Ball.UBound).Top > 195 + Ball(Ball.UBound).Height) Then
    If Amount < 9 Then
        Command8.Enabled = True
    End If
Else
Command8.Enabled = False
End If

For a = 0 To Ball.UBound
c = 0
colBalls(a + 1).Move

For b = 0 To Ball.UBound   ' Do they hit ? if they do calc direction changes/speed changes
If (a + 1) <> (b + 1) Then
    If Abs(Sqr((colBalls(a + 1).CentreX - colBalls(b + 1).CentreX) ^ 2 + (colBalls(a + 1).CentreY - colBalls(b + 1).CentreY) ^ 2)) < Ball(a).Width Then
        temp = Abs(Sqr((colBalls(a + 1).CentreX - colBalls(b + 1).CentreX) ^ 2 + (colBalls(a + 1).CentreY - colBalls(b + 1).CentreY) ^ 2))
        If temp < 210 Then ' if a ball overlaps it stops it...
            colBalls(a + 1).Moveback
            While Bool = True
            If (colBalls(a + 1).Left > colBalls(a + 1).PrevLeft) Then
                colBalls(a + 1).Left = colBalls(a + 1).PrevLeft + (colBalls(a + 1).Left + colBalls(a + 1).PrevLeft) / (3 + c)
            Else
                colBalls(a + 1).Left = colBalls(a + 1).PrevLeft - (colBalls(a + 1).PrevLeft - colBalls(a + 1).Left) / (3 + c)
            End If
            If colBalls(a + 1).Top > colBalls(a + 1).Prevtop Then
                colBalls(a + 1).Top = colBalls(a + 1).Prevtop + (colBalls(a + 1).Top + colBalls(a + 1).Prevtop) / (3 + c)
            Else
                colBalls(a + 1).Top = colBalls(a + 1).Prevtop - (colBalls(a + 1).Prevtop - colBalls(a + 1).Top) / (3 + c)
            End If
            temp = Abs(Sqr((colBalls(a + 1).CentreX - colBalls(b + 1).CentreX) ^ 2 + (colBalls(a + 1).CentreY - colBalls(b + 1).CentreY) ^ 2))
            If temp < 210 Then
                If (colBalls(a + 1).HSpeed > colBalls(b + 1).HSpeed) Then
                    colBalls(a + 1).HSpeed = -(colBalls(a + 1).HSpeed)
                    Bool = False
                ElseIf (colBalls(a + 1).VSpeed > colBalls(b + 1).VSpeed) Then
                    colBalls(a + 1).VSpeed = -(colBalls(a + 1).VSpeed)
                    Bool = False
                Else
                    Bool = True
                    c = c + 2
                End If
            Else
                Bool = False
            End If
            Wend
        End If ' ----
      
    'same direction going left
    If colBalls(a + 1).HSpeed > 0 And colBalls(b + 1).HSpeed > 0 Then
        If colBalls(a + 1).Left < colBalls(b + 1).Left Then
            colBalls(a + 1).HSpeed = -(colBalls(a + 1).HSpeed * 0.9)
            colBalls(b + 1).HSpeed = (colBalls(b + 1).HSpeed * 0.9) + (colBalls(a + 1).HSpeed / 4)
        Else
            colBalls(b + 1).HSpeed = -(colBalls(b + 1).HSpeed * 0.9)
            colBalls(a + 1).HSpeed = (colBalls(a + 1).HSpeed * 0.9) + (colBalls(b + 1).HSpeed / 4)
        End If
    'same direction going right
      ElseIf colBalls(a + 1).HSpeed < 0 And colBalls(b + 1).HSpeed < 0 Then
        If colBalls(a + 1).Left > colBalls(b + 1).Left Then
            colBalls(a + 1).HSpeed = -(colBalls(a + 1).HSpeed * 0.9)
            colBalls(b + 1).HSpeed = (colBalls(b + 1).HSpeed * 0.9) - (colBalls(a + 1).HSpeed / 4)
        Else
            colBalls(b + 1).HSpeed = -(colBalls(b + 1).HSpeed * 0.9)
            colBalls(a + 1).HSpeed = (colBalls(a + 1).HSpeed * 0.9) - (colBalls(b + 1).HSpeed / 4)
        End If
    'balls meeting from opposite x directions
    ElseIf colBalls(a + 1).HSpeed > 0 And colBalls(b + 1).HSpeed < 0 Or colBalls(a + 1).HSpeed < 0 And colBalls(b + 1).HSpeed > 0 Then
        If colBalls(a + 1).HSpeed < 0 Then
        colBalls(a + 1).HSpeed = -(colBalls(a + 1).HSpeed * 0.9 - colBalls(b + 1).HSpeed / 4)
        Else
        colBalls(a + 1).HSpeed = -(colBalls(a + 1).HSpeed * 0.9 + colBalls(b + 1).HSpeed / 4)
        End If
        If colBalls(b + 1).HSpeed < 0 Then
        colBalls(b + 1).HSpeed = -(colBalls(b + 1).HSpeed * 0.9 - colBalls(a + 1).HSpeed / 4)
        Else
        colBalls(b + 1).HSpeed = -(colBalls(b + 1).HSpeed * 0.9 + colBalls(a + 1).HSpeed / 4)
        End If
    End If
    
 
    'balls both from top
    If colBalls(a + 1).VSpeed > 0 And colBalls(b + 1).VSpeed > 0 Then
        If colBalls(a + 1).Top > colBalls(b + 1).Top Then
            colBalls(b + 1).VSpeed = -(colBalls(b + 1).VSpeed * 0.9)
            colBalls(a + 1).VSpeed = colBalls(a + 1).VSpeed + (colBalls(b + 1).VSpeed / 4)
        Else
            colBalls(a + 1).VSpeed = -(colBalls(a + 1).VSpeed * 0.9)
            colBalls(b + 1).VSpeed = colBalls(b + 1).VSpeed + (colBalls(a + 1).VSpeed / 4)
        End If
    'balls both from bottom
    ElseIf colBalls(a + 1).VSpeed < 0 And colBalls(b + 1).VSpeed < 0 Then
        If colBalls(a + 1).Top < colBalls(b + 1).Top Then
            colBalls(b + 1).VSpeed = -(colBalls(b + 1).VSpeed * 0.9)
            colBalls(a + 1).VSpeed = colBalls(a + 1).VSpeed - (colBalls(b + 1).VSpeed / 4)
        Else
            colBalls(a + 1).VSpeed = -(colBalls(a + 1).VSpeed * 0.9)
            colBalls(b + 1).VSpeed = colBalls(b + 1).VSpeed - (colBalls(a + 1).VSpeed / 4)
        End If
    'balls from opposite y values
    ElseIf colBalls(a + 1).VSpeed > 0 And colBalls(b + 1).VSpeed < 0 Or colBalls(a + 1).VSpeed < 0 And colBalls(b + 1).VSpeed > 0 Then
        If colBalls(a + 1).VSpeed < 0 Then
        colBalls(a + 1).VSpeed = -(colBalls(a + 1).VSpeed * 0.9 - colBalls(b + 1).VSpeed / 4)
        Else
        colBalls(a + 1).VSpeed = -(colBalls(a + 1).VSpeed * 0.9 + colBalls(b + 1).VSpeed / 4)
        End If
        If colBalls(b + 1).VSpeed < 0 Then
        colBalls(b + 1).VSpeed = -(colBalls(b + 1).VSpeed * 0.9 - colBalls(a + 1).VSpeed / 4)
        Else
        colBalls(b + 1).VSpeed = -(colBalls(b + 1).VSpeed * 0.9 + colBalls(a + 1).VSpeed / 4)
        End If
    End If
    
    

    Ball(a).FillColor = colBalls(a + 1).Color
    Ball(b).FillColor = colBalls(b + 1).Color
    End If
End If
Next b
'Shows the move
Picture1.Visible = False
Ball(a).Left = colBalls(a + 1).Left
Ball(a).Top = colBalls(a + 1).Top


Next a
Picture1.Visible = True


End Sub


