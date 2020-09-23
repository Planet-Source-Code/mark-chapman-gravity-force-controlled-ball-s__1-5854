VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassBall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private lPrevTop As Long
Private lPrevLeft As Long
Private lPrevHspeed As Long
Private lPrevVspeed As Long

Private lColor As Long
Private lWidth As Long
Private lTop As Long
Private lLeft As Long
Private lHeight As Long
Private mlngHSpeed As Long
Private mlngVSpeed As Long

Property Get Color() As Long
Randomize
lColor = Int((16777215 - 0 + 1) * Rnd + 0)
Color = lColor
End Property
Property Get PrevHspeed() As Long
PrevHspeed = lPrevHspeed
End Property
Property Get PrevVspeed() As Long
PrevVspeed = lPrevVspeed
End Property
Property Get Prevtop() As Long
Prevtop = lPrevTop
End Property
Property Get PrevLeft() As Long
PrevLeft = lPrevLeft
End Property
Property Get Width() As Long
Width = lWidth
End Property
Property Get Top() As Long
Top = lTop
End Property
Property Get Left() As Long
Left = lLeft
End Property
Property Get Height() As Long
Height = lHeight
End Property
Property Let Top(Stuff As Long)
lTop = Stuff
End Property
Property Let Left(Stuff As Long)
lLeft = Stuff
End Property
Property Let HSpeed(Stuff As Long)
mlngHSpeed = Stuff
End Property
Property Get HSpeed() As Long
HSpeed = mlngHSpeed
End Property
Property Let VSpeed(Stuff As Long)
mlngVSpeed = Stuff
End Property
Property Get VSpeed() As Long
VSpeed = mlngVSpeed
End Property

Property Get CentreX()
CentreX = lLeft + (lWidth / 2)
End Property

Property Get CentreY()
CentreY = lTop + (lHeight / 2)
End Property

Public Sub Move()
Dim a As Integer

lPrevVspeed = VSpeed
lPrevHspeed = HSpeed
lPrevTop = lTop
lPrevLeft = lLeft

VSpeed = VSpeed + Gravity.mlngGravity
    
    If lTop + VSpeed + lHeight > 5000 Then
        VSpeed = -(VSpeed * 0.9)
        lTop = 5000 - lHeight
    Else
        lTop = lTop + VSpeed
    End If
    
    If HSpeed > 0 Then
        If lLeft + lWidth + HSpeed > 5000 Then
            HSpeed = -(HSpeed * 0.9)
            lLeft = 5000 - lWidth
        Else
        lLeft = lLeft + HSpeed
        End If
    ElseIf HSpeed < 0 Then
        If Left + HSpeed < 0 Then
            HSpeed = -(HSpeed * 0.9)
            lLeft = 20
        Else
            lLeft = lLeft + HSpeed
        End If
    End If
    
    If lTop <= 0 Then
    lTop = 0
    VSpeed = -Int(VSpeed * 0.8)
    End If

End Sub

Public Sub Moveback()
VSpeed = lPrevVspeed
HSpeed = lPrevHspeed
lTop = lPrevTop
lLeft = lPrevLeft
End Sub


Private Sub Class_Initialize()
HSpeed = 50
VSpeed = 0
lLeft = 195
lTop = 195
lWidth = 210
lHeight = 210
End Sub


