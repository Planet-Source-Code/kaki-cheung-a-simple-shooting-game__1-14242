VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPlane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private virPlane As Rect
Private MaskImage As PictureBox
Private SpriteImage As PictureBox
Private CopyImage As PictureBox
Private pKey As Integer
Private pSpeed As Integer
Private intHP As Integer
Private pTag As Integer

Private virBullet(maxBullet) As New CBullet

Public Property Get Key() As Integer
    Key = pKey
End Property

Public Property Get Count() As Integer
    Count = number_of_planes
End Property

Public Property Let Left(L As Integer)
    virPlane.Left = L
End Property

Public Property Get Left() As Integer
    Left = virPlane.Left
End Property

Public Property Let Top(T As Integer)
    virPlane.Top = T
End Property

Public Property Get Top() As Integer
    Top = virPlane.Top
End Property

Public Property Let Width(W As Integer)
    virPlane.Width = W
End Property

Public Property Get Width() As Integer
    Width = virPlane.Width
End Property

Public Property Let Height(H As Integer)
    virPlane.Height = H
End Property

Public Property Get Height() As Integer
    Height = virPlane.Height
End Property

Public Property Get Speed() As Integer
    Speed = pSpeed
End Property

Public Property Let HP(HP As Integer)
    intHP = HP
    If intHP <= 0 Then
        virPlane.Left = -999
        virPlane.Top = -999
        virPlane.Width = 0
        virPlane.Height = 0
    End If
End Property

Public Property Get HP() As Integer
    HP = intHP
End Property

Public Property Let Tag(T As Integer)
    pTag = T
End Property

Public Property Get Tag() As Integer
    Tag = pTag
End Property

Public Sub Initialize_Me(Index As Integer)
    With BBP(Index)
        Set MaskImage = .MaskPic
        Set SpriteImage = .SpritePic
        Set CopyImage = .CopyPic
        pSpeed = .Speed
        intHP = .HP
    End With
    virPlane.Width = CopyImage.ScaleWidth
    virPlane.Height = CopyImage.ScaleHeight
    
    number_of_planes = number_of_planes + 1
    pKey = number_of_planes
    
    Set virBullet(0) = New CBullet
    virBullet(0).Initialize_Me pKey
End Sub

Public Sub Terminate_Me()
    number_of_planes = number_of_planes - 1
Dim b As Integer, numB As Integer
    numB = virBullet(0).Count
    For b = 0 To numB
        virBullet(b).Terminate_Me
        Set virBullet(b) = Nothing
    Next
End Sub

Public Sub SetPos(StartX As Integer, StartY As Integer, ContainerHDC As Long)
Dim rc As Long
    rc = BitBlt(CopyImage.hDC, 0, 0, virPlane.Width, virPlane.Height, ContainerHDC, StartX, StartY, vbSrcCopy)
    virPlane.Left = StartX
    virPlane.Top = StartY
End Sub

Public Sub Move(Xdir As Integer, Ydir As Integer, ContainerHDC As Long)
    If virPlane.Left + virPlane.Width <= 0 Then Exit Sub

Dim NewX As Integer, NewY As Integer
Dim rc As Long
    
    NewX = virPlane.Left + Xdir * pSpeed
    NewY = virPlane.Top + Ydir * pSpeed

'Now perform "transparent" BitBlts:
'1 Copy old background back over sprite's old position
'2 Copy the background where the sprite will go
'3 Draw the mask
'4 Draw the sprite

    'rc = BitBlt(ContainerHDC, virPlane.Left - ccBgMove, virPlane.Top, virPlane.Width, virPlane.Height, frmMain.picPlane.hDC, 0, 0, vbSrcCopy)
    'rc = BitBlt(CopyImage.hDC, 0, 0, virPlane.Width, virPlane.Height, ContainerHDC, NewX, NewY, vbSrcCopy)
    rc = BitBlt(ContainerHDC, NewX, NewY, virPlane.Width, virPlane.Height, MaskImage.hDC, 0, 0, vbSrcAnd)
    rc = BitBlt(ContainerHDC, NewX, NewY, virPlane.Width, virPlane.Height, SpriteImage.hDC, 0, 0, vbSrcInvert)
    
    virPlane.Left = NewX
    virPlane.Top = NewY
End Sub

Public Sub Shoot(NewBullet As Integer, Xdir As Integer, PicIndex As Integer, ContainerHDC As Long)
Dim totalB As Integer
Dim b As Integer
    totalB = virBullet(0).Count
    If NewBullet Then
        If (totalB = 1 Or gameElement Mod 2 <> 0) Then
            virBullet(totalB).Initialize_Me pKey
            virBullet(totalB).SetPic PicIndex
            
            Dim tmpX As Integer
            If Xdir = 1 Then
                tmpX = virPlane.Left + virPlane.Width
            ElseIf Xdir = -1 Then
                tmpX = virPlane.Left
            End If
            virBullet(totalB).SetPos Xdir, tmpX, _
                virPlane.Top + (virPlane.Height - virBullet(totalB).Height) / 2, ContainerHDC
            totalB = virBullet(totalB).Count
        End If
    End If
    For b = 1 To totalB - 1
        virBullet(b).Move Xdir, ccBAMove, ContainerHDC
    Next
Debug.Print totalB
End Sub
