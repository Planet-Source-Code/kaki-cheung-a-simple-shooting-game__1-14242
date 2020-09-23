VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBullet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private virBullet As Rect
Private MaskImage As PictureBox
Private SpriteImage As PictureBox
Private CopyImage As PictureBox
Private bKey As Integer
Private bSpeed As Integer

Public Property Get Count() As Integer
    Count = number_of_bullets(bKey)
End Property

Public Property Get Height() As Integer
    Height = virBullet.Height
End Property

Public Property Get Speed() As Integer
    Speed = bSpeed
End Property

Public Sub Initialize_Me(Key As Integer)
    bKey = Key
    number_of_bullets(bKey) = number_of_bullets(bKey) + 1
    If number_of_bullets(bKey) = maxBullet Then number_of_bullets(bKey) = 1
End Sub

Public Sub SetPic(Index As Integer)
    Set MaskImage = BBP(Index).MaskPic
    Set SpriteImage = BBP(Index).SpritePic
    Set CopyImage = BBP(Index).CopyPic
    bSpeed = BBP(Index).Speed
    virBullet.Width = CopyImage.ScaleWidth
    virBullet.Height = CopyImage.ScaleHeight
End Sub

Public Sub Terminate_Me()
    number_of_bullets(bKey) = number_of_bullets(bKey) - 1
End Sub

Public Sub SetPos(Xdir As Integer, StartX As Integer, StartY As Integer, ContainerHDC As Long)
Dim rc As Long
    rc = BitBlt(CopyImage.hDC, 0, 0, virBullet.Width, virBullet.Height, ContainerHDC, StartX, StartY, vbSrcCopy)
    virBullet.Left = StartX
    virBullet.Top = StartY
End Sub

Public Sub Move(Xdir As Integer, Speed As Integer, ContainerHDC As Long)

If virBullet.Left > intActWidth Then Exit Sub

Dim NewX As Integer
Dim rc As Long
    rc = BitBlt(ContainerHDC, virBullet.Left, virBullet.Top, virBullet.Width, virBullet.Height, MaskImage.hDC, 0, 0, vbSrcAnd)
    rc = BitBlt(ContainerHDC, virBullet.Left, virBullet.Top, virBullet.Width, virBullet.Height, SpriteImage.hDC, 0, 0, vbSrcInvert)
    virBullet.Left = virBullet.Left + bSpeed * Xdir
    
    
    '///////////////////////////////////////////////////////////////
Dim r2 As Rect
Dim e As Integer
Static boom(maxPlane) As Integer

    If bKey = 1 Then
        For e = 1 To PlaneA.Count - 1
            '-Set r2
            r2.Left = PlaneE(e).Left
            r2.Top = PlaneE(e).Top
            r2.Width = PlaneE(e).Width
            r2.Height = PlaneE(e).Height
                
            If rectCollide(virBullet, r2) Or boom(e) Then
                Draw_Explore r2, frmMain.picMaskO(boom(e)), frmMain.picSpriteO(boom(e)), frmMain.picBuffer.hDC
                boom(e) = boom(e) + 1
                If boom(e) = 3 Then
                    sndPlaySound "boom.wav", &H1
                    boom(e) = 0
                    PlaneE(e).HP = PlaneE(e).HP - 1
                End If
            End If
        Next
    Else
        '-Set r2
        r2.Left = PlaneA.Left
        r2.Top = PlaneA.Top
        r2.Width = PlaneA.Width
        r2.Height = PlaneA.Height
        
        If rectCollide(virBullet, r2) Or boom(0) Then
            Draw_Explore r2, frmMain.picMaskO(boom(0)), frmMain.picSpriteO(boom(0)), frmMain.picBuffer.hDC
            boom(0) = boom(0) + 1
            If boom(0) = 3 Then
                sndPlaySound "boom.wav", &H1
                boom(0) = 0
                PlaneA.HP = PlaneA.HP - 1
                
                If PlaneA.HP = 0 Then
                    MsgBox "Game over"
                    End
                End If
                
            End If
        End If
            
    End If
    
End Sub
