Attribute VB_Name = "basGame"
Option Explicit

Public Const maxBullet = 32, maxPlane = 32

'-These 2 variables must not be modify outside class CPlane and CBullet
Public number_of_planes As Integer     '-As shared class property
Public number_of_bullets(maxPlane) As Integer     '-As shared class property

Public PlaneA As New CPlane
Public PlaneE(maxPlane) As New CPlane

Public Type Rect
    Left As Integer
    Top As Integer
    Width As Integer
    Height As Integer
End Type

Public Type BilBltPic
    MaskPic As PictureBox
    SpritePic As PictureBox
    CopyPic As PictureBox
    Speed As Integer
    HP As Integer
End Type
Public BBP(9) As BilBltPic
Public Const PA = 1, PA2 = 2, BA = 3, E1 = 4, E2 = 5, E3 = 6, E3r = 7, E4 = 8, BE = 9
Public Const PAHP = 3, E1HP = 2, E2HP = 5, E3HP = 6, E4HP = 6

Public gameElement As Integer

Public intActWidth As Integer, intActHeight As Integer

Public Const ccBgMove = 2
Public Const ccPAMove = 10
Public Const ccE1Move = 15, ccE2Move = 7, ccE3Move = 5, ccE4Move = 10
Public Const ccBAMove = 40, ccBEMove = 36

Public Const ccMargin = 20

Public Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, _
    ByVal YSrc As Long, ByVal dwRop As Long) As Long
    
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


Public Sub Draw_Explore(rectO As Rect, pbxBoomM As PictureBox, pbxBoomS As PictureBox, ContainerHDC As Long)
Dim x As Integer
Dim rc As Long

    rc = BitBlt(ContainerHDC, rectO.Left, rectO.Top, pbxBoomM.ScaleWidth, pbxBoomM.ScaleHeight, pbxBoomM.hDC, 0, 0, vbSrcAnd)
    rc = BitBlt(ContainerHDC, rectO.Left, rectO.Top, pbxBoomS.ScaleWidth, pbxBoomS.ScaleHeight, pbxBoomS.hDC, 0, 0, vbSrcInvert)

End Sub

Public Function rectCollide(Rect1 As Rect, Rect2 As Rect) As Boolean
    rectCollide = False
    If Rect1.Left + Rect1.Width > Rect2.Left Then
        If Rect1.Left < Rect2.Left + Rect2.Width Then
            If Rect1.Top + Rect1.Height > Rect2.Top Then
                If Rect1.Top < Rect2.Top + Rect2.Height Then
                    rectCollide = True
                End If
            End If
        End If
    End If
End Function
