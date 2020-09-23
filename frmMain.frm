VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00F8DDC0&
   BorderStyle     =   0  'None
   Caption         =   "SkyBlue"
   ClientHeight    =   7560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrOpen 
      Interval        =   1
      Left            =   480
      Top             =   7080
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   6660
      Left            =   7440
      ScaleHeight     =   444
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   -3960
      Visible         =   0   'False
      Width           =   9000
   End
   Begin VB.Timer tmrMove 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   7080
   End
   Begin VB.PictureBox picMaskO 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   720
      Index           =   2
      Left            =   960
      Picture         =   "frmMain.frx":000C
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   56
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   6240
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.PictureBox picSpriteO 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   720
      Index           =   2
      Left            =   0
      Picture         =   "frmMain.frx":01D6
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   56
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   6240
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.PictureBox picMaskO 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   1
      Left            =   840
      Picture         =   "frmMain.frx":1098
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   5520
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox picSpriteO 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   1
      Left            =   0
      Picture         =   "frmMain.frx":123A
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   5520
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox picMaskO 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   0
      Left            =   720
      Picture         =   "frmMain.frx":1F38
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   46
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   4680
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.PictureBox picSpriteO 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   0
      Left            =   0
      Picture         =   "frmMain.frx":210A
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   46
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   4680
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.PictureBox picEnemy3r 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   4800
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   82
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.PictureBox picMaskE3r 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   4800
      Picture         =   "frmMain.frx":2E7C
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   82
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.PictureBox picSpriteE3r 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   4800
      Picture         =   "frmMain.frx":7C3E
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   82
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2640
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.PictureBox picPlane2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1560
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox picSpriteP2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      Picture         =   "frmMain.frx":CA00
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2760
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox picMaskP2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      Picture         =   "frmMain.frx":E27E
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2160
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox picSpace 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   6660
      Left            =   7440
      ScaleHeight     =   444
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5160
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.PictureBox picBullet 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   0
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox picEnemy4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   6120
      ScaleHeight     =   58
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox picMaskE4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   6120
      Picture         =   "frmMain.frx":FAFC
      ScaleHeight     =   58
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox picSpriteE4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   6120
      Picture         =   "frmMain.frx":11DAE
      ScaleHeight     =   58
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox picSpriteBE 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1200
      Picture         =   "frmMain.frx":14060
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3360
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picBulletE 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1200
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2880
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picMaskBE 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1200
      Picture         =   "frmMain.frx":14252
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picPlane 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox picSpriteP 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      Picture         =   "frmMain.frx":14444
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox picMaskP 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      Picture         =   "frmMain.frx":1571A
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox picEnemy3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   3480
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   82
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.PictureBox picSpriteE3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   3480
      Picture         =   "frmMain.frx":169F0
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   82
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2640
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.PictureBox picMaskE3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   3480
      Picture         =   "frmMain.frx":1B7B2
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   82
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.PictureBox picEnemy2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1020
      Left            =   2280
      ScaleHeight     =   68
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox picSpriteE2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1020
      Left            =   2280
      Picture         =   "frmMain.frx":20574
      ScaleHeight     =   68
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2160
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox picMaskE2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1020
      Left            =   2280
      Picture         =   "frmMain.frx":24466
      ScaleHeight     =   68
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox picMaskB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   0
      Picture         =   "frmMain.frx":28358
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox picSpriteB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   0
      Picture         =   "frmMain.frx":28F5A
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox picEnemy1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   1200
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picSpriteE1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   1200
      Picture         =   "frmMain.frx":29B5C
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picMaskE1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   1200
      Picture         =   "frmMain.frx":2C67E
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picSky 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   6660
      Left            =   7440
      Picture         =   "frmMain.frx":2F1A0
      ScaleHeight     =   444
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2640
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.PictureBox picOpen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6660
      Left            =   0
      Picture         =   "frmMain.frx":328AC
      ScaleHeight     =   6660
      ScaleWidth      =   9000
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   0
      Width           =   9000
      Begin VB.Label lblOpen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":73D8E
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   0
         TabIndex        =   38
         Top             =   3120
         Width           =   14340
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "If any of the pictures is copyrighted or cannot be used freely in this program, PLEASE PLEASE let me know and I will remove it."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   720
      Left            =   1920
      TabIndex        =   37
      Top             =   6720
      Visible         =   0   'False
      Width           =   5415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbxBackGround As PictureBox

Private intBgW As Integer, intBgH As Integer

Private intBgMove As Integer

Private boolLeft As Integer, boolRight As Integer, boolUp As Integer, boolDown As Integer
Private boolShoot As Integer

Private xdir2 As Integer, ydir2(maxPlane) As Integer        '-For enemy2 move

Private frmMove As Integer, OldX As Single, OldY As Single

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyLeft
            boolLeft = True
        Case vbKeyRight
            boolRight = True
        Case vbKeyUp
            boolUp = True
        Case vbKeyDown
            boolDown = True
        Case vbKeyD
            boolShoot = True
        Case vbKeyPause
            tmrMove.Enabled = IIf(tmrMove.Enabled, False, True)
            tmrOpen.Enabled = IIf(tmrOpen.Enabled, False, True)
            Cls
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Dim rc As Long
        If tmrMove.Enabled = False Then
            PlaneA.SetPos PlaneA.Width, (intActHeight - PlaneA.Height) / 2, pbxBackGround.hDC
        End If
        tmrMove.Enabled = True
        tmrOpen.Enabled = False
'picBuffer.AutoRedraw = False
'picBuffer.Move 0, 0
'picBuffer.Visible = True
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyLeft
            boolLeft = False
        Case vbKeyRight
            boolRight = False
        Case vbKeyUp
            boolUp = False
        Case vbKeyDown
            boolDown = False
        Case vbKeyD
            boolShoot = False
    End Select
End Sub

Private Sub Form_Load()
    ChDir App.Path
    ChDrive App.Path
    Randomize

    Set pbxBackGround = picSky
    'Set pbxBackGround = picSpace

    Width = 9000
    Height = 6660
    picBuffer.Width = Width
    picBuffer.Height = Height
    intActWidth = picBuffer.ScaleWidth
    intActHeight = picBuffer.ScaleHeight

    intBgW = pbxBackGround.ScaleWidth
    intBgH = pbxBackGround.ScaleHeight
    
    lblOpen.Left = picOpen.ScaleWidth
    
    Initialize_BBP
    
    PlaneA.Initialize_Me PA2
End Sub

Private Sub Initialize_BBP()
    With BBP(PA)
        Set .MaskPic = picMaskP
        Set .SpritePic = picSpriteP
        Set .CopyPic = picPlane
        .Speed = ccPAMove
        .HP = PAHP
    End With
    With BBP(PA2)
        Set .MaskPic = picMaskP2
        Set .SpritePic = picSpriteP2
        Set .CopyPic = picPlane2
        .Speed = ccPAMove
        .HP = PAHP
    End With
    With BBP(BA)
        Set .MaskPic = picMaskB
        Set .SpritePic = picSpriteB
        Set .CopyPic = picBullet
        .Speed = ccBAMove
    End With
    With BBP(E1)
        Set .MaskPic = picMaskE1
        Set .SpritePic = picSpriteE1
        Set .CopyPic = picEnemy1
        .Speed = ccE1Move
        .HP = E1HP
    End With
    With BBP(E2)
        Set .MaskPic = picMaskE2
        Set .SpritePic = picSpriteE2
        Set .CopyPic = picEnemy2
        .Speed = ccE2Move
        .HP = E2HP
    End With
    With BBP(E3)
        Set .MaskPic = picMaskE3
        Set .SpritePic = picSpriteE3
        Set .CopyPic = picEnemy3
        .Speed = ccE3Move
        .HP = E3HP
    End With
    With BBP(E3r)
        Set .MaskPic = picMaskE3r
        Set .SpritePic = picSpriteE3r
        Set .CopyPic = picEnemy3r
        .Speed = ccE3Move
        .HP = E3HP
    End With
    With BBP(E4)
        Set .MaskPic = picMaskE4
        Set .SpritePic = picSpriteE4
        Set .CopyPic = picEnemy4
        .Speed = ccE4Move
        .HP = E4HP
    End With
    With BBP(BE)
        Set .MaskPic = picMaskBE
        Set .SpritePic = picSpriteBE
        Set .CopyPic = picBulletE
        .Speed = ccBEMove
    End With
End Sub

Private Sub picOpen_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If vbLeftButton = Button Then
        frmMove = True
        OldX = x
        OldY = Y
    End If
End Sub

Private Sub picOpen_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If vbLeftButton = Button And frmMove Then
        Move Left + (x - OldX), Top + (Y - OldY)
    End If
End Sub

Private Sub picOpen_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    frmMove = False
End Sub

Private Sub tmrMove_Timer()
    gameElement = gameElement + 1
    If gameElement = 1000 Then
        gameElement = 0
        Dim p As Integer, pTotal As Integer
        pTotal = PlaneA.Count
        For p = 1 To pTotal - 1
            PlaneE(p).Terminate_Me
            Set PlaneE(p) = Nothing
        Next
Debug.Print PlaneA.Count & "<-"
    End If
    Background_Move
    Enemy_Move
    Plane_Move
    Check_PlaneCollide
    Copy_Buffer
End Sub

Private Sub Background_Move()
Dim rc As Long
    '--Background
    rc = BitBlt(picBuffer.hDC, intBgMove, 0, intBgW, intBgH, pbxBackGround.hDC, 0, 0, vbSrcCopy)
    rc = BitBlt(picBuffer.hDC, intBgMove + intBgW, 0, intBgW, intBgH, pbxBackGround.hDC, 0, 0, vbSrcCopy)
    intBgMove = intBgMove - ccBgMove
    If intBgMove < -intBgW Then intBgMove = 0
End Sub

Private Sub Enemy_Move()
Dim totalP As Integer
Dim e As Integer, newB As Integer
Dim tmpT As Integer, tmpH As Integer

    totalP = PlaneA.Count
    Select Case gameElement
        Case 10, 30, 50, 70, 130, 150, 210, 250, 350, 450, 550, 650
            PlaneE(totalP).Initialize_Me E1
            PlaneE(totalP).SetPos intActWidth, (intActHeight - PlaneE(totalP).Height) * Rnd, picBuffer.hDC
        Case 80, 180, 320, 440, 480, 560, 600
            PlaneE(totalP).Initialize_Me E2
            PlaneE(totalP).SetPos _
                intActWidth - PlaneE(totalP).Width - 70 * Rnd, intActHeight, picBuffer.hDC
        Case 0
            PlaneE(totalP).Initialize_Me E3
            PlaneE(totalP).SetPos intActWidth, (intActHeight - PlaneE(totalP).Height) * Rnd, picBuffer.hDC
        Case 0
            PlaneE(totalP).Initialize_Me E4
            PlaneE(totalP).SetPos intActWidth, (intActHeight - PlaneE(totalP).Height) * Rnd, picBuffer.hDC
        Case Is > 750
            
    End Select
'/////////////////////////////////////////////////////////
    totalP = PlaneA.Count
    For e = 1 To totalP - 1
        If PlaneE(e).HP > 0 Then
    
            tmpT = PlaneE(e).Top
            tmpH = PlaneE(e).Height
            '-Determine if new bullet should be fired
            newB = IIf(tmpT + tmpH / 2 >= PlaneA.Top And _
                tmpT + tmpH / 2 <= PlaneA.Top + PlaneA.Height, True, False)
                
            Select Case PlaneE(e).Speed
            
                Case ccE1Move
                    PlaneE(e).Move -1, 0, picBuffer.hDC
                    
                Case ccE2Move
                    If gameElement < 750 Then
                        xdir2 = 0
                    Else
                        xdir2 = -1
                        ydir2(e) = 0
                    End If
                    If tmpT >= intActHeight Then ydir2(e) = -1
                    If tmpT < 0 Then
                        ydir2(e) = 1
                    End If
                    If tmpT + tmpH > intActHeight Then
                        ydir2(e) = -1
                    End If
                    PlaneE(e).Move xdir2, ydir2(e), picBuffer.hDC
                    If gameElement < 800 Then PlaneE(e).Shoot newB, -1, BE, picBuffer.hDC
                    
                Case ccE3Move
                    PlaneE(e).Move -1, 0, picBuffer.hDC
                    
                Case ccE4Move
                    PlaneE(e).Move -1, 0, picBuffer.hDC
                    
            End Select
        End If
    Next
End Sub

Private Sub Plane_Move()

    If PlaneA.HP = 0 Then Exit Sub

Dim intPMoveX As Integer, intPMoveY As Integer
Dim rc As Long
    '-Determine X dir
    If boolLeft Then
        intPMoveX = -1
        If PlaneA.Left + intPMoveX - ccMargin < 0 Then
            intPMoveX = 0
        End If
    ElseIf boolRight Then
        intPMoveX = 1
        If PlaneA.Left + PlaneA.Width + intPMoveX + ccMargin > intActWidth Then
            intPMoveX = 0
        End If
    Else
        intPMoveX = 0
    End If
    '-Determine Y dir
    If boolUp Then
        intPMoveY = -1
        If PlaneA.Top + intPMoveY - ccMargin < 0 Then
            intPMoveY = 0
        End If
    ElseIf boolDown Then
        intPMoveY = 1
        If PlaneA.Top + PlaneA.Height + intPMoveY + ccMargin > intBgH Then
            intPMoveY = 0
        End If
    Else
        intPMoveY = 0
    End If
    
    PlaneA.Move intPMoveX, intPMoveY, picBuffer.hDC
    PlaneA.Shoot boolShoot, 1, BA, picBuffer.hDC
    
End Sub

Private Sub Check_PlaneCollide()
Dim r1 As Rect, r2 As Rect
    '-Set Rect1
    r1.Height = PlaneA.Height
    r1.Left = PlaneA.Left
    r1.Top = PlaneA.Top
    r1.Width = PlaneA.Width
Dim e As Integer
    For e = 1 To PlaneA.Count - 1
        '-Set Rect2
        r2.Height = PlaneE(e).Height
        r2.Left = PlaneE(e).Left
        r2.Top = PlaneE(e).Top
        r2.Width = PlaneE(e).Width
        
        '//////////////////////////////////////////////////
Static boom(maxPlane) As Integer
        If rectCollide(r1, r2) Or boom(e) Then
            Draw_Explore r1, picMaskO(boom(e)), picSpriteO(boom(e)), picBuffer.hDC
            Draw_Explore r2, picMaskO(boom(e)), picSpriteO(boom(e)), picBuffer.hDC
            boom(e) = boom(e) + 1
            If boom(e) = 3 Then
                sndPlaySound "boom.wav", &H1
                boom(e) = 0
                PlaneA.HP = PlaneA.HP - 1
                PlaneE(e).HP = PlaneE(e).HP - 1
                
                If PlaneA.HP = 0 Then
                    MsgBox "Game over"
                    End
                End If
                
            End If
        End If
    Next
End Sub

Private Sub Copy_Buffer()
Dim rc As Long
    rc = BitBlt(hDC, 0, 0, Width, Height, picBuffer.hDC, 0, 0, vbSrcCopy)
End Sub

Private Sub tmrOpen_Timer()
    lblOpen.Left = lblOpen.Left - 50
    If lblOpen.Left + lblOpen.Width <= 0 Then lblOpen.Left = picOpen.ScaleWidth
End Sub
