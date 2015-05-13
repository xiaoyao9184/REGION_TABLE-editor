VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSetting 
   Caption         =   "设置"
   ClientHeight    =   4425
   ClientLeft      =   4125
   ClientTop       =   3690
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   5730
   Begin VB.CommandButton Command1 
      Caption         =   "应用"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   3840
      Width           =   975
   End
   Begin VB.CheckBox chkPicture 
      Caption         =   "加载外部图片"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Frame fraPicture 
      Caption         =   "图片目录"
      Height          =   2775
      Left            =   360
      TabIndex        =   7
      Top             =   960
      Width           =   5055
      Begin VB.TextBox txtIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox txtWallpaper 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1680
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   11
         Top             =   1440
         Width           =   3135
      End
      Begin VB.CommandButton cmdIcon 
         Caption         =   "状态栏图标"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   2040
         Width           =   1215
      End
      Begin VB.OptionButton optPictureSuff 
         Caption         =   "默认"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton optPictureOther 
         Caption         =   "其他目录"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdWallpaper 
         Caption         =   "背景图片"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "程序目录\icon\"
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "仅选择背景图片即可           其他图标将在背景图片目录下搜索"
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   840
         Width           =   2775
      End
   End
   Begin VB.CheckBox chkAutosize 
      Caption         =   "根据图片大小 自动调整图标框大小"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4920
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PictureFT As Byte, PictureAS As Byte, PicturePath As Byte
Dim WallpaperPath$, IconPath$, nowPath$

Private Sub Form_Load()
Open App.Path & "\Config.cfg" For Binary As #3

Get #3, 81, PictureFT
Get #3, 82, PictureAS
Get #3, 83, PicturePath
Seek #3, 88
Line Input #3, WallpaperPath
Seek #3, Seek(3) + 4
Line Input #3, IconPath

chkPicture.Value = PictureFT
If PictureFT = 0 Then
    chkAutosize.Enabled = False
    optPictureSuff.Enabled = False
    optPictureOther.Enabled = False
    cmdWallpaper.Enabled = False
    cmdIcon.Enabled = False
End If
chkAutosize.Value = PictureAS
If PicturePath = 0 Then
    optPictureSuff.Value = True
    txtWallpaper.Text = App.Path & "\icon"
    txtIcon.Text = App.Path & "\icon"
Else
    optPictureOther.Value = True
    txtWallpaper.Text = WallpaperPath
    txtIcon.Text = IconPath
End If
End Sub
Private Sub chkPicture_Click()
If chkPicture.Value = 1 Then
    chkAutosize.Enabled = True
    optPictureSuff.Enabled = True
    optPictureOther.Enabled = True
    If optPictureSuff.Value = False Then
        cmdWallpaper.Enabled = True
        cmdIcon.Enabled = True
    End If
Else
    chkAutosize.Enabled = False
    optPictureSuff.Enabled = False
    optPictureOther.Enabled = False
    cmdWallpaper.Enabled = False
    cmdIcon.Enabled = False
End If
End Sub

Private Sub optPictureSuff_Click()
If optPictureSuff.Value = True Then
    PicturePath = 0
    cmdWallpaper.Enabled = False
    cmdIcon.Enabled = False
End If
End Sub
Private Sub optPictureOther_Click()
If optPictureOther.Value = True Then
    PicturePath = 1
    cmdWallpaper.Enabled = True
    cmdIcon.Enabled = True
End If
End Sub
Private Sub cmdWallpaper_Click()
CommonDialog1.Filter = "位图文件(*.bmp)|*.bmp|GIF(*.gif)|*.gif|JPEG|*.jpg"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
    txtWallpaper.Text = CommonDialog1.FileName
    txtWallpaper.ToolTipText = txtWallpaper.Text
    WallpaperPath = CommonDialog1.FileName
End If
End Sub
Private Sub cmdIcon_Click()
CommonDialog1.Filter = "任何一个图标文件(*.gif)|*.gif"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
    txtIcon.Text = CurDir()
    txtIcon.ToolTipText = txtIcon.Text
    IconPath = CurDir()
End If
End Sub
Private Sub cmdSave_Click()
Put #3, 81, CByte(chkPicture.Value)
Put #3, 82, CByte(chkAutosize.Value)
Put #3, 83, CByte(PicturePath)
Put #3, 84, WallpaperPath & Chr(13) & Chr(10)
Put #3, , IconPath & Chr(13) & Chr(10)
cmdSave.Enabled = False
End Sub
Private Sub Command1_Click()
If PictureFT = 1 Then
    If PictureAS = 1 Then
    '    FrmMain.Imgicon().Stretch = True
    Else
    '    FrmMain.Imgicon().Stretch = False
    End If
    If PicturePath = 1 Then
        nowPath = App.Path & "\icon"
    Else
        nowPath = OtherPath
    End If
    LoadAllPicture
End If
End Sub

Public Sub LoadAllPicture()
Wallpaper.Picture = LoadPicture(nowPath & "\Wallpaper.jpg")
'Imgicon(0).Picture = LoadPicture(nowPath & "\415.gif")
'Imgicon(1).Picture = LoadPicture(nowPath & "\404.gif")
'Imgicon(2).Picture = LoadPicture(nowPath & "\407.gif")
'Imgicon(3).Picture = LoadPicture(nowPath & "\473.gif")
'Imgicon(4).Picture = LoadPicture(nowPath & "\391.gif")
'Imgicon(5).Picture = LoadPicture(nowPath & "\457.gif")
'Imgicon(6).Picture = LoadPicture(nowPath & "\394.gif")
'Imgicon(8).Picture = LoadPicture(nowPath & "\416.gif")
'Imgicon(9).Picture = LoadPicture(nowPath & "\335.gif")
End Sub
