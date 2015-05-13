VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSetting 
   Caption         =   "设置"
   ClientHeight    =   5685
   ClientLeft      =   4185
   ClientTop       =   3330
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   6870
   Begin VB.CommandButton cmdapply 
      Caption         =   "应用"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   5160
      Width           =   975
   End
   Begin VB.CheckBox chkPicture 
      Caption         =   "加载外部图片"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Frame fraPicture 
      Caption         =   "图片目录"
      Height          =   4335
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   6375
      Begin VB.TextBox txtWallpaper2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   1680
         MousePointer    =   1  'Arrow
         TabIndex        =   17
         Top             =   3000
         Width           =   4455
      End
      Begin VB.CommandButton cmdWallpaper2 
         Caption         =   "外屏背景图片"
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtIcon2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1680
         MousePointer    =   1  'Arrow
         TabIndex        =   15
         Top             =   3720
         Width           =   4455
      End
      Begin VB.CommandButton cmdIcon2 
         Caption         =   "外屏目录"
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox txtIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   1680
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   11
         Top             =   2160
         Width           =   4455
      End
      Begin VB.TextBox txtWallpaper 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1680
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   10
         Top             =   1320
         Width           =   4455
      End
      Begin VB.CommandButton cmdIcon 
         Caption         =   "状态栏目录"
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   2040
         Width           =   1215
      End
      Begin VB.OptionButton optPictureSuff_Other 
         Caption         =   "默认"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton optPictureSuff_Other 
         Caption         =   "其他目录"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdWallpaper 
         Caption         =   "背景图片"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "仅用于V3/V3I"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2640
         Width           =   5895
      End
      Begin VB.Label Label3 
         Caption         =   "选择目录下任何一个GIF文件即可指定此目录为状态栏目录"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1800
         Width           =   6015
      End
      Begin VB.Label Label2 
         Caption         =   "程序目录\icon\"
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   "必须同时指定背景图片和图标目录"
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   840
         Width           =   4455
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   5040
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
Private Sub Form_Load()
chkPicture.Value = GetINI("Setting", "PictureFT")
End Sub
Private Sub chkPicture_Click()
If chkPicture.Value = 1 Then
    optPictureSuff_Other(0).Enabled = True
    optPictureSuff_Other(1).Enabled = True
    optPictureSuff_Other(1).Value = GetINI("Setting", "PictureSuff_Other")
Else
    optPictureSuff_Other(0).Enabled = False
    optPictureSuff_Other(1).Enabled = False
    cmdWallpaper.Enabled = False
    cmdIcon.Enabled = False
    cmdIcon2.Enabled = False
End If
End Sub

Private Sub optPictureSuff_Other_Click(Index As Integer)
If Index = 0 Then
    cmdWallpaper.Enabled = False
    cmdWallpaper2.Enabled = False
    cmdIcon.Enabled = False
    cmdIcon2.Enabled = False
    
    txtWallpaper.Text = App.Path & "\icon\Wallpaper.jpg"
    txtWallpaper2.Text = App.Path & "\icon\V3I\cl.gif"
    txtIcon.Text = App.Path & "\icon\"
    txtIcon2.Text = App.Path & "\icon\V3I\"
Else
    cmdWallpaper.Enabled = True
    cmdWallpaper2.Enabled = True
    cmdIcon.Enabled = True
    cmdIcon2.Enabled = True
    
    txtWallpaper.Text = GetINI("Setting", "WallpaperPath")
    txtWallpaper2.Text = GetINI("Setting", "Wallpaper2Path")
    txtIcon.Text = GetINI("Setting", "IconPath")
    txtIcon2.Text = GetINI("Setting", "Icon2Path")
End If

txtWallpaper.ToolTipText = txtWallpaper.Text
txtWallpaper2.ToolTipText = txtWallpaper2.Text
txtIcon.ToolTipText = txtIcon.Text
txtIcon2.ToolTipText = txtIcon2.Text
End Sub
Private Sub cmdWallpaper_Click()
CommonDialog1.Filter = "JPEG(*.jpg)|*.jpg|BMP(*.bmp)|*.bmp|GIF(*.gif)|*.gif"
CommonDialog1.CancelError = True
On Error Resume Next
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Or Err.Number = 32755 Then Exit Sub
    txtWallpaper.Text = CommonDialog1.FileName
    txtWallpaper.ToolTipText = CommonDialog1.FileName
End Sub
Private Sub cmdWallpaper2_Click()
CommonDialog1.Filter = "GIF(*.gif)|*.gif|JPEG(*.jpg)|*.jpg|BMP(*.bmp)|*.bmp"
CommonDialog1.CancelError = True
On Error Resume Next
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Or Err.Number = 32755 Then Exit Sub
    txtWallpaper2.Text = CommonDialog1.FileName
    txtWallpaper2.ToolTipText = CommonDialog1.FileName
End Sub
Private Sub cmdIcon_Click()
CommonDialog1.Filter = GetINI("lng", "cmdIcon_CF")
CommonDialog1.CancelError = True
On Error Resume Next
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Or Err.Number = 32755 Then Exit Sub
    txtIcon.Text = CurDir()
    txtIcon.ToolTipText = CurDir()
End Sub
Private Sub cmdIcon2_Click()
CommonDialog1.Filter = GetINI("lng", "cmdIcon_CF")
CommonDialog1.CancelError = True
On Error Resume Next
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Or Err.Number = 32755 Then Exit Sub
    txtIcon2.Text = CurDir()
    txtIcon2.ToolTipText = CurDir()
End Sub

Private Sub cmdSave_Click()
WriteINI "Setting", "PictureFT", chkPicture.Value
WriteINI "Setting", "PictureSuff_Other", CInt(optPictureSuff_Other(1).Value)
WriteINI "Setting", "WallpaperPath", txtWallpaper.Text
WriteINI "Setting", "Wallpaper2Path", txtWallpaper2.Text
WriteINI "Setting", "IconPath", txtIcon.Text
WriteINI "Setting", "Icon2Path", txtIcon2.Text
cmdapply.Enabled = True
End Sub
Private Sub cmdapply_Click()
apply_Picture
cmdapply.Enabled = False
End Sub
