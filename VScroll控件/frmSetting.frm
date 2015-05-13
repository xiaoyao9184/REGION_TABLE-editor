VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSetting 
   Caption         =   "����"
   ClientHeight    =   4425
   ClientLeft      =   4185
   ClientTop       =   3330
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   5730
   Begin VB.CommandButton cmdapply 
      Caption         =   "Ӧ��"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   3840
      Width           =   975
   End
   Begin VB.CheckBox chkPicture 
      Caption         =   "�����ⲿͼƬ"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Frame fraPicture 
      Caption         =   "ͼƬĿ¼"
      Height          =   2775
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   5055
      Begin VB.TextBox txtIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   1320
         Width           =   3135
      End
      Begin VB.CommandButton cmdIcon 
         Caption         =   "״̬��Ŀ¼"
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   2040
         Width           =   1095
      End
      Begin VB.OptionButton optPictureSuff 
         Caption         =   "Ĭ��"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton optPictureOther 
         Caption         =   "����Ŀ¼"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdWallpaper 
         Caption         =   "����ͼƬ"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "ѡ��Ŀ¼���κ�һ��GIF�ļ�����ָ����Ŀ¼Ϊ״̬��Ŀ¼"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1800
         Width           =   4695
      End
      Begin VB.Label Label2 
         Caption         =   "����Ŀ¼\icon\"
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "����ͬʱָ������ͼƬ��ͼ��Ŀ¼"
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   840
         Width           =   2775
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   3840
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
Close #3
load_cfg (App.Path & "\Config.cfg")

chkPicture.Value = PictureFT
If PictureFT = 0 Then
    optPictureSuff.Enabled = False
    optPictureOther.Enabled = False
    cmdWallpaper.Enabled = False
    cmdIcon.Enabled = False
    cmdapply.Enabled = False
End If
If PicturePath = 0 Then
    optPictureSuff.Value = True
    txtWallpaper.Text = App.Path & "\icon\Wallpaper.jpg"
    txtIcon.Text = App.Path & "\icon"
Else
    optPictureOther.Value = True
    txtWallpaper.Text = WallpaperPath
    txtIcon.Text = IconPath
End If
End Sub
Private Sub chkPicture_Click()
PictureFT = chkPicture.Value
If chkPicture.Value = 1 Then
    optPictureSuff.Enabled = True
    optPictureOther.Enabled = True
    If optPictureSuff.Value = False Then
        cmdWallpaper.Enabled = True
        cmdIcon.Enabled = True
    End If
    cmdapply.Enabled = True
Else
    optPictureSuff.Enabled = False
    optPictureOther.Enabled = False
    cmdWallpaper.Enabled = False
    cmdIcon.Enabled = False
    cmdapply.Enabled = False
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
CommonDialog1.Filter = "BMP(*.bmp)|*.bmp|GIF(*.gif)|*.gif|JPEG(*.jpg)|*.jpg"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
    txtWallpaper.Text = CommonDialog1.FileName
    txtWallpaper.ToolTipText = CommonDialog1.FileName
    WallpaperPath = CommonDialog1.FileName
End If
End Sub
Private Sub cmdIcon_Click()
CommonDialog1.Filter = "�κ�һ��״̬��ͼ��(*.gif)|*.gif"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
    txtIcon.Text = CurDir()
    txtIcon.ToolTipText = CurDir()
    IconPath = CurDir()
End If
End Sub
Private Sub cmdSave_Click()
Put #3, 82, CByte(PictureFT)
Put #3, 83, CByte(PicturePath)
Put #3, 84, WallpaperPath & Chr(13) & Chr(10)
Put #3, , IconPath & Chr(13) & Chr(10)
End Sub
Private Sub cmdapply_Click()
apply_Picture (PictureFT)
End Sub
