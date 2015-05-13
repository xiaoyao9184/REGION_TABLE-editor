VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   8295
   StartUpPosition =   3  '窗口缺省
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "图片路径"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Check3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Option2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "其他"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Check2"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Caption         =   "图标路径"
         Height          =   1215
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   2160
         Width           =   2295
         Begin VB.OptionButton Option1 
            Caption         =   "同背景图片目录"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   7
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton Option3 
            Caption         =   "默认(\icon)"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.OptionButton Option2 
         Caption         =   "加载外部图片"
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   840
         Width           =   1695
      End
      Begin VB.Frame Frame2 
         Caption         =   "背景图片"
         Height          =   2775
         Left            =   2880
         TabIndex        =   3
         Top             =   600
         Width           =   2655
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   255
         Left            =   -74640
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.CheckBox Check3 
         Caption         =   "根据图片大小 自动调整图标框大小"
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   1440
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

