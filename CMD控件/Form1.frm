VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   165
   ClientTop       =   885
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   6570
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdminus 
      Caption         =   "-"
      Height          =   180
      Index           =   3
      Left            =   6000
      TabIndex        =   23
      Top             =   1870
      Width           =   255
   End
   Begin VB.CommandButton cmdminus 
      Caption         =   "-"
      Height          =   180
      Index           =   2
      Left            =   6000
      TabIndex        =   22
      Top             =   1150
      Width           =   255
   End
   Begin VB.CommandButton cmdminus 
      Caption         =   "-"
      Height          =   180
      Index           =   1
      Left            =   4200
      TabIndex        =   21
      Top             =   1870
      Width           =   255
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "+"
      Height          =   180
      Index           =   3
      Left            =   6000
      TabIndex        =   20
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "+"
      Height          =   180
      Index           =   2
      Left            =   6000
      TabIndex        =   19
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "+"
      Height          =   180
      Index           =   1
      Left            =   4200
      TabIndex        =   18
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton cmdminus 
      Caption         =   "-"
      Height          =   180
      Index           =   0
      Left            =   4200
      TabIndex        =   17
      Top             =   1150
      Width           =   255
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "+"
      Height          =   180
      Index           =   0
      Left            =   4200
      TabIndex        =   16
      Top             =   960
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      FillColor       =   &H8000000F&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   300
      TabIndex        =   15
      Top             =   480
      Width           =   330
   End
   Begin VB.TextBox txttime 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3360
      TabIndex        =   14
      Text            =   "time"
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox txthint 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5040
      TabIndex        =   13
      Text            =   "调用字体"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox txtfontNO 
      Height          =   375
      Left            =   5040
      TabIndex        =   12
      Text            =   "fontNO."
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdtimeColor 
      Caption         =   "颜色"
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   3000
      Width           =   615
   End
   Begin VB.Frame Fra_time_correlation 
      Caption         =   "时间相关"
      Height          =   1335
      Left            =   3120
      TabIndex        =   10
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox txtXY 
      Height          =   375
      Index           =   3
      Left            =   5400
      TabIndex        =   5
      Text            =   "0"
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox txtXY 
      Height          =   375
      Index           =   2
      Left            =   5400
      TabIndex        =   4
      Text            =   "0"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtXY 
      Height          =   375
      Index           =   1
      Left            =   3840
      TabIndex        =   3
      Text            =   "0"
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox txtXY 
      Height          =   375
      Index           =   0
      Left            =   3840
      TabIndex        =   2
      Text            =   "0"
      Top             =   960
      Width           =   375
   End
   Begin VB.ComboBox cbbSelect 
      Height          =   300
      Left            =   3120
      TabIndex        =   1
      Text            =   "请选择图标"
      Top             =   480
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Fra_icon_area 
      Height          =   3015
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label LblXY 
      Caption         =   "高度"
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   9
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label LblXY 
      Caption         =   "宽度"
      Height          =   255
      Index           =   2
      Left            =   4800
      TabIndex        =   8
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label LblXY 
      Caption         =   "顶部"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   7
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label LblXY 
      Caption         =   "左侧"
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   6
      Top             =   1080
      Width           =   495
   End
   Begin VB.Menu munFile 
      Caption         =   "文件"
      Begin VB.Menu munOpen 
         Caption         =   "打开"
      End
      Begin VB.Menu munBuild 
         Caption         =   "生成"
      End
      Begin VB.Menu munExit 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu munAbout 
      Caption         =   "关于"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
cbbSelect.AddItem "信号"
cbbSelect.AddItem "GPRS"
cbbSelect.AddItem "数据"
cbbSelect.AddItem "漫游"
cbbSelect.AddItem "拨号"
cbbSelect.AddItem "JAVA"
cbbSelect.AddItem "短信"
cbbSelect.AddItem "时间"
cbbSelect.AddItem "铃音"
cbbSelect.AddItem "电量"
txttime.Text = Time
End Sub
Private Sub munOpen_Click()
Dim i%, ONE_right As Byte, ONE_bottom As Byte, TRGB_Color$
CommonDialog1.Filter = "状态栏时间各种颜色的位置文件|REGION_TABLE"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Binary As #1
    '读取图标位置数据
    For i = 1 To 76 Step 8
        Get #1, i + 1, one(0) '_left
        Get #1, i + 3, one(1) '_top
        Get #1, i + 5, ONE_right
        Get #1, i + 7, ONE_bottom
        one(2) = ONE_right - one(0)
        one(3) = ONE_bottom - one(1)
        load_data (i)
    Next i
    '读取时间颜色，显示时间颜色
    For i = 1 To 3
        Get #1, 84 + i, RGBt(i)
    Next i
    Get #1, 88, RGBt(0)
    TRGB_Color = "&H00000000"
    For i = 0 To 3
     If RGBt(i) <= 15 Then
        Mid(TRGB_Color, (i + 1) * 2 + 1 + 1, 2) = Hex(RGBt(i))
     Else
        Mid(TRGB_Color, (i + 1) * 2 + 1, 2) = Hex(RGBt(i))
     End If
    Next i
    txttime.ForeColor = TRGB_Color
    '读取时间字体编号，显示时间字体编号
    Get #1, 96, fontNO
    txtfontNO.Text = Hex(fontNO)
End If
End Sub
'更改时间颜色
Private Sub cmdtimeColor_Click()
CommonDialog1.ShowColor
txttime.ForeColor = CommonDialog1.Color
End Sub
'更改图标位置数据+
Private Sub cmdadd_Click(Index As Integer)
Select Case Index
Case 0
icon10(SelectListNO, 0) = icon10(SelectListNO, 0) + 1
txtXY(0).Text = txtXY(0).Text + 1
Case 1
icon10(SelectListNO, 1) = icon10(SelectListNO, 1) + 1
txtXY(1).Text = txtXY(1).Text + 1
Case 2
icon10(SelectListNO, 2) = icon10(SelectListNO, 2) + 1
txtXY(2).Text = txtXY(2).Text + 1
Case 3
icon10(SelectListNO, 3) = icon10(SelectListNO, 3) + 1
txtXY(3).Text = txtXY(3).Text + 1
End Select
End Sub
'更改图标位置数据-
Private Sub cmdminus_Click(Index As Integer)
Select Case Index
Case 0
    If icon10(SelectListNO, 0) >= 1 Then
        icon10(SelectListNO, 0) = icon10(SelectListNO, 0) - 1
        txtXY(0).Text = txtXY(0).Text - 1
    End If
Case 1
    If icon10(SelectListNO, 1) >= 1 Then
        icon10(SelectListNO, 1) = icon10(SelectListNO, 1) - 1
        txtXY(1).Text = txtXY(1).Text - 1
    End If
Case 2
    If icon10(SelectListNO, 2) >= 1 Then
        icon10(SelectListNO, 2) = icon10(SelectListNO, 2) - 1
        txtXY(2).Text = txtXY(2).Text - 1
    End If
Case 3
    If icon10(SelectListNO, 3) >= 1 Then
        icon10(SelectListNO, 3) = icon10(SelectListNO, 3) - 1
        txtXY(3).Text = txtXY(3).Text - 1
    End If
End Select
End Sub
'显示位置图标数据
Private Sub cbbSelect_Click()
Dim q%
Select Case cbbSelect.ListIndex
Case 0
    SelectListNO = 0
    show_data (SelectListNO)
Case 1
    SelectListNO = 1
    show_data (SelectListNO)
Case 2
    SelectListNO = 2
    show_data (SelectListNO)
Case 3
    SelectListNO = 3
    show_data (SelectListNO)
Case 4
    SelectListNO = 4
    show_data (SelectListNO)
Case 5
    SelectListNO = 5
    show_data (SelectListNO)
Case 6
    SelectListNO = 6
    show_data (SelectListNO)
Case 7
    SelectListNO = 7
    show_data (SelectListNO)
Case 8
    SelectListNO = 8
    show_data (SelectListNO)
Case 9
    SelectListNO = 9
    show_data (SelectListNO)
End Select
End Sub
'显示/更改位置图标数据（子过程）
Public Sub show_data(SelectListNO)
Dim q%
For q = 0 To 3
    txtXY(q).Text = icon10(SelectListNO, q)
Next q
End Sub
