VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REGION_TABLE 编辑器"
   ClientHeight    =   4335
   ClientLeft      =   3795
   ClientTop       =   3540
   ClientWidth     =   6570
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6570
   Begin VB.PictureBox Wallpaper2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   3000
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   1200
      ScaleWidth      =   1470
      TabIndex        =   20
      Top             =   2520
      Width           =   1470
      Begin VB.Image Imgicon2 
         Appearance      =   0  'Flat
         Height          =   165
         Index           =   6
         Left            =   1235
         Picture         =   "frmMain.frx":0DE2
         Top             =   0
         Width           =   210
      End
      Begin VB.Image Imgicon2 
         Appearance      =   0  'Flat
         Height          =   165
         Index           =   5
         Left            =   950
         Picture         =   "frmMain.frx":0E72
         Top             =   0
         Width           =   285
      End
      Begin VB.Image Imgicon2 
         Appearance      =   0  'Flat
         Height          =   165
         Index           =   4
         Left            =   769
         Picture         =   "frmMain.frx":0F16
         Top             =   0
         Width           =   180
      End
      Begin VB.Image Imgicon2 
         Appearance      =   0  'Flat
         Height          =   165
         Index           =   3
         Left            =   618
         Picture         =   "frmMain.frx":0F89
         Top             =   0
         Width           =   150
      End
      Begin VB.Image Imgicon2 
         Appearance      =   0  'Flat
         Height          =   165
         Index           =   2
         Left            =   452
         Picture         =   "frmMain.frx":0FEB
         Top             =   0
         Width           =   165
      End
      Begin VB.Image Imgicon2 
         Appearance      =   0  'Flat
         Height          =   165
         Index           =   1
         Left            =   286
         Picture         =   "frmMain.frx":1047
         Top             =   0
         Width           =   165
      End
      Begin VB.Image Imgicon2 
         Appearance      =   0  'Flat
         Height          =   165
         Index           =   0
         Left            =   0
         Picture         =   "frmMain.frx":10B4
         Top             =   0
         Width           =   285
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5400
      Top             =   0
   End
   Begin VB.CheckBox chkIconBS 
      Caption         =   "图标显示边框"
      Height          =   255
      Left            =   4560
      TabIndex        =   15
      Top             =   480
      Width           =   1815
   End
   Begin VB.PictureBox Wallpaper 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawMode        =   1  'Blackness
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   3300
      Left            =   240
      Picture         =   "frmMain.frx":1132
      ScaleHeight     =   3300
      ScaleWidth      =   2640
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   480
      Width           =   2640
      Begin VB.Image Imgicon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   9
         Left            =   2310
         Picture         =   "frmMain.frx":4841
         Top             =   0
         Width           =   330
      End
      Begin VB.Image Imgicon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   8
         Left            =   2025
         Picture         =   "frmMain.frx":4901
         Top             =   0
         Width           =   285
      End
      Begin VB.Image Imgicon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   7
         Left            =   1755
         Top             =   0
         Width           =   270
      End
      Begin VB.Image Imgicon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   6
         Left            =   1515
         Picture         =   "frmMain.frx":4991
         Top             =   0
         Width           =   240
      End
      Begin VB.Image Imgicon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   5
         Left            =   1245
         Picture         =   "frmMain.frx":4A3C
         Top             =   0
         Width           =   270
      End
      Begin VB.Image Imgicon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   4
         Left            =   975
         Picture         =   "frmMain.frx":4AEC
         Top             =   0
         Width           =   270
      End
      Begin VB.Image Imgicon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   3
         Left            =   780
         Picture         =   "frmMain.frx":4B77
         Top             =   0
         Width           =   195
      End
      Begin VB.Image Imgicon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   2
         Left            =   540
         Picture         =   "frmMain.frx":4BE4
         Top             =   0
         Width           =   240
      End
      Begin VB.Image Imgicon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   1
         Left            =   330
         Picture         =   "frmMain.frx":4C62
         Top             =   0
         Width           =   210
      End
      Begin VB.Image Imgicon 
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   0
         Left            =   0
         Picture         =   "frmMain.frx":4CE6
         Top             =   0
         Width           =   330
      End
   End
   Begin VB.VScrollBar VSXY 
      Height          =   375
      Index           =   3
      Left            =   6000
      TabIndex        =   8
      Top             =   1680
      Width           =   255
   End
   Begin VB.VScrollBar VSXY 
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   4
      Top             =   1680
      Width           =   255
   End
   Begin VB.VScrollBar VSXY 
      Height          =   375
      Index           =   2
      Left            =   6000
      TabIndex        =   6
      Top             =   960
      Width           =   255
   End
   Begin VB.VScrollBar VSXY 
      Height          =   375
      Index           =   0
      Left            =   4200
      TabIndex        =   2
      Top             =   960
      Width           =   255
   End
   Begin VB.Frame Fra_time_area 
      Caption         =   "时间相关"
      Height          =   1335
      Left            =   3000
      TabIndex        =   13
      Top             =   2400
      Width           =   3375
      Begin VB.TextBox txthint 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "调用字体"
         Top             =   840
         Width           =   735
      End
      Begin VB.ComboBox cbbfontNO 
         Height          =   300
         Left            =   960
         TabIndex        =   18
         Text            =   "请选择字体编号"
         Top             =   840
         Width           =   2295
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
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "time"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdtimeColor 
         Caption         =   "颜色"
         Height          =   375
         Left            =   1920
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox txtXY 
      Height          =   390
      Index           =   3
      Left            =   5520
      TabIndex        =   7
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtXY 
      Height          =   375
      Index           =   2
      Left            =   5520
      TabIndex        =   5
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtXY 
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   3
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtXY 
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   1
      Top             =   960
      Width           =   495
   End
   Begin VB.ComboBox cbbSelect 
      Height          =   300
      Left            =   3000
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5880
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label LblXY 
      Caption         =   "高度"
      Height          =   255
      Index           =   3
      Left            =   4920
      TabIndex        =   12
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label LblXY 
      Caption         =   "宽度"
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   11
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label LblXY 
      Caption         =   "顶部"
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   10
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label LblXY 
      Caption         =   "左侧"
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   9
      Top             =   1080
      Width           =   495
   End
   Begin VB.Menu munFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu munNew 
         Caption         =   "新建(&N)"
         Begin VB.Menu munL7E398 
            Caption         =   "L7/E398"
         End
         Begin VB.Menu munV3I 
            Caption         =   "V3I"
         End
      End
      Begin VB.Menu munOpen 
         Caption         =   "打开(&O)"
      End
      Begin VB.Menu munSave 
         Caption         =   "保存(&S)"
      End
      Begin VB.Menu munSaveAs 
         Caption         =   "另存为(&A)"
      End
      Begin VB.Menu mun_ 
         Caption         =   "-"
      End
      Begin VB.Menu munExit 
         Caption         =   "退出(&E)"
      End
   End
   Begin VB.Menu munOther 
      Caption         =   "其他(&O)"
      Begin VB.Menu munSetting 
         Caption         =   "设置(&S)"
      End
      Begin VB.Menu munAbout 
         Caption         =   "关于(&A)"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DragX As Single, DragY As Single
Private Sub Form_UnLoad(Cancel As Integer)
    End
End Sub
Private Sub Form_Load()
cbbfontNO.AddItem GetINI("lng", "cbbfontNO00")
cbbfontNO.AddItem GetINI("lng", "cbbfontNO01")
cbbfontNO.AddItem GetINI("lng", "cbbfontNO02")
cbbfontNO.AddItem GetINI("lng", "cbbfontNO03")
cbbfontNO.AddItem GetINI("lng", "cbbfontNO04")
cbbfontNO.AddItem GetINI("lng", "cbbfontNO05")
cbbfontNO.AddItem GetINI("lng", "cbbfontNO06")
cbbfontNO.AddItem GetINI("lng", "cbbfontNO07")
cbbfontNO.AddItem GetINI("lng", "cbbfontNO08")
cbbfontNO.AddItem GetINI("lng", "cbbfontNO09")
cbbfontNO.AddItem GetINI("lng", "cbbfontNO0A")
cbbfontNO.AddItem GetINI("lng", "cbbfontNO0B")
cbbfontNO.AddItem GetINI("lng", "cbbfontNO0C")
cbbfontNO.AddItem GetINI("lng", "cbbfontNO0D")
cbbfontNO.AddItem GetINI("lng", "cbbfontNO0E")
cbbfontNO.AddItem GetINI("lng", "cbbfontNO0F")
cbbfontNO.AddItem GetINI("lng", "cbbfontNO10")
cbbfontNO.AddItem GetINI("lng", "cbbfontNO11")
cbbfontNO.AddItem GetINI("lng", "cbbfontNO12")
cbbfontNO.AddItem GetINI("lng", "cbbfontNO13")
cbbfontNO.AddItem GetINI("lng", "cbbfontNO14")
txttime.Text = Time
munSave.Enabled = False
munSaveAs.Enabled = False
Wallpaper.Enabled = False
Dim i As Byte
For i = 0 To 3
    txtXY(i).Enabled = False
    VSXY(i).Enabled = False
Next
Wallpaper2.Visible = False
Fra_time_area.Visible = False

apply_Picture
End Sub

Private Sub Timer1_Timer()
    txttime.Text = Time
End Sub

Private Sub munAbout_Click()
    frmAbout.Show
End Sub
Private Sub munSetting_Click()
    frmSetting.Show
End Sub
Private Sub munExit_Click()
    If MsgBox(GetINI("lng", "SavePrompt_MG"), vbYesNo, GetINI("lng", "Savetitle_MG")) = vbYes Then Call munSave_Click
    End
End Sub
Private Sub munSaveAs_Click()
'保存图标位置数据
CommonDialog1.FileName = "REGION_TABLE"
CommonDialog1.Filter = GetINI("lng", "munSaveAs_CF")
CommonDialog1.CancelError = True
On Error Resume Next
CommonDialog1.ShowSave
If CommonDialog1.FileName = "" Or Err.Number = 32755 Then Exit Sub
    Savepath = CommonDialog1.FileName
    SaveDAT (Savepath)
End Sub
Private Sub munSave_Click()
    If Savepath = "" Then Call munSaveAs_Click: Exit Sub
    SaveDAT (Savepath)
End Sub

Private Sub munOpen_Click()
CommonDialog1.Filter = GetINI("lng", "munOpen_CF")
CommonDialog1.CancelError = True
On Error Resume Next
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Or Err.Number = 32755 Then Exit Sub
'确定机型
Open CommonDialog1.FileName For Binary As #1
    Platform = IIf(LOF(1) <= 96, 1, 2)
Close #1
    OpenDAT (CommonDialog1.FileName)
    Savepath = CommonDialog1.FileName
    munSave.Enabled = True
    munSaveAs.Enabled = True
End Sub
Private Sub munL7E398_Click()
    Platform = 1
    OpenDAT (App.Path & "\L7E398.cfg")
    Savepath = ""
    munSaveAs.Enabled = True
    munSave.Enabled = True
End Sub
Private Sub munV3I_Click()
    Platform = 2
    OpenDAT (App.Path & "\V3I.cfg")
    Savepath = ""
    munSaveAs.Enabled = True
    munSave.Enabled = True
End Sub

'更改时间颜色
Private Sub cmdtimeColor_Click()
CommonDialog1.ShowColor
txttime.ForeColor = CommonDialog1.Color
End Sub
'更改时间字体编号提示
Private Sub cbbfontNO_Click()
Dim Response, scarcity$, nono$, warn$, advise$
scarcity = GetINI("lng", "scarcity1_MG") & Chr(13) & Chr(10) & GetINI("lng", "scarcity2_MG") & Chr(13) & Chr(10) & Chr(13) & Chr(10) & GetINI("lng", "scarcity3_MG")
'不建议设置此字体！"（原CG4字体包中）此组字体可能缺少以下字符：数字0-9、标点：（字母A M P）
nono = GetINI("lng", "nono_MG") '不能设置此字体！此编号不含任何字符
warn = GetINI("lng", "warn_MG") '警告！
advise = GetINI("lng", "advise_MG") '建议！

Select Case cbbfontNO.ListIndex
Case 4
    Response = MsgBox(scarcity, 0, advise)
Case 5
    Response = MsgBox(scarcity, 0, advise)
Case 9
    Response = MsgBox(scarcity, 0, advise)
Case 11
    Response = MsgBox(scarcity, 0, advise)
Case 14
    Response = MsgBox(scarcity, 0, advise)
Case 15
    Response = MsgBox(scarcity, 0, advise)
Case 16
    Response = MsgBox(nono, 0, warn)
    cbbfontNO.ListIndex = 0
Case 17
    Response = MsgBox(nono, 0, warn)
    cbbfontNO.ListIndex = 0
Case 18
    Response = MsgBox(scarcity, 0, advise)
Case 19
    Response = MsgBox(scarcity, 0, advise)
Case 20
    Response = MsgBox(scarcity, 0, advise)
End Select
End Sub
'显示位置图标数据
Private Sub cbbSelect_Click()
Dim i%
nowform = IIf(cbbSelect.ListIndex > 9, 2, 1)
For i = 0 To 3 '选择第几个，就把第几个数据从数组中读入TXT
    If nowform = 2 Then
        txtXY(i).Text = icon7(cbbSelect.ListIndex - 10, i)
    Else
        txtXY(i).Text = icon10(cbbSelect.ListIndex, i)
    End If
Next i
End Sub

'限制输入
Private Sub txtXY_KeyPress(Index As Integer, KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
If KeyAscii = 8 Then Exit Sub
If CInt(Mid(txtXY(Index).Text, 1, txtXY(Index).SelStart) & Chr(KeyAscii) & Mid(txtXY(Index).Text, txtXY(Index).SelStart + 1)) > 255 Then KeyAscii = 0
End Sub

'更改位置图标数据-文本调整体现-连接VSXY
Private Sub txtXY_Change(Index As Integer)
    If txtXY(Index).Text <> "" Then
        If txtXY(Index).Text >= 0 Then
            VSXY(Index).Value = txtXY(Index).Text
        End If
    End If
End Sub



'更改位置图标数据-箭头调整体现-连接txtXY,Imgicon1\2,数据
Private Sub VSXY_Change(Index As Integer)
txtXY(Index).Text = VSXY(Index).Value
If nowform = 2 Then
    Select Case Index
    Case 0
        Imgicon2(cbbSelect.ListIndex - 10).Left = VSXY(0).Value * 15
    Case 1
        Imgicon2(cbbSelect.ListIndex - 10).Top = VSXY(1).Value * 15
    Case 2
        Imgicon2(cbbSelect.ListIndex - 10).Width = VSXY(2).Value * 15
    Case 3
        Imgicon2(cbbSelect.ListIndex - 10).Height = VSXY(3).Value * 15
    End Select
    icon7(cbbSelect.ListIndex - 10, Index) = VSXY(Index).Value         '数据保存到数组
Else
    Select Case Index
    Case 0
        Imgicon(cbbSelect.ListIndex).Left = VSXY(0).Value * 15
    Case 1
        Imgicon(cbbSelect.ListIndex).Top = VSXY(1).Value * 15
    Case 2
        Imgicon(cbbSelect.ListIndex).Width = VSXY(2).Value * 15
    Case 3
        Imgicon(cbbSelect.ListIndex).Height = VSXY(3).Value * 15
    End Select
    icon10(cbbSelect.ListIndex, Index) = VSXY(Index).Value          '数据保存到数组
End If
End Sub
'更改位置图标数据-图标调整体现-连接txtXY
Private Sub Imgicon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
nowform = 1
cbbSelect.ListIndex = Index '下拉列表响应

Dim i%
For i = 0 To 9
    Imgicon(i).Enabled = False '设置图标不可用（这样拖动到其他图标上时不会影响坐标）
Next i

Imgicon(Index).Drag 1       '设置可拖动
DragX = X                   '鼠标在此图标上的X坐标
DragY = Y                   '鼠标在此图标上的Y坐标

'限制拖动区域：开始
With CurrentPoint
    .X = 0
    .Y = 0
End With
' find position on the screen (not the window)
RetValue = ClientToScreen(Wallpaper.hwnd, CurrentPoint) 'CurrentPoint是代表Wallpaper的坐标
With ClipRect
    .Top = CurrentPoint.Y + DragY \ Screen.TwipsPerPixelY '单位是像素（+）
    .Left = CurrentPoint.X + DragX \ Screen.TwipsPerPixelX
    .Right = CurrentPoint.X + 176 - (Imgicon(Index).Width - DragX) \ Screen.TwipsPerPixelX + 1 '+1修正边界问题
    .Bottom = CurrentPoint.Y + 220 - (Imgicon(Index).Height - DragY) \ Screen.TwipsPerPixelY + 1 '+1修正边界问题
End With ' clip it
RetValue = ClipCursor(ClipRect)

End Sub
Private Sub Wallpaper_DragDrop(Source As Control, X As Single, Y As Single)
Dim i%
For i = 0 To 9
    Imgicon(i).Enabled = True
Next i
'读入TXT
txtXY(0).Text = (X - DragX) \ 15
txtXY(1).Text = (Y - DragY) \ 15
    '边界超出问题
If (X - DragX) >= 0 And (Y - DragY) <= 0 Then
    txtXY(1).Text = 0
ElseIf (X - DragX) <= 0 And (Y - DragY) >= 0 Then
    txtXY(0).Text = 0
ElseIf (X - DragX) <= 0 And (Y - DragY) <= 0 Then
    txtXY(0).Text = 0
    txtXY(1).Text = 0
End If
RetValue = ClipCursorClear(0)
End Sub

Private Sub Wallpaper2_DragDrop(Source As Control, X As Single, Y As Single)
Dim i%
For i = 0 To 6
    Imgicon2(i).Enabled = True
Next i
'读入TXT
txtXY(0).Text = (X - DragX) \ 15
txtXY(1).Text = (Y - DragY) \ 15
    '边界超出问题
If (X - DragX) >= 0 And (Y - DragY) <= 0 Then
    txtXY(1).Text = 0
ElseIf (X - DragX) <= 0 And (Y - DragY) >= 0 Then
    txtXY(0).Text = 0
ElseIf (X - DragX) <= 0 And (Y - DragY) <= 0 Then
    txtXY(0).Text = 0
    txtXY(1).Text = 0
End If
RetValue = ClipCursorClear(0)
End Sub
'更改位置图标数据-图标调整体现-连接txtXY
Private Sub Imgicon2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
nowform = 2
cbbSelect.ListIndex = Index + 10 '下拉列表响应

Dim i%
For i = 0 To 6
    Imgicon2(i).Enabled = False '设置图标不可用（这样拖动到其他图标上时不会影响坐标）
Next i

Imgicon2(Index).Drag 1       '设置可拖动
DragX = X                   '鼠标在此图标上的X坐标
DragY = Y                   '鼠标在此图标上的Y坐标

'限制拖动区域：开始
With CurrentPoint
    .X = 0
    .Y = 0
End With
' find position on the screen (not the window)
RetValue = ClientToScreen(Wallpaper2.hwnd, CurrentPoint) 'CurrentPoint是代表Wallpaper的坐标
With ClipRect
    .Top = CurrentPoint.Y + DragY \ Screen.TwipsPerPixelY '单位是像素（+）
    .Left = CurrentPoint.X + DragX \ Screen.TwipsPerPixelX
    .Right = CurrentPoint.X + 98 - (Imgicon2(Index).Width - DragX) \ Screen.TwipsPerPixelX + 1 '+1修正边界问题
    .Bottom = CurrentPoint.Y + 80 - (Imgicon2(Index).Height - DragY) \ Screen.TwipsPerPixelY + 1 '+1修正边界问题
End With ' clip it
RetValue = ClipCursor(ClipRect)

End Sub


'图标显示边框
Private Sub chkIconBS_Click()
Dim i As Byte
For i = 0 To 9
    Imgicon(i).BorderStyle = chkIconBS.Value
Next i
For i = 0 To 6
    Imgicon2(i).BorderStyle = chkIconBS.Value
Next i
End Sub
