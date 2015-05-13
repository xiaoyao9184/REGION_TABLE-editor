Attribute VB_Name = "Module1"
Option Explicit

Public Platform As Byte '机型（1=L7/E398，2=V3/V3I）
Public Savepath$ '保存路径
Public RGBt(3) As Byte
Public one(3) As Byte
Public icon10(9, 3) As Integer '主屏
Public icon7(6, 3) As Integer '外屏

Public nowform As Byte

'限制拖动区域
Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Declare Function ClipCursorClear Lib "user32" Alias "ClipCursor" (ByVal lpRect As Long) As Long
Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Type POINTAPI
    X As Long
    Y As Long
End Type
Public CurrentPoint As POINTAPI '坐标
Public ClipRect As RECT
Public RetValue As Long


Public Sub Main()
    FrmMain.Show
    '取得命令行参数
    If Len(Command()) <> 0 Then
        Dim iName As String '路径
        iName = Replace(Command(), Chr(34), "") '替换"为空
        OpenDAT (iName)
        Savepath = iName
        FrmMain.munSave.Enabled = True
        FrmMain.munSaveAs.Enabled = True
    End If
FrmMain.Caption = GetINI("lng", "MainCaption")
FrmMain.munFile.Caption = GetINI("lng", "munFile")
FrmMain.munNew.Caption = GetINI("lng", "munNew")
FrmMain.munL7E398.Caption = GetINI("lng", "munL7E398")
FrmMain.munV3I.Caption = GetINI("lng", "munV3I")
FrmMain.munOpen.Caption = GetINI("lng", "munOpen")
FrmMain.munSave.Caption = GetINI("lng", "munSave")
FrmMain.munSaveAs.Caption = GetINI("lng", "munSaveAs")
FrmMain.munExit.Caption = GetINI("lng", "munExit")
FrmMain.munOther.Caption = GetINI("lng", "munOther")
FrmMain.munSetting.Caption = GetINI("lng", "munSetting")
FrmMain.munAbout.Caption = GetINI("lng", "munAbout")
FrmMain.chkIconBS.Caption = GetINI("lng", "chkIconBS")
FrmMain.LblXY(0).Caption = GetINI("lng", "LblXY_Left")
FrmMain.LblXY(1).Caption = GetINI("lng", "LblXY_Width")
FrmMain.LblXY(2).Caption = GetINI("lng", "LblXY_Top")
FrmMain.LblXY(3).Caption = GetINI("lng", "LblXY_Hight")
FrmMain.Fra_time_area.Caption = GetINI("lng", "Fra_time_area")
FrmMain.cmdtimeColor.Caption = GetINI("lng", "cmdtimeColor")
FrmMain.txthint.Text = GetINI("lng", "txthint")


frmSetting.Caption = GetINI("lng", "SettingCaption")
frmSetting.chkPicture.Caption = GetINI("lng", "chkPicture")
frmSetting.fraPicture.Caption = GetINI("lng", "fraPicture")
frmSetting.optPictureSuff_Other(0).Caption = GetINI("lng", "optPictureSuff")
frmSetting.optPictureSuff_Other(1).Caption = GetINI("lng", "optPictureOther")
frmSetting.Label1.Caption = GetINI("lng", "Label1")
frmSetting.Label2.Caption = GetINI("lng", "Label2")
frmSetting.Label3.Caption = GetINI("lng", "Label3")
frmSetting.Label4.Caption = GetINI("lng", "Label4")
frmSetting.cmdWallpaper.Caption = GetINI("lng", "cmdWallpaper")
frmSetting.cmdIcon.Caption = GetINI("lng", "cmdIcon")
frmSetting.cmdWallpaper2.Caption = GetINI("lng", "cmdWallpaper2")
frmSetting.cmdIcon2.Caption = GetINI("lng", "cmdIcon2")
frmSetting.cmdSave.Caption = GetINI("lng", "cmdSave")
frmSetting.cmdapply.Caption = GetINI("lng", "cmdapply")
End Sub


Public Sub apply_Picture()
Dim nowPath$, nowPath2$
If GetINI("Setting", "PictureFT") = 1 Then
    If GetINI("Setting", "PictureSuff_Other") = 0 Then
        nowPath = App.Path & "\icon\"
        nowPath2 = App.Path & "\icon\V3I\"
        FrmMain.Wallpaper.Picture = LoadPicture(LoadP(App.Path & "\icon\Wallpaper.jpg"))
        FrmMain.Wallpaper2.Picture = LoadPicture(LoadP(App.Path & "\icon\V3I\cl.gif"))
    Else
        nowPath = GetINI("Setting", "IconPath")
        nowPath2 = GetINI("Setting", "Icon2Path")
        FrmMain.Wallpaper.Picture = LoadPicture(LoadP(GetINI("Setting", "WallpaperPath")))
        FrmMain.Wallpaper2.Picture = LoadPicture(LoadP(GetINI("Setting", "Wallpaper2Path")))
    End If
FrmMain.Imgicon(0).Picture = LoadPicture(LoadP(nowPath & "\415.gif"))
FrmMain.Imgicon(1).Picture = LoadPicture(LoadP(nowPath & "\404.gif"))
FrmMain.Imgicon(2).Picture = LoadPicture(LoadP(nowPath & "\407.gif"))
FrmMain.Imgicon(3).Picture = LoadPicture(LoadP(nowPath & "\473.gif"))
FrmMain.Imgicon(4).Picture = LoadPicture(LoadP(nowPath & "\391.gif"))
FrmMain.Imgicon(5).Picture = LoadPicture(LoadP(nowPath & "\457.gif"))
FrmMain.Imgicon(6).Picture = LoadPicture(LoadP(nowPath & "\394.gif"))
FrmMain.Imgicon(8).Picture = LoadPicture(LoadP(nowPath & "\416.gif"))
FrmMain.Imgicon(9).Picture = LoadPicture(LoadP(nowPath & "\335.gif"))
FrmMain.Imgicon2(0).Picture = LoadPicture(LoadP(nowPath2 & "\595.gif"))
FrmMain.Imgicon2(1).Picture = LoadPicture(LoadP(nowPath2 & "\623.gif"))
FrmMain.Imgicon2(2).Picture = LoadPicture(LoadP(nowPath2 & "\627.gif"))
FrmMain.Imgicon2(3).Picture = LoadPicture(LoadP(nowPath2 & "\626.gif"))
FrmMain.Imgicon2(4).Picture = LoadPicture(LoadP(nowPath2 & "\618.gif"))
FrmMain.Imgicon2(5).Picture = LoadPicture(LoadP(nowPath2 & "\630.gif"))
FrmMain.Imgicon2(6).Picture = LoadPicture(LoadP(nowPath2 & "\607.gif"))
End If
End Sub
Public Function LoadP(Path As String)
If Len(Dir(Path)) = 0 Then Path = App.Path & "\icon\Error.gif"
LoadP = Replace(Path, "\\", "\")
End Function

Public Sub SaveDAT(Savepath As String)
Dim Offset%, p As Byte
Open Savepath For Binary As #1
'保存图标位置数据
For Offset = 1 To 73 Step 8
    For p = 0 To 3
        one(p) = icon10(Offset \ 8, p)
    Next p
    Put #1, Offset + 1, CInt(one(0)) '_left
    Put #1, Offset + 3, CInt(one(1)) '_top
    Put #1, Offset + 5, CInt(one(0) + one(2) - 1) '_right
    Put #1, Offset + 7, CInt(one(1) + one(3) - 1) '_bottom
Next Offset
If Platform = 1 Then
    '保存字体颜色
    RGBt(1) = FrmMain.txttime.ForeColor \ 65536
    RGBt(2) = FrmMain.txttime.ForeColor \ 256 - (FrmMain.txttime.ForeColor \ 65536) * 256
    RGBt(3) = FrmMain.txttime.ForeColor Mod 256
    Put #1, 87, RGBt(1)
    Put #1, 86, RGBt(2)
    Put #1, 85, RGBt(3)
    '保存字体编号
    Put #1, 96, CByte(FrmMain.cbbfontNO.ListIndex)
ElseIf Platform = 2 Then
    '保存外屏图标位置数据
    For Offset = 81 To 129 Step 8
        For p = 0 To 3
            one(p) = icon7((Offset - 80) \ 8, p)
        Next p
        Put #1, Offset + 1, CInt(one(0)) '_left
        Put #1, Offset + 3, CInt(one(1)) '_top
        Put #1, Offset + 5, CInt(one(0) + one(2) - 1) '_right
        Put #1, Offset + 7, CInt(one(1) + one(3) - 1) '_bottom
    Next Offset
    Put #1, 159, 0
End If
Close #1
End Sub

Public Sub OpenDAT(Path As String)
Dim Offset%, p As Byte
Dim ONE_right As Byte, ONE_bottom As Byte, TBGR_Color$, fontNO As Byte

Open Path For Binary As #1
'读取图标位置数据
    For Offset = 1 To 76 Step 8
        Get #1, Offset + 1, one(0)            'left
        Get #1, Offset + 3, one(1)            'top
        Get #1, Offset + 5, ONE_right
        Get #1, Offset + 7, ONE_bottom
        one(2) = ONE_right - one(0) + 1  'Width
        one(3) = ONE_bottom - one(1) + 1 'Height
        '读到图片
        FrmMain.Imgicon(Offset \ 8).Left = one(0) * 15
        FrmMain.Imgicon(Offset \ 8).Top = one(1) * 15
        FrmMain.Imgicon(Offset \ 8).Width = one(2) * 15
        FrmMain.Imgicon(Offset \ 8).Height = one(3) * 15
        '读到数组
        For p = 0 To 3
            icon10(Offset \ 8, p) = one(p)
        Next p
    Next Offset
    
    FrmMain.Wallpaper.Enabled = True
    For p = 0 To 3
        FrmMain.txtXY(p).Enabled = True
        FrmMain.VSXY(p).Enabled = True
    Next
If Platform = 1 Then
    FrmMain.cbbSelect.Clear
    FrmMain.cbbSelect.AddItem GetINI("lng", "cbbSelect1")
    FrmMain.cbbSelect.AddItem GetINI("lng", "cbbSelect2")
    FrmMain.cbbSelect.AddItem GetINI("lng", "cbbSelect3")
    FrmMain.cbbSelect.AddItem GetINI("lng", "cbbSelect4")
    FrmMain.cbbSelect.AddItem GetINI("lng", "cbbSelect5")
    FrmMain.cbbSelect.AddItem GetINI("lng", "cbbSelect6")
    FrmMain.cbbSelect.AddItem GetINI("lng", "cbbSelect7")
    FrmMain.cbbSelect.AddItem GetINI("lng", "cbbSelect8")
    FrmMain.cbbSelect.AddItem GetINI("lng", "cbbSelect9")
    FrmMain.cbbSelect.AddItem GetINI("lng", "cbbSelect10")
    FrmMain.cbbSelect.Text = GetINI("lng", "cbbSelectText")
    FrmMain.Fra_time_area.Visible = True
    FrmMain.Wallpaper2.Visible = False
    '读取时间颜色，显示时间颜色
    Get #1, 85, RGBt(3) '红
    Get #1, 86, RGBt(2) '绿
    Get #1, 87, RGBt(1) '蓝
    'Get #1, 88, RGBt(0) '透明
    TBGR_Color = "&H00000000"
    For p = 1 To 3
        If RGBt(p) <= 15 Then   '一/两位
            Mid(TBGR_Color, (p + 1) * 2 + 1 + 1, 2) = Hex(RGBt(p))
        Else
            Mid(TBGR_Color, (p + 1) * 2 + 1, 2) = Hex(RGBt(p))
        End If
    Next p
    FrmMain.txttime.ForeColor = TBGR_Color
'读取时间字体编号，显示时间字体编号
    Get #1, 96, fontNO
    FrmMain.cbbfontNO.ListIndex = fontNO
ElseIf Platform = 2 Then
    FrmMain.cbbSelect.Clear
    FrmMain.cbbSelect.AddItem GetINI("lng", "cbbSelect1")
    FrmMain.cbbSelect.AddItem GetINI("lng", "cbbSelect2")
    FrmMain.cbbSelect.AddItem GetINI("lng", "cbbSelect3")
    FrmMain.cbbSelect.AddItem GetINI("lng", "cbbSelect4")
    FrmMain.cbbSelect.AddItem GetINI("lng", "cbbSelect5")
    FrmMain.cbbSelect.AddItem GetINI("lng", "cbbSelect6")
    FrmMain.cbbSelect.AddItem GetINI("lng", "cbbSelect7")
    FrmMain.cbbSelect.AddItem GetINI("lng", "cbbSelect8")
    FrmMain.cbbSelect.AddItem GetINI("lng", "cbbSelect9")
    FrmMain.cbbSelect.AddItem GetINI("lng", "cbbSelect10")
    FrmMain.cbbSelect.AddItem GetINI("lng", "cbbSelect11")
    FrmMain.cbbSelect.AddItem GetINI("lng", "cbbSelect12")
    FrmMain.cbbSelect.AddItem GetINI("lng", "cbbSelect13")
    FrmMain.cbbSelect.AddItem GetINI("lng", "cbbSelect14")
    FrmMain.cbbSelect.AddItem GetINI("lng", "cbbSelect15")
    FrmMain.cbbSelect.AddItem GetINI("lng", "cbbSelect16")
    FrmMain.cbbSelect.AddItem GetINI("lng", "cbbSelect17")
    FrmMain.cbbSelect.Text = GetINI("lng", "cbbSelectText")
    FrmMain.Fra_time_area.Visible = False
    FrmMain.Wallpaper2.Visible = True
    For Offset = 81 To 129 Step 8
        Get #1, Offset + 1, one(0)            'left
        Get #1, Offset + 3, one(1)            'top
        Get #1, Offset + 5, ONE_right
        Get #1, Offset + 7, ONE_bottom
        one(2) = ONE_right - one(0) + 1  'Width
        one(3) = ONE_bottom - one(1) + 1 'Height
        '读到图片
        FrmMain.Imgicon2((Offset - 80) \ 8).Left = one(0) * 15
        FrmMain.Imgicon2((Offset - 80) \ 8).Top = one(1) * 15
        FrmMain.Imgicon2((Offset - 80) \ 8).Width = one(2) * 15
        FrmMain.Imgicon2((Offset - 80) \ 8).Height = one(3) * 15
        '读到数组
        For p = 0 To 3
            icon7((Offset - 80) \ 8, p) = one(p)
        Next p
    Next Offset
End If

Close #1
End Sub
