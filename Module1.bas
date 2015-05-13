Attribute VB_Name = "Module1"
Option Explicit
Public PictureFT As Byte, PicturePath As Byte
Public WallpaperPath$, IconPath$

Public Savepath$
Public RGBt(3) As Byte
Public one(3) As Byte
Public icon10(9, 3) As Integer

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
End Sub


Public Sub load_cfg(cfgPath)
Open cfgPath For Binary As #3
    Get #3, 98, PictureFT
    Get #3, 99, PicturePath
    Seek #3, 104
    Line Input #3, WallpaperPath
    Seek #3, Seek(3) + 4
    Line Input #3, IconPath
Close #3
End Sub
Public Sub apply_Picture(FT)
Dim i As Byte, nowPath$
If FT = 1 Then
    If PicturePath = 0 Then
        nowPath = App.path & "\icon"
        FrmMain.Wallpaper.Picture = LoadPicture(nowPath & "\Wallpaper.jpg")
    Else
        nowPath = IconPath
        FrmMain.Wallpaper.Picture = LoadPicture(WallpaperPath)
    End If
    load_Picture (nowPath)
End If
End Sub

Public Sub load_Picture(nowPath)
FrmMain.Imgicon(0).Picture = LoadPicture(nowPath & "\415.gif")
FrmMain.Imgicon(1).Picture = LoadPicture(nowPath & "\404.gif")
FrmMain.Imgicon(2).Picture = LoadPicture(nowPath & "\407.gif")
FrmMain.Imgicon(3).Picture = LoadPicture(nowPath & "\473.gif")
FrmMain.Imgicon(4).Picture = LoadPicture(nowPath & "\391.gif")
FrmMain.Imgicon(5).Picture = LoadPicture(nowPath & "\457.gif")
FrmMain.Imgicon(6).Picture = LoadPicture(nowPath & "\394.gif")
FrmMain.Imgicon(8).Picture = LoadPicture(nowPath & "\416.gif")
FrmMain.Imgicon(9).Picture = LoadPicture(nowPath & "\335.gif")
End Sub

Public Sub load_data(i)
Dim q%
Select Case i
Case 1
        For q = 0 To 3
            icon10(0, q) = one(q)
        Next q
Case 9
        For q = 0 To 3
            icon10(1, q) = one(q)
        Next q
Case 17
        For q = 0 To 3
            icon10(2, q) = one(q)
        Next q
Case 25
        For q = 0 To 3
            icon10(3, q) = one(q)
        Next q
Case 33
        For q = 0 To 3
            icon10(4, q) = one(q)
        Next q
Case 41
        For q = 0 To 3
            icon10(5, q) = one(q)
        Next q
Case 49
        For q = 0 To 3
            icon10(6, q) = one(q)
        Next q
Case 57
        For q = 0 To 3
            icon10(7, q) = one(q)
        Next q
Case 65
        For q = 0 To 3
            icon10(8, q) = one(q)
        Next q
Case 73
        For q = 0 To 3
            icon10(9, q) = one(q)
        Next q
End Select
End Sub
Public Sub save_data(i)
Dim q%
Select Case i
Case 1
        For q = 0 To 3
            one(q) = icon10(0, q)
        Next q
Case 9
        For q = 0 To 3
            one(q) = icon10(1, q)
        Next q
Case 17
        For q = 0 To 3
            one(q) = icon10(2, q)
        Next q
Case 25
        For q = 0 To 3
            one(q) = icon10(3, q)
        Next q
Case 33
        For q = 0 To 3
            one(q) = icon10(4, q)
        Next q
Case 41
        For q = 0 To 3
            one(q) = icon10(5, q)
        Next q
Case 49
        For q = 0 To 3
            one(q) = icon10(6, q)
        Next q
Case 57
        For q = 0 To 3
            one(q) = icon10(7, q)
        Next q
Case 65
        For q = 0 To 3
            one(q) = icon10(8, q)
        Next q
Case 73
        For q = 0 To 3
             one(q) = icon10(9, q)
        Next q
End Select
End Sub



Public Sub SaveDAT(Savepath As String)
Dim Offset%
Open Savepath For Binary As #1
'保存图标位置数据
For Offset = 1 To 73 Step 8
    save_data (Offset)
    Put #1, Offset + 1, CInt(one(0)) '_left
    Put #1, Offset + 3, CInt(one(1)) '_top
    Put #1, Offset + 5, CInt(one(0) + one(2) - 1) '_right
    Put #1, Offset + 7, CInt(one(1) + one(3) - 1) '_bottom
Next Offset
'保存字体颜色
RGBt(1) = FrmMain.txttime.ForeColor \ 65536
RGBt(2) = FrmMain.txttime.ForeColor \ 256 - (FrmMain.txttime.ForeColor \ 65536) * 256
RGBt(3) = FrmMain.txttime.ForeColor Mod 256
Put #1, 87, RGBt(1)
Put #1, 86, RGBt(2)
Put #1, 85, RGBt(3)
'保存字体编号
Put #1, 96, CByte(FrmMain.cbbfontNO.ListIndex)
Close #1
End Sub

Public Sub OpenDAT(path As String)
Dim i%, q%, ONE_right As Byte, ONE_bottom As Byte, TBGR_Color$, fontNO As Byte
q = 0
Open path For Binary As #1
'读取图标位置数据
    For i = 1 To 76 Step 8
        Get #1, i + 1, one(0)            'left
        Get #1, i + 3, one(1)            'top
        Get #1, i + 5, ONE_right
        Get #1, i + 7, ONE_bottom
        one(2) = ONE_right - one(0) + 1  'Width
        one(3) = ONE_bottom - one(1) + 1 'Height
        '读到图片
        FrmMain.Imgicon(q).Left = one(0) * 15
        FrmMain.Imgicon(q).Top = one(1) * 15
        FrmMain.Imgicon(q).Width = one(2) * 15
        FrmMain.Imgicon(q).Height = one(3) * 15
        q = q + 1
        '读到数组
        load_data (i)
    Next i
'读取时间颜色，显示时间颜色
    Get #1, 85, RGBt(3) '红
    Get #1, 86, RGBt(2) '绿
    Get #1, 87, RGBt(1) '蓝
    'Get #1, 88, RGBt(0) '透明
    TBGR_Color = "&H00000000"
    For i = 1 To 3
        If RGBt(i) <= 15 Then   '一/两位
            Mid(TBGR_Color, (i + 1) * 2 + 1 + 1, 2) = Hex(RGBt(i))
        Else
            Mid(TBGR_Color, (i + 1) * 2 + 1, 2) = Hex(RGBt(i))
        End If
    Next i
    FrmMain.txttime.ForeColor = TBGR_Color
'读取时间字体编号，显示时间字体编号
    Get #1, 96, fontNO
    FrmMain.cbbfontNO.ListIndex = fontNO
Close #1
End Sub
