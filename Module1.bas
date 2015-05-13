Attribute VB_Name = "Module1"
Option Explicit
Public PictureFT As Byte, PicturePath As Byte '����ͼƬ��Ĭ��/����
Public WallpaperPath$, Wallpaper2Path$, IconPath$, IconPath2$ '����������������

Public Platform As Byte '���ͣ�1=L7/E398��2=V3/V3I��
Public Savepath$ '����·��
Public RGBt(3) As Byte
Public one(3) As Byte
Public icon10(9, 3) As Integer '����
Public icon7(6, 3) As Integer '����

Public nowform As Byte

'�����϶�����
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
Public CurrentPoint As POINTAPI '����
Public ClipRect As RECT
Public RetValue As Long


Public Sub Main()
    FrmMain.Show
    'ȡ�������в���
    If Len(Command()) <> 0 Then
        Dim iName As String '·��
        iName = Replace(Command(), Chr(34), "") '�滻"Ϊ��
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
    Line Input #3, Wallpaper2Path
    Seek #3, Seek(3) + 4
    Line Input #3, IconPath
    Seek #3, Seek(3) + 4
    Line Input #3, IconPath2
Close #3
End Sub
Public Sub apply_Picture(FT)
Dim nowPath$, nowPath2$
If FT = 1 Then
    If PicturePath = 0 Then
        nowPath = App.Path & "\icon"
        nowPath2 = App.Path & "\icon\V3I"
        FrmMain.Wallpaper.Picture = LoadPicture(LoadP(App.Path & "\icon\Wallpaper.jpg"))
        FrmMain.Wallpaper2.Picture = LoadPicture(LoadP(App.Path & "\icon\V3I\cl.gif"))
    Else
        nowPath = IconPath
        nowPath2 = IconPath2
        FrmMain.Wallpaper.Picture = LoadPicture(LoadP(WallpaperPath))
        FrmMain.Wallpaper2.Picture = LoadPicture(LoadP(Wallpaper2Path))
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
LoadP = Path
End Function

Public Sub SaveDAT(Savepath As String)
Dim Offset%, p As Byte
Open Savepath For Binary As #1
'����ͼ��λ������
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
    '����������ɫ
    RGBt(1) = FrmMain.txttime.ForeColor \ 65536
    RGBt(2) = FrmMain.txttime.ForeColor \ 256 - (FrmMain.txttime.ForeColor \ 65536) * 256
    RGBt(3) = FrmMain.txttime.ForeColor Mod 256
    Put #1, 87, RGBt(1)
    Put #1, 86, RGBt(2)
    Put #1, 85, RGBt(3)
    '����������
    Put #1, 96, CByte(FrmMain.cbbfontNO.ListIndex)
ElseIf Platform = 2 Then
    '��������ͼ��λ������
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
'��ȡͼ��λ������
    For Offset = 1 To 76 Step 8
        Get #1, Offset + 1, one(0)            'left
        Get #1, Offset + 3, one(1)            'top
        Get #1, Offset + 5, ONE_right
        Get #1, Offset + 7, ONE_bottom
        one(2) = ONE_right - one(0) + 1  'Width
        one(3) = ONE_bottom - one(1) + 1 'Height
        '����ͼƬ
        FrmMain.Imgicon(Offset \ 8).Left = one(0) * 15
        FrmMain.Imgicon(Offset \ 8).Top = one(1) * 15
        FrmMain.Imgicon(Offset \ 8).Width = one(2) * 15
        FrmMain.Imgicon(Offset \ 8).Height = one(3) * 15
        '��������
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
    FrmMain.cbbSelect.AddItem "�ź�"
    FrmMain.cbbSelect.AddItem "GPRS"
    FrmMain.cbbSelect.AddItem "����"
    FrmMain.cbbSelect.AddItem "����"
    FrmMain.cbbSelect.AddItem "����"
    FrmMain.cbbSelect.AddItem "JAVA"
    FrmMain.cbbSelect.AddItem "����"
    FrmMain.cbbSelect.AddItem "ʱ��"
    FrmMain.cbbSelect.AddItem "����"
    FrmMain.cbbSelect.AddItem "����"
    FrmMain.cbbSelect.Text = "��ѡ��ͼ��"
    FrmMain.Fra_time_area.Visible = True
    FrmMain.Wallpaper2.Visible = False
    '��ȡʱ����ɫ����ʾʱ����ɫ
    Get #1, 85, RGBt(3) '��
    Get #1, 86, RGBt(2) '��
    Get #1, 87, RGBt(1) '��
    'Get #1, 88, RGBt(0) '͸��
    TBGR_Color = "&H00000000"
    For p = 1 To 3
        If RGBt(p) <= 15 Then   'һ/��λ
            Mid(TBGR_Color, (p + 1) * 2 + 1 + 1, 2) = Hex(RGBt(p))
        Else
            Mid(TBGR_Color, (p + 1) * 2 + 1, 2) = Hex(RGBt(p))
        End If
    Next p
    FrmMain.txttime.ForeColor = TBGR_Color
'��ȡʱ�������ţ���ʾʱ��������
    Get #1, 96, fontNO
    FrmMain.cbbfontNO.ListIndex = fontNO
ElseIf Platform = 2 Then
    FrmMain.cbbSelect.Clear
    FrmMain.cbbSelect.AddItem "�ź�"
    FrmMain.cbbSelect.AddItem "GPRS"
    FrmMain.cbbSelect.AddItem "����"
    FrmMain.cbbSelect.AddItem "����"
    FrmMain.cbbSelect.AddItem "����"
    FrmMain.cbbSelect.AddItem "JAVA"
    FrmMain.cbbSelect.AddItem "����"
    FrmMain.cbbSelect.AddItem "ʱ��"
    FrmMain.cbbSelect.AddItem "����"
    FrmMain.cbbSelect.AddItem "����"
    FrmMain.cbbSelect.AddItem "�����ź�"
    FrmMain.cbbSelect.AddItem "����GPRS"
    FrmMain.cbbSelect.AddItem "��������"
    FrmMain.cbbSelect.AddItem "��������"
    FrmMain.cbbSelect.AddItem "����JAVA"
    FrmMain.cbbSelect.AddItem "��������"
    FrmMain.cbbSelect.AddItem "��������"
    FrmMain.cbbSelect.Text = "��ѡ��ͼ��"
    FrmMain.Fra_time_area.Visible = False
    FrmMain.Wallpaper2.Visible = True
    For Offset = 81 To 129 Step 8
        Get #1, Offset + 1, one(0)            'left
        Get #1, Offset + 3, one(1)            'top
        Get #1, Offset + 5, ONE_right
        Get #1, Offset + 7, ONE_bottom
        one(2) = ONE_right - one(0) + 1  'Width
        one(3) = ONE_bottom - one(1) + 1 'Height
        '����ͼƬ
        FrmMain.Imgicon2((Offset - 80) \ 8).Left = one(0) * 15
        FrmMain.Imgicon2((Offset - 80) \ 8).Top = one(1) * 15
        FrmMain.Imgicon2((Offset - 80) \ 8).Width = one(2) * 15
        FrmMain.Imgicon2((Offset - 80) \ 8).Height = one(3) * 15
        '��������
        For p = 0 To 3
            icon7((Offset - 80) \ 8, p) = one(p)
        Next p
    Next Offset
End If

Close #1
End Sub
