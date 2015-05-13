Attribute VB_Name = "Module1"
Option Explicit
Public PictureFT As Byte, PicturePath As Byte
Public WallpaperPath$, IconPath$

Public RGBt(3) As Byte
Public one(3) As Byte
Public icon10(9, 3) As Integer
Public Sub load_cfg(cfgPath)
Open cfgPath For Binary As #3
Get #3, 82, PictureFT
Get #3, 83, PicturePath
Seek #3, 88
Line Input #3, WallpaperPath
Seek #3, Seek(3) + 4
Line Input #3, IconPath
End Sub
Public Sub apply_Picture(FT)
Dim i As Byte, nowPath$
If FT = 1 Then
    If PicturePath = 0 Then
        nowPath = App.Path & "\icon"
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

Public Function HEXtoDEC(NO1, NO2)
If NO1 = "A" Then
    NO1 = 10
    ElseIf NO1 = "B" Then
    NO1 = 11
    ElseIf NO1 = "C" Then
    NO1 = 12
    ElseIf NO1 = "D" Then
    NO1 = 13
    ElseIf NO1 = "E" Then
    NO1 = 14
    ElseIf NO1 = "F" Then
    NO1 = 15
    Else
    NO1 = NO1
End If
If NO2 = "A" Then
    NO2 = 10
    ElseIf NO2 = "B" Then
    NO2 = 11
    ElseIf NO2 = "C" Then
    NO2 = 12
    ElseIf NO2 = "D" Then
    NO2 = 13
    ElseIf NO2 = "E" Then
    NO2 = 14
    ElseIf NO2 = "F" Then
    NO2 = 15
    Else
    NO2 = NO2
End If
HEXtoDEC = NO1 * 16 + NO2
End Function
