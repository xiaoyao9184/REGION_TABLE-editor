VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REGION_TABLE �༭��"
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
      Left            =   3120
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
      Interval        =   1000
      Left            =   5400
      Top             =   0
   End
   Begin VB.CheckBox chkIconBS 
      Caption         =   "ͼ����ʾ�߿�"
      Height          =   255
      Left            =   4800
      TabIndex        =   15
      Top             =   480
      Width           =   1455
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
      Left            =   5880
      TabIndex        =   8
      Top             =   1680
      Width           =   255
   End
   Begin VB.VScrollBar VSXY 
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   4
      Top             =   1680
      Width           =   255
   End
   Begin VB.VScrollBar VSXY 
      Height          =   375
      Index           =   2
      Left            =   5880
      TabIndex        =   6
      Top             =   960
      Width           =   255
   End
   Begin VB.VScrollBar VSXY 
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   2
      Top             =   960
      Width           =   255
   End
   Begin VB.Frame Fra_time_area 
      Caption         =   "ʱ�����"
      Height          =   1335
      Left            =   3120
      TabIndex        =   13
      Top             =   2400
      Width           =   3135
      Begin VB.TextBox txthint 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "��������"
         Top             =   840
         Width           =   735
      End
      Begin VB.ComboBox cbbfontNO 
         Height          =   300
         Left            =   960
         TabIndex        =   18
         Text            =   "��ѡ��������"
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txttime 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "��ɫ"
         Height          =   375
         Left            =   1560
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.TextBox txtXY 
      Height          =   390
      Index           =   3
      Left            =   5400
      TabIndex        =   7
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtXY 
      Height          =   375
      Index           =   2
      Left            =   5400
      TabIndex        =   5
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtXY 
      Height          =   375
      Index           =   1
      Left            =   3840
      TabIndex        =   3
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtXY 
      Height          =   375
      Index           =   0
      Left            =   3840
      TabIndex        =   1
      Top             =   960
      Width           =   495
   End
   Begin VB.ComboBox cbbSelect 
      Height          =   300
      Left            =   3120
      TabIndex        =   0
      Text            =   "��ѡ��ͼ��"
      Top             =   480
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5880
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label LblXY 
      Caption         =   "�߶�"
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   12
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label LblXY 
      Caption         =   "���"
      Height          =   255
      Index           =   2
      Left            =   4800
      TabIndex        =   11
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label LblXY 
      Caption         =   "����"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   10
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label LblXY 
      Caption         =   "���"
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   9
      Top             =   1080
      Width           =   495
   End
   Begin VB.Menu munFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu munNew 
         Caption         =   "�½�(&N)"
         Begin VB.Menu L7E398 
            Caption         =   "L7/E398"
         End
         Begin VB.Menu V3I 
            Caption         =   "V3I"
         End
      End
      Begin VB.Menu munOpen 
         Caption         =   "��(&O)"
      End
      Begin VB.Menu munSave 
         Caption         =   "����(&S)"
      End
      Begin VB.Menu munSaveAs 
         Caption         =   "���Ϊ(&A)"
      End
      Begin VB.Menu mun_ 
         Caption         =   "-"
      End
      Begin VB.Menu munExit 
         Caption         =   "�˳�(&E)"
      End
   End
   Begin VB.Menu munOther 
      Caption         =   "����(&O)"
      Begin VB.Menu munSetting 
         Caption         =   "����(&S)"
      End
      Begin VB.Menu munAbout 
         Caption         =   "����(&A)"
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
Private Sub Form_Load()
cbbfontNO.AddItem "00|??"
cbbfontNO.AddItem "01|��õ�����"
cbbfontNO.AddItem "02|(��01һ��������)"
cbbfontNO.AddItem "03|(ʱ��)"
cbbfontNO.AddItem "04|����,����3��ʱ"
cbbfontNO.AddItem "05|����"
cbbfontNO.AddItem "06|??"
cbbfontNO.AddItem "07|Ӣ���ճ̱�"
cbbfontNO.AddItem "08|�ⲿ��Ļ"
cbbfontNO.AddItem "09|(��խ������)"
cbbfontNO.AddItem "0A|(����� ��)"
cbbfontNO.AddItem "0B|̩��"
cbbfontNO.AddItem "0C|����"
cbbfontNO.AddItem "0D|(������)"
cbbfontNO.AddItem "0E|(�ܴ������+ð��)"
cbbfontNO.AddItem "0F|(�ܴ��AMP)"
cbbfontNO.AddItem "10|���ַ�"
cbbfontNO.AddItem "11|���ַ�"
cbbfontNO.AddItem "12|����??"
cbbfontNO.AddItem "13|���ıʻ�����"
cbbfontNO.AddItem "14|����ƴ������"
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

load_cfg (App.Path & "\Config.cfg")
apply_Picture (PictureFT)
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
    If MsgBox("������", vbYesNo, "��ʾ") = vbYes Then Call munSave_Click
    End
End Sub
Private Sub munSaveAs_Click()
'����ͼ��λ������
CommonDialog1.FileName = "REGION_TABLE"
CommonDialog1.Filter = "״̬��ʱ�������ɫ��λ���ļ�"
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
CommonDialog1.Filter = "״̬��ʱ�������ɫ��λ���ļ�|REGION_TABLE"
CommonDialog1.CancelError = True
On Error Resume Next
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Or Err.Number = 32755 Then Exit Sub
'ȷ������
Open CommonDialog1.FileName For Binary As #1
    Platform = IIf(LOF(1) <= 96, 1, 2)
Close #1
    OpenDAT (CommonDialog1.FileName)
    Savepath = CommonDialog1.FileName
    munSave.Enabled = True
    munSaveAs.Enabled = True
End Sub
Private Sub L7E398_Click()
    Platform = 1
    OpenDAT (App.Path & "\Config.cfg")
    Savepath = ""
    munSaveAs.Enabled = True
    munSave.Enabled = True
End Sub
Private Sub V3I_Click()
    Platform = 2
    OpenDAT (App.Path & "\Config2.cfg")
    Savepath = ""
    munSaveAs.Enabled = True
    munSave.Enabled = True
End Sub

'����ʱ����ɫ
Private Sub cmdtimeColor_Click()
CommonDialog1.ShowColor
txttime.ForeColor = CommonDialog1.Color
End Sub
'����ʱ����������ʾ
Private Sub cbbfontNO_Click()
Dim Response, scarcity$, nono$, warn$, advise$
scarcity = "���������ô����壡" & Chr(13) & Chr(10) & "��ԭCG4������У������������ȱ�������ַ���" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "����0-9����㣺����ĸA M P��"
nono = "�������ô����壡�˱�Ų����κ��ַ�"
warn = "���棡"
advise = "���飡"

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
'��ʾλ��ͼ������
Private Sub cbbSelect_Click()
Dim i%
nowform = IIf(cbbSelect.ListIndex > 9, 2, 1)
For i = 0 To 3 'ѡ��ڼ������Ͱѵڼ������ݴ������ж���TXT
    If nowform = 2 Then
        txtXY(i).Text = icon7(cbbSelect.ListIndex - 10, i)
    Else
        txtXY(i).Text = icon10(cbbSelect.ListIndex, i)
    End If
Next i
End Sub

'��������
Private Sub txtXY_KeyPress(Index As Integer, KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
If KeyAscii = 8 Then Exit Sub
If CInt(Mid(txtXY(Index).Text, 1, txtXY(Index).SelStart) & Chr(KeyAscii) & Mid(txtXY(Index).Text, txtXY(Index).SelStart + 1)) > 255 Then KeyAscii = 0
End Sub

'����λ��ͼ������-�ı���������-����VSXY
Private Sub txtXY_Change(Index As Integer)
    If txtXY(Index).Text <> "" Then
        If txtXY(Index).Text >= 0 Then
            VSXY(Index).Value = txtXY(Index).Text
        End If
    End If
End Sub



'����λ��ͼ������-��ͷ��������-����txtXY,Imgicon1\2,����
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
    icon7(cbbSelect.ListIndex - 10, Index) = VSXY(Index).Value         '���ݱ��浽����
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
    icon10(cbbSelect.ListIndex, Index) = VSXY(Index).Value          '���ݱ��浽����
End If
End Sub
'����λ��ͼ������-ͼ���������-����txtXY
Private Sub Imgicon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
nowform = 1
cbbSelect.ListIndex = Index '�����б���Ӧ

Dim i%
For i = 0 To 9
    Imgicon(i).Enabled = False '����ͼ�겻���ã������϶�������ͼ����ʱ����Ӱ�����꣩
Next i

Imgicon(Index).Drag 1       '���ÿ��϶�
DragX = X                   '����ڴ�ͼ���ϵ�X����
DragY = Y                   '����ڴ�ͼ���ϵ�Y����

'�����϶����򣺿�ʼ
With CurrentPoint
    .X = 0
    .Y = 0
End With
' find position on the screen (not the window)
RetValue = ClientToScreen(Wallpaper.hwnd, CurrentPoint) 'CurrentPoint�Ǵ���Wallpaper������
With ClipRect
    .Top = CurrentPoint.Y + DragY \ Screen.TwipsPerPixelY '��λ�����أ�+��
    .Left = CurrentPoint.X + DragX \ Screen.TwipsPerPixelX
    .Right = CurrentPoint.X + 176 - (Imgicon(Index).Width - DragX) \ Screen.TwipsPerPixelX + 1 '+1�����߽�����
    .Bottom = CurrentPoint.Y + 220 - (Imgicon(Index).Height - DragY) \ Screen.TwipsPerPixelY + 1 '+1�����߽�����
End With ' clip it
RetValue = ClipCursor(ClipRect)

End Sub
Private Sub Wallpaper_DragDrop(Source As Control, X As Single, Y As Single)
Dim i%
For i = 0 To 9
    Imgicon(i).Enabled = True
Next i
'����TXT
txtXY(0).Text = (X - DragX) \ 15
txtXY(1).Text = (Y - DragY) \ 15
    '�߽糬������
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
'����TXT
txtXY(0).Text = (X - DragX) \ 15
txtXY(1).Text = (Y - DragY) \ 15
    '�߽糬������
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
'����λ��ͼ������-ͼ���������-����txtXY
Private Sub Imgicon2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
nowform = 2
cbbSelect.ListIndex = Index + 10 '�����б���Ӧ

Dim i%
For i = 0 To 6
    Imgicon2(i).Enabled = False '����ͼ�겻���ã������϶�������ͼ����ʱ����Ӱ�����꣩
Next i

Imgicon2(Index).Drag 1       '���ÿ��϶�
DragX = X                   '����ڴ�ͼ���ϵ�X����
DragY = Y                   '����ڴ�ͼ���ϵ�Y����

'�����϶����򣺿�ʼ
With CurrentPoint
    .X = 0
    .Y = 0
End With
' find position on the screen (not the window)
RetValue = ClientToScreen(Wallpaper2.hwnd, CurrentPoint) 'CurrentPoint�Ǵ���Wallpaper������
With ClipRect
    .Top = CurrentPoint.Y + DragY \ Screen.TwipsPerPixelY '��λ�����أ�+��
    .Left = CurrentPoint.X + DragX \ Screen.TwipsPerPixelX
    .Right = CurrentPoint.X + 98 - (Imgicon2(Index).Width - DragX) \ Screen.TwipsPerPixelX + 1 '+1�����߽�����
    .Bottom = CurrentPoint.Y + 80 - (Imgicon2(Index).Height - DragY) \ Screen.TwipsPerPixelY + 1 '+1�����߽�����
End With ' clip it
RetValue = ClipCursor(ClipRect)

End Sub


'ͼ����ʾ�߿�
Private Sub chkIconBS_Click()
Dim i As Byte
For i = 0 To 9
    Imgicon(i).BorderStyle = chkIconBS.Value
Next i
For i = 0 To 6
    Imgicon2(i).BorderStyle = chkIconBS.Value
Next i
End Sub
