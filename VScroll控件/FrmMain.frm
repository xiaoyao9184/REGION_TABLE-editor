VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   Caption         =   "REGION_TABLE �༭��"
   ClientHeight    =   4335
   ClientLeft      =   3810
   ClientTop       =   3555
   ClientWidth     =   6570
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   6570
   Begin VB.CheckBox chkIconBS 
      Caption         =   "ͼ����ʾ�߿�"
      Height          =   255
      Left            =   4800
      TabIndex        =   19
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
      Picture         =   "FrmMain.frx":0000
      ScaleHeight     =   3300
      ScaleWidth      =   2640
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   480
      Width           =   2640
      Begin VB.Image Imgicon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   9
         Left            =   2310
         Picture         =   "FrmMain.frx":370F
         Top             =   0
         Width           =   330
      End
      Begin VB.Image Imgicon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   8
         Left            =   2025
         Picture         =   "FrmMain.frx":37CF
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
         Picture         =   "FrmMain.frx":385F
         Top             =   0
         Width           =   240
      End
      Begin VB.Image Imgicon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   5
         Left            =   1245
         Picture         =   "FrmMain.frx":390A
         Top             =   0
         Width           =   270
      End
      Begin VB.Image Imgicon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   4
         Left            =   975
         Picture         =   "FrmMain.frx":39BA
         Top             =   0
         Width           =   270
      End
      Begin VB.Image Imgicon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   3
         Left            =   780
         Picture         =   "FrmMain.frx":3A45
         Top             =   0
         Width           =   195
      End
      Begin VB.Image Imgicon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   2
         Left            =   540
         Picture         =   "FrmMain.frx":3AB2
         Top             =   0
         Width           =   240
      End
      Begin VB.Image Imgicon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   1
         Left            =   330
         Picture         =   "FrmMain.frx":3B30
         Top             =   0
         Width           =   210
      End
      Begin VB.Image Imgicon 
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   0
         Left            =   0
         Picture         =   "FrmMain.frx":3BB4
         Top             =   0
         Width           =   330
      End
   End
   Begin VB.ComboBox cbbfontNO 
      Height          =   300
      Left            =   4080
      TabIndex        =   10
      Text            =   "��ѡ��������"
      Top             =   3240
      Width           =   2055
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
      Left            =   3240
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "time"
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txthint 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3240
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "��������"
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdtimeColor 
      Caption         =   "��ɫ"
      Height          =   375
      Left            =   4560
      TabIndex        =   9
      Top             =   2640
      Width           =   615
   End
   Begin VB.Frame Fra_time_area 
      Caption         =   "ʱ�����"
      Height          =   1335
      Left            =   3120
      TabIndex        =   15
      Top             =   2400
      Width           =   3135
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
      TabIndex        =   14
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label LblXY 
      Caption         =   "���"
      Height          =   255
      Index           =   2
      Left            =   4800
      TabIndex        =   13
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label LblXY 
      Caption         =   "����"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   12
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label LblXY 
      Caption         =   "���"
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   11
      Top             =   1080
      Width           =   495
   End
   Begin VB.Menu munFile 
      Caption         =   "�ļ�"
      Begin VB.Menu munNew 
         Caption         =   "�½�"
      End
      Begin VB.Menu munOpen 
         Caption         =   "��"
      End
      Begin VB.Menu munSave 
         Caption         =   "����"
      End
      Begin VB.Menu munSaveAs 
         Caption         =   "���Ϊ"
      End
      Begin VB.Menu mun_ 
         Caption         =   "-"
      End
      Begin VB.Menu munExit 
         Caption         =   "�˳�"
      End
   End
   Begin VB.Menu munOther 
      Caption         =   "����"
      Begin VB.Menu munSetting 
         Caption         =   "����"
      End
      Begin VB.Menu munAbout 
         Caption         =   "����"
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
cbbSelect.AddItem "�ź�"
cbbSelect.AddItem "GPRS"
cbbSelect.AddItem "����"
cbbSelect.AddItem "����"
cbbSelect.AddItem "����"
cbbSelect.AddItem "JAVA"
cbbSelect.AddItem "����"
cbbSelect.AddItem "ʱ��"
cbbSelect.AddItem "����"
cbbSelect.AddItem "����"
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

load_cfg (App.Path & "\Config.cfg")
apply_Picture (PictureFT)
End Sub

Private Sub munAbout_Click()
Load frmAbout
frmAbout.Show
End Sub
Private Sub munSetting_Click()
Load frmSetting
frmSetting.Show
End Sub
Private Sub munExit_Click()
End
End Sub
Private Sub munSaveAs_Click()
Dim Offset%
'����ͼ��λ������
CommonDialog1.FileName = "REGION_TABLE"
CommonDialog1.Filter = "״̬��ʱ�������ɫ��λ���ļ�"
CommonDialog1.ShowSave
Open CommonDialog1.FileName For Binary As #2
For Offset = 1 To 73 Step 8
    save_data (Offset)
    Put #2, Offset + 1, CInt(one(0)) 'left
    Put #2, Offset + 3, CInt(one(1)) 'top
    Put #2, Offset + 5, CInt(one(0) + one(2) - 1) 'right
    Put #2, Offset + 7, CInt(one(1) + one(3) - 1) 'bottom
Next Offset
'����������ɫ
RGBt(1) = txttime.ForeColor \ 65536
RGBt(2) = txttime.ForeColor \ 256 - (txttime.ForeColor \ 65536) * 256
RGBt(3) = txttime.ForeColor Mod 256
Put #2, 87, RGBt(1)
Put #2, 86, RGBt(2)
Put #2, 85, RGBt(3)
'����������
Put #2, 96, CByte(cbbfontNO.ListIndex)
End Sub
Private Sub munSave_Click()
Dim Offset%
'����ͼ��λ������
For Offset = 1 To 73 Step 8
    save_data (Offset)
    Put #1, Offset + 1, CInt(one(0)) '_left
    Put #1, Offset + 3, CInt(one(1)) '_top
    Put #1, Offset + 5, CInt(one(0) + one(2) - 1) '_right
    Put #1, Offset + 7, CInt(one(1) + one(3) - 1) '_bottom
Next Offset
'����������ɫ
RGBt(1) = txttime.ForeColor \ 65536
RGBt(2) = txttime.ForeColor \ 256 - (txttime.ForeColor \ 65536) * 256
RGBt(3) = txttime.ForeColor Mod 256
Put #1, 87, RGBt(1)
Put #1, 86, RGBt(2)
Put #1, 85, RGBt(3)
'����������
Put #1, 96, CByte(cbbfontNO.ListIndex)
End Sub
Private Sub munNew_Click()
munSaveAs.Enabled = True
munSave.Enabled = False
Dim q%, i%, ONE_right As Byte, ONE_bottom As Byte
txttime.ForeColor = &H0
cbbfontNO.ListIndex = 0
For i = 1 To 73 Step 8
        Get #3, i + 1, one(0)            'left
        Get #3, i + 3, one(1)            'top
        Get #3, i + 5, ONE_right
        Get #3, i + 7, ONE_bottom
        one(2) = ONE_right - one(0) + 1  'Width
        one(3) = ONE_bottom - one(1) + 1 'Height
        '����ͼƬ
        Imgicon(q).Left = one(0) * 15
        Imgicon(q).Top = one(1) * 15
        Imgicon(q).Width = one(2) * 15
        Imgicon(q).Height = one(3) * 15
        q = q + 1
        '��������
        load_data (i)
    Next i
End Sub
Private Sub munOpen_Click()
Close #1
munSave.Enabled = True
munSaveAs.Enabled = True
Dim i%, q%, ONE_right As Byte, ONE_bottom As Byte, TBGR_Color$, fontNO As Byte
q = 0
CommonDialog1.Filter = "״̬��ʱ�������ɫ��λ���ļ�|REGION_TABLE"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Binary As #1
'��ȡͼ��λ������
    For i = 1 To 76 Step 8
        Get #1, i + 1, one(0)            'left
        Get #1, i + 3, one(1)            'top
        Get #1, i + 5, ONE_right
        Get #1, i + 7, ONE_bottom
        one(2) = ONE_right - one(0) + 1  'Width
        one(3) = ONE_bottom - one(1) + 1 'Height
        '����ͼƬ
        Imgicon(q).Left = one(0) * 15
        Imgicon(q).Top = one(1) * 15
        Imgicon(q).Width = one(2) * 15
        Imgicon(q).Height = one(3) * 15
        q = q + 1
        '��������
        load_data (i)
    Next i
'��ȡʱ����ɫ����ʾʱ����ɫ
    Get #1, 85, RGBt(3) '��
    Get #1, 86, RGBt(2) '��
    Get #1, 87, RGBt(1) '��
    'Get #1, 88, RGBt(0) '͸��
    TBGR_Color = "&H00000000"
    For i = 1 To 3
        If RGBt(i) <= 15 Then   'һ/��λ
            Mid(TBGR_Color, (i + 1) * 2 + 1 + 1, 2) = Hex(RGBt(i))
        Else
            Mid(TBGR_Color, (i + 1) * 2 + 1, 2) = Hex(RGBt(i))
        End If
    Next i
    txttime.ForeColor = TBGR_Color
'��ȡʱ�������ţ���ʾʱ��������
    Get #1, 96, fontNO
    cbbfontNO.ListIndex = fontNO
End If
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
For i = 0 To 3 'ѡ��ڼ������Ͱѵڼ������ݴ������ж���TXT
    txtXY(i).Text = icon10(cbbSelect.ListIndex, i)
Next i
End Sub

'����λ��ͼ������-�ı���������-����VSXY
Private Sub txtXY_Change(Index As Integer)
    If txtXY(Index).Text <> "" Then
        If txtXY(Index).Text >= 0 Then
            VSXY(Index).Value = txtXY(Index).Text
        End If
    End If
End Sub
'����λ��ͼ������-��ͷ��������-����txtXY,Imgicon
Private Sub VSXY_Change(Index As Integer)
txtXY(Index).Text = VSXY(Index).Value
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
End Sub
'����λ��ͼ������-ͼ���������-����txtXY
Private Sub Imgicon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
cbbSelect.ListIndex = Index '�����б���Ӧ

Dim i%
For i = 0 To 9
    Imgicon(i).Enabled = False '����ͼ�겻���ã������϶�������ͼ����ʱ����Ӱ�����꣩
Next i

Imgicon(Index).Drag 1       '���ÿ��϶�
DragX = X                   '����ڴ�ͼ���ϵ�X����
DragY = y                   '����ڴ�ͼ���ϵ�Y����
End Sub
Private Sub Wallpaper_DragDrop(Source As Control, X As Single, y As Single)
Dim i%
For i = 0 To 9
    Imgicon(i).Enabled = True
Next i
'����TXT
txtXY(0).Text = (X - DragX) \ 15
txtXY(1).Text = (y - DragY) \ 15
    '�߽糬������
If (X - DragX) >= 0 And (y - DragY) <= 0 Then
    txtXY(1).Text = 0
ElseIf (X - DragX) <= 0 And (y - DragY) >= 0 Then
    txtXY(0).Text = 0
ElseIf (X - DragX) <= 0 And (y - DragY) <= 0 Then
    txtXY(0).Text = 0
    txtXY(1).Text = 0
End If
End Sub
Private Sub chkIconBS_Click()
Dim i As Byte
For i = 0 To 9
    Imgicon(i).BorderStyle = chkIconBS.Value
Next i
End Sub
