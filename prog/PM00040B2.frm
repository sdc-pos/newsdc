VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form PM00040B2 
   Caption         =   "�p�[�c���x�����s"
   ClientHeight    =   10290
   ClientLeft      =   1920
   ClientTop       =   2430
   ClientWidth     =   14715
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10290
   ScaleWidth      =   14715
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.ListBox List2 
      Height          =   780
      Left            =   5940
      Sorted          =   -1  'True
      TabIndex        =   71
      Top             =   5160
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   4
      ItemData        =   "PM00040B2.frx":0000
      Left            =   1800
      List            =   "PM00040B2.frx":0002
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   25
      Top             =   4380
      Width           =   2805
   End
   Begin VB.CheckBox Check1 
      Caption         =   "���Y���󎚂���"
      Height          =   375
      Index           =   4
      Left            =   7470
      TabIndex        =   27
      Top             =   4380
      Width           =   2115
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   2  '��
      Index           =   11
      Left            =   4845
      MaxLength       =   20
      TabIndex        =   26
      Top             =   4380
      Width           =   2490
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   1470
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   0
      Top             =   120
      Width           =   1725
   End
   Begin VB.Frame Frame1 
      Caption         =   "���َw��"
      Height          =   2895
      Left            =   5775
      TabIndex        =   63
      Top             =   6480
      Width           =   3615
      Begin VB.TextBox Text1 
         Alignment       =   1  '�E����
         Height          =   375
         IMEMode         =   2  '��
         Index           =   14
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   375
         IMEMode         =   2  '��
         Index           =   18
         Left            =   2940
         MaxLength       =   1
         TabIndex        =   8
         Top             =   2280
         Width           =   405
      End
      Begin VB.TextBox Text1 
         Height          =   375
         IMEMode         =   2  '��
         Index           =   17
         Left            =   1320
         MaxLength       =   11
         TabIndex        =   7
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   375
         IMEMode         =   2  '��
         Index           =   15
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   5
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   375
         IMEMode         =   2  '��
         Index           =   16
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   6
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '�E����
         Height          =   375
         IMEMode         =   2  '��
         Index           =   13
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "����"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   70
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "���t"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   67
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "���ް��"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   66
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "���ч�"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   65
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "����"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   64
         Top             =   480
         Width           =   615
      End
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4575
      Index           =   0
      Left            =   1755
      TabIndex        =   28
      Top             =   4800
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   8070
      _Version        =   393217
      TextRTF         =   $"PM00040B2.frx":0004
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   2  '��
      Index           =   4
      Left            =   5520
      MaxLength       =   5
      TabIndex        =   14
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   2  '��
      Index           =   3
      Left            =   1800
      MaxLength       =   30
      TabIndex        =   13
      Top             =   2040
      Width           =   3735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "�������x��"
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   6360
      TabIndex        =   21
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "�K�p�@�탉�x��"
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   20
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Caption         =   "�v��"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   19
      Top             =   3480
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "��"
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   18
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   2  '��
      Index           =   7
      Left            =   5400
      MaxLength       =   20
      TabIndex        =   17
      Top             =   3000
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   2  '��
      Index           =   6
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   16
      Top             =   3000
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   2  '��
      Index           =   5
      Left            =   1800
      MaxLength       =   8
      TabIndex        =   15
      Top             =   2520
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   3
      Left            =   1800
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   12
      Top             =   1560
      Width           =   5325
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   2
      Left            =   1800
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   11
      Top             =   1080
      Width           =   5325
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   2  '��
      Index           =   10
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   24
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   2  '��
      Index           =   9
      Left            =   4440
      MaxLength       =   10
      TabIndex        =   23
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   2  '��
      Index           =   8
      Left            =   1785
      MaxLength       =   10
      TabIndex        =   22
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   2  '��
      Index           =   12
      Left            =   9480
      MaxLength       =   25
      TabIndex        =   29
      Top             =   960
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   10800
      MaxLength       =   30
      TabIndex        =   10
      Top             =   480
      Width           =   3735
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   1
      Left            =   1470
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   1
      Top             =   600
      Width           =   885
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   2  '��
      Index           =   1
      Left            =   5640
      MaxLength       =   40
      TabIndex        =   9
      Top             =   600
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   2400
      MaxLength       =   20
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�I ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   10320
      TabIndex        =   44
      Top             =   9480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   9480
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   9480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   8640
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   9480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   7800
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   9480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�O��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6480
      TabIndex        =   40
      Top             =   9480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "JAN"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5640
      TabIndex        =   39
      Top             =   9480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4800
      TabIndex        =   38
      Top             =   9480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3960
      TabIndex        =   37
      Top             =   9480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2640
      TabIndex        =   36
      Top             =   9480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   9480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   9480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�X �V"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   33
      Top             =   9480
      Width           =   855
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4575
      Index           =   2
      Left            =   9600
      TabIndex        =   31
      Top             =   2640
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   8070
      _Version        =   393217
      TextRTF         =   $"PM00040B2.frx":00C2
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1575
      Index           =   3
      Left            =   9600
      TabIndex        =   32
      Top             =   7680
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2778
      _Version        =   393217
      TextRTF         =   $"PM00040B2.frx":0180
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   735
      Index           =   1
      Left            =   9480
      TabIndex        =   30
      Top             =   1560
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1296
      _Version        =   393217
      TextRTF         =   $"PM00040B2.frx":023E
   End
   Begin VB.Label lblUpd_DateTime 
      Height          =   255
      Left            =   11610
      TabIndex        =   73
      Top             =   9840
      Width           =   2535
   End
   Begin VB.Label lblUpd_Tanto 
      Height          =   255
      Left            =   11610
      TabIndex        =   72
      Top             =   9420
      Width           =   2535
   End
   Begin VB.Label Label 
      Caption         =   "���Y��"
      Height          =   255
      Index           =   18
      Left            =   735
      TabIndex        =   69
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "���ƕ�"
      Height          =   255
      Index           =   17
      Left            =   525
      TabIndex        =   68
      Top             =   240
      Width           =   795
   End
   Begin VB.Label LabJIGYO 
      Appearance      =   0  '�ׯ�
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   240
      TabIndex        =   62
      Top             =   9840
      Width           =   180
   End
   Begin VB.Label Label 
      Caption         =   "���l"
      Height          =   255
      Index           =   16
      Left            =   1200
      TabIndex        =   61
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label 
      Caption         =   "��Ǝw��"
      Height          =   255
      Index           =   15
      Left            =   720
      TabIndex        =   60
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "�K�p�@����l"
      Height          =   255
      Index           =   14
      Left            =   9600
      TabIndex        =   59
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Label Label 
      Caption         =   "�I��(2)"
      Height          =   255
      Index           =   13
      Left            =   4440
      TabIndex        =   58
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "�I��(1)"
      Height          =   255
      Index           =   12
      Left            =   840
      TabIndex        =   57
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "���萔"
      Height          =   255
      Index           =   11
      Left            =   840
      TabIndex        =   56
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "���ƕ���"
      Height          =   255
      Index           =   10
      Left            =   960
      TabIndex        =   55
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "��Ж�"
      Height          =   255
      Index           =   9
      Left            =   960
      TabIndex        =   54
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "���i(3)"
      Height          =   255
      Index           =   8
      Left            =   6000
      TabIndex        =   53
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "���i(2)"
      Height          =   255
      Index           =   7
      Left            =   3360
      TabIndex        =   52
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "���i(1)"
      Height          =   255
      Index           =   6
      Left            =   720
      TabIndex        =   51
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "�@��(3)"
      Height          =   255
      Index           =   5
      Left            =   8520
      TabIndex        =   50
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "�@��(2)"
      Height          =   255
      Index           =   4
      Left            =   8520
      TabIndex        =   49
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "�@��(1)"
      Height          =   255
      Index           =   3
      Left            =   8520
      TabIndex        =   48
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "PART�@NAME"
      Height          =   255
      Index           =   1
      Left            =   9480
      TabIndex        =   47
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label 
      Caption         =   "�i�ڃR�[�h"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   46
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label 
      Caption         =   "�i��"
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   45
      Top             =   720
      Width           =   495
   End
End
Attribute VB_Name = "PM00040B2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�e�L�X�g�p�Y��
Private Const ptxHIN_GAI% = 0               '�i��
Private Const ptxHIN_NAME% = 1              '�i��
Private Const ptxL_HIN_NAME_E% = 2          '�i��E
Private Const ptxL_BIKOU% = 3               '���l
Private Const ptxL_BIKOU3% = 4              '���l�R
Private Const ptxL_IRI_QTY% = 5             '���萔
Private Const ptxL_TANA1% = 6               '�I��(1)
Private Const ptxL_TANA2% = 7               '�I��(2)
Private Const ptxL_URIKIN1% = 8             '���i(1)
Private Const ptxL_URIKIN2% = 9             '���i(2)
Private Const ptxL_URIKIN3% = 10            '���i(3)

Private Const ptxGENSANKOKU% = 11           '���Y�� 2008.06.12



Private Const ptxL_KISHU1% = 12             '�@��(1)
'Private Const ptxL_KISHU2% = 12             '�@��(2)




Private Const ptxL_MAISU% = 13              '���ٖ���

Private Const ptxL_QTY% = 14                '����   2008.10.03


Private Const ptxL_ORDERNO% = 15            '���ް��
Private Const ptxL_ITEMNO% = 16             '���ч�
Private Const ptxL_PRI_DATE% = 17           '������t

Private Const ptxL_MARK% = 18               '�č���ϰ�  2007.11.08

'���b�`�e�L�X�g�p�Y��
Private Const prchL_SAGYO_SHIJI% = 0        '��Ǝw��
Private Const prchL_KISHU2% = 1             '�@��(2)
Private Const prchL_KISHU3% = 2             '�@��(3)
Private Const prchL_KISHU_BIKOU% = 3        '�K�p�@����l


'�R���{�p�Y��
Private Const pcmbJGYOBU% = 0               '���ƕ�     '2008.06.12


Private Const pcmbNAIGAI% = 1               '�����O
Private Const pcmbL_KAISHA% = 2             '��Ж�
Private Const pcmbL_JGYOBU% = 3             '���ƕ���
Private Const pcmbGENSAN% = 4               '���Y��



'�`�F�b�N�p�Y��
Private Const pchkL_PAPER% = 0              '��
Private Const pchkL_PLASTIC% = 1            '��׽���
Private Const pchkL_LABEL% = 2              '�K�p�@������
Private Const pchkL_MAISU% = 3              '��������

Private Const pchkGENSANKOKU% = 4           '���Y���󎚗L��


'��������ݓ��ꏈ��
Private Const pcmdLabel% = 4                '���و���w��
Private Const pcmdItem% = 5                 '���ш���w��
Private Const pcmdJan% = 6                  'JAN����w��
Private Const pcmdGAISO% = 7                '�O������w��


Private GENSANKOKU_FLG  As String * 1       '���Y���@�󎚗L��


Private INIT_FLG        As Boolean



Private KAISYA_CHK_F    As Boolean          '��Ё^���ƕ��̃G���[�����L�� 2008.09.26

Private PRINT_CHECK_F   As Boolean          '�������L��   2008.09.26



Private GENSANKOKU_CHECK_TBL _
                        As Variant          '���Y�������L���i���ƕ��j 2009.03.28



Private TANKA_SPACE_F   As String           '2010.02.02


Private Const Last_Update_Day$ = "[���Y���Ή�](PM00040 2010.08.02 16:30)"

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    PM00040B2.MousePointer = vbHourglass

    Call Ctrl_Lock(PM00040B2)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PM00040B2)


    PM00040B2.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   ���͍��ڂ̃G���[�`�F�b�N
'----------------------------------------------------------------------------
Dim com     As Integer
Dim ans     As Integer
Dim sts     As Integer

Dim i       As Integer
Dim j       As Integer
Dim k       As Integer
    
    Error_Check_Proc = True
    
    
    
    Select Case Mode
        
        Case ptxHIN_GAI      '�i��
            
            If Trim(Text1(ptxHIN_GAI).Text) = "" Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxHIN_GAI).SetFocus
                Exit Function
            End If
            
        
        
            If Last_JGYOBU = StrConv(ITEM_BREC.JGYOBU, vbUnicode) And _
                Right(Combo1(pcmbNAIGAI), 1) = StrConv(ITEM_BREC.NAIGAI, vbUnicode) And _
                Trim(Text1(ptxHIN_GAI).Text) = Trim(StrConv(ITEM_BREC.HIN_GAI, vbUnicode)) Then
            Else
                Call UniCode_Conv(K0_ITEM_B.JGYOBU, Last_JGYOBU)
                Call UniCode_Conv(K0_ITEM_B.NAIGAI, Right(Combo1(pcmbNAIGAI), 1))
                Call UniCode_Conv(K0_ITEM_B.HIN_GAI, Text1(ptxHIN_GAI).Text)
            
                sts = BTRV(BtOpGetEqual, ITEM_B_POS, ITEM_BREC, Len(ITEM_BREC), K0_ITEM_B, Len(K0_ITEM_B), 0)
                Select Case sts
                    Case BtNoErr
                    
                        Call Item_Disp_Proc(Last_JGYOBU & Right(Combo1(pcmbNAIGAI), 1) & Text1(ptxHIN_GAI).Text)
                    
                    Case BtErrKeyNotFound
                    
                    
                    
                        For i = 0 To UBound(JGYOBU_T)
                            For j = 0 To Combo1(pcmbNAIGAI).ListCount - 1
                                Call UniCode_Conv(K0_ITEM_B.JGYOBU, JGYOBU_T(i).CODE)
                                Call UniCode_Conv(K0_ITEM_B.NAIGAI, Right(Combo1(pcmbNAIGAI).List(j), 1))
                                Call UniCode_Conv(K0_ITEM_B.HIN_GAI, Text1(ptxHIN_GAI).Text)
        
                                sts = BTRV(BtOpGetEqual, ITEM_B_POS, ITEM_BREC, Len(ITEM_BREC), K0_ITEM_B, Len(K0_ITEM_B), 0)
                                Select Case sts
                                    Case BtNoErr
        
                                        
                                        
                                        For k = 0 To Combo1(pcmbJGYOBU).ListCount - 1
                                        
                                            
                                            If Right(Combo1(pcmbJGYOBU).List(k), 1) = JGYOBU_T(i).CODE Then
                                            
                                                Combo1(pcmbJGYOBU).ListIndex = k
                                                
                                                Last_JGYOBU = JGYOBU_T(i).CODE
                                                Exit For
                                            
                                            End If
                                        
                                        Next k
                                    
                                    
                                        For k = 0 To Combo1(pcmbNAIGAI).ListCount - 1
                                        
                                            
                                            If Right(Combo1(pcmbNAIGAI).List(k), 1) = StrConv(ITEM_BREC.NAIGAI, vbUnicode) Then
                                            
                                                Combo1(pcmbNAIGAI).ListIndex = k
                                                Exit For
                                            
                                            End If
                                        
                                        Next k
                                        
                                        Call Item_Disp_Proc(Last_JGYOBU & Right(Combo1(pcmbNAIGAI), 1) & Text1(ptxHIN_GAI).Text)
                                        Exit For
        
                                    Case BtErrKeyNotFound
                                        Exit For
        
                                    Case Else
                                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                        Exit Function
                                End Select
        
        
                            Next j
                    
                    
                            If sts = BtNoErr Then
                    
                            
                                Exit For
                            
                            End If
                    
                    
                        Next i
                    
                    
                        
                        If i > UBound(JGYOBU_T) Then
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> MT_2009.06.01
                            'MsgBox "���͂����R�[�h�́A���o�^�ł��B"
                            'Exit Function
                                
                            If PN_CHK(Text1(ptxHIN_GAI), "G", "PLABEL", 1) Then
                                ''MsgBox "���͂����R�[�h�́A���o�^�ł��B"
                                
                                Exit Function
                            End If
                            
                            Call Item_Disp_Proc(Last_JGYOBU & Right(Combo1(pcmbNAIGAI), 1) & Text1(ptxHIN_GAI).Text)
                            
                        End If
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                        Exit Function
                End Select
            End If
        
            
        
        Case ptxL_IRI_QTY          '���萔
        
            If Trim(Text1(ptxL_IRI_QTY).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxL_IRI_QTY).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(ptxL_IRI_QTY).SetFocus
                    Exit Function
                Else
                
                    Text1(ptxL_IRI_QTY).Text = Format(CLng(Text1(ptxL_IRI_QTY).Text), "#0")
                
                End If
            End If
        
        Case ptxL_URIKIN1          '���i(1)
        
            If Trim(Text1(ptxL_URIKIN1).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxL_URIKIN1).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(ptxL_URIKIN1).SetFocus
                    Exit Function
                Else
                
                    Text1(ptxL_URIKIN1).Text = Format(CLng(Text1(ptxL_URIKIN1).Text), "#0")
                
                End If
            End If
        
        Case ptxL_URIKIN2          '���i(2)
        
            If Trim(Text1(ptxL_URIKIN2).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxL_URIKIN2).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(ptxL_URIKIN2).SetFocus
                    Exit Function
                Else
                
                    Text1(ptxL_URIKIN2).Text = Format(CLng(Text1(ptxL_URIKIN2).Text), "#0")
                
                End If
            End If
        
        Case ptxL_URIKIN3          '���i(3)
        
            If Trim(Text1(ptxL_URIKIN3).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxL_URIKIN3).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(ptxL_URIKIN3).SetFocus
                    Exit Function
                Else
                
                    Text1(ptxL_URIKIN3).Text = Format(CLng(Text1(ptxL_URIKIN3).Text), "#0")
                
                End If
            End If
        
        
        
    End Select
        
    Error_Check_Proc = False


End Function
Private Function Item_Disp_Proc(CODE As String) As Integer
'----------------------------------------------------------------------------
'                   ��ʕ\��
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim i       As Integer

Dim L_CODE  As String

    Item_Disp_Proc = True
    
    '�i��Ͻ��ǂݍ���
    Call UniCode_Conv(K0_ITEM_B.JGYOBU, Right(Combo1(pcmbJGYOBU).Text, 1))
    Call UniCode_Conv(K0_ITEM_B.NAIGAI, Right(Combo1(pcmbNAIGAI).Text, 1))
    Call UniCode_Conv(K0_ITEM_B.HIN_GAI, Text1(ptxHIN_GAI).Text)
    
    sts = BTRV(BtOpGetEqual, ITEM_B_POS, ITEM_BREC, Len(ITEM_BREC), K0_ITEM_B, Len(K0_ITEM_B), 0)
    Select Case sts
        Case BtNoErr
            'ں��ޓ��e�̕\��
                                            '�i�ں���
            Text1(ptxHIN_GAI).Text = Trim(StrConv(ITEM_BREC.HIN_GAI, vbUnicode))
                                            '�i��
            Text1(ptxHIN_NAME).Text = Trim(StrConv(ITEM_BREC.HIN_NAME, vbUnicode))
                                            '�i��E
            Text1(ptxL_HIN_NAME_E).Text = Trim(StrConv(ITEM_BREC.L_HIN_NAME_E, vbUnicode))
                                            '��Ж�
            If Trim(StrConv(ITEM_BREC.L_KAISHA_CODE, vbUnicode)) = "" Then
                Combo1(pcmbL_KAISHA).ListIndex = -1
            Else
                
                
                For i = 0 To Combo1(pcmbL_KAISHA).ListCount - 1
                    
                    L_CODE = Left(Right(Combo1(pcmbL_KAISHA).List(i), 4), 2)
                    If Trim(L_CODE) = "" Then
                        L_CODE = Right(Combo1(pcmbL_KAISHA).List(i), 2)
                    End If
                    
                    
                    If StrConv(ITEM_BREC.L_KAISHA_CODE, vbUnicode) = L_CODE Then
                        Combo1(pcmbL_KAISHA).ListIndex = i
                        Exit For
                        
                    End If
                
                
                Next i
            End If
                                            '���ƕ�
            If Trim(StrConv(ITEM_BREC.L_JGYOBU_CODE, vbUnicode)) = "" Then
                Combo1(pcmbL_JGYOBU).ListIndex = -1
            Else
                For i = 0 To Combo1(pcmbL_JGYOBU).ListCount - 1
                    L_CODE = Left(Right(Combo1(pcmbL_JGYOBU).List(i), 4), 2)
                    If Trim(L_CODE) = "" Then
                        L_CODE = Right(Combo1(pcmbL_JGYOBU).List(i), 2)
                    End If
                    
                    
                    If StrConv(ITEM_BREC.L_JGYOBU_CODE, vbUnicode) = L_CODE Then
                        Combo1(pcmbL_JGYOBU).ListIndex = i
                        Exit For
                        
                    End If
                
                
                Next i
            End If
                                            '���l
            Text1(ptxL_BIKOU).Text = Trim(StrConv(ITEM_BREC.L_BIKOU, vbUnicode))
                                            '���l3
            Text1(ptxL_BIKOU3).Text = Trim(StrConv(ITEM_BREC.L_BIKOU3, vbUnicode))
                                            '���萔
            If Not IsNumeric(Trim(StrConv(ITEM_BREC.L_IRI_QTY, vbUnicode))) Then
                Text1(ptxL_IRI_QTY).Text = ""
            Else
                Text1(ptxL_IRI_QTY).Text = CLng(StrConv(ITEM_BREC.L_IRI_QTY, vbUnicode))
            End If
                                            '�I��(1)
            Text1(ptxL_TANA1).Text = Trim(StrConv(ITEM_BREC.L_TANA1, vbUnicode))
                                            '�I��(2)
            Text1(ptxL_TANA2).Text = Trim(StrConv(ITEM_BREC.L_TANA2, vbUnicode))
                                            '��
'            If StrConv(ITEM_BREC.L_PAPER, vbUnicode) = L_PAPER_OFF Then
'                Check1(pchkL_PAPER).Value = vbUnchecked
'            Else
'                Check1(pchkL_PAPER).Value = vbChecked
'            End If
                                            
                                            
            If StrConv(ITEM_BREC.L_PAPER, vbUnicode) = L_PAPER_ON Then
                Check1(pchkL_PAPER).Value = vbChecked
            Else
                Check1(pchkL_PAPER).Value = vbUnchecked
            End If
                                            
                                            '�v��
'            If StrConv(ITEM_BREC.L_PLASTIC, vbUnicode) = L_PLASTIC_OFF Or StrConv(ITEM_BREC.L_PLASTIC, vbUnicode) <= " " Then
'                Check1(pchkL_PLASTIC).Value = vbUnchecked
'            Else
'                Check1(pchkL_PLASTIC).Value = vbChecked
'            End If
                                            
                                            
            If StrConv(ITEM_BREC.L_PLASTIC, vbUnicode) = L_PLASTIC_ON Then
                Check1(pchkL_PLASTIC).Value = vbChecked
            Else
                Check1(pchkL_PLASTIC).Value = vbUnchecked
            End If
                                            
                                            
                                            '�K�p�@�탉�x��
'            If StrConv(ITEM_BREC.L_LABEL, vbUnicode) = L_LABEL_OFF Or StrConv(ITEM_BREC.L_LABEL, vbUnicode) <= " " Then
'                Check1(pchkL_LABEL).Value = vbUnchecked
'            Else
'                Check1(pchkL_LABEL).Value = vbChecked
'            End If
                                            
                                            
            If StrConv(ITEM_BREC.L_LABEL, vbUnicode) = L_LABEL_ON Then
                Check1(pchkL_LABEL).Value = vbChecked
            Else
                Check1(pchkL_LABEL).Value = vbUnchecked

            End If
                                            
                                            '�������x��
'            If StrConv(ITEM_BREC.L_MAISU, vbUnicode) = L_MAISU_OFF Or StrConv(ITEM_BREC.L_MAISU, vbUnicode) <= " " Then
'                Check1(pchkL_MAISU).Value = vbUnchecked
'            Else
'                Check1(pchkL_MAISU).Value = vbChecked
'            End If
                                            
            If StrConv(ITEM_BREC.L_MAISU, vbUnicode) = L_MAISU_ON Then
                Check1(pchkL_MAISU).Value = vbChecked
            Else
                Check1(pchkL_MAISU).Value = vbUnchecked
            End If
                                            
                                            
                                            '���i(1)
            If Not IsNumeric(Trim(StrConv(ITEM_BREC.L_URIKIN1, vbUnicode))) Then
                Text1(ptxL_URIKIN1).Text = ""
            Else
                Text1(ptxL_URIKIN1).Text = Format(CDbl(StrConv(ITEM_BREC.L_URIKIN1, vbUnicode)), "#0")
            End If
                                            '���i(2)
            If Not IsNumeric(Trim(StrConv(ITEM_BREC.L_URIKIN2, vbUnicode))) Then
                Text1(ptxL_URIKIN2).Text = ""
            Else
                Text1(ptxL_URIKIN2).Text = Format(CDbl(StrConv(ITEM_BREC.L_URIKIN2, vbUnicode)), "#0")
            End If
                                            '���i(3)
            If Not IsNumeric(Trim(StrConv(ITEM_BREC.L_URIKIN3, vbUnicode))) Then
                Text1(ptxL_URIKIN3).Text = ""
            Else
                Text1(ptxL_URIKIN3).Text = Format(CDbl(StrConv(ITEM_BREC.L_URIKIN3, vbUnicode)), "#0")
            End If
                                            
                                            
                                            
            '���Y��     2008.06.12
            Text1(ptxGENSANKOKU).Text = Trim(StrConv(ITEM_BREC.GENSANKOKU, vbUnicode))
            
            If GENSANKOKU_SET_Proc() Then
                Exit Function
            End If
            
            If GENSANKOKU_FLG = "1" Then
                Check1(pchkGENSANKOKU).Value = vbChecked
            Else
                Check1(pchkGENSANKOKU).Value = vbUnchecked
            End If
                                            
                                            
                                            
                                            
                                            
                                            
                                            
                                            
                                            '��Ǝw��
            RichTextBox1(prchL_SAGYO_SHIJI).Text = IIf(Len(RTrim(StrConv(ITEM_BREC.L_SAGYO_SHIJI, vbUnicode))) = 450, "", Trim(StrConv(ITEM_BREC.L_SAGYO_SHIJI, vbUnicode)))
                                            '�@��(1)
            Text1(ptxL_KISHU1).Text = Trim(StrConv(ITEM_BREC.L_KISHU1, vbUnicode))
                                            '�@��(2)
'            Text1(ptxL_KISHU2).Text = Trim(StrConv(ITEM_BREC.L_KISHU2, vbUnicode))
            ' 2006.02.06 KUBOTA IIF�Ń������s���G���[�����
            RichTextBox1(prchL_KISHU2).Text = IIf(Len(RTrim(StrConv(ITEM_BREC.L_KISHU2, vbUnicode))) = 52, "", RTrim(StrConv(ITEM_BREC.L_KISHU2, vbUnicode)))
                                            '�@��(3)
'            RichTextBox1(prchL_KISHU3).Text = Trim(StrConv(ITEM_BREC.L_KISHU3, vbUnicode))
            RichTextBox1(prchL_KISHU3).Text = IIf(Len(RTrim(StrConv(ITEM_BREC.L_KISHU_BIKOU, vbUnicode))) = 450, "", Trim(StrConv(ITEM_BREC.L_KISHU_BIKOU, vbUnicode)))
                                            '�K�p�@����l
'            RichTextBox1(prchL_KISHU_BIKOU).Text = Trim(StrConv(ITEM_BREC.L_KISHU_BIKOU, vbUnicode))
            RichTextBox1(prchL_KISHU_BIKOU).Text = IIf(Len(RTrim(StrConv(ITEM_BREC.L_KISHU3, vbUnicode))) = 150, "", Trim(StrConv(ITEM_BREC.L_KISHU3, vbUnicode)))
            '������t
            Text1(ptxL_PRI_DATE).Text = Format(Now, "YYYY/mm/DD")
        
        
        
            lblUpd_Tanto.Caption = StrConv(ITEM_BREC.UPD_TANTO, vbUnicode)
            lblUpd_DateTime.Caption = StrConv(ITEM_BREC.UPD_DATETIME, vbUnicode)
        
        
        Case BtErrKeyNotFound
        
            MsgBox "���[���ŕύX����Ă��܂��B�O��ʂɖ߂�܂��B"
            PM00040B2.Visible = False
            INIT_FLG = False
            
            Exit Function
                    
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Function
        
    End Select

    Item_Disp_Proc = False

End Function

Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   �i�ڃ}�X�^�o��
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer
Dim com     As Integer
Dim i       As Integer

Dim L_CODE  As String

    Update_Proc = True
    
    '�i�ڃ}�X�^�@�ǂݍ���
    Call UniCode_Conv(K0_ITEM_B.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM_B.NAIGAI, Right(Combo1(pcmbNAIGAI).Text, 1))
    Call UniCode_Conv(K0_ITEM_B.HIN_GAI, Text1(ptxHIN_GAI).Text)
    
    
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_B_POS, ITEM_BREC, Len(ITEM_BREC), K0_ITEM_B, Len(K0_ITEM_B), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Update_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
                Exit Function
        
        End Select
    
    
    Loop
    '--------------------------------------------------------���R�[�h���e�ҏW
    
    If com = BtOpInsert Then
        Call UniCode_Conv(ITEM_BREC.JGYOBU, Last_JGYOBU)                              '���ƕ�
        Call UniCode_Conv(ITEM_BREC.NAIGAI, Right(Combo1(pcmbNAIGAI).Text, 1))        '�����O
        Call UniCode_Conv(ITEM_BREC.HIN_GAI, Text1(ptxHIN_GAI).Text)                  '�i�ں���
        
        Call UniCode_Conv(ITEM_BREC.ST_SET_DT, "")                                    '�W���I�Ԑݒ���t
        Call UniCode_Conv(ITEM_BREC.ST_SOKO, "")                                      '�W�����Ɂ@�q��
        Call UniCode_Conv(ITEM_BREC.ST_RETU, "")                                      '�W�����Ɂ@��
        Call UniCode_Conv(ITEM_BREC.ST_REN, "")                                       '�W�����Ɂ@�A
        Call UniCode_Conv(ITEM_BREC.ST_DAN, "")                                       '�W�����Ɂ@�i
        Call UniCode_Conv(ITEM_BREC.BEF_SOKO, "")                                     '�O����Ɂ@�q��
        Call UniCode_Conv(ITEM_BREC.BEF_RETU, "")                                     '�O����Ɂ@��
        Call UniCode_Conv(ITEM_BREC.BEF_REN, "")                                      '�O����Ɂ@�A
        Call UniCode_Conv(ITEM_BREC.BEF_DAN, "")                                      '�O����Ɂ@�i
        Call UniCode_Conv(ITEM_BREC.LAST_NYU_DT, "")                                  '�ŏI���ɓ�
        Call UniCode_Conv(ITEM_BREC.LAST_SYU_DT, "")                                  '�ŏI�o�ɓ�
        Call UniCode_Conv(ITEM_BREC.HIN_NAI, "")                                      '�i�ԁi���j
        Call UniCode_Conv(ITEM_BREC.BIKOU_SOKO, "")                                   'νđq��
        Call UniCode_Conv(ITEM_BREC.BIKOU_TANA, "")                                   'νĒI��
        Call UniCode_Conv(ITEM_BREC.HOJYU_P, "00000000")                              '��[�_
        Call UniCode_Conv(ITEM_BREC.AVE_SYUKA, "00000000")                            '�����Ϗo�א�
        Call UniCode_Conv(ITEM_BREC.SAMPLE_QTY, "0")                                  '����ِ�
        Call UniCode_Conv(ITEM_BREC.SAMPLE_QTY, "0")                                  '����ِ�
        Call UniCode_Conv(ITEM_BREC.LAST_INP_DT, "")                                  '�ŏI���ד��t
        Call UniCode_Conv(ITEM_BREC.LAST_CHK_DT, "")                                  '�ŏI�ƍ����t
        Call UniCode_Conv(ITEM_BREC.LAST_CHK_QTY, "00000000")                         '�ƍ����݌ɐ�
        Call UniCode_Conv(ITEM_BREC.BIKOU, "")                                        '������l
        Call UniCode_Conv(ITEM_BREC.IRI_QTY, "")                                      '������萔
        Call UniCode_Conv(ITEM_BREC.JAN_CODE, "")                                     'JAN����
        Call UniCode_Conv(ITEM_BREC.HIN_CHANGE, "")                                   '�i�ԓǂݑւ�����
        Call UniCode_Conv(ITEM_BREC.GOODS_KBN, "1")                                   '���i���L��
        Call UniCode_Conv(ITEM_BREC.PACKING_NO, "")                                   '������
        Call UniCode_Conv(ITEM_BREC.RANK, "")                                         '�����ݸ
        Call UniCode_Conv(ITEM_BREC.NEW_RANK, "")                                     '�V�ݸ
        Call UniCode_Conv(ITEM_BREC.GLICS1_TANA, "")                                  '��د���I��1
        Call UniCode_Conv(ITEM_BREC.GLICS2_TANA, "")                                  '��د���I��2
        Call UniCode_Conv(ITEM_BREC.GLICS3_TANA, "")                                  '��د���I��3
    
        Call UniCode_Conv(ITEM_BREC.G_SHIIRE_KBN, "")                                 '�Ɩ��Ǘ��@ �d���敪
        Call UniCode_Conv(ITEM_BREC.G_HANBAI_KBN, "")                                 '           �̔��敪
        Call UniCode_Conv(ITEM_BREC.G_SYUSHI, "")                                     '           ���x�P��
        Call UniCode_Conv(ITEM_BREC.G_KUMITATE, "")                                   '           �g�����i
        Call UniCode_Conv(ITEM_BREC.G_ST_URITAN, "")                                  '           �W���e�������P���@9(8)V99
        Call UniCode_Conv(ITEM_BREC.G_ST_URITAN_DT, "")                               '           �W���e�������ݒ��
        Call UniCode_Conv(ITEM_BREC.G_ST_SHITAN, "")                                  '           �W���e�������P��  9(8)V99
        Call UniCode_Conv(ITEM_BREC.G_ST_SHITAN_DT, "")                               '           �W���e�������ݒ��
        
        For i = 0 To 2                                                              '�d������
            Call UniCode_Conv(ITEM_BREC.G_SHIIRE_TBL(i).CODE, "")                     '           �d����R�[�h
            Call UniCode_Conv(ITEM_BREC.G_SHIIRE_TBL(i).TANKA, "")                    '           �P��
            Call UniCode_Conv(ITEM_BREC.G_SHIIRE_TBL(i).TANKA_DT, "")                 '           �P���ݒ��
            Call UniCode_Conv(ITEM_BREC.G_SHIIRE_TBL(i).LOT, "")                      '           �P���ݒ��
            Call UniCode_Conv(ITEM_BREC.G_SHIIRE_TBL(i).LEAD_TIME, "")                '           ���[�h�^�C��
            Call UniCode_Conv(ITEM_BREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, "")            '           �ŏI������
            Call UniCode_Conv(ITEM_BREC.G_SHIIRE_TBL(i).LAST_ORDER_QTY, "")           '           �ŏI������
        
        Next i
    
        Call UniCode_Conv(ITEM_BREC.G_ZEN_ZAIKO_KIN, "")                              '           �O���݌ɋ��z
        Call UniCode_Conv(ITEM_BREC.G_SHIIRE_KBN, "")                                 '           ���ދ敪
        Call UniCode_Conv(ITEM_BREC.G_LABEL_NON, P_G_LABEL_ON)                        '           ���ٓ\��t��
        Call UniCode_Conv(ITEM_BREC.S_TANTO, "")                                      '���P�^�S����
        
        Call UniCode_Conv(ITEM_BREC.FILLER, "")                                       'Filler
    
    End If
    
    Call UniCode_Conv(ITEM_BREC.HIN_NAME, Text1(ptxHIN_NAME).Text)                    '�i��
        
    Call UniCode_Conv(ITEM_BREC.L_HIN_NAME_E, Text1(ptxL_HIN_NAME_E).Text)            '�i��E
                                                                                        
                                                                                    '��Ж�
'    Call UniCode_Conv(ITEM_BREC.L_KAISHA_CODE, Left(Right(Combo1(pcmbL_KAISHA).Text, 4), 2))
                                                                                    '���ƕ���
'    Call UniCode_Conv(ITEM_BREC.L_JGYOBU_CODE, Left(Right(Combo1(pcmbL_JGYOBU).Text, 4), 2))
    
    
     L_CODE = Left(Right(Combo1(pcmbL_KAISHA).Text, 4), 2)
     If Trim(L_CODE) = "" Then
         L_CODE = Right(Combo1(pcmbL_KAISHA).Text, 2)
     End If
     Call UniCode_Conv(ITEM_BREC.L_KAISHA_CODE, L_CODE)
    
     L_CODE = Left(Right(Combo1(pcmbL_JGYOBU).Text, 4), 2)
     If Trim(L_CODE) = "" Then
         L_CODE = Right(Combo1(pcmbL_JGYOBU).Text, 2)
     End If
     Call UniCode_Conv(ITEM_BREC.L_JGYOBU_CODE, L_CODE)
    
    
    
    
    Call UniCode_Conv(ITEM_BREC.L_BIKOU, Text1(ptxL_BIKOU).Text)                      '���l
    Call UniCode_Conv(ITEM_BREC.L_BIKOU3, Text1(ptxL_BIKOU3).Text)                    '���l3
    
    If Trim(Text1(ptxL_IRI_QTY).Text) = "" Then                                     '���萔
        Call UniCode_Conv(ITEM_BREC.L_IRI_QTY, "")
    Else
        Call UniCode_Conv(ITEM_BREC.L_IRI_QTY, Format(CLng((Text1(ptxL_IRI_QTY).Text)), "00000000"))
    End If
    
    Call UniCode_Conv(ITEM_BREC.L_TANA1, Text1(ptxL_TANA1).Text)                      '�I��(1)
    Call UniCode_Conv(ITEM_BREC.L_TANA2, Text1(ptxL_TANA2).Text)                      '�I��(2)
    
    If Check1(pchkL_PAPER).Value = vbChecked Then                                   '��
        Call UniCode_Conv(ITEM_BREC.L_PAPER, L_PAPER_ON)
    Else
        Call UniCode_Conv(ITEM_BREC.L_PAPER, L_PAPER_OFF)
    End If
    
    If Check1(pchkL_PLASTIC).Value = vbChecked Then                                 '�v���X�`�b�N
        Call UniCode_Conv(ITEM_BREC.L_PLASTIC, L_PLASTIC_ON)
    Else
        Call UniCode_Conv(ITEM_BREC.L_PLASTIC, L_PLASTIC_OFF)
    End If
    
    If Check1(pchkL_LABEL).Value = vbChecked Then                                   '�K�p�@�탉�x��
        Call UniCode_Conv(ITEM_BREC.L_LABEL, L_LABEL_ON)
    Else
        Call UniCode_Conv(ITEM_BREC.L_LABEL, L_LABEL_OFF)
    End If
    
    If Check1(pchkL_MAISU).Value = vbChecked Then                                   '�������x��
        Call UniCode_Conv(ITEM_BREC.L_MAISU, L_MAISU_ON)
    Else
        Call UniCode_Conv(ITEM_BREC.L_MAISU, L_MAISU_OFF)
    End If
    
    If Trim(Text1(ptxL_URIKIN1).Text) = "" Then                                     '���i(1)
        Call UniCode_Conv(ITEM_BREC.L_URIKIN1, "")
    Else
        Call UniCode_Conv(ITEM_BREC.L_URIKIN1, Format(CDbl((Text1(ptxL_URIKIN1).Text)), "0000000000"))
    End If
    
    If Trim(Text1(ptxL_URIKIN2).Text) = "" Then                                     '���i(2)
        Call UniCode_Conv(ITEM_BREC.L_URIKIN2, "")
    Else
        Call UniCode_Conv(ITEM_BREC.L_URIKIN2, Format(CDbl((Text1(ptxL_URIKIN2).Text)), "0000000000"))
    End If
    
    If Trim(Text1(ptxL_URIKIN3).Text) = "" Then                                     '���i(3)
        Call UniCode_Conv(ITEM_BREC.L_URIKIN3, "")
    Else
        Call UniCode_Conv(ITEM_BREC.L_URIKIN3, Format(CDbl((Text1(ptxL_URIKIN3).Text)), "0000000000"))
    End If
    
    '���Y�� 2008.06.12
    Call UniCode_Conv(ITEM_BREC.GENSANKOKU, Text1(ptxGENSANKOKU).Text)
        
    
    Call UniCode_Conv(ITEM_BREC.L_SAGYO_SHIJI, RichTextBox1(prchL_SAGYO_SHIJI).Text)         '��Ǝw��
    Call UniCode_Conv(ITEM_BREC.L_KISHU1, Text1(ptxL_KISHU1).Text)                    '�@��(1)
    Call UniCode_Conv(ITEM_BREC.xL_KISHU2, "")                                        '���@��(2)
    Call UniCode_Conv(ITEM_BREC.L_KISHU2, RichTextBox1(prchL_KISHU2).Text)            '�@��(2)
 '   Call UniCode_Conv(ITEM_BREC.L_KISHU3, RichTextBox1(prchL_KISHU3).Text)           '�@��(3)
    Call UniCode_Conv(ITEM_BREC.L_KISHU3, RichTextBox1(prchL_KISHU_BIKOU).Text)       '�@��(3)
'    Call UniCode_Conv(ITEM_BREC.L_KISHU_BIKOU, RichTextBox1(prchL_KISHU_BIKOU).Text)  '�K�p�@��
    Call UniCode_Conv(ITEM_BREC.L_KISHU_BIKOU, RichTextBox1(prchL_KISHU3).Text)  '�K�p�@��
    
    Call UniCode_Conv(ITEM_BREC.UPD_TANTO, "")                                        '�X�V�S���Һ���
                                                                                    '�X�V����
    Call UniCode_Conv(ITEM_BREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    
    Do
        
        DoEvents
        
        sts = BTRV(com, ITEM_B_POS, ITEM_BREC, Len(ITEM_BREC), K0_ITEM_B, Len(K0_ITEM_B), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, ITEM_B_POS, ITEM_BREC, Len(ITEM_BREC), K0_ITEM_B, Len(K0_ITEM_B), 0)
                    Update_Proc = False
                    Exit Do
                End If
            Case Else
                Call File_Error(sts, com, "�i�ڃ}�X�^")
                Exit Function
        End Select
    
    Loop
    
    
    Update_Proc = False


End Function
Private Function Delete_Proc() As Integer
'----------------------------------------------------------------------------
'                   �i�ڃ}�X�^�폜
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer

    Delete_Proc = True
    
    '�i�ڃ}�X�^�@�ǂݍ���
    Call UniCode_Conv(K0_ITEM_B.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM_B.NAIGAI, Right(Combo1(pcmbNAIGAI).Text, 1))
    Call UniCode_Conv(K0_ITEM_B.HIN_GAI, Text1(ptxHIN_GAI).Text)

    
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_B_POS, ITEM_BREC, Len(ITEM_BREC), K0_ITEM_B, Len(K0_ITEM_B), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
                Delete_Proc = False
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Delete_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
                Exit Function
        
        End Select
    
    
    Loop
    
    
    Do
        
        DoEvents
        
        sts = BTRV(BtOpDelete, ITEM_B_POS, ITEM_BREC, Len(ITEM_BREC), K0_ITEM_B, Len(K0_ITEM_B), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, ITEM_B_POS, ITEM_BREC, Len(ITEM_BREC), K0_ITEM_B, Len(K0_ITEM_B), 0)
                    Delete_Proc = False
                    Exit Do
                End If
            Case Else
                Call File_Error(sts, BtOpDelete, "�i�ڃ}�X�^")
                Exit Function
        End Select
    Loop


    Delete_Proc = False


End Function


Private Sub Check1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
        
    Call Tab_Ctrl(Shift)        '�ړ�

End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    
Dim i   As Integer
    
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
    
    
    Select Case Index
    
        Case pcmbJGYOBU
    
            
            For i = 0 To UBound(JGYOBU_T)
                If Right(Combo1(pcmbJGYOBU).Text, 1) = JGYOBU_T(i).CODE Then
                
                    
                    Last_JGYOBU = JGYOBU_T(i).CODE
                    Exit For
                
                End If
            Next i
    
    
    End Select
    
    
    Call Tab_Ctrl(Shift)        '�ړ�

End Sub


Private Sub Combo1_LostFocus(Index As Integer)
Dim i   As Integer
    
    
    
    Select Case Index
    
        Case pcmbJGYOBU
    
            For i = 0 To UBound(JGYOBU_T)
                If Right(Combo1(pcmbJGYOBU).Text, 1) = JGYOBU_T(i).CODE Then
                
                    
                    Last_JGYOBU = JGYOBU_T(i).CODE
                    Exit For
                
                End If
            Next i
    
    
    
    End Select

End Sub

Private Sub Command1_Click(Index As Integer)

Dim ans         As Integer
Dim i           As Integer

Dim Parts       As String   '�i��
Dim ID          As Long     '�w����

Dim PartsLabel  As Integer  '�i������ 0:�Ȃ� �ȊO�F����
Dim KisyuLabel  As Integer  '�@������ 0:�Ȃ�
Dim JanLabel    As Integer  'JAN���� 0:�Ȃ�
Dim GLabel      As Integer  '�O������ 0:�Ȃ�
Dim ItemLabel   As Integer  '�������ٖ���

Dim OrderNo     As String
Dim ItemNo      As String

Dim Pri_Date    As String

Dim objAccess       As Access.Application
Dim strAccessPath   As String

Dim com         As Integer
Dim sts         As Integer

Dim L_Print_Flg As Boolean
    
Dim check_flg   As Boolean
    
    
Dim check_flg1  As Boolean      '2008.09.26
    
    
Dim L_CODE      As String
    
Dim FileNo      As Long         '2008.05.30
    
    
Dim L_QTY       As Long         '2008.10.03
    
    
    Select Case Index
        Case P_CMD_Upd                      '�X�V
            
            
            For i = ptxHIN_GAI To ptxL_KISHU1
            
                If Error_Check_Proc(i) Then     '�G���[�`�F�b�N
                    Exit Sub
                End If
            
            Next i
            
            
            Beep
            ans = MsgBox("�X�V���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
            Else
                Exit Sub
            End If
'            PM000402.Visible = False
'            INIT_FLG = False
                    
            Call Clear_Proc
        
        Case P_CMD_DEL                      '�폜
            ans = MsgBox("�폜���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Delete_Proc() Then
                    Unload Me
                End If
            Else
                Exit Sub
            End If
'            PM000402.Visible = False
'            INIT_FLG = False
 
            Call Clear_Proc
 
 
 
 '       Case P_CMD_DSP                      '����/�\��
 '       Case P_CMD_OUT                      '�ް��o��
 '       Case P_CMD_PRT                      '���
        
        Case pcmdLabel, pcmdItem, pcmdJan, pcmdGAISO
            If Not IsNumeric(Text1(ptxL_MAISU).Text) Then
        
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxL_MAISU).SetFocus
                Exit Sub
        
            Else
                If CInt(Text1(ptxL_MAISU).Text) <= 0 Then
                
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(ptxL_MAISU).SetFocus
                    Exit Sub
                
                End If
            
            End If
            
            If Trim(Text1(ptxL_PRI_DATE).Text) = "" Then
            Else
                If Not IsDate(Text1(ptxL_PRI_DATE).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(ptxL_MAISU).SetFocus
                    Exit Sub
                End If
            End If
        
            L_Print_Flg = True
        
        
        
        
        
        
            check_flg1 = False                              '2008.09.26
            If Trim(Combo1(pcmbL_KAISHA).Text) = "" Then    '2008.09.26
            Else                                            '2008.09.26
                check_flg1 = True                           '2008.09.26
            End If                                          '2008.09.26
            check_flg1 = False                              '2008.09.26
            If Trim(Combo1(pcmbL_JGYOBU).Text) = "" Then    '2008.09.26
            Else                                            '2008.09.26
                check_flg1 = True                           '2008.09.26
            End If                                          '2008.09.26
        
        
            If Not check_flg1 Then       '2008.09.26
            
                If KAISYA_CHK_F Then
            
'                    MsgBox "��Ж��������͎��ƕ����󔒂ׁ̈A����ł��܂���"
'                    Text1(ptxHIN_GAI).SetFocus
'
'                    Exit Sub
                
                
                
                
                
                    ans = MsgBox("��Ж�/���ƕ� ���w�肳��Ă��܂���B(�n�j�����s�A��ݾ�=���s���Ȃ�)", vbOKCancel + vbQuestion + vbDefaultButton2, "�m�F����")
                    If ans = vbCancel Then
                        Text1(ptxHIN_GAI).SetFocus
                        Exit Sub
                    End If
                
                
                
                
                
                
                End If
            
            End If
        
        
        
            If KAISYA_CHK_F Then        '2008.09.26
            
            
            
                If Not IsNumeric(Text1(ptxL_URIKIN2).Text) Or _
                     Not IsNumeric(Text1(ptxL_URIKIN3).Text) Then
                
            '��2010.02.08
                    
                    
                    
                    If TANKA_SPACE_F = "1" Then
                    
                        ans = MsgBox("�P�����o�^�ł��B(�n�j���������s�A��ݾ�=���s���Ȃ�)", vbOKCancel + vbQuestion + vbDefaultButton2, "�m�F����")
                        If ans = vbCancel Then
                            Text1(ptxHIN_GAI).SetFocus
                            Exit Sub
                        End If
                    Else

                        MsgBox "�P�����o�^�ׁ̈A���s�ł��܂���"
                        Text1(ptxHIN_GAI).SetFocus
                        Exit Sub
            
                    End If
            '��2010.02.08
                
                End If
            
            
            
            
                check_flg = True
            
            
            Else
                check_flg = False
                If Not IsNumeric(Text1(ptxL_URIKIN1).Text) Then
                Else
                    If CDbl(Text1(ptxL_URIKIN1).Text) <> 0 Then
                        check_flg = True
                    End If
                End If
                
                If Not IsNumeric(Text1(ptxL_URIKIN2).Text) Then
                Else
                    If CDbl(Text1(ptxL_URIKIN2).Text) <> 0 Then
                        check_flg = True
                    End If
                End If
                If Not IsNumeric(Text1(ptxL_URIKIN3).Text) Then
                Else
                    If CDbl(Text1(ptxL_URIKIN3).Text) <> 0 Then
                        check_flg = True
                    End If
                End If
            End If
            
            
            If PRINT_CHECK_F Then       '2008.09.26
            
            
                '��2008.05.30
                Do
                    On Error Resume Next
    
                    FileNo = FreeFile
    
                    Open LabelPrint_F For Input As FileNo
    
                    Select Case Err.Number
                        Case 0
    
    
                            Close #FileNo
    
                            ans = MsgBox("���x�����s���ł�", vbAbortRetryIgnore + vbQuestion, "�m�F����")
    
                            Select Case ans
                            
                                Case vbAbort    '���~
    
                                    Exit Sub
                            
                                Case vbIgnore   '����
                            
                                    Exit Do
                            
                            End Select
    
    
    
    
                        Case 53
                            Exit Do
    
    
                        Case Else
    
                            Unload Me
    
    
                    End Select
    
                    On Error GoTo 0
    
                Loop
                
                Open LabelPrint_F For Output As FileNo
                Close #FileNo
            
            End If
            '��2008.05.30
            
            
            
            
            
            If Not check_flg Then
                ans = MsgBox("�P�����ݒ�ł��B���x��������܂����H", vbYesNo + vbQuestion, "�m�F����")
                If ans = vbYes Then
                Else
                    L_Print_Flg = False
                End If
            End If
            
            '2009.03.28
            For i = 0 To UBound(GENSANKOKU_CHECK_TBL)
            
            
                If Last_JGYOBU = GENSANKOKU_CHECK_TBL(i) Then
                    Exit For
                End If
            
            Next i
            '2009.03.28
            If i > UBound(GENSANKOKU_CHECK_TBL) Then
            Else
                
                
                If Trim(Text1(ptxGENSANKOKU).Text) = "" Then
                    

                    ans = MsgBox("���Y�����󔒂ł��B(�n�j��������~�A��ݾ�=�p��)", vbOKCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                    Else
                        L_Print_Flg = False
                    End If
                End If
            End If
                
                
                
                
                
            If L_Print_Flg Then
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
'-----------------  ں��ނ̒��g����ւ�
                Call UniCode_Conv(ITEM_BREC.HIN_NAME, Text1(ptxHIN_NAME).Text)                    '�i��
                    
                Call UniCode_Conv(ITEM_BREC.L_HIN_NAME_E, Text1(ptxL_HIN_NAME_E).Text)            '�i��E
                                                                                                    
                                                                                                '��Ж�
                
                        
                L_CODE = Left(Right(Combo1(pcmbL_KAISHA).Text, 4), 2)
                If Trim(L_CODE) = "" Then
                    L_CODE = Right(Combo1(pcmbL_KAISHA).Text, 2)
                End If
                Call UniCode_Conv(ITEM_BREC.L_KAISHA_CODE, L_CODE)
               
                L_CODE = Left(Right(Combo1(pcmbL_JGYOBU).Text, 4), 2)
                If Trim(L_CODE) = "" Then
                    L_CODE = Right(Combo1(pcmbL_JGYOBU).Text, 2)
                End If
                Call UniCode_Conv(ITEM_BREC.L_JGYOBU_CODE, L_CODE)
                
                Call UniCode_Conv(ITEM_BREC.L_BIKOU, Text1(ptxL_BIKOU).Text)                      '���l
                Call UniCode_Conv(ITEM_BREC.L_BIKOU3, Text1(ptxL_BIKOU3).Text)                    '���l3
                
                If Trim(Text1(ptxL_IRI_QTY).Text) = "" Then                                     '���萔
                    Call UniCode_Conv(ITEM_BREC.L_IRI_QTY, "")
                Else
                    Call UniCode_Conv(ITEM_BREC.L_IRI_QTY, Format(CLng((Text1(ptxL_IRI_QTY).Text)), "00000000"))
                End If
                
                Call UniCode_Conv(ITEM_BREC.L_TANA1, Text1(ptxL_TANA1).Text)                      '�I��(1)
                
                '2008.10.29 �I��(1)�ɕW���I�Ԃ��Z�b�g
                Call UniCode_Conv(ITEM_BREC.L_TANA1, StrConv(ITEM_BREC.ST_SOKO, vbUnicode) & "-" & _
                                                    StrConv(ITEM_BREC.ST_RETU, vbUnicode) & "-" & _
                                                    StrConv(ITEM_BREC.ST_REN, vbUnicode) & "-" & _
                                                    StrConv(ITEM_BREC.ST_DAN, vbUnicode))
                
                '2008.10.29
                
                
                Call UniCode_Conv(ITEM_BREC.L_TANA2, Text1(ptxL_TANA2).Text)                      '�I��(2)
                
                If Check1(pchkL_PAPER).Value = vbChecked Then                                   '��
                    Call UniCode_Conv(ITEM_BREC.L_PAPER, L_PAPER_ON)
                Else
                    Call UniCode_Conv(ITEM_BREC.L_PAPER, L_PAPER_OFF)
                End If
                
                If Check1(pchkL_PLASTIC).Value = vbChecked Then                                 '�v���X�`�b�N
                    Call UniCode_Conv(ITEM_BREC.L_PLASTIC, L_PLASTIC_ON)
                Else
                    Call UniCode_Conv(ITEM_BREC.L_PLASTIC, L_PLASTIC_OFF)
                End If
                
                If Check1(pchkL_LABEL).Value = vbChecked Then                                   '�K�p�@�탉�x��
                    Call UniCode_Conv(ITEM_BREC.L_LABEL, L_LABEL_ON)
                Else
                    Call UniCode_Conv(ITEM_BREC.L_LABEL, L_LABEL_OFF)
                End If
                
                If Check1(pchkL_MAISU).Value = vbChecked Then                                   '�������x��
                    Call UniCode_Conv(ITEM_BREC.L_MAISU, L_MAISU_ON)
                Else
                    Call UniCode_Conv(ITEM_BREC.L_MAISU, L_MAISU_OFF)
                End If
                
                If Trim(Text1(ptxL_URIKIN1).Text) = "" Then                                     '���i(1)
                    Call UniCode_Conv(ITEM_BREC.L_URIKIN1, "")
                Else
                    Call UniCode_Conv(ITEM_BREC.L_URIKIN1, Format(CDbl((Text1(ptxL_URIKIN1).Text)), "0000000000"))
                End If
                
                If Trim(Text1(ptxL_URIKIN2).Text) = "" Then                                     '���i(2)
                    Call UniCode_Conv(ITEM_BREC.L_URIKIN2, "")
                Else
                    Call UniCode_Conv(ITEM_BREC.L_URIKIN2, Format(CDbl((Text1(ptxL_URIKIN2).Text)), "0000000000"))
                End If
                
                If Trim(Text1(ptxL_URIKIN3).Text) = "" Then                                     '���i(3)
                    Call UniCode_Conv(ITEM_BREC.L_URIKIN3, "")
                Else
                    Call UniCode_Conv(ITEM_BREC.L_URIKIN3, Format(CDbl((Text1(ptxL_URIKIN3).Text)), "0000000000"))
                End If
                
                
                '���Y�� 2008.06.12
                If Check1(pchkGENSANKOKU).Value = vbChecked Then
                    
                    
                    If Text1(ptxGENSANKOKU).Enabled Then
                        
                        Call UniCode_Conv(ITEM_BREC.GENSANKOKU, Text1(ptxGENSANKOKU).Text)
                    Else
                                
                        If Combo1(pcmbGENSAN).Enabled Then
                            Call UniCode_Conv(ITEM_BREC.GENSANKOKU, Trim(Left(Combo1(pcmbGENSAN).Text, 20)))
                        End If
                    End If
                Else
                    Call UniCode_Conv(ITEM_BREC.GENSANKOKU, "")
                End If
                
                
                
                Call UniCode_Conv(ITEM_BREC.L_SAGYO_SHIJI, RichTextBox1(prchL_SAGYO_SHIJI).Text)  '��Ǝw��
                Call UniCode_Conv(ITEM_BREC.L_KISHU1, Text1(ptxL_KISHU1).Text)                    '�@��(1)
                Call UniCode_Conv(ITEM_BREC.xL_KISHU2, "")                                        '���@��(2)
                Call UniCode_Conv(ITEM_BREC.L_KISHU2, RichTextBox1(prchL_KISHU2).Text)            '�@��(2)
'                Call UniCode_Conv(ITEM_BREC.L_KISHU3, RichTextBox1(prchL_KISHU3).Text)           '�@��(3)
'                Call UniCode_Conv(ITEM_BREC.L_KISHU_BIKOU, RichTextBox1(prchL_KISHU_BIKOU).Text) '�K�p�@��

                Call UniCode_Conv(ITEM_BREC.L_KISHU3, RichTextBox1(prchL_KISHU_BIKOU).Text)       '�@��(3)
                Call UniCode_Conv(ITEM_BREC.L_KISHU_BIKOU, RichTextBox1(prchL_KISHU3).Text)       '�K�p�@��


'-----------------  ں��ނ̒��g����ւ�
                                
                                
                PartsLabel = 0
                KisyuLabel = 0
                JanLabel = 0
                GLabel = 0
                ItemLabel = 0

                Parts = Text1(ptxHIN_GAI).Text     '�i��
    
                    
                Select Case Index
                    Case pcmdLabel
                        If Check1(pchkL_LABEL).Value = vbChecked Then
                        
                            KisyuLabel = CInt(Text1(ptxL_MAISU).Text)
                        Else
                            PartsLabel = CInt(Text1(ptxL_MAISU).Text)
                        
                        
                        End If
                    Case pcmdItem
                    
                        ItemLabel = CInt(Text1(ptxL_MAISU).Text)
                                            
                    
                    Case pcmdJan
                        JanLabel = CInt(Text1(ptxL_MAISU).Text)
                    Case pcmdGAISO
                        GLabel = CInt(Text1(ptxL_MAISU).Text)
                End Select
                OrderNo = Text1(ptxL_ORDERNO).Text
                ItemNo = Text1(ptxL_ITEMNO).Text
                Pri_Date = Text1(ptxL_PRI_DATE).Text
                
                On Error Resume Next
                Set objAccess = GetObject(, "Access.Application")
                If Err().Number <> 0 Then
                    
                    MsgBox "���̒[���ł͏��i���x�����s�͍s���܂���B"
'                        MsgBox "GetObject(Access.Application)" & Err().Number & " " & Err().Description
                Else
'                        MsgBox Err.Number
                        
                    strAccessPath = App.Path
                    If Right(strAccessPath, 1) <> "\" Then
                        strAccessPath = strAccessPath & "\"
                    End If
                    
                    strAccessPath = strAccessPath & "litem.mdb"
                    Set objAccess = GetObject(strAccessPath)

                
                
                    
                    com = BtOpGetFirst
                    Do
                    
                    
                    
                        sts = BTRV(com, L_ITEM_POS, L_ITEMREC, Len(L_ITEMREC), K0_L_ITEM, Len(K0_L_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            
                                sts = BTRV(BtOpDelete, L_ITEM_POS, L_ITEMREC, Len(L_ITEMREC), K0_L_ITEM, Len(K0_L_ITEM), 0)
                                
                                Select Case sts
                                
                                    Case BtNoErr
                                    Case Else
                                        Call File_Error(sts, com, "���ٗp�i��Ͻ�")
                                        Exit Sub
                                End Select
                            
                            Case BtErrEOF
                                Exit Do
                            Case Else
                                Call File_Error(sts, com, "���ٗp�i��Ͻ�")
                                Exit Sub
                        End Select
                        
                        com = BtOpGetNext
                    
                    
                    Loop
                        
                    '�č���ϰ��ǉ�  2007.11.08
                    Call UniCode_Conv(ITEM_BREC.L_MARK, Text1(ptxL_MARK).Text)
                        
                        
                    sts = BTRV(BtOpInsert, L_ITEM_POS, ITEM_BREC, Len(ITEM_BREC), K0_L_ITEM, Len(K0_L_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                
                    
                
                
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                            Exit Sub
                        
                
                    End Select
                            
                    If IsNumeric(Text1(ptxL_QTY).Text) Then     '2008.10.03
                        L_QTY = CLng(Text1(ptxL_QTY).Text)      '2008.10.03
                    Else                                        '2008.10.03
                        L_QTY = "1"                             '2008.10.03
                    End If                                      '2008.10.03
                            
                            
                    ID = 0
'                    objAccess.Run "NewPosPrintLabel", _
'                                        Trim(Parts), _
'                                        PartsLabel, _
'                                        KisyuLabel, _
'                                        JanLabel, _
'                                        GLabel, _
'                                        ID, _
'                                        ItemLabel, _
'                                        Trim(OrderNo), _
'                                        Trim(ItemNo), _
'                                        Pri_Date

                    '2008.10.03 �����ǉ�(L_QTY)
                    objAccess.Run "NewPosPrintLabel", _
                                        Trim(Parts), _
                                        PartsLabel, _
                                        KisyuLabel, _
                                        JanLabel, _
                                        GLabel, _
                                        ID, _
                                        ItemLabel, _
                                        Trim(OrderNo), _
                                        Trim(ItemNo), _
                                        Pri_Date, _
                                        L_QTY
                
                
                End If
                
                
                
                
                
                Set objAccess = Nothing
            End If
            
            
            
            
            '2008.12.19
            Text1(ptxL_QTY).Text = "1"

                    
        
            'PM000402.Visible = False
            'INIT_FLG = False
        
        
        
        
        
        
        
        
        Case P_CMD_End                      '�I��
    
            Unload Me
    End Select

End Sub

Private Sub Form_Activate()
    
'Dim i       As Integer
'Dim CODE    As String
    
'    If INIT_FLG Then
'        Exit Sub
'    End If

'    If JGYOBU_T(i).CODE = Last_JGYOBU Then
'        PM000402.Caption = "���i���V�X�e���@�i�ڃ}�X�^�����e�i���X�i���i���x�����ځj�i" + RTrim(JGYOBU_T(i).NAME) + ")"
'        LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
'        LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'    End If



'    Select Case G_SCREEN_FLG
'        Case G_SCREEN_INS       '�V�K
'
'            Combo1(pcmbNAIGAI).BackColor = G_INPUT_OK
'            Combo1(pcmbNAIGAI).TabStop = True
'            Combo1(pcmbNAIGAI).Locked = False
'
'
'            Text1(ptxHIN_GAI).BackColor = G_INPUT_OK
'            Text1(ptxHIN_GAI).TabStop = True
'            Text1(ptxHIN_GAI).Locked = False
'
'            Text1(ptxHIN_NAME).BackColor = G_INPUT_OK
'            Text1(ptxHIN_NAME).TabStop = True
'            Text1(ptxHIN_NAME).Locked = False
'
'
'            For i = ptxHIN_GAI To ptxL_ITEMNO
'                Text1(i).Text = ""
'            Next i
'
'            For i = prchL_SAGYO_SHIJI To prchL_KISHU_BIKOU
'                RichTextBox1(i).Text = ""
'            Next i
'
'
'            For i = pcmbNAIGAI To pcmbL_JGYOBU
'
'                Combo1(i).ListIndex = -1
'            Next i
'
'
'
'
'            Combo1(pcmbNAIGAI).SetFocus
'            Combo1(pcmbNAIGAI).ListIndex = 0
'
'
'
'
'        Case G_SCREEN_UPD       '�X�V
'
'            Combo1(pcmbNAIGAI).BackColor = G_INPUT_NG
'            Combo1(pcmbNAIGAI).TabStop = False
'            Combo1(pcmbNAIGAI).Locked = True
'
'
'
'            Text1(ptxHIN_GAI).BackColor = G_INPUT_NG
'            Text1(ptxHIN_GAI).TabStop = False
'            Text1(ptxHIN_GAI).Locked = True
'
'            Text1(ptxHIN_NAME).BackColor = G_INPUT_OK
'            Text1(ptxHIN_NAME).TabStop = True
'            Text1(ptxHIN_NAME).Locked = False
'
'
'            CODE = PM000401.txSEL_KEY.Text
'
'            If Item_Disp_Proc(CODE) Then
'                Exit Sub
'            End If
'
'            For i = ptxL_MAISU To ptxL_ITEMNO
'                Text1(i).Text = ""
'            Next i
'
'            '========================================================= 2007/03/19 =====
'''            Text1(ptxL_HIN_NAME_E).SetFocus
'            Text1(ptxL_MAISU).SetFocus
'            '==========================================================================
'
'    End Select
'
'
'    INIT_FLG = True
'
End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   �j���� �c������ �O����
'----------------------------------------------------------------------------
    Select Case KeyCode
        Case vbKeyF1 To vbKeyF12
            Command1(KeyCode - vbKeyF1).Value = True
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()

Dim com     As Integer
Dim sts     As Integer




Dim c       As String * 128
Dim i       As Integer

    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If

                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
                                
                                
                                
    PRINT_CHECK_F = True        '2008.09.26
                                '���x������p�R���g���[���e�l��2008.05.30
    If GetIni("FILE", "labelprint", "SYS", c) Then
'        Beep
'        MsgBox "���x������p�R���g���[���e�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
'        End
        PRINT_CHECK_F = False   '2008.09.26
    Else
        LabelPrint_F = RTrim(c)
    End If
'    LabelPrint_F = RTrim(c)
                                
                                
                                '���Y���󎚗L�� 2008.06.12
    If GetIni(App.EXEName, "GENSANKOKU_DEF_F", "P_SYS", c) Then
        GENSANKOKU_FLG = "0"
    Else
        GENSANKOKU_FLG = RTrim(c)
    End If
                                
                                
                                '��Ў��ƕ��G���[�����L�� 2008.09.26
    If GetIni(App.EXEName, "KAISYA_CHECK", "P_SYS", c) Then
        KAISYA_CHK_F = False
    Else
        
        If Trim(c) = "1" Then
            KAISYA_CHK_F = True
        Else
            KAISYA_CHK_F = False
        End If
    End If
                                '���Y�������� 2009.03.28
    If GetIni(App.EXEName, "GENSANKOKU_CHECK", "P_SYS", c) Then
        ReDim GENSANKOKU_CHECK_TBL(0 To 0)
        GENSANKOKU_CHECK_TBL(0) = "*"
    Else
        GENSANKOKU_CHECK_TBL = Split(Trim(c))
    End If
                                
                                
                                
                                '�P�������� 2010.02.02
    If GetIni(App.EXEName, "TANKA_SPACE_F", "P_SYS", c) Then
        TANKA_SPACE_F = "0"
    Else
        If Trim(c) = "1" Then
            TANKA_SPACE_F = "1"
        Else
            TANKA_SPACE_F = "0"
        End If
    End If
                                
                                
                                
                                
                                
                                
                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
        
    Combo1(pcmbJGYOBU).Clear
    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If

        Combo1(pcmbJGYOBU).AddItem RTrim(JGYOBU_T(i).NAME) & "                             " & JGYOBU_T(i).CODE

        
    Next i
        
        
    For i = 0 To Combo1(pcmbJGYOBU).ListCount - 1
    
        
        If Right(Combo1(pcmbJGYOBU).List(i), 1) = Last_JGYOBU Then
        
            Combo1(pcmbJGYOBU).ListIndex = i
            Exit For
        
        End If
    
    Next i
        
        
        
        
'    For i = 0 To UBound(JGYOBU_T)
'        If JGYOBU_T(i).CODE = " " Then
'            Unload SubMenu(i)
'            Exit For
'        End If
'
'        Load SubMenu(i + 1)
'        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)
'
'        If JGYOBU_T(i).CODE = Last_JGYOBU Then
'            PM000402.Caption = "���i���V�X�e���@�i�ڃ}�X�^�����e�i���X�i���i���x�����ځj�i" + RTrim(JGYOBU_T(i).NAME) + ")"
'            SubMenu(i).Checked = True
'            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
'            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'        Else
'            SubMenu(i).Checked = False
'        End If
'    Next i
'
'    Unload SubMenu(i)
                                
                                
    PM00040B2.Caption = PM00040B2.Caption & " " & Last_Update_Day
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_B_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '�i�ڃ}�X�^�n�o�d�m
    If L_ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '���Y���}�X�^�n�o�d�m
    If GENSAN_Open(BtOpenNomal) Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        Unload Me
    End If
                                '�R�[�h�}�X�^�n�o�d�m
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�󕥐�}�X�^�i�d����j�n�o�d�m
    If P_UKEHARAI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
    Call P_CODE_TBL_Proc
                                
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  MT_2009.06.01
                                'PN�}�X�^�n�o�d�m
    If PN_M_Open(0) Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        Unload Me
    End If
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    

    Combo1(pcmbNAIGAI).AddItem NAIGAI1 & "   " & NAIGAI_NAI
    Combo1(pcmbNAIGAI).AddItem NAIGAI2 & "   " & NAIGAI_GAI
    Combo1(pcmbNAIGAI).ListIndex = 0
    
    '��Ж��̃Z�b�g
    If Code_Set_Proc(pcmbL_KAISHA, P_KBN07_CD) Then
        Unload Me
    End If
    
    '���ƕ����̃Z�b�g
    If Code_Set_Proc(pcmbL_JGYOBU, P_KBN07_CD) Then
        Unload Me
    End If
    
    Text1(ptxL_QTY).Text = "1"              '2008.10.03
    
    
    INIT_FLG = False
    
    
    
End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer



    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  MT_2009.06.01
                                            'PN�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, PN_M_POS, PN_MREC, Len(PN_MREC), K0_PN_M, Len(K0_PN_M), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "PN�}�X�^")
            Beep
            MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly
        End If
    End If
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
                                            
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_B_POS, ITEM_BREC, Len(ITEM_BREC), K0_ITEM_B, Len(K0_ITEM_B), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            
                                            '�R�[�h�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�R�[�h�}�X�^")
        End If
    End If
                                            '�󕥐�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�󕥐�}�X�^")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PM00040B2 = Nothing

    End
End Sub


Private Sub Text1_GotFocus(Index As Integer)
    
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
        
        
    If Index = ptxHIN_GAI Then
        Text1(ptxHIN_GAI).Text = StrConv(RTrim(Text1(ptxHIN_GAI).Text), vbUpperCase)
    End If
        
        
    If Error_Check_Proc(Index) Then     '�G���[�`�F�b�N
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '�ړ�

End Sub

Private Function Code_Set_Proc(Index As Integer, KBN As String) As Integer
'----------------------------------------------------------------------------
'                   �R�[�h�}�X�^���R���{�ɃZ�b�g����B
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim Key_Len     As Integer
Dim OPTION1     As Integer
Dim OPTION2     As Integer

Dim wkOption    As String



Dim i           As Integer
    
    Code_Set_Proc = True
    
    Combo1(Index).Clear
    
    For i = 0 To UBound(P_KBN_TBL)
    
        If KBN = P_KBN_TBL(i).KBN_CD Then
            Key_Len = P_KBN_TBL(i).KBN_Len
            Exit For
        End If
    
    Next i
    
    If i > UBound(P_KBN_TBL) Then
        Exit Function
    End If
    
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, KBN)
    Call UniCode_Conv(K0_P_CODE.C_Code, "")

    com = BtOpGetGreater

    Do
        DoEvents
    
        sts = BTRV(com, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            
        Select Case sts
            Case BtNoErr
            
                                
                If StrConv(P_CODEREC.DATA_KBN, vbUnicode) <> KBN Then
                    
                    Exit Do
                
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�R�[�h�}�X�^")
                Exit Function
        
        End Select

        wkOption = ""
        If P_KBN_TBL(i).KBN_OP1 Then
            wkOption = Trim(StrConv(P_CODEREC.OPTION1, vbUnicode))
        End If
        If P_KBN_TBL(i).KBN_OP2 Then
            wkOption = wkOption & Trim(StrConv(P_CODEREC.OPTION2, vbUnicode))
        End If
        
        
        
        Combo1(Index).AddItem StrConv(P_CODEREC.C_NAME, vbUnicode) & "                                        " & _
                                Left(StrConv(P_CODEREC.C_Code, vbUnicode), Key_Len) & wkOption
        
        
        com = BtOpGetNext
    
    Loop

    Code_Set_Proc = False
    



End Function



Private Sub Clear_Proc()
    
    
Dim i   As Integer
    
    
    For i = ptxHIN_GAI To ptxL_MARK
        Text1(i).Text = ""
    Next i

    For i = prchL_SAGYO_SHIJI To prchL_KISHU_BIKOU
        RichTextBox1(i).Text = ""
    Next i


    For i = pcmbL_KAISHA To pcmbL_JGYOBU

        Combo1(i).ListIndex = -1
    Next i

    Text1(ptxL_QTY).Text = "1"

    '2008.12.19
    Text1(ptxL_MAISU).Text = "1"

    
    Call UniCode_Conv(ITEM_BREC.HIN_GAI, "")


    Text1(ptxHIN_GAI).SetFocus

End Sub

Private Sub Text1_LostFocus(Index As Integer)

    If Index = ptxHIN_GAI Then
        Text1(ptxHIN_GAI).Text = StrConv(RTrim(Text1(ptxHIN_GAI).Text), vbUpperCase)
    End If

End Sub
Private Function GENSANKOKU_SET_Proc() As Integer
'----------------------------------------------------------------------------
'                   ���Y���}�X�^���R���{�ɃZ�b�g����B
'----------------------------------------------------------------------------

Dim sts     As Integer
Dim com     As Integer
Dim i       As Integer

    GENSANKOKU_SET_Proc = True
    
    
    
    
    
    Combo1(pcmbGENSAN).Clear
    List2.Clear
    
    
    
    Call UniCode_Conv(K0_GENSAN.JGYOBU, StrConv(ITEM_BREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_GENSAN.NAIGAI, StrConv(ITEM_BREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_GENSAN.HIN_GAI, StrConv(ITEM_BREC.HIN_GAI, vbUnicode))

    Call UniCode_Conv(K0_GENSAN.GENSANKOKU, "")
    
    com = BtOpGetGreater
    
    Do
        
        DoEvents
        
        sts = BTRV(com, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(ITEM_BREC.JGYOBU, vbUnicode) <> StrConv(GENSANREC.JGYOBU, vbUnicode) Or _
                    StrConv(ITEM_BREC.NAIGAI, vbUnicode) <> StrConv(GENSANREC.NAIGAI, vbUnicode) Or _
                    StrConv(ITEM_BREC.HIN_GAI, vbUnicode) <> StrConv(GENSANREC.HIN_GAI, vbUnicode) Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Exit Function
        End Select
    
        
        List2.AddItem StrConv(GENSANREC.Ins_DateTime, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)
        
        com = BtOpGetNext
    Loop
    
        
    Combo1(pcmbGENSAN).Enabled = False
    Text1(ptxGENSANKOKU).Enabled = False
        
    If List2.ListCount > 0 Then
        Combo1(pcmbGENSAN).Enabled = True
        For i = 0 To List2.ListCount - 1
        
            Combo1(pcmbGENSAN).AddItem Right(List2.List(i), 20)
        
        Next i
    
        Combo1(pcmbGENSAN).ListIndex = 0
    Else
        Text1(ptxGENSANKOKU).Enabled = True
    End If
    
    GENSANKOKU_SET_Proc = False


End Function


