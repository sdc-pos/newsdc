VERSION 5.00
Begin VB.Form F1010511 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�i�ڃ}�X�^�����e�i���X�i�폜�@�\�t���j"
   ClientHeight    =   6915
   ClientLeft      =   1920
   ClientTop       =   2295
   ClientWidth     =   14055
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   14055
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   33
      Left            =   11160
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   89
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   32
      Left            =   12480
      MaxLength       =   4
      TabIndex        =   87
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   31
      Left            =   8880
      MaxLength       =   1
      TabIndex        =   84
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   30
      Left            =   5760
      MaxLength       =   13
      TabIndex        =   82
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   29
      Left            =   2280
      MaxLength       =   13
      TabIndex        =   80
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   21
      Left            =   10200
      MaxLength       =   8
      TabIndex        =   22
      Top             =   1560
      Width           =   1095
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   1800
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   20
      Left            =   9720
      MaxLength       =   2
      TabIndex        =   21
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   25
      Left            =   4200
      MaxLength       =   2
      TabIndex        =   26
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   24
      Left            =   3360
      MaxLength       =   2
      TabIndex        =   25
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   23
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   24
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   22
      Left            =   12840
      MaxLength       =   1
      TabIndex        =   23
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   19
      Left            =   8520
      MaxLength       =   2
      TabIndex        =   20
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   18
      Left            =   7680
      MaxLength       =   2
      TabIndex        =   19
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   17
      Left            =   6600
      MaxLength       =   4
      TabIndex        =   18
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   16
      Left            =   4200
      MaxLength       =   2
      TabIndex        =   17
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   15
      Left            =   3360
      MaxLength       =   2
      TabIndex        =   16
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   14
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   15
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   13
      Left            =   13080
      MaxLength       =   2
      TabIndex        =   14
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   12
      Left            =   12360
      MaxLength       =   2
      TabIndex        =   13
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   11
      Left            =   11640
      MaxLength       =   2
      TabIndex        =   12
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   10
      Left            =   10920
      MaxLength       =   2
      TabIndex        =   11
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   9
      Left            =   8280
      MaxLength       =   2
      TabIndex        =   10
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   8
      Left            =   7440
      MaxLength       =   2
      TabIndex        =   9
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   7
      Left            =   6360
      MaxLength       =   4
      TabIndex        =   8
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   5160
      MaxLength       =   13
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   4440
      MaxLength       =   2
      TabIndex        =   7
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   3720
      MaxLength       =   2
      TabIndex        =   6
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   5
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   6  '���p����
      Index           =   2
      Left            =   8040
      MaxLength       =   25
      TabIndex        =   3
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   1800
      MaxLength       =   13
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command 
      Caption         =   "�I  ��"
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
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�f�[�^"
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
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Index           =   7
      Left            =   6480
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Index           =   6
      Left            =   5640
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Index           =   5
      Left            =   4800
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Index           =   4
      Left            =   3960
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "��  ��"
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
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�X  �V"
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
      Index           =   0
      Left            =   120
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   2700
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   28
      Top             =   3000
      Width           =   8655
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   26
      Left            =   6720
      MaxLength       =   4
      TabIndex        =   76
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   27
      Left            =   7800
      MaxLength       =   2
      TabIndex        =   77
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   28
      Left            =   8640
      MaxLength       =   2
      TabIndex        =   78
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����Ϗo�א�"
      Height          =   255
      Index           =   24
      Left            =   9360
      TabIndex        =   88
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "������"
      Height          =   255
      Index           =   42
      Left            =   11280
      TabIndex        =   86
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "(0:�v�@1:�s�v)"
      Height          =   255
      Index           =   41
      Left            =   9240
      TabIndex        =   85
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���i���L��"
      Height          =   255
      Index           =   40
      Left            =   7560
      TabIndex        =   83
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�Ǒւ��R�[�h"
      Height          =   255
      Index           =   39
      Left            =   4080
      TabIndex        =   81
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�����R�[�h"
      Height          =   255
      Index           =   38
      Left            =   600
      TabIndex        =   79
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�ŐV�ƍ����t"
      Height          =   255
      Index           =   37
      Left            =   4920
      TabIndex        =   75
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�^"
      Height          =   255
      Index           =   36
      Left            =   7440
      TabIndex        =   74
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�^"
      Height          =   255
      Index           =   35
      Left            =   8280
      TabIndex        =   73
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�j"
      Height          =   255
      Index           =   34
      Left            =   9120
      TabIndex        =   72
      Top             =   2160
      Width           =   135
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
      Left            =   120
      TabIndex        =   71
      Top             =   6480
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����O"
      Height          =   255
      Index           =   33
      Left            =   840
      TabIndex        =   70
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���l"
      Height          =   255
      Index           =   32
      Left            =   9120
      TabIndex        =   69
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�j"
      Height          =   255
      Index           =   31
      Left            =   4680
      TabIndex        =   68
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�^"
      Height          =   255
      Index           =   30
      Left            =   3840
      TabIndex        =   67
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�^"
      Height          =   255
      Index           =   29
      Left            =   3000
      TabIndex        =   66
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�ŏI���ד��t"
      Height          =   255
      Index           =   28
      Left            =   480
      TabIndex        =   65
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�T���v����"
      Height          =   255
      Index           =   27
      Left            =   11400
      TabIndex        =   64
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�j"
      Height          =   255
      Index           =   23
      Left            =   9000
      TabIndex        =   63
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�^"
      Height          =   255
      Index           =   22
      Left            =   8160
      TabIndex        =   62
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�^"
      Height          =   255
      Index           =   21
      Left            =   7320
      TabIndex        =   61
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�ŏI�o�ɓ��t"
      Height          =   255
      Index           =   20
      Left            =   4800
      TabIndex        =   60
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�j"
      Height          =   255
      Index           =   19
      Left            =   4680
      TabIndex        =   59
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�^"
      Height          =   255
      Index           =   18
      Left            =   3840
      TabIndex        =   58
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�^"
      Height          =   255
      Index           =   17
      Left            =   3000
      TabIndex        =   57
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�ŏI���ɓ��t"
      Height          =   255
      Index           =   16
      Left            =   480
      TabIndex        =   56
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�j"
      Height          =   255
      Index           =   15
      Left            =   13560
      TabIndex        =   55
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   13
      Left            =   12840
      TabIndex        =   54
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   12
      Left            =   12120
      TabIndex        =   53
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   11
      Left            =   11400
      TabIndex        =   52
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�O����ɒI"
      Height          =   255
      Index           =   10
      Left            =   9000
      TabIndex        =   51
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�j"
      Height          =   255
      Index           =   9
      Left            =   8760
      TabIndex        =   50
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�^"
      Height          =   255
      Index           =   8
      Left            =   7920
      TabIndex        =   49
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�^"
      Height          =   255
      Index           =   7
      Left            =   7080
      TabIndex        =   48
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�ԁi�����j"
      Height          =   255
      Index           =   14
      Left            =   3600
      TabIndex        =   47
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�ݒ���t"
      Height          =   255
      Index           =   6
      Left            =   5040
      TabIndex        =   46
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   5
      Left            =   4200
      TabIndex        =   45
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   44
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   43
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�W�����ɒI"
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   42
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i ��"
      Height          =   255
      Index           =   1
      Left            =   7200
      TabIndex        =   41
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�ԁi�O���j"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   40
      Top             =   720
      Width           =   1575
   End
   Begin VB.Menu MainMenu 
      Caption         =   "���ƕ�"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1010511"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim LIST_MAX    As Integer              '���X�g�{�b�N�X�ő�\������

Dim Text_Max    As Integer              '��ʍ��ڕʍő���ޯ��
Dim Combo_Max   As Integer
Dim Command_Max As Integer
''Dim JIGYOBU_BEF As String * 1         '���݌����ƕ�
Dim ITEM_CSV    As String



Private Function List_Disp()
Dim sts As Integer
Dim com As Integer
Dim i As Integer
Dim Sv_Naigai As String * 1
Dim Edit As String
    
    List_Disp = False
    
    List1.Clear
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    If Combo(0).Text = NAIGAI1$ Then
        Sv_Naigai = NAIGAI_NAI$
    Else
        Sv_Naigai = NAIGAI_GAI$
    End If
    Call UniCode_Conv(K0_ITEM.NAIGAI, Sv_Naigai)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, RTrim(Text(0).Text))
    
    com = BtOpGetGreaterEqual
    For i = 0 To LIST_MAX - 1
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(ITEMREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(ITEMREC.NAIGAI, vbUnicode) <> Sv_Naigai Then
                    Exit For
                End If
            Case BtErrEOF
                Exit Function
            Case Else
                Call File_Error(sts, com, "�i�ڃ}�X�^")
                List_Disp = True
                Exit Function
        End Select
        
        Edit = StrConv(ITEMREC.HIN_GAI, vbUnicode) & " " & StrConv(ITEMREC.HIN_NAI, vbUnicode) & " " & StrConv(ITEMREC.HIN_NAME, vbUnicode) & " "
        Edit = Edit & StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" + StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & StrConv(ITEMREC.ST_DAN, vbUnicode)
        List1.AddItem Edit
        
        com = BtOpGetNext
    Next i
    
End Function
                                    '��ʏ�����Ԃ�ݒ肷��
Private Sub Clear_Field(Mode As Integer)
Dim i As Integer

    If (Mode = 0) Then
        Text(0).Text = ""
    End If
    
    For i = 1 To 32
        Text(i).Text = ""
    Next i
End Sub

'                                       ���͍��ڂ̃G���[�`�F�b�N
Private Function Del_Chk() As Integer
            
Dim RetBuf As String
Dim i As Integer
Dim sts As Integer


    Del_Chk = False
    
    
    Call UniCode_Conv(K1_ZAIKO.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K1_ZAIKO.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K1_ZAIKO.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
    Call UniCode_Conv(K1_ZAIKO.NYUKA_DT, "")
    Call UniCode_Conv(K1_ZAIKO.Soko_No, "")
    Call UniCode_Conv(K1_ZAIKO.Retu, "")
    Call UniCode_Conv(K1_ZAIKO.Ren, "")
    Call UniCode_Conv(K1_ZAIKO.Dan, "")

    sts = BTRV(BtOpGetGreaterEqual, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
    Select Case sts
        Case BtNoErr
            If StrConv(ZAIKOREC.JGYOBU, vbUnicode) = StrConv(ZAIKOREC.JGYOBU, vbUnicode) And _
                StrConv(ZAIKOREC.NAIGAI, vbUnicode) = StrConv(ITEMREC.NAIGAI, vbUnicode) And _
                StrConv(ZAIKOREC.HIN_GAI, vbUnicode) = StrConv(ITEMREC.HIN_GAI, vbUnicode) Then
                Beep
                MsgBox "�L���݌Ɏc�L��I�I�폜�ł��܂���B"
                Text(0).SelStart = 0
                Text(0).SelLength = Len(Text(0).Text)
                Text(0).SetFocus
                Del_Chk = True
                Exit Function
            End If
        Case BtErrEOF
        Case Else
            Call File_Error(sts, BtOpGetGreaterEqual, "�݌Ƀf�[�^")
            Del_Chk = SYS_ERR
    End Select

End Function
'                                       ���͍��ڂ̃G���[�`�F�b�N
Private Function Err_Chk() As Integer
            
Dim RetBuf As String
Dim i As Integer
Dim sts As Integer


    Err_Chk = False
    
    If Len(RTrim(Text(0).Text)) = 0 Then
        Beep
        MsgBox "���͂������ڂ̓G���[�ł��B"
        Text(0).SetFocus
        Err_Chk = True
        Exit Function
    End If
    
                                            '�W�����ɒI�`�F�b�N
    If Len(RTrim(Text(3).Text)) = 0 Then
        Text(4).Text = ""
        Text(5).Text = ""
        Text(6).Text = ""
    Else
        For i = 4 To 6
            If Not IsNumeric(Text(i).Text) Then
                Beep
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text(i).SetFocus
                Err_Chk = True
                Exit Function
            Else
                Text(i).Text = Format(CInt(Text(i).Text), "00")
            End If
        Next i
        Call UniCode_Conv(K0_SOKO.Soko_No, Text(3).Text)
        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(SOKOREC.KONS_KBN, vbUnicode) = KONS_KBN_NG$ Then
                    If StrConv(SOKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                        Beep
                        MsgBox "���͂������ڂ̓G���[�ł��B�i���ڃG���[�j"
                        Text(3).SetFocus
                        Err_Chk = True
                        Exit Function
                    End If
                End If
            Case BtErrKeyNotFound
                Beep
                MsgBox "���͂������ڂ̓G���[�ł��B�i���o�^�G���[�j"
                Text(3).SetFocus
                Err_Chk = True
                Exit Function
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^")
                Err_Chk = SYS_ERR
                Exit Function
        End Select
                                                '�I�}�X�^�ǂݍ���
        Call UniCode_Conv(K0_TANA.Soko_No, Text(3).Text)
        Call UniCode_Conv(K0_TANA.Retu, Text(4).Text)
        Call UniCode_Conv(K0_TANA.Ren, Text(5).Text)
        Call UniCode_Conv(K0_TANA.Dan, Text(6).Text)
        sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Beep
                MsgBox "���͂������ڂ̓G���[�ł��B�i���o�^�G���[�j"
                Text(3).SetFocus
                Err_Chk = True
                Exit Function
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�I�}�X�^")
                Err_Chk = SYS_ERR
                Exit Function
        End Select
    End If
                                        '�T���v����
    If Len(RTrim(Text(22).Text)) = 0 Then
        Text(22).Text = "1"
    End If
    If Not IsNumeric(Text(22).Text) Then
        Beep
        MsgBox "���͂������ڂ̓G���[�ł��B"
        Text(22).SetFocus
        Err_Chk = True
        Exit Function
    Else
        Text(22).Text = Format(CLng(Text(22).Text), "0")
    End If
                                        'JAN�R�[�h
    If Len(Trim(Text(29).Text)) <> 0 Then
        If Len(RTrim(Text(29).Text)) <> Text(29).MaxLength Then
            Beep
            MsgBox "���͂������ڂ̓G���[�ł��B"
            Text(22).SetFocus
            Err_Chk = True
            Exit Function
        End If

        If Not IsNumeric(Text(29).Text) Then
            Beep
            MsgBox "���͂������ڂ̓G���[�ł��B"
            Text(22).SetFocus
            Err_Chk = True
            Exit Function
        End If


        Call UniCode_Conv(K4_ITEM.JGYOBU, Last_JGYOBU)
        If Combo(0).Text = NAIGAI1$ Then
            Call UniCode_Conv(K4_ITEM.NAIGAI, NAIGAI_NAI$)
        Else
            Call UniCode_Conv(K4_ITEM.NAIGAI, NAIGAI_GAI$)
        End If
        Call UniCode_Conv(K4_ITEM.JAN_CODE, Text(29).Text)
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K4_ITEM, Len(K4_ITEM), 4)
        Select Case sts
            Case BtNoErr
                If Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) <> Trim(Text(0).Text) Then
                    Beep
                    MsgBox "���͂������ڂ̓G���[�ł��B(�o�^�ς�)"
                    Text(29).SetFocus
                    Err_Chk = True
                    Exit Function
                End If
            Case BtErrKeyNotFound
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                Err_Chk = SYS_ERR
                Exit Function
        End Select
    End If
                                        '�Ǒւ��R�[�h
    If Len(Trim(Text(30).Text)) <> 0 Then
        Call UniCode_Conv(K5_ITEM.JGYOBU, Last_JGYOBU)
        If Combo(0).Text = NAIGAI1$ Then
            Call UniCode_Conv(K5_ITEM.NAIGAI, NAIGAI_NAI$)
        Else
            Call UniCode_Conv(K5_ITEM.NAIGAI, NAIGAI_GAI$)
        End If
        Call UniCode_Conv(K5_ITEM.HIN_CHANGE, Text(30).Text)
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K5_ITEM, Len(K5_ITEM), 5)
        Select Case sts
            Case BtNoErr
                If Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) <> Trim(Text(0).Text) Then
                    Beep
                    MsgBox "���͂������ڂ̓G���[�ł��B(�o�^�ς�)"
                    Text(30).SetFocus
                    Err_Chk = True
                    Exit Function
                End If
            Case BtErrKeyNotFound
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                Err_Chk = SYS_ERR
                Exit Function
        End Select
    End If
                                        '���i���L��
    If Text(31).Text = "" Then
        Text(31).Text = "0"
    Else
        If Text(31).Text <> "0" And Text(31).Text <> "1" Then
            Beep
            MsgBox "���͂������ڂ̓G���[�ł��B"
            Text(31).SetFocus
            Err_Chk = True
            Exit Function
        End If
    End If
                                        '������
    If Text(32).Text <> "" Then
        Call UniCode_Conv(K0_PACKING.PACKING_NO, Text(32).Text)
        sts = BTRV(BtOpGetEqual, PACKING_POS, PACKINGREC, Len(PACKINGREC), K0_PACKING, Len(K0_PACKING), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Beep
                MsgBox "���͂������ڂ̓G���[�ł��B�i���o�^�G���[�j"
                Text(32).SetFocus
                Err_Chk = True
                Exit Function
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�����}�X�^")
                Err_Chk = SYS_ERR
                Exit Function
        End Select
    End If

End Function

Private Sub Item_Dsp()
Dim sts         As Integer
Dim Work_Date   As String * 8
Dim RetBuf      As String
            
    Text(0).Text = RTrim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
    Text(1).Text = RTrim(StrConv(ITEMREC.HIN_NAI, vbUnicode))
    Text(2).Text = RTrim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
    Text(3).Text = StrConv(ITEMREC.ST_SOKO, vbUnicode)
    Text(4).Text = StrConv(ITEMREC.ST_RETU, vbUnicode)
    Text(5).Text = StrConv(ITEMREC.ST_REN, vbUnicode)
    Text(6).Text = StrConv(ITEMREC.ST_DAN, vbUnicode)
    Work_Date = StrConv(ITEMREC.ST_SET_DT, vbUnicode)
    Text(7).Text = Mid(Work_Date, 1, 4)
    Text(8).Text = Mid(Work_Date, 5, 2)
    Text(9).Text = Mid(Work_Date, 7, 2)
    Text(10).Text = StrConv(ITEMREC.BEF_SOKO, vbUnicode)
    Text(11).Text = StrConv(ITEMREC.BEF_RETU, vbUnicode)
    Text(12).Text = StrConv(ITEMREC.BEF_REN, vbUnicode)
    Text(13).Text = StrConv(ITEMREC.BEF_DAN, vbUnicode)
    Work_Date = StrConv(ITEMREC.LAST_NYU_DT, vbUnicode)
    Text(14).Text = Mid(Work_Date, 1, 4)
    Text(15).Text = Mid(Work_Date, 5, 2)
    Text(16).Text = Mid(Work_Date, 7, 2)
    Work_Date = StrConv(ITEMREC.LAST_SYU_DT, vbUnicode)
    Text(17).Text = Mid(Work_Date, 1, 4)
    Text(18).Text = Mid(Work_Date, 5, 2)
    Text(19).Text = Mid(Work_Date, 7, 2)
    Text(20).Text = RTrim(StrConv(ITEMREC.BIKOU_SOKO, vbUnicode))
    Text(21).Text = RTrim(StrConv(ITEMREC.BIKOU_TANA, vbUnicode))
    Text(22).Text = StrConv(ITEMREC.SAMPLE_QTY, vbUnicode)
    Work_Date = StrConv(ITEMREC.LAST_INP_DT, vbUnicode)
    Text(23).Text = Mid(Work_Date, 1, 4)
    Text(24).Text = Mid(Work_Date, 5, 2)
    Text(25).Text = Mid(Work_Date, 7, 2)
    Work_Date = StrConv(ITEMREC.LAST_CHK_DT, vbUnicode)
    Text(26).Text = Mid(Work_Date, 1, 4)
    Text(27).Text = Mid(Work_Date, 5, 2)
    Text(28).Text = Mid(Work_Date, 7, 2)

    Text(29).Text = StrConv(ITEMREC.JAN_CODE, vbUnicode)
    Text(30).Text = StrConv(ITEMREC.HIN_CHANGE, vbUnicode)
    Text(31).Text = StrConv(ITEMREC.GOODS_KBN, vbUnicode)
    Text(32).Text = StrConv(ITEMREC.PACKING_NO, vbUnicode)
    If IsNumeric(CLng(StrConv(ITEMREC.AVE_SYUKA, vbUnicode))) Then
        Text(33).Text = Format(CDbl(StrConv(ITEMREC.AVE_SYUKA, vbUnicode)), "#0.0")
    Else
        Text(33).Text = ""

    End If
End Sub

                                            '�i�ڃ}�X�^�̒ǉ��^����
Private Function Update_Proc() As Integer
Dim sts As Integer
Dim ans As Integer
Dim com As Integer
Dim Sv_Naigai As String * 1
Dim RetBuf As String
Dim Edit As String
Dim i As Integer

    Update_Proc = False

    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    If Combo(0).Text = NAIGAI1 Then
        Sv_Naigai = NAIGAI_NAI$
    Else
        Sv_Naigai = NAIGAI_GAI$
    End If
    Call UniCode_Conv(K0_ITEM.NAIGAI, Sv_Naigai)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text(0).Text)
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Call Clear_Field(0)
                    Text(0).SetFocus
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                Update_Proc = True
        End Select
    Loop
                                            '���R�[�h���e�ҏW
    Call UniCode_Conv(ITEMREC.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(ITEMREC.NAIGAI, Sv_Naigai)
    Call UniCode_Conv(ITEMREC.HIN_GAI, Text(0).Text)
    Call UniCode_Conv(ITEMREC.HIN_NAI, Text(1).Text)
    Call UniCode_Conv(ITEMREC.HIN_NAME, Text(2).Text)
    If Len(RTrim(Text(3).Text)) = 0 Then
        Call UniCode_Conv(ITEMREC.ST_SET_DT, "")
    Else
        If com = BtOpUpdate Then
            If (StrConv(ITEMREC.ST_SOKO, vbUnicode) <> Text(3).Text) Or _
                (StrConv(ITEMREC.ST_RETU, vbUnicode) <> Text(4).Text) Or _
                (StrConv(ITEMREC.ST_REN, vbUnicode) <> Text(5).Text) Or _
                (StrConv(ITEMREC.ST_DAN, vbUnicode) <> Text(6).Text) Then
                Call UniCode_Conv(ITEMREC.ST_SET_DT, (Left(Format(Date, "yyyymmdd"), 4) + Mid(Format(Date, "yyyymmdd"), 5, 2) + Mid(Format(Date, "yyyymmdd"), 7, 2)))
            End If
        Else
            Call UniCode_Conv(ITEMREC.ST_SET_DT, (Left(Format(Date, "yyyymmdd"), 4) + Mid(Format(Date, "yyyymmdd"), 5, 2) + Mid(Format(Date, "yyyymmdd"), 7, 2)))
        End If
    End If
    Call UniCode_Conv(ITEMREC.ST_SOKO, Text(3).Text)
    Call UniCode_Conv(ITEMREC.ST_RETU, Text(4).Text)
    Call UniCode_Conv(ITEMREC.ST_REN, Text(5).Text)
    Call UniCode_Conv(ITEMREC.ST_DAN, Text(6).Text)
    If com = BtOpInsert Then
        Call UniCode_Conv(ITEMREC.BEF_SOKO, "")
        Call UniCode_Conv(ITEMREC.BEF_RETU, "")
        Call UniCode_Conv(ITEMREC.BEF_REN, "")
        Call UniCode_Conv(ITEMREC.BEF_DAN, "")
    
        Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")
        Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")
        Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")
        
        Call UniCode_Conv(ITEMREC.LOCK_F, LOCK_OFF)
        Call UniCode_Conv(ITEMREC.WEL_ID, "")
        Call UniCode_Conv(ITEMREC.PRG_ID, "")
        Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")
        Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "00000000")
        Call UniCode_Conv(ITEMREC.FILLER, "")
    
        Call UniCode_Conv(ITEMREC.SIZAI_CD, "")
        Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")
        Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")
    
    End If
    
    Call UniCode_Conv(ITEMREC.HIN_NAI, Text(1).Text)
    Call UniCode_Conv(ITEMREC.BIKOU_SOKO, Text(20).Text)
    Call UniCode_Conv(ITEMREC.BIKOU_TANA, Text(21).Text)
    Call UniCode_Conv(ITEMREC.SAMPLE_QTY, Format(CInt(Text(22).Text), "0"))
    Call UniCode_Conv(ITEMREC.JAN_CODE, Text(29).Text)
    Call UniCode_Conv(ITEMREC.HIN_CHANGE, Text(30).Text)
    Call UniCode_Conv(ITEMREC.GOODS_KBN, Text(31).Text)
    Call UniCode_Conv(ITEMREC.PACKING_NO, Text(32).Text)
            
            
    Do
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Call Clear_Field(0)
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, com, "�i�ڃ}�X�^")
                Beep
                MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly + vbCritical
                Update_Proc = True
        End Select
    Loop
                                        '���X�g�{�b�N�X�ǉ�
    If com = BtOpUpdate Then
        For i = 0 To List1.ListCount - 1
            If RTrim(Left$(List1.List(i), 13)) = RTrim(Text(0).Text) Then
                List1.RemoveItem i
            End If
        Next i
    End If
        
    Edit = StrConv(ITEMREC.HIN_GAI, vbUnicode) + " " + StrConv(ITEMREC.HIN_NAI, vbUnicode) + " " + StrConv(ITEMREC.HIN_NAME, vbUnicode) + " "
    Edit = Edit + StrConv(ITEMREC.ST_SOKO, vbUnicode) + "-" + StrConv(ITEMREC.ST_RETU, vbUnicode) + "-" + StrConv(ITEMREC.ST_REN, vbUnicode) + "-" + StrConv(ITEMREC.ST_DAN, vbUnicode)
    List1.AddItem Edit
                                        '��ʃN���A�[
    Call Clear_Field(0)
'

End Function
                                            '�i�ڃ}�X�^�̍폜
Private Function Delete_Proc() As Integer
Dim sts As Integer
Dim ans As Integer
Dim com As Integer
Dim Sv_Naigai As String * 1
Dim RetBuf As String
Dim Edit As String
Dim i As Integer
    
    Delete_Proc = False

    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    If Combo(0).Text = NAIGAI1 Then
        Sv_Naigai = NAIGAI_NAI$
    Else
        Sv_Naigai = NAIGAI_GAI$
    End If
    Call UniCode_Conv(K0_ITEM.NAIGAI, Sv_Naigai)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text(0).Text)
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
'                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
'                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Call Clear_Field(0)
                    Text(0).SetFocus
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                Delete_Proc = True
                Exit Function
        End Select
    Loop
            
    If sts = BtNoErr Then
            
        Do
            sts = BTRV(BtOpDelete, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        sts = BTRV(BtOpUnlock, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Call Clear_Field(0)
                        Text(0).SetFocus
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpDelete, "�i�ڃ}�X�^")
                    Beep
                    MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly + vbCritical
                    Delete_Proc = True
                    Exit Function
            End Select
        Loop
    End If
                                        '���X�g�{�b�N�X�폜
'    For i = 0 To List1.ListCount - 1
'        If RTrim(Left$(List1.List(i), 13)) = RTrim(Text(0).Text) Then
'            List1.RemoveItem i
'        End If
'    Next i
                                        '��ʃN���A�[
    Call Clear_Field(0)
'

End Function

Private Sub Combo_DblClick(Index As Integer)
    
    Call Clear_Field(0)
    List1.Clear

End Sub


Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Select Case Index
                Case 0
                    Call Clear_Field(0)
                    List1.Clear
                    Text(0).SetFocus
            End Select
    End Select

End Sub


Private Sub Combo_LostFocus(Index As Integer)
'        Call Clear_Field(0)
'        List1.Clear

End Sub

Private Sub Command_Click(Index As Integer)

Dim yn  As Integer
Dim i   As Integer
Dim sts As Integer
    
    Select Case Index
        Case 0                              '�f�[�^�X�V
                                            
            
            Beep
            yn = MsgBox("�X�V���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
                Text(0).SetFocus
            End If
        Case 3                              '�폜
            sts = Del_Chk()
                
            Select Case sts
                Case False
                Case True
                    Exit Sub
                Case SYS_ERR
                    Unload Me
            End Select
            
            Beep
            yn = MsgBox("�폜���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                If Delete_Proc() Then
                    Unload Me
                End If
                Text(0).SetFocus
            End If
        Case 8                              '�f�[�^�o��
                        
            If CSV_OUTPUT_Proc() Then
                Unload Me
            End If
        Case 11                             '�I��
            Unload Me
        Case Else
            Beep
    End Select
    
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
            Command(KeyCode - vbKeyF1).Value = True
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
Dim i       As Integer
Dim c       As String * 128
Dim sts     As Integer


    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If
    
    Text_Max = 1                '����w����ʍ��ڕʍő���ޯ��
    Command_Max = 2
    
    Show
    
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
                                
''                                '���ݎ��ƕ�����荞��
''    If GetIni("SENTAKU", "JIGYOBU_BEF", "SYS", c) Then
''        Beep
''        MsgBox "���ݎ��ƕ����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
''        End
''    End If
''    JIGYOBU_BEF = RTrim(c)
    
                                '�b�r�u�t�@�C������荞��
    If GetIni("FILE", "ITEM_CSV", "SYS", c) Then
        Beep
        MsgBox "�i�ڃ}�X�^�f�[�^�o�͗p�t�@�C��[ITEM_CSV]�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    ITEM_CSV = Trim(c)
                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
        
    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).Code = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).Code = Last_JGYOBU Then
            F1010511.Caption = "�i�ڃ}�X�^�����e�i���X�i�폜�@�\�t���j�i" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
                                '���X�g�{�b�N�X�ő�\�������\��
    If GetIni("F101051", "LISTMAX", "SYS", c) Then
        Beep
        MsgBox "�ő�\�������̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LIST_MAX = CInt(RTrim(c))
                                '�����O��荞��
    Combo(0).AddItem NAIGAI1$
    Combo(0).AddItem NAIGAI2$
    Combo(0).Text = NAIGAI1$
                                
                                '�q�Ƀ}�X�^�n�o�d�m
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�I�}�X�^�n�o�d�m
    If TANA_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�����}�X�^�n�o�d�m
    If PACKING_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ƀf�[�^�n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '��ʏ����ݒ�
    Call Clear_Field(0)
    
    Combo(0).SetFocus
    
    End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '�q�Ƀ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�q�Ƀ}�X�^")
        End If
    End If
                                            '�I�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�I�}�X�^")
        End If
    End If
                                            '�����}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, PACKING_POS, PACKINGREC, Len(PACKINGREC), K0_PACKING, Len(K0_PACKING), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�����}�X�^")
        End If
    End If
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        End If
    End If
    
    sts = BTRV(BtOpReset, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1010511 = Nothing

    End
End Sub

Private Sub List1_DblClick()
Dim sts As Integer
                                        '���X�g�{�b�N�X��荀�ڕ\��
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Mid$(List1.List(List1.ListIndex), 1, 13))
                                                '�O���i�Ԃœǂݍ���
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            Call Item_Dsp
            Text(1).SetFocus
        Case BtErrKeyNotFound           '����͖����͂�
            MsgBox "�}�X�^���e���ύX����Ă��܂��B�ŐV�����ĕ\�����܂��B"
            If List_Disp() Then
                Unload Me
            End If
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Beep
            MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
            Unload Me
    End Select

End Sub


Private Sub List1_GotFocus()
    
    If List1.ListCount > 0 Then
        If List1.ListIndex <= 0 Then
            List1.ListIndex = 0
        End If
    End If
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sts As Integer
                                        
    If List1.ListCount = 0 Then
        Exit Sub
    End If
                                        '���X�g�{�b�N�X��荀�ڕ\��
    Select Case KeyCode
        Case vbKeyReturn
            Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Mid$(List1.List(List1.ListIndex), 1, 13))
                                                '�O���i�Ԃœǂݍ���
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Call Item_Dsp
                    Text(1).SetFocus
                Case BtErrKeyNotFound           '����͖����͂�
                    MsgBox "�}�X�^���e���ύX����Ă��܂��B�ŐV�����ĕ\�����܂��B"
                    If List_Disp() Then
                        Unload Me
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                    Unload Me
            End Select
    End Select
End Sub

Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    
                                    '���j���[���I���v��
'    If JGYOBU_T(Index).CODE = " " Then
'        Unload Me
'    End If
    
    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).Code = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '���ƕ��؂�ւ�
    F1010511.Caption = "�i�ڃ}�X�^�����e�i���X�i�폜�@�\�t���j�i" + RTrim(JGYOBU_T(Index).NAME) + "�j"
    Last_JGYOBU = JGYOBU_T(Index).Code
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)
End Sub

Private Sub Text_GotFocus(Index As Integer)
    
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If

End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim RetBuf As String
Dim i As Integer
Dim sts As Integer

    Select Case KeyCode
        Case vbKeyReturn
            Select Case Index
                Case 0
                        If List_Disp() Then
                            Unload Me
                        End If
                                                '�O���i�Ԃœǂݍ���
                        Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
                        If Combo(0).Text = NAIGAI1$ Then
                            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI$)
                        Else
                            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_GAI$)
                        End If
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, RTrim(Text(0).Text))
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                                Call Item_Dsp
                            Case BtErrKeyNotFound
                                Call Clear_Field(1)
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                Unload Me
                        End Select
                Case 1
                Case 29             'Jan�R�[�h
                Case 30             '�Ǒւ��R�[�h
            End Select
            For i = Index + 1 To 32
                If Text(i).Enabled Then
                    Text(i).SetFocus
                    Exit For
                End If
            Next i
    End Select
End Sub



Private Function CSV_OUTPUT_Proc() As Integer

Dim FileNo          As Integer
Dim fileName        As String
Dim Ret             As Integer

Dim com             As Integer
Dim sts             As Integer

Dim c               As String * 128

Dim Soko_No         As String * 2

    Call Input_Lock
    
    FileNo = FreeFile
    fileName = ITEM_CSV
    
    Ret = InStr(1, Trim(fileName), ".") - 1
    fileName = Left(Trim(fileName), Ret) & "_" & Last_JGYOBU & Right(Trim(fileName), Len(Trim(fileName)) - Ret)

    On Error GoTo Error_Proc

    Open (fileName) For Output As FileNo
    On Error GoTo 0
    
'    Write #FileNo, "���ƕ�", "���O", "�i�ԁi�O���j", "�i��", "�W���I�Ԑݒ��", "�W���I��", "�O��I��", "�ŏI���ɓ�", "�ŏI�o�ɓ�", "�i�ԁi�����j", "�z�X�g�q��", "�z�X�g�I��", "�T���v����", "�ŏI���ד�", "�ŏI�ƍ���", "�ƍ����݌ɐ�", "������l", "������萔", "�i�`�m", "�ǂݑւ�", "���i���L��", "������"
    Write #FileNo, "���O", "�i�ԁi�O���j", "�i��", "�W���I�Ԑݒ��", "�W���I��", "�O��I��", "�ŏI���ɓ�", "�ŏI�o�ɓ�", "�i�ԁi�����j", "�z�X�g�q��", "�z�X�g�I��", "�T���v����", "�ŏI���ד�", "�ŏI�ƍ���", "�ƍ����݌ɐ�", "������l", "������萔", "�i�`�m", "�ǂݑւ�", "���i���L��", "������"

    com = BtOpGetFirst
    Do
        DoEvents
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(ITEMREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�i�ڃ}�X�^")
                
                Call Input_UnLock
                Exit Function
        End Select
    
'        Write #FileNo, StrConv(ITEMREC.JGYOBU, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.NAIGAI, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.HIN_GAI, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.HIN_NAME, vbUnicode),
        
        If IsDate(StrConv(ITEMREC.ST_SET_DT, vbUnicode)) Then
            Write #FileNo, Format(StrConv(ITEMREC.ST_SET_DT, vbUnicode), "YYYY/MM/DD"),
        Else
            Write #FileNo, ,
        End If
        
        
        If GetIni("SOKO_NO", StrConv(ITEMREC.ST_SOKO, vbUnicode), "SYS", c) Then
            Soko_No = StrConv(ITEMREC.ST_SOKO, vbUnicode)
        Else
            Soko_No = Trim(c)
        End If
        Write #FileNo, Soko_No & "-" & StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & StrConv(ITEMREC.ST_DAN, vbUnicode),
        
        If GetIni("SOKO_NO", StrConv(ITEMREC.BEF_SOKO, vbUnicode), "SYS", c) Then
            Soko_No = StrConv(ITEMREC.BEF_SOKO, vbUnicode)
        Else
            Soko_No = Trim(c)
        End If
        Write #FileNo, Soko_No & "-" & StrConv(ITEMREC.BEF_RETU, vbUnicode) & "-" & StrConv(ITEMREC.BEF_REN, vbUnicode) & "-" & StrConv(ITEMREC.BEF_DAN, vbUnicode),
    
        Write #FileNo, StrConv(ITEMREC.LAST_NYU_DT, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.LAST_SYU_DT, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.HIN_NAI, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.BIKOU_SOKO, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.BIKOU_TANA, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.SAMPLE_QTY, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.LAST_INP_DT, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.LAST_CHK_DT, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.LAST_CHK_QTY, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.BIKOU, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.IRI_QTY, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.JAN_CODE, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.HIN_CHANGE, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.GOODS_KBN, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.PACKING_NO, vbUnicode)
    
    
        com = BtOpGetNext
    Loop

    Close #FileNo
    
    Call Input_UnLock
    
    
    Beep
    MsgBox "�u" & fileName & "�v�͐���ɏo�͂���܂����B"


    Exit Function
Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox fileName & "���g�p���ł��B"
        Call Input_UnLock
        CSV_OUTPUT_Proc = False
    Else
        MsgBox "Err.Number= " & Err.Number
        CSV_OUTPUT_Proc = True
    End If

    Call Input_UnLock

End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------
Dim i As Integer

    F1010511.MousePointer = vbHourglass

    Call Ctrl_Lock(F1010511)

End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(F1010511)

    F1010511.MousePointer = vbDefault

End Sub

