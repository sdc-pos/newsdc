VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form PI000151 
   Caption         =   "���i���w�}�[���s(�󒍋@�\�t��)"
   ClientHeight    =   9525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16545
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
   ScaleHeight     =   9525
   ScaleWidth      =   16545
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.ComboBox Combo2 
      Height          =   360
      Index           =   8
      Left            =   1200
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   115
      Top             =   8400
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   360
      Index           =   7
      Left            =   1200
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   106
      Top             =   8040
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   360
      Index           =   6
      Left            =   1200
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   97
      Top             =   7680
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   360
      Index           =   5
      Left            =   1200
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   88
      Top             =   7320
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   360
      Index           =   4
      Left            =   1200
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   79
      Top             =   6960
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   360
      Index           =   3
      Left            =   1200
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   70
      Top             =   6600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   1590
      MaxLength       =   10
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   12
      Left            =   14160
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   11
      Left            =   13200
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1080
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   1800
      TabIndex        =   176
      Top             =   2160
      Width           =   5925
      Begin VB.OptionButton Option1 
         Caption         =   "�č���"
         Height          =   375
         Index           =   3
         Left            =   4410
         TabIndex        =   180
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "���i����"
         Height          =   375
         Index           =   2
         Left            =   2880
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "���O"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   178
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "�X�|�b�g"
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   98
      Left            =   11520
      TabIndex        =   122
      Top             =   8400
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   97
      Left            =   10920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   121
      TabStop         =   0   'False
      Top             =   8400
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   96
      Left            =   9480
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   120
      TabStop         =   0   'False
      Top             =   8400
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   95
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   119
      TabStop         =   0   'False
      Top             =   8400
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   94
      Left            =   7560
      MaxLength       =   6
      TabIndex        =   118
      Top             =   8400
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   93
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   117
      TabStop         =   0   'False
      Top             =   8400
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   92
      Left            =   2520
      MaxLength       =   20
      TabIndex        =   116
      Top             =   8400
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   91
      Left            =   11520
      TabIndex        =   113
      Top             =   8040
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   90
      Left            =   10920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   112
      TabStop         =   0   'False
      Top             =   8040
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   89
      Left            =   9480
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   111
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   88
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   110
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   87
      Left            =   7560
      MaxLength       =   6
      TabIndex        =   109
      Top             =   8040
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   86
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   108
      TabStop         =   0   'False
      Top             =   8040
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   85
      Left            =   2520
      MaxLength       =   20
      TabIndex        =   107
      Top             =   8040
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   84
      Left            =   11520
      TabIndex        =   104
      Top             =   7680
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   83
      Left            =   10920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   103
      TabStop         =   0   'False
      Top             =   7680
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   82
      Left            =   9480
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   102
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   81
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   101
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   80
      Left            =   7560
      MaxLength       =   6
      TabIndex        =   100
      Top             =   7680
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   79
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   7680
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   78
      Left            =   2520
      MaxLength       =   20
      TabIndex        =   98
      Top             =   7680
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   77
      Left            =   11520
      TabIndex        =   95
      Top             =   7320
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   76
      Left            =   10920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   94
      TabStop         =   0   'False
      Top             =   7320
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   75
      Left            =   9480
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   7320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   74
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   73
      Left            =   7560
      MaxLength       =   6
      TabIndex        =   91
      Top             =   7320
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   72
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   7320
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   71
      Left            =   2520
      MaxLength       =   20
      TabIndex        =   89
      Top             =   7320
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   70
      Left            =   11520
      TabIndex        =   86
      Top             =   6960
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   69
      Left            =   10920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   6960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   68
      Left            =   9480
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   6960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   67
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   6960
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   66
      Left            =   7560
      MaxLength       =   6
      TabIndex        =   82
      Top             =   6960
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   65
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   6960
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   64
      Left            =   2520
      MaxLength       =   20
      TabIndex        =   80
      Top             =   6960
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   56
      Left            =   13320
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   55
      Left            =   12240
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   54
      Left            =   11400
      MaxLength       =   6
      TabIndex        =   66
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   53
      Left            =   9240
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   5040
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   52
      Left            =   7920
      MaxLength       =   20
      TabIndex        =   64
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   51
      Left            =   13320
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   50
      Left            =   12240
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   49
      Left            =   11400
      MaxLength       =   6
      TabIndex        =   61
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   48
      Left            =   9240
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   4680
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   47
      Left            =   7920
      MaxLength       =   20
      TabIndex        =   59
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   46
      Left            =   13320
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   45
      Left            =   12240
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   44
      Left            =   11400
      MaxLength       =   6
      TabIndex        =   56
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   43
      Left            =   9240
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   42
      Left            =   7920
      MaxLength       =   20
      TabIndex        =   54
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   41
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   40
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   39
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   51
      Top             =   5760
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   38
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   5760
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   37
      Left            =   600
      MaxLength       =   20
      TabIndex        =   49
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   36
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   35
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   34
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   46
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   33
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   5400
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   32
      Left            =   600
      MaxLength       =   20
      TabIndex        =   44
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   31
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   30
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   29
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   41
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   28
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   5040
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   27
      Left            =   600
      MaxLength       =   20
      TabIndex        =   39
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   26
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   25
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   24
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   36
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   16
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   19
      Top             =   1800
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1095
      Index           =   0
      Left            =   8040
      TabIndex        =   28
      Top             =   2640
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   1931
      _Version        =   393217
      TextRTF         =   $"PI000151.frx":0000
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   8
      Left            =   240
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   114
      Top             =   8400
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   7
      Left            =   240
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   105
      Top             =   8040
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   6
      Left            =   240
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   96
      Top             =   7680
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   5
      Left            =   240
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   87
      Top             =   7320
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   4
      Left            =   240
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   78
      Top             =   6960
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   3
      Left            =   240
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   69
      Top             =   6600
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "�o�͑Ώ�"
      Height          =   855
      Left            =   240
      TabIndex        =   158
      Top             =   2880
      Width           =   6735
      Begin VB.CheckBox Check1 
         Caption         =   "�@�탉�x��"
         Height          =   375
         Index           =   4
         Left            =   6240
         TabIndex        =   27
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "�O�����x��"
         Height          =   375
         Index           =   3
         Left            =   4320
         TabIndex        =   26
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "�p�[�c���x��"
         Height          =   375
         Index           =   2
         Left            =   2040
         TabIndex        =   25
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "�w�}�["
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   23
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   4680
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   22
      Left            =   600
      MaxLength       =   20
      TabIndex        =   34
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   21
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   20
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   19
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   31
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   18
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   17
      Left            =   600
      MaxLength       =   20
      TabIndex        =   29
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   6
      Left            =   8850
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   13
      Left            =   240
      MaxLength       =   5
      TabIndex        =   15
      Top             =   1800
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "���{�쐬"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   21
      Top             =   2400
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   2
      Left            =   9960
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   20
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   15
      Left            =   5640
      MaxLength       =   10
      TabIndex        =   18
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   14
      Left            =   4200
      MaxLength       =   10
      TabIndex        =   17
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   1
      Left            =   960
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   16
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   10
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   9
      Left            =   9960
      MaxLength       =   8
      TabIndex        =   10
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   8
      Left            =   5640
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   7
      Left            =   3120
      MaxLength       =   20
      TabIndex        =   8
      Top             =   1080
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   240
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   7
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   8130
      MaxLength       =   5
      TabIndex        =   5
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   4
      Left            =   5970
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   5250
      MaxLength       =   5
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   3450
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   240
      MaxLength       =   8
      TabIndex        =   0
      Top             =   360
      Width           =   1155
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
      Left            =   10440
      TabIndex        =   134
      Top             =   9000
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�� ��"
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
      Left            =   9600
      TabIndex        =   133
      Top             =   9000
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
      Left            =   8760
      TabIndex        =   132
      TabStop         =   0   'False
      Top             =   9000
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�� ��"
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
      Left            =   7920
      TabIndex        =   131
      Top             =   9000
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
      Index           =   7
      Left            =   6600
      TabIndex        =   130
      TabStop         =   0   'False
      Top             =   9000
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
      Index           =   6
      Left            =   5760
      TabIndex        =   129
      TabStop         =   0   'False
      Top             =   9000
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�\�����i"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4920
      TabIndex        =   128
      Top             =   9000
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4080
      TabIndex        =   127
      TabStop         =   0   'False
      Top             =   9000
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "M�X�V"
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
      Left            =   2760
      TabIndex        =   126
      Top             =   9000
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
      Left            =   1920
      TabIndex        =   125
      TabStop         =   0   'False
      Top             =   9000
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
      Left            =   1080
      TabIndex        =   124
      TabStop         =   0   'False
      Top             =   9000
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
      Left            =   240
      TabIndex        =   123
      Top             =   9000
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   63
      Left            =   11520
      TabIndex        =   77
      Top             =   6600
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   62
      Left            =   10920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   6600
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   61
      Left            =   9480
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   6600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   60
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   59
      Left            =   7560
      MaxLength       =   6
      TabIndex        =   73
      Top             =   6600
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   58
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   6600
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   57
      Left            =   2520
      MaxLength       =   20
      TabIndex        =   71
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���ƕ�"
      Height          =   255
      Index           =   25
      Left            =   1200
      TabIndex        =   181
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "������"
      Height          =   255
      Index           =   24
      Left            =   1560
      TabIndex        =   179
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���i����"
      Height          =   255
      Index           =   23
      Left            =   14040
      TabIndex        =   14
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�����i"
      Height          =   255
      Index           =   17
      Left            =   13200
      TabIndex        =   177
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�I��"
      Height          =   255
      Index           =   16
      Left            =   9840
      TabIndex        =   175
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "����"
      Height          =   255
      Index           =   22
      Left            =   8880
      TabIndex        =   174
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�i��"
      Height          =   255
      Index           =   21
      Left            =   5040
      TabIndex        =   173
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���E�׽"
      Height          =   255
      Index           =   20
      Left            =   7320
      TabIndex        =   172
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���l"
      Height          =   255
      Index           =   19
      Left            =   12240
      TabIndex        =   171
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�݌�"
      Height          =   255
      Index           =   18
      Left            =   11040
      TabIndex        =   170
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "����"
      Height          =   255
      Index           =   15
      Left            =   7800
      TabIndex        =   169
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�i��"
      Height          =   255
      Index           =   14
      Left            =   2760
      TabIndex        =   168
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���"
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   167
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�O�����އ�"
      Enabled         =   0   'False
      Height          =   375
      Index           =   17
      Left            =   7560
      TabIndex        =   166
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�@"
      Enabled         =   0   'False
      Height          =   375
      Index           =   16
      Left            =   7560
      TabIndex        =   165
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�A"
      Enabled         =   0   'False
      Height          =   375
      Index           =   15
      Left            =   7560
      TabIndex        =   164
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�B"
      Enabled         =   0   'False
      Height          =   375
      Index           =   14
      Left            =   7560
      TabIndex        =   163
      Top             =   5040
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�i��"
      Enabled         =   0   'False
      Height          =   375
      Index           =   13
      Left            =   9240
      TabIndex        =   162
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "����"
      Enabled         =   0   'False
      Height          =   375
      Index           =   12
      Left            =   11400
      TabIndex        =   161
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "����"
      Enabled         =   0   'False
      Height          =   375
      Index           =   11
      Left            =   12240
      TabIndex        =   160
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�I��"
      Enabled         =   0   'False
      Height          =   375
      Index           =   10
      Left            =   13320
      TabIndex        =   159
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�I��"
      Enabled         =   0   'False
      Height          =   375
      Index           =   9
      Left            =   6000
      TabIndex        =   157
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "����"
      Enabled         =   0   'False
      Height          =   375
      Index           =   8
      Left            =   4920
      TabIndex        =   156
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "����"
      Enabled         =   0   'False
      Height          =   375
      Index           =   7
      Left            =   4080
      TabIndex        =   155
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�i��"
      Enabled         =   0   'False
      Height          =   375
      Index           =   6
      Left            =   1920
      TabIndex        =   154
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�D"
      Enabled         =   0   'False
      Height          =   375
      Index           =   5
      Left            =   240
      TabIndex        =   153
      Top             =   5760
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�C"
      Enabled         =   0   'False
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   152
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�B"
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   151
      Top             =   5040
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�A"
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   150
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�@"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   149
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�����އ�"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   148
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���l"
      Height          =   255
      Index           =   13
      Left            =   8040
      TabIndex        =   147
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���P/�S����"
      Height          =   255
      Index           =   11
      Left            =   9960
      TabIndex        =   146
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�t���׽"
      Height          =   255
      Index           =   10
      Left            =   5880
      TabIndex        =   145
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���i���׽"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   9
      Left            =   4320
      TabIndex        =   144
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "��z��"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   143
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�W���I��"
      Height          =   255
      Index           =   7
      Left            =   11520
      TabIndex        =   142
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "����"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   6
      Left            =   10560
      TabIndex        =   141
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�i��"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   3120
      TabIndex        =   140
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�d������"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   139
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���F"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   8130
      TabIndex        =   138
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�S����"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   5250
      TabIndex        =   137
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���s��"
      Height          =   255
      Index           =   1
      Left            =   3330
      TabIndex        =   136
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�w�}�[��"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   135
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "PI000151"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private MTS_CODE    As String * 8
Private SS_CODE     As String * 8
Private CYU_KBN     As String * 1
Private CYU_KBN_N   As String * 1

Private G_Kisyu_F   As Integer
Private L_URIKIN1   As Double
Private L_URIKIN2   As Double
Private L_URIKIN3   As Double

Private SAIKON_F    As String * 1           '�č���F    2007.11.09

Private TEHAI       As String
    
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.09

Private Type ZAIKO_FUSOKU_T
    JGYOBU          As String * 1
    NAIGAI          As String * 1
    HIN_GAI         As String * 20
    USE_QTY         As Double
    ZAIKO_QTY       As Double

    SAI_QTY         As Double


    IDO_SUMI        As String * 1
    HIKIATE_QTY     As Double


    IDO_SUMI_QTY    As Double               '2012.04.13

End Type


Private ZAIKO_FUSOKU() _
                    As ZAIKO_FUSOKU_T       '�݌ɕs���i��




'--------------------------------------------------- ���  ���ޑΉ��@2012.03.09
    
    
'�e�L�X�g�p�Y��
Private Const ptxSHIJI_NO% = 0              '�w�}�[��
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.18
'Private Const ptxORDER_DT% = 1              '�󒍓�
Private Const ptxORDER_NO% = 1              '������
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.18
Private Const ptxHAKKO_DT% = 2              '���s��
Private Const ptxTANTO_CODE% = 3            '�S���Һ���
Private Const ptxTANTO_NAME% = 4            '�S���Җ���
Private Const ptxSHONIN_CODE% = 5           '���F�Һ���
Private Const ptxSHONIN_NAME% = 6           '���F�Җ���
Private Const ptxHIN_GAI% = 7               '�i��
Private Const ptxHIN_NAME% = 8              '�i��
Private Const ptxSHIJI_QTY% = 9             '����
Private Const ptxST_LOCATION% = 10          '�W���I��
Private Const ptxMI_QTY% = 11               '�����i
Private Const ptxSUMI_QTY% = 12             '���i����
Private Const ptxUKEHARAI_CODE% = 13        '��z�溰��
Private Const ptxS_CLASS_CODE% = 14         '���i���׽
Private Const ptxF_CLASS_CODE% = 15         '�t���׽
Private Const ptxN_CLASS_CODE% = 16         '���E�׽

    
Private Const ptxK_HIN_GAI01% = 17          '�@�@�����އ�
Private Const ptxK_HIN_NAME01% = 18         '�@�@�����ޖ���
Private Const ptxK_QTY01% = 19              '�@�@����
Private Const ptxK_SHIJI_QTY01% = 20        '�@�@����
Private Const ptxK_ST_LOCATION01% = 21      '�@�@�I��

Private Const ptxK_HIN_GAI02% = 22          '�A�@�����އ�
Private Const ptxK_HIN_NAME02% = 23         '�A�@�����ޖ���
Private Const ptxK_QTY02% = 24              '�A�@����
Private Const ptxK_SHIJI_QTY02% = 25        '�A�@����
Private Const ptxK_ST_LOCATION02% = 26      '�A�@�I��
    
Private Const ptxK_HIN_GAI03% = 27          '�B�@�����އ�
Private Const ptxK_HIN_NAME03% = 28         '�B�@�����ޖ���
Private Const ptxK_QTY03% = 29              '�B�@����
Private Const ptxK_SHIJI_QTY03% = 30        '�B�@����
Private Const ptxK_ST_LOCATION03% = 31      '�B�@�I��
    
Private Const ptxK_HIN_GAI04% = 32          '�C�@�����އ�
Private Const ptxK_HIN_NAME04% = 33         '�C�@�����ޖ���
Private Const ptxK_QTY04% = 34              '�C�@����
Private Const ptxK_SHIJI_QTY04% = 35        '�C�@����
Private Const ptxK_ST_LOCATION04% = 36      '�C�@�I��
    
Private Const ptxK_HIN_GAI05% = 37          '�D�@�����އ�
Private Const ptxK_HIN_NAME05% = 38         '�D�@�����ޖ���
Private Const ptxK_QTY05% = 39              '�D�@����
Private Const ptxK_SHIJI_QTY05% = 40        '�D�@����
Private Const ptxK_ST_LOCATION05% = 41      '�D�@�I��
    
    
Private Const ptxG_HIN_GAI01% = 42          '�@�@�O�����އ�
Private Const ptxG_HIN_NAME01% = 43         '�@�@�O�����ޖ���
Private Const ptxG_QTY01% = 44              '�@�@����
Private Const ptxG_SHIJI_QTY01% = 45        '�@�@����
Private Const ptxG_ST_LOCATION01% = 46      '�@�@�I��
    
Private Const ptxG_HIN_GAI02% = 47          '�A�@�O�����އ�
Private Const ptxG_HIN_NAME02% = 48         '�A�@�O�����ޖ���
Private Const ptxG_QTY02% = 49              '�A�@����
Private Const ptxG_SHIJI_QTY02% = 50        '�A�@����
Private Const ptxG_ST_LOCATION02% = 51      '�A�@�I��
    
Private Const ptxG_HIN_GAI03% = 52          '�B�@�O�����އ�
Private Const ptxG_HIN_NAME03% = 53         '�B�@�O�����ޖ���
Private Const ptxG_QTY03% = 54              '�B�@����
Private Const ptxG_SHIJI_QTY03% = 55        '�B�@����
Private Const ptxG_ST_LOCATION03% = 56      '�B�@�I��
    
Private Const ptxD_HIN_GAI01% = 57          '�@�@�����^�\���i��
Private Const ptxD_HIN_NAME01% = 58         '�@�@�����^�\���i��
Private Const ptxD_QTY01% = 59              '�@�@����
Private Const ptxD_SHIJI_QTY01% = 60        '�@�@����
Private Const ptxD_ST_LOCATION01% = 61      '�@�@�I��
Private Const ptxD_ZAIKO_QTY01% = 62        '�@�@�݌ɐ�
Private Const ptxD_BIKOU01% = 63            '�@�@���l
    
Private Const ptxD_HIN_GAI02% = 64          '�A�@�����^�\���i��
Private Const ptxD_HIN_NAME02% = 65         '�A�@�����^�\���i��
Private Const ptxD_QTY02% = 66              '�A�@����
Private Const ptxD_SHIJI_QTY02% = 67        '�A�@����
Private Const ptxD_ST_LOCATION02% = 68      '�A�@�I��
Private Const ptxD_ZAIKO_QTY02% = 69        '�A�@�݌ɐ�
Private Const ptxD_BIKOU02% = 70            '�A�@���l
    
Private Const ptxD_HIN_GAI03% = 71          '�B�@�����^�\���i��
Private Const ptxD_HIN_NAME03% = 72         '�B�@�����^�\���i��
Private Const ptxD_QTY03% = 73              '�B�@����
Private Const ptxD_SHIJI_QTY03% = 74        '�B�@����
Private Const ptxD_ST_LOCATION03% = 75      '�B�@�I��
Private Const ptxD_ZAIKO_QTY03% = 76        '�B�@�݌ɐ�
Private Const ptxD_BIKOU03% = 77            '�B�@���l
    
Private Const ptxD_HIN_GAI04% = 78          '�C�@�����^�\���i��
Private Const ptxD_HIN_NAME04% = 79         '�C�@�����^�\���i��
Private Const ptxD_QTY04% = 80              '�C�@����
Private Const ptxD_SHIJI_QTY04% = 81        '�C�@����
Private Const ptxD_ST_LOCATION04% = 82      '�C�@�I��
Private Const ptxD_ZAIKO_QTY04% = 83        '�C�@�݌ɐ�
Private Const ptxD_BIKOU04% = 84            '�C�@���l
    
Private Const ptxD_HIN_GAI05% = 85          '�D�@�����^�\���i��
Private Const ptxD_HIN_NAME05% = 86         '�D�@�����^�\���i��
Private Const ptxD_QTY05% = 87              '�D�@����
Private Const ptxD_SHIJI_QTY05% = 88        '�D�@����
Private Const ptxD_ST_LOCATION05% = 89      '�D�@�I��
Private Const ptxD_ZAIKO_QTY05% = 90        '�D�@�݌ɐ�
Private Const ptxD_BIKOU05% = 91            '�D�@���l
    
Private Const ptxD_HIN_GAI06% = 92          '�E�@�����^�\���i��
Private Const ptxD_HIN_NAME06% = 93         '�E�@�����^�\���i��
Private Const ptxD_QTY06% = 94              '�E�@����
Private Const ptxD_SHIJI_QTY06% = 95        '�E�@����
Private Const ptxD_ST_LOCATION06% = 96      '�E�@�I��
Private Const ptxD_ZAIKO_QTY06% = 97        '�E�@�݌ɐ�
Private Const ptxD_BIKOU06% = 98            '�E�@���l
    
    
    
 


'�R���{�p�Y��
Private Const pcmbSHIMUKE% = 0              '�d������
Private Const pcmbUKEHARAI% = 1             '��z��
Private Const pcmbS_TANTO% = 2              '���P�^�S���҃R�[�h

Private Const pcmbD_SYUBETSU01% = 3         '�@�@���
Private Const pcmbD_SYUBETSU02% = 4         '�A�@���
Private Const pcmbD_SYUBETSU03% = 5         '�B�@���
Private Const pcmbD_SYUBETSU04% = 6         '�C�@���
Private Const pcmbD_SYUBETSU05% = 7         '�D�@���
Private Const pcmbD_SYUBETSU06% = 8         '�E�@���


Private Const pcmbD_JGYOBU01% = 3           '�@�@���ƕ� 2016.01.27
Private Const pcmbD_JGYOBU02% = 4           '�A�@���ƕ� 2016.01.27
Private Const pcmbD_JGYOBU03% = 5           '�B�@���ƕ� 2016.01.27
Private Const pcmbD_JGYOBU04% = 6           '�C�@���ƕ� 2016.01.27
Private Const pcmbD_JGYOBU05% = 7           '�D�@���ƕ� 2016.01.27
Private Const pcmbD_JGYOBU06% = 8           '�E�@���ƕ� 2016.01.27



'�`�F�b�N�p�Y��
Private Const pchkSAMPLE_F% = 0             '���{�쐬
Private Const pchkPRI_SHIJI% = 1            '�o�͑Ώہ@�w�}�[
Private Const pchkPRI_PARTS% = 2            '�o�͑Ώہ@�߰�����
Private Const pchkPRI_GAISOU% = 3           '�o�͑Ώہ@�O������
Private Const pchkPRI_KISHU% = 4            '�o�͑Ώہ@�@������

'��߼�����ݗp�Y��
Private Const poptSHIJI_NORMAL% = 0         '�ʏ�
Private Const poptSHIJI_SPOT% = 1           '�X�|�b�g
Private Const poptSHIJI_KEPPIN% = 2         '���i����
Private Const poptSHIJI_SAIKON% = 3         '�č��� 2007.11.09


'���b�`�e�L�X�g�p�Y��
Private Const prchBIKOU% = 0                '���l





Private Const cmdMUPDATE% = 3               'Ͻ��X�V

Private Const cmdNext% = 5                  '�\�����i��ʂ�
Private Const cmdCen% = 10                  '������


'Private Const LAST_UPDATE_DAY$ = "([PI00015] 2017.10.17 09:30)"
'Private Const LAST_UPDATE_DAY$ = "([PI00015] 2017.12.15 10:15)"
Private Const LAST_UPDATE_DAY$ = "([PI00015] 2020.05.07 12:30) ��Ǝ��э��ږ��ύX"


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    PI000151.MousePointer = vbHourglass

    Call Ctrl_Lock(PI000151)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PI000151)


    PI000151.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer, chk As Integer, flg As Integer) As Integer
'----------------------------------------------------------------------------
'                   ���͍��ڂ̃G���[�`�F�b�N
'----------------------------------------------------------------------------
    
Dim sts         As Integer
    
    
Dim Mi_Qty      As Long
Dim Sumi_Qty    As Long
    
Dim i           As Integer
Dim j           As Integer
Dim k           As Integer
    
Dim yn          As Integer
    
Dim wkJgyobu    As String * 1
    
    Error_Check_Proc = True
    
    Select Case Mode
    
        Case ptxSHIJI_NO    '�w�}�[��
        
            If Text1(ptxSHIJI_NO).Locked Then
            Else
            
            
                    
            
            
                If Trim(Text1(ptxSHIJI_NO).text) = "" Then
                Else
                    
                    If IsNumeric(Text1(ptxSHIJI_NO).text) Then
                        Text1(ptxSHIJI_NO).text = Format(CLng(Text1(ptxSHIJI_NO).text), "00000000")
                    End If
                    
                    '�w�}�[�ް�������
                    sts = P_SSHIJI_Read_Proc()
                    Select Case sts
                        Case False, BtNoErr
''                            If CLng(StrConv(P_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)) <> 0 Then
''                                MsgBox "��������������͎�����ł��B���̉�ʂł͏����ł��܂���"
''                                Text1(Mode).SetFocus
''                                Exit Function
''                            End If
                            If CLng(StrConv(P_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)) <> 0 Then
                                yn = MsgBox("��������������͎�����ł��B" & Chr(13) & Chr(10) & _
                                        "�����ҏW����ꍇ�́A�u�͂��v���N���b�N�B", vbYesNo + vbDefaultButton2, "�m�F����")
                                If yn = vbNo Then
                                    Text1(Mode).SetFocus
                                    Exit Function
                                End If
                            End If
                            
                            If StrConv(P_SSHIJI_O_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
                                MsgBox "�L�����Z���ςł��B���̉�ʂł͏����ł��܂���"
                                Text1(Mode).SetFocus
                                Exit Function
                            End If
                        
                        Case BtErrKeyNotFound
                            MsgBox "���͂������ڂ̓G���[�ł��B"
                            Text1(Mode).SetFocus
                            Exit Function
                        Case Else
                            Exit Function
                    End Select
                End If
            
            
            
                Text1(ptxSHIJI_NO).BackColor = G_INPUT_NG
                Text1(ptxSHIJI_NO).Locked = True
                Text1(ptxSHIJI_NO).TabStop = False
            
            
            
            End If
        
        
        Case ptxORDER_NO    '�󒍓�     2012.03.18 �󒍇��ɕύX
            
            If chk = 1 Then
            Else
                
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.18
'                If Trim(Text1(ptxORDER_DT).text) = "" Then
'                Else
'
'                    If Not IsDate(Text1(ptxORDER_DT).text) Then
'                        MsgBox "���͂������ڂ̓G���[�ł��B(�󒍓�)"
'                        Text1(Mode).SetFocus
'                        Exit Function
'                    Else
'                        Text1(ptxORDER_DT).text = Format(CDate(Text1(ptxORDER_DT).text), "YYYY/MM/DD")
'                    End If
'                End If
            
            
                If Trim(Text1(ptxORDER_NO).text) = "" Then
                    Call Disp_Lock_Proc(False)
                Else
                    Call UniCode_Conv(K6_ODR_ORDER.ORDER_NO, Text1(ptxORDER_NO).text)
                
                    sts = BTRV(BtOpGetEqual, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K6_ODR_ORDER, Len(K6_ODR_ORDER), 6)
                    Select Case sts
                        Case BtNoErr
                        
                            If Trim(StrConv(ODR_ORDER_REC.FIN_DT, vbUnicode)) <> "" Then
                                MsgBox "���͂������ڂ̓G���[�ł��B(�e�i�Ԓ�����:�����ς݂ł�)"
                                Text1(Mode).SetFocus
                                Exit Function
                            End If
                        
                            Text1(ptxHIN_GAI).text = RTrim(StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode))
                        
                        
                            Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
                            Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
                            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).text)
                    
                    
                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                                    
                                    Text1(ptxHIN_NAME).text = ""
                                    Text1(ptxST_LOCATION).text = ""
                                    Text1(ptxMI_QTY).text = ""
                                    Text1(ptxSUMI_QTY).text = ""
                            
                                    MsgBox "���͂������ڂ̓G���[�ł��B(�i��)"
                                    Text1(Mode).SetFocus
                                    Exit Function
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                    Exit Function
                            
                            End Select
                        
                        
                        
                            Text1(ptxHIN_NAME).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                            
                            Text1(ptxSHIJI_QTY).text = Val(StrConv(ODR_ORDER_REC.ODR_QTY, vbUnicode))
                            
                            
                            If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                                Text1(ptxST_LOCATION).text = ""
                            Else
                                Text1(ptxST_LOCATION).text = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                                StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                                StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                                StrConv(ITEMREC.ST_DAN, vbUnicode)
                            End If
                
                            '>>>>>>>>>>>>>>>>>> 2013.01.07
                            'If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIZAI Then
                            '    wkJgyobu = BUZAI
                            'Else
                            '    'wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)  '2012.04.04
                            '    wkJgyobu = YUKO_JGYOBU                          '2012.04.04
                            'End If
                            wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)
                            '>>>>>>>>>>>>>>>>>> 2013.01.07


                            If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, wkJgyobu, _
                                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                                                    , , , Jyogai_Soko_umu) Then
                                Exit Function

                            End If

                            Text1(ptxMI_QTY).text = Format(Mi_Qty, "#0")
                            Text1(ptxSUMI_QTY).text = Format(Sumi_Qty, "#0")

                            If flg = 1 Then
                            Else

                                If Trim(Text1(ptxSHIJI_NO).text) = "" Then
                                    If StrConv(P_COMPO_O_REC.SHIMUKE_CODE, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) And _
                                        StrConv(P_COMPO_O_REC.JGYOBU, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) And _
                                        StrConv(P_COMPO_O_REC.NAIGAI, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) And _
                                        Trim(StrConv(P_COMPO_O_REC.HIN_GAI, vbUnicode)) = Trim(Text1(ptxHIN_GAI).text) Then
                                    Else
                                        sts = P_COMPO_Disp_Proc()
                                        Select Case sts
                                            Case BtNoErr
                                            Case BtErrKeyNotFound
                    '                            MsgBox "���͂������ڂ̓G���[�ł��B"
                    '                            Text1(Mode).SetFocus
                    '                            Exit Function
                                            Case Else
                                                Call File_Error(sts, BtOpGetEqual, "�\���}�X�^")
                                                Exit Function
                                        End Select
                                            
                                    End If
                                Else
                                    If StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) And _
                                        StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) And _
                                        StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) And _
                                        Trim(StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode)) = Trim(Text1(ptxHIN_GAI).text) Then
                                    Else
                                        sts = P_COMPO_Disp_Proc()
                                        Select Case sts
                                            Case BtNoErr
                                            Case BtErrKeyNotFound
                    '                            MsgBox "���͂������ڂ̓G���[�ł��B"
                    '                            Text1(Mode).SetFocus
                    '                            Exit Function
                                            Case Else
                                                Call File_Error(sts, BtOpGetEqual, "�\���}�X�^")
                                                Exit Function
                                        End Select
                                    End If
                                End If
                            End If
                        
                        
                                             '�����ލČv�Z
                        For i = ptxK_QTY01 To ptxK_QTY05 Step 5
                        
                            If IsNumeric(Text1(i).text) Then
                                Text1(i + 1).text = Format(CDbl(CLng(Text1(ptxSHIJI_QTY).text) * CDbl(Text1(i).text)), "#0.00")
                            Else
                                Text1(i + 1).text = ""
                            End If
                        Next i
                    
                    
                        '�O�����ލČv�Z
                        For i = ptxG_QTY01 To ptxG_QTY03 Step 5
                        
                            If IsNumeric(Text1(i).text) Then
                                Text1(i + 1).text = Format(Int(CDbl(CLng(Text1(ptxSHIJI_QTY).text) / CDbl(Text1(i).text))), "#0")
                            Else
                                Text1(i + 1).text = ""
                            End If
                        Next i
                    
                        '�����^�\���Čv�Z
                        For i = ptxD_QTY01 To ptxD_QTY06 Step 7
                        
                            k = 0
                            j = Mode - ptxD_QTY01
                            Do
                                j = j - 5
                                If j < 0 Then
                                    Exit Do
                                End If
                                k = k + 1
                            Loop
                        
                            
                            
                            
                            If IsNumeric(Text1(i).text) Then
                                Text1(i + 1).text = Format(CDbl(CLng(Text1(ptxSHIJI_QTY).text) * CDbl(Text1(i).text)), "#0.00")
                            
                            
                            Else
                                Text1(i + 1).text = ""
                            
                            
                            
                            End If
                        Next i
                    
                    
                        For i = 0 To UBound(D_Item_Tbl)
                            If Trim(D_Item_Tbl(i).JGYOBU) = "" Then
                            Else
                                D_Item_Tbl(i).SHIJI_QTY = CDbl(Text1(ptxSHIJI_QTY).text) * D_Item_Tbl(i).QTY
                            End If
                        Next i
   
                        
                        
                        
                        Case BtErrKeyNotFound
                            MsgBox "���͂������ڂ̓G���[�ł��B(�e�i�Ԓ�����)"
                            Text1(Mode).SetFocus
                            Exit Function
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�e�i�Ԓ���F")
                            Exit Function
                    
                    End Select
                
                    Call Disp_Lock_Proc(True)
                
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    �Ĕ��s���̌x��  2012.04.13
                    If StrConv(ODR_ORDER_REC.PRT_FLG, vbUnicode) = "F" Then
                        yn = MsgBox("�w�}�[���s�ς݂ł��B�������p�����܂����H", vbYesNo + vbDefaultButton2, "�m�F����")
                        If yn = vbNo Then
                            Call Disp_Lock_Proc(False)
                            Exit Function
                        End If
                    End If
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    �Ĕ��s���̌x��  2012.04.13
                
                
                
                End If
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.18
            End If
        
        
        
        Case ptxHAKKO_DT    '���s��
            
            If chk = 1 Then
            Else
                If Trim(Text1(ptxHAKKO_DT).text) = "" Then
                Else
                    If Not IsDate(Text1(ptxHAKKO_DT).text) Then
                        MsgBox "���͂������ڂ̓G���[�ł��B(���s��)"
                        Text1(Mode).SetFocus
                        Exit Function
                    Else
                        Text1(ptxHAKKO_DT).text = Format(CDate(Text1(ptxHAKKO_DT).text), "YYYY/MM/DD")
                    End If
                End If
            End If
        
        Case ptxTANTO_CODE      '�S����
        
            If chk = 1 Then
            Else
                Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxTANTO_CODE).text)
    
                sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
                Select Case sts
                    Case BtNoErr
                        Text1(ptxTANTO_NAME).text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
                    Case BtErrKeyNotFound
                        Text1(ptxTANTO_NAME).text = ""
                
                        MsgBox "���͂������ڂ̓G���[�ł��B(�S����)"
                        Text1(Mode).SetFocus
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
                        Exit Function
                
                End Select
            End If
    
        Case ptxSHONIN_CODE     '���F��
        
            If chk = 1 Then
            Else
                Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxSHONIN_CODE).text)
    
                sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
                Select Case sts
                    Case BtNoErr
                        Text1(ptxSHONIN_NAME).text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
                    Case BtErrKeyNotFound
                        Text1(ptxSHONIN_NAME).text = ""
                
                        MsgBox "���͂������ڂ̓G���[�ł��B(���F��)"
                        Text1(Mode).SetFocus
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
                        Exit Function
                    
                
                
                End Select
            End If
        Case ptxHIN_GAI         '�i��
    
                    
            Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
            Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).text)
    
    
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    
                    Text1(ptxHIN_NAME).text = ""
                    Text1(ptxST_LOCATION).text = ""
                    Text1(ptxMI_QTY).text = ""
                    Text1(ptxSUMI_QTY).text = ""
            
                    MsgBox "���͂������ڂ̓G���[�ł��B(�i��)"
                    Text1(Mode).SetFocus
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                    Exit Function
            
            End Select
            
            Text1(ptxHIN_NAME).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
            
            If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                Text1(ptxST_LOCATION).text = ""
            Else
                Text1(ptxST_LOCATION).text = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_DAN, vbUnicode)
            End If

            '>>>>>>>>>>>>>> 2013.01.07
            'If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIZAI Then
            '    wkJgyobu = BUZAI
            'Else
            '    'wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)  2012.04.04
            '    wkJgyobu = YUKO_JGYOBU                          '2012.04.04
            'End If
            wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)
            '>>>>>>>>>>>>>> 2013.01.07

            If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, wkJgyobu, _
                                                   StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                   StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                                   , , , Jyogai_Soko_umu) Then
                Exit Function
            
            End If

            Text1(ptxMI_QTY).text = Format(Mi_Qty, "#0")
            Text1(ptxSUMI_QTY).text = Format(Sumi_Qty, "#0")
    
            
            If flg = 1 Then
            Else
            
                If Trim(Text1(ptxSHIJI_NO).text) = "" Then
                    If StrConv(P_COMPO_O_REC.SHIMUKE_CODE, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) And _
                        StrConv(P_COMPO_O_REC.JGYOBU, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) And _
                        StrConv(P_COMPO_O_REC.NAIGAI, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) And _
                        Trim(StrConv(P_COMPO_O_REC.HIN_GAI, vbUnicode)) = Trim(Text1(ptxHIN_GAI).text) Then
                    Else
                        sts = P_COMPO_Disp_Proc()
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
    '                            MsgBox "���͂������ڂ̓G���[�ł��B"
    '                            Text1(Mode).SetFocus
    '                            Exit Function
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "�\���}�X�^")
                                Exit Function
                        End Select
                            
                    End If
                Else
                    If StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) And _
                        StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) And _
                        StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) And _
                        Trim(StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode)) = Trim(Text1(ptxHIN_GAI).text) Then
                    Else
                        sts = P_COMPO_Disp_Proc()
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
    '                            MsgBox "���͂������ڂ̓G���[�ł��B"
    '                            Text1(Mode).SetFocus
    '                            Exit Function
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "�\���}�X�^")
                                Exit Function
                        End Select
                    End If
                End If
            End If
        Case ptxSHIJI_QTY       '����
    
            If chk = 1 Then
            Else
                If Not IsNumeric(Text1(ptxSHIJI_QTY).text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(����)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxSHIJI_QTY).text = Format(CLng(Text1(ptxSHIJI_QTY).text), "#0")
                
                    '�����ލČv�Z
                    For i = ptxK_QTY01 To ptxK_QTY05 Step 5
                    
                        If IsNumeric(Text1(i).text) Then
                            Text1(i + 1).text = Format(CDbl(CLng(Text1(ptxSHIJI_QTY).text) * CDbl(Text1(i).text)), "#0.00")
                        Else
                            Text1(i + 1).text = ""
                        End If
                    Next i
                
                
                    '�O�����ލČv�Z
                    For i = ptxG_QTY01 To ptxG_QTY03 Step 5
                    
                        If IsNumeric(Text1(i).text) Then
                            Text1(i + 1).text = Format(Int(CDbl(CLng(Text1(ptxSHIJI_QTY).text) / CDbl(Text1(i).text))), "#0")
                        Else
                            Text1(i + 1).text = ""
                        End If
                    Next i
                
                    '�����^�\���Čv�Z
                    For i = ptxD_QTY01 To ptxD_QTY06 Step 7
                    
                        k = 0
                        j = Mode - ptxD_QTY01
                        Do
                            j = j - 5
                            If j < 0 Then
                                Exit Do
                            End If
                            k = k + 1
                        Loop
                    
                        
                        
                        
                        If IsNumeric(Text1(i).text) Then
                            Text1(i + 1).text = Format(CDbl(CLng(Text1(ptxSHIJI_QTY).text) * CDbl(Text1(i).text)), "#0.00")
                        
                        
                        Else
                            Text1(i + 1).text = ""
                        
                        
                        
                        End If
                    Next i
                
                
                    For i = 0 To UBound(D_Item_Tbl)
                        If Trim(D_Item_Tbl(i).JGYOBU) = "" Then
                        Else
                            D_Item_Tbl(i).SHIJI_QTY = CDbl(Text1(ptxSHIJI_QTY).text) * D_Item_Tbl(i).QTY
                        End If
                    Next i
                End If
            End If
    
        Case ptxUKEHARAI_CODE   '��z��
            
            If chk = 1 Then
            Else
               Combo1(pcmbUKEHARAI).ListIndex = -1
               For i = 0 To Combo1(pcmbUKEHARAI).ListCount - 1
                   If Text1(ptxUKEHARAI_CODE).text = Trim(Right(Combo1(pcmbUKEHARAI).List(i), 5)) Then
                       Combo1(pcmbUKEHARAI).ListIndex = i
                       Exit For
                   End If
               
               Next i
        
               If i > Combo1(pcmbUKEHARAI).ListCount - 1 Then
                   MsgBox "���͂������ڂ̓G���[�ł��B(��z��)"
                   Text1(Mode).SetFocus
                   Exit Function
               End If
            End If
    
    
    
        Case ptxS_CLASS_CODE    '���i���׽
        
            If Trim(Text1(ptxS_CLASS_CODE).text) = "" Then
            Else
            
            
                Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 1, 2))
                Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, Text1(ptxS_CLASS_CODE).text)
    
                sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                
                        MsgBox "���͂������ڂ̓G���[�ł��B(���i���׽)"
                        Text1(Mode).SetFocus
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "���i���׽")
                        Exit Function
                
                End Select
            End If
        Case ptxF_CLASS_CODE    '�t���׽
        
            If Trim(Text1(ptxF_CLASS_CODE).text) = "" Then
            Else
                Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 1, 2))
                Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, Text1(ptxF_CLASS_CODE).text)
    
                sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                
                        MsgBox "���͂������ڂ̓G���[�ł��B(�t���׽)"
                        Text1(Mode).SetFocus
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "���i���׽")
                        Exit Function
                
                End Select
            End If
    
        Case ptxN_CLASS_CODE    '���E�׽
        
            If Trim(Text1(ptxN_CLASS_CODE).text) = "" Then
            Else
                Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 1, 2))
                Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, Text1(ptxN_CLASS_CODE).text)
    
                sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                
                        MsgBox "���͂������ڂ̓G���[�ł��B(���E�׽)"
                        Text1(Mode).SetFocus
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "���i���׽")
                        Exit Function
                
                End Select
            End If
    
                                '�����އ�
        Case ptxK_HIN_GAI01, ptxK_HIN_GAI02, ptxK_HIN_GAI03, ptxK_HIN_GAI04, ptxK_HIN_GAI05
            If Trim(Text1(Mode).text) = "" Then
                Text1(Mode + 1).text = ""
                Text1(Mode + 2).text = ""
                Text1(Mode + 3).text = ""
                Text1(Mode + 4).text = ""
            Else
                
                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2014.03.24
'                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1))
'                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1))
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)
'
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                Select Case sts
'                    Case BtNoErr
'                    Case BtErrKeyNotFound
'                        '���ޕi�œǂݑւ�
'
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
'                        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
'                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)
'
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                        Select Case sts
'                            Case BtNoErr
'                            Case BtErrKeyNotFound
'
'                                If HIN_INV Then
'                                    '���o�^�i�ԁ@�@���ނƂ��Ă���
'                                    Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)
'                                    Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
'                                    Call UniCode_Conv(ITEMREC.HIN_NAME, "���o�^�i��")
'                                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
'                                Else
'                                    MsgBox "���͂������ڂ̓G���[�ł��B(�����ށ@�i��)"
'                                    Text1(Mode).SetFocus
'                                    Exit Function
'                                End If
'                            Case Else
'                                Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
'                                Exit Function
'
'                        End Select
'
'                    Case Else
'                        Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
'                        Exit Function
'
'                End Select

                sts = Item_Read_Proc(Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1), Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1), Text1(Mode).text)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        MsgBox "���͂������ڂ̓G���[�ł��B(�����ށ@�i��)"
                        Text1(Mode).SetFocus
                        Exit Function
                    Case Else
                        Exit Function
                End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2014.03.24
    
                '�i��
                Text1(Mode + 1).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                '�W���I��
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                Else
                    Text1(Mode + 4).text = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
            
                
                i = 0
                j = Mode - ptxK_HIN_GAI01
                Do
                    j = j - 5
                    If j < 0 Then
                        Exit Do
                    End If
                    i = i + 1
                Loop
            
                K_Item_Tbl(i).JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
                K_Item_Tbl(i).NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
            
            
            End If
                                '�����ށ@����
        Case ptxK_QTY01, ptxK_QTY02, ptxK_QTY03, ptxK_QTY04, ptxK_QTY05
            
            If Trim(Text1(Mode).text) = "" Then
                If Trim(Text1(Mode - 1).text) <> "" Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(�����ށ@����)"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            Else
                If Trim(Text1(Mode - 1).text) = "" Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(�����ށ@����)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    If Not IsNumeric(Text1(Mode).text) Then
                        MsgBox "���͂������ڂ̓G���[�ł��B(�����ށ@����)"
                        Text1(Mode).SetFocus
                        Exit Function
                    Else
                        Text1(Mode).text = Format(CDbl(Text1(Mode).text), "#0.00")
                        '����
                        If IsNumeric(Text1(ptxSHIJI_QTY).text) Then
                            Text1(Mode + 1).text = Format(CDbl(CLng(Text1(ptxSHIJI_QTY).text) * CDbl(Text1(Mode).text)), "#0.00")
                        
                        
                        
                        Else
                            Text1(Mode + 1).text = ""
                        End If
                    
                    End If
                End If
            End If
    
    
    
    
    
                                '�O�����އ�
        Case ptxG_HIN_GAI01, ptxG_HIN_GAI02, ptxG_HIN_GAI03
            If Trim(Text1(Mode).text) = "" Then
                Text1(Mode + 1).text = ""
                Text1(Mode + 2).text = ""
                Text1(Mode + 3).text = ""
                Text1(Mode + 4).text = ""
            Else
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2014.03.24
'                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1))
'                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1))
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)
'
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                Select Case sts
'                    Case BtNoErr
'                    Case BtErrKeyNotFound
'                        '���ޕi�œǂݑւ�
'
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
'                        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
'                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)
'
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                        Select Case sts
'                            Case BtNoErr
'                            Case BtErrKeyNotFound
'
'                                If HIN_INV Then
'                                    '���o�^�i�ԁ@�@���ނƂ��Ă���
'                                    Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)
'                                    Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
'                                    Call UniCode_Conv(ITEMREC.HIN_NAME, "���o�^�i��")
'                                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
'                                Else
'
'                                    MsgBox "���͂������ڂ̓G���[�ł��B(�O�����ށ@�i��)"
'                                    Text1(Mode).SetFocus
'                                    Exit Function
'                                End If
'                            Case Else
'                                Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
'                                Exit Function
'
'                        End Select
'
'                    Case Else
'                        Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
'                        Exit Function
'
'                End Select
                sts = Item_Read_Proc(Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1), Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1), Text1(Mode).text)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        MsgBox "���͂������ڂ̓G���[�ł��B(�O�����ށ@�i��)"
                        Text1(Mode).SetFocus
                        Exit Function
                    Case Else
                        Exit Function
                End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2014.03.24
    
                '�i��
                Text1(Mode + 1).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                '�W���I��
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                Else
                    Text1(Mode + 4).text = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
            
            
                i = 0
                j = Mode - ptxG_HIN_GAI01
                Do
                    j = j - 5
                    If j < 0 Then
                        Exit Do
                    End If
                    i = i + 1
                Loop
            
                G_Item_Tbl(i).JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
                G_Item_Tbl(i).NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
            
            
            End If
                                '�O�����ށ@����
        Case ptxG_QTY01, ptxG_QTY02, ptxG_QTY03
            
            If Trim(Text1(Mode).text) = "" Then
                If Trim(Text1(Mode - 1).text) <> "" Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(�O�����ށ@����)"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            Else
                If Trim(Text1(Mode - 1).text) = "" Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(�O�����ށ@����)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    If Not IsNumeric(Text1(Mode).text) Then
                        MsgBox "���͂������ڂ̓G���[�ł��B(�O�����ށ@����)"
                        Text1(Mode).SetFocus
                        Exit Function
                    Else
                        Text1(Mode).text = Format(CDbl(Text1(Mode).text), "#0.00")
                        '����
                        If IsNumeric(Text1(ptxSHIJI_QTY).text) Then
                            Text1(Mode + 1).text = Format(Int(CDbl(CLng(Text1(ptxSHIJI_QTY).text) / CDbl(Text1(Mode).text))), "#0")
                        
                        Else
                            Text1(Mode + 1).text = ""
                        End If
                    
                    End If
                End If
            End If
    
                                '�����^�\���@�i��
        Case ptxD_HIN_GAI01, ptxD_HIN_GAI02, ptxD_HIN_GAI03, ptxD_HIN_GAI04, ptxD_HIN_GAI05, ptxD_HIN_GAI06
            If Trim(Text1(Mode).text) = "" Then
                Text1(Mode + 1).text = ""
                Text1(Mode + 2).text = ""
                Text1(Mode + 3).text = ""
                Text1(Mode + 4).text = ""
                Text1(Mode + 5).text = ""
                Text1(Mode + 6).text = ""
            
            
                i = 0
                j = Mode - ptxD_HIN_GAI01
                Do
                    j = j - 7
                    If j < 0 Then
                        Exit Do
                    End If
                    i = i + 1
                Loop
            
                D_Item_Tbl(i).JGYOBU = ""
                D_Item_Tbl(i).NAIGAI = ""
                D_Item_Tbl(i).HIN_GAI = ""
                D_Item_Tbl(i).QTY = 0
                D_Item_Tbl(i).SHIJI_QTY = 0
                D_Item_Tbl(i).BIKOU = ""
            
            
            Else
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2014.03.24
'                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1))
'                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1))
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)
'
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                Select Case sts
'                    Case BtNoErr
'                    Case BtErrKeyNotFound
'
'                        '�i�ԁi���j�œǂݑւ�
'                        Call UniCode_Conv(K2_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1))
'                        Call UniCode_Conv(K2_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1))
'                        Call UniCode_Conv(K2_ITEM.HIN_NAI, Text1(Mode).text)
'
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
'                        Select Case sts
'                            Case BtNoErr
'                            Case BtErrKeyNotFound
'
'
'
'
'
'
'
'                                '���ޕi�œǂݑւ�
'                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
'
'                                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
'                                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)
'
'                                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                                Select Case sts
'                                    Case BtNoErr
'                                    Case BtErrKeyNotFound
'
'                                        If HIN_INV Then
'                                            '���o�^�i�ԁ@�@���ނƂ��Ă���
'                                            Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)
'                                            Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
'                                            Call UniCode_Conv(ITEMREC.HIN_GAI, Text1(Mode).text)
'                                            Call UniCode_Conv(ITEMREC.HIN_NAME, "���o�^�i��")
'                                            Call UniCode_Conv(ITEMREC.ST_SOKO, "")
'
'                                        Else
'
'                                            MsgBox "���͂������ڂ̓G���[�ł��B(�����^�\���@�i��)"
'                                            Text1(Mode).SetFocus
'                                            Exit Function
'                                        End If
'                                    Case Else
'                                        Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
'                                        Exit Function
'
'                                End Select
'
'                            Case Else
'                                Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
'                                Exit Function
'                       End Select
'
'                    Case Else
'                        Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
'                        Exit Function
'
'                End Select
                
                
                
'>>>>>>>>>>>>>>>>>>>>>>>>>  �i�ԓǍ��ݕύX 2016.01.27
'                sts = Item_Read_Proc(Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1), Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1), Text1(Mode).text)
                
                Select Case Mode
                    Case ptxD_HIN_GAI01
                        Call UniCode_Conv(K0_ITEM.JGYOBU, Right(Combo2(pcmbD_JGYOBU01).text, 1))
                    Case ptxD_HIN_GAI02
                        Call UniCode_Conv(K0_ITEM.JGYOBU, Right(Combo2(pcmbD_JGYOBU02).text, 1))
                    Case ptxD_HIN_GAI03
                        Call UniCode_Conv(K0_ITEM.JGYOBU, Right(Combo2(pcmbD_JGYOBU03).text, 1))
                    Case ptxD_HIN_GAI04
                        Call UniCode_Conv(K0_ITEM.JGYOBU, Right(Combo2(pcmbD_JGYOBU04).text, 1))
                    Case ptxD_HIN_GAI05
                        Call UniCode_Conv(K0_ITEM.JGYOBU, Right(Combo2(pcmbD_JGYOBU05).text, 1))
                    Case ptxD_HIN_GAI06
                        Call UniCode_Conv(K0_ITEM.JGYOBU, Right(Combo2(pcmbD_JGYOBU06).text, 1))
                End Select
                
                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)
                
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'>>>>>>>>>>>>>>>>>>>>>>>>>  �i�ԓǍ��ݕύX 2016.01.27
                
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        MsgBox "���͂������ڂ̓G���[�ł��B(�����^�\���@�i��)"
                        Text1(Mode).SetFocus
                        Exit Function
                    Case Else
                        Exit Function
                End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2014.03.24
    
                '�i��
                Text1(Mode + 1).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
'                '�W���I��
'                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
'                Else
'                    Text1(Mode + 4).text = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
'                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
'                                            StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
'                                            StrConv(ITEMREC.ST_DAN, vbUnicode)
'                End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
            
                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.01.07
                '�݌ɐ�
                'If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIZAI Then
                '    wkJgyobu = BUZAI
                'Else
                '    'wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)  2012.04.04
                '    wkJgyobu = YUKO_JGYOBU                          '2012.04.04
                'End If
                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2016.01.27
'                Select Case StrConv(ITEMREC.JGYOBU, vbUnicode)
'                    Case SHIZAI
'                        wkJgyobu = BUZAI
'                    Case SETSUBI
'                        wkJgyobu = YUKO_JGYOBU
'                    Case Else
'                        wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)
'                End Select

                wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2016.01.27

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.01.07
                
                '�W���I�Ԃ�ݒ� 2013.01.11
                If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, wkJgyobu, _
                                                        StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                        StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                                        (StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                        StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                        StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                        StrConv(ITEMREC.ST_DAN, vbUnicode)), _
                                                        , , Jyogai_Soko_umu) Then
                    Exit Function
                
                End If
            
                Text1(Mode + 5).text = Format(Sumi_Qty + Mi_Qty, "#0")
            
            
            
                i = 0
                j = Mode - ptxD_HIN_GAI01
                Do
                    j = j - 7
                    If j < 0 Then
                        Exit Do
                    End If
                    i = i + 1
                Loop
            
' 2013.01.07 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                D_Item_Tbl(i).JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
'                D_Item_Tbl(i).JGYOBU = BUZAI            '����/�\���̎��ƕ����u���ށv�Œ�ɕύX
' 2013.01.07 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                D_Item_Tbl(i).NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
                D_Item_Tbl(i).HIN_GAI = StrConv(ITEMREC.HIN_GAI, vbUnicode)
            
            
            
            
            
            
            
            
'>>>>>>>>>>>>>>>>>>>>>  2013.01.07 �W���I�Ԃ̕\��������--�����޾���
                
'>>>>>>>>>> �p�~�@2016.01.27
'                Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)
'
'                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)
'
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                Select Case sts
'                    Case BtNoErr
'                    Case BtErrKeyNotFound
'                        Call UniCode_Conv(ITEMREC.ST_SOKO, "")
'                    Case Else
'                        Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
'                        Exit Function
'
'                End Select
'
'                '�W���I��
'                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
'                Else
'                    Text1(Mode + 4).text = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
'                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
'                                            StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
'                                            StrConv(ITEMREC.ST_DAN, vbUnicode)
'                End If
'>>>>>>>>>> �p�~�@2016.01.27


'>>>>>>>>>>>>>>>>>>>>>  2013.01.07

            
            
            End If
                                '�����^�\���@����
        Case ptxD_QTY01, ptxD_QTY02, ptxD_QTY03, ptxD_QTY04, ptxD_QTY05, ptxD_QTY06
            
            If Trim(Text1(Mode).text) = "" Then
                If Trim(Text1(Mode - 2).text) <> "" Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(�����^�\���@����)"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            Else
                If Trim(Text1(Mode - 2).text) = "" Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(�����^�\���@����)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    If Not IsNumeric(Text1(Mode).text) Then
                        MsgBox "���͂������ڂ̓G���[�ł��B(�����^�\���@����)"
                        Text1(Mode).SetFocus
                        Exit Function
                    Else
                        Text1(Mode).text = Format(CDbl(Text1(Mode).text), "#0.00")
                        
                        
                        i = 0
                        j = Mode - ptxD_QTY01
                        Do
                            j = j - 7
                            If j < 0 Then
                                Exit Do
                            End If
                            i = i + 1
                        Loop
                    
                        D_Item_Tbl(i).QTY = CDbl(Text1(Mode).text)
                        
                        
                        '����
                        If IsNumeric(Text1(ptxSHIJI_QTY).text) Then
                            Text1(Mode + 1).text = Format(CDbl(CLng(Text1(ptxSHIJI_QTY).text) * CDbl(Text1(Mode).text)), "#0.00")
                            D_Item_Tbl(i).SHIJI_QTY = CDbl(Text1(Mode + 1).text)
                        Else
                            Text1(Mode + 1).text = ""
                            D_Item_Tbl(i).SHIJI_QTY = 0
                        End If
                    
                    End If
                End If
            End If
                                '�����^�\���@���l
        Case ptxD_BIKOU01, ptxD_BIKOU02, ptxD_BIKOU03, ptxD_BIKOU04, ptxD_BIKOU05, ptxD_BIKOU06
            If Trim(Text1(Mode).text) <> "" Then
                If Trim(Text1(Mode - 6).text) = "" Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(�����^�\���@���l)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    i = 0
                    j = Mode - ptxD_BIKOU01
                    Do
                        j = j - 7
                        If j < 0 Then
                            Exit Do
                        End If
                        i = i + 1
                    Loop
                
                    D_Item_Tbl(i).BIKOU = Text1(Mode).text
                End If
            End If
    End Select
        
        
    Error_Check_Proc = False
    

End Function

Private Function Item_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   ��ʕ\��
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim k           As Integer
Dim g           As Integer
Dim d           As Integer

Dim K_Index     As Integer
Dim G_Index     As Integer
Dim DT_Index    As Integer
Dim DC_Index    As Integer


Dim Mi_Qty      As Long
Dim Sumi_Qty    As Long

Dim wkJgyobu    As String * 1


Dim Wk_LOC      As String   '2013.01.07


    Item_Disp_Proc = True
    
    For i = ptxK_HIN_GAI01 To ptxD_BIKOU06
        Text1(i).text = ""
    Next i
            
    For i = pcmbD_JGYOBU01 To pcmbD_JGYOBU06        '2016.04.07
        Combo2(i).ListIndex = -1                    '2016.04.07
    Next                                            '2016.04.07
            
            
    '�o�͑Ώ�
    Erase K_Item_Tbl
    Erase G_Item_Tbl
    Erase D_Item_Tbl
    
    ReDim K_Item_Tbl(0 To 4)
    ReDim G_Item_Tbl(0 To 2)
    ReDim D_Item_Tbl(0 To 49)

    For i = 0 To UBound(K_Item_Tbl)
        K_Item_Tbl(i).JGYOBU = ""
        K_Item_Tbl(i).NAIGAI = ""
    Next i

    For i = 0 To UBound(G_Item_Tbl)
        G_Item_Tbl(i).JGYOBU = ""
        G_Item_Tbl(i).NAIGAI = ""
    Next i

    For i = 0 To UBound(D_Item_Tbl)
        D_Item_Tbl(i).JGYOBU = ""
        D_Item_Tbl(i).NAIGAI = ""
        D_Item_Tbl(i).HIN_GAI = ""
    
    Next i
    
    Text1(ptxS_CLASS_CODE).text = ""
    Text1(ptxF_CLASS_CODE).text = ""
    Text1(ptxN_CLASS_CODE).text = ""
    
    
    
    '--------------------------------   �u�e�v���
        
    
    Text1(ptxSHIJI_NO).text = StrConv(P_SSHIJI_O_REC.SHIJI_No, vbUnicode)           '�w�}�[��
                                                                                    
                                                                                    '�󒍓�
    
    If Trim(StrConv(P_SSHIJI_O_REC.ORDER_DT, vbUnicode)) = "" Then
        Text1(ptxORDER_NO).text = ""
    Else
'---------------------------    2012.03.27
'        Text1(ptxORDER_NO).text = Mid(StrConv(P_SSHIJI_O_REC.ORDER_DT, vbUnicode), 1, 4) & "/" & _
'                                    Mid(StrConv(P_SSHIJI_O_REC.ORDER_DT, vbUnicode), 5, 2) & "/" & _
'                                    Mid(StrConv(P_SSHIJI_O_REC.ORDER_DT, vbUnicode), 7, 2)

        Text1(ptxORDER_NO).text = StrConv(P_SSHIJI_O_REC.ORDER_DT, vbUnicode) & StrConv(P_SSHIJI_O_REC.ORDER_DT_SEQ, vbUnicode)
'---------------------------    2012.03.27
    End If
                                                                                    
                                                                                    
                                                                                    
                                                                                    '���s��
    
    If Trim(StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode)) = "" Then
        Text1(ptxHAKKO_DT).text = ""
    Else
    
        Text1(ptxHAKKO_DT).text = Mid(StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 7, 2)
    
    End If
    
    
    Text1(ptxTANTO_CODE).text = StrConv(P_SSHIJI_O_REC.TANTO_CODE, vbUnicode)       '�S���Һ��ށ^����
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxTANTO_CODE).text)

    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
            Text1(ptxTANTO_NAME).text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Text1(ptxTANTO_NAME).text = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
            Exit Function
    
    End Select
    
    Text1(ptxSHONIN_CODE).text = StrConv(P_SSHIJI_O_REC.SHONIN_CODE, vbUnicode)     '���F�Һ��ށ^����
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxSHONIN_CODE).text)

    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
            Text1(ptxSHONIN_NAME).text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Text1(ptxSHONIN_NAME).text = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
            Exit Function
    
    End Select
    
    For i = 0 To Combo1(pcmbSHIMUKE).ListCount - 1                                  '�d�����溰��
    
        If StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE).List(i), 4), 1, 2) Then
            Combo1(pcmbSHIMUKE).ListIndex = i
            Exit For
        End If
    
    Next i
    
    
    Text1(ptxHIN_GAI).text = Trim(StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode))       '�i�ԁ^�i���^�W���I�ԁ^�����i�^���i����
        
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).text)

    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            Text1(ptxHIN_NAME).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
            Text1(ptxST_LOCATION).text = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_DAN, vbUnicode)
        
        
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
'            If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIZAI Then
'                wkJgyobu = BUZAI
'            Else
'                'wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)  2012.04.04
'                wkJgyobu = YUKO_JGYOBU                          '2012.04.04
'            End If

            wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
        
            If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, wkJgyobu, _
                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode), , , , Jyogai_Soko_umu) Then
                Exit Function
            
            End If

            Text1(ptxMI_QTY).text = Format(Mi_Qty, "#0")
            Text1(ptxSUMI_QTY).text = Format(Sumi_Qty, "#0")
        
        
        
        Case BtErrKeyNotFound
            Text1(ptxHIN_NAME).text = ""
            Text1(ptxST_LOCATION).text = ""
            Text1(ptxMI_QTY).text = ""
            Text1(ptxSUMI_QTY).text = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Function
    
    End Select
                                                                                    '�w������
    Text1(ptxSHIJI_QTY).text = Format(CLng(StrConv(P_SSHIJI_O_REC.SHIJI_QTY, vbUnicode)), "#0")
    
    Text1(ptxHIN_GAI).text = Trim(StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode))       '�i�ԁ^�i���^�W���I�ԁ^�����i�^���i����
                                                                                        
    Text1(ptxUKEHARAI_CODE).text = Trim(StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode))   '��z��
    For i = 0 To Combo1(pcmbUKEHARAI).ListCount - 1
    
        If Text1(ptxUKEHARAI_CODE).text = Trim(Right(Combo1(pcmbUKEHARAI).List(i), 5)) Then
            Combo1(pcmbUKEHARAI).ListIndex = i
            Exit For
        End If
    
    Next i
    
    Text1(ptxS_CLASS_CODE).text = Trim(StrConv(P_SSHIJI_O_REC.S_CLASS_CODE, vbUnicode)) '���i���׽
    Text1(ptxF_CLASS_CODE).text = Trim(StrConv(P_SSHIJI_O_REC.F_CLASS_CODE, vbUnicode)) '�t���׽
    Text1(ptxN_CLASS_CODE).text = Trim(StrConv(P_SSHIJI_O_REC.N_CLASS_CODE, vbUnicode)) '���E�׽
    

    If Combo1(pcmbS_TANTO).ListCount = 0 Then                                       '���P�^�S����
    Else
        For i = 0 To Combo1(pcmbS_TANTO).ListCount - 1
            If StrConv(P_SSHIJI_O_REC.S_TANTO, vbUnicode) = Right(Combo1(pcmbS_TANTO).List(i), 2) Then
                Combo1(pcmbS_TANTO).ListIndex = i
                Exit For
            End If
        Next i
    End If
    
    
    If StrConv(P_SSHIJI_O_REC.SAMPLE_F, vbUnicode) = P_SAMPLE_F_OFF Then            '���{�쐬
        Check1(pchkSAMPLE_F).Value = vbUnchecked
    Else
        Check1(pchkSAMPLE_F).Value = vbChecked
    End If
    
    Select Case StrConv(P_SSHIJI_O_REC.SHIJI_F, vbUnicode)                          '�ʏ�/��߯�/���i����
        Case P_SHIJI_F_NORMAL
            Option1(poptSHIJI_NORMAL).Value = True
            Option1(poptSHIJI_SPOT).Value = False
            Option1(poptSHIJI_KEPPIN).Value = False
        Case P_SHIJI_F_SPOT
            Option1(poptSHIJI_NORMAL).Value = False
            Option1(poptSHIJI_SPOT).Value = True
            Option1(poptSHIJI_KEPPIN).Value = False
        Case P_SHIJI_F_KEPPIN
            Option1(poptSHIJI_NORMAL).Value = False
            Option1(poptSHIJI_SPOT).Value = False
            Option1(poptSHIJI_KEPPIN).Value = True
    End Select
    
    
    If StrConv(P_SSHIJI_O_REC.PRI_SHIJI, vbUnicode) = P_PRI_SHIJI_OFF Then          '�o�͑Ώہ@�w�}�[
        Check1(pchkPRI_SHIJI).Value = vbUnchecked
    Else
        Check1(pchkPRI_SHIJI).Value = vbChecked
    End If
    
    If StrConv(P_SSHIJI_O_REC.PRI_PARTS, vbUnicode) = P_PRI_PARTS_OFF Then          '�o�͑Ώہ@�߰�����
        Check1(pchkPRI_PARTS).Value = vbUnchecked
    Else
        Check1(pchkPRI_PARTS).Value = vbChecked
    End If
    
    RichTextBox1(prchBIKOU).text = StrConv(P_SSHIJI_O_REC.BIKOU, vbUnicode)         '���l
    
    '--------------------------------   �u�q�v���
    Erase K_Item_Tbl
    Erase G_Item_Tbl
    Erase D_Item_Tbl
    
    ReDim K_Item_Tbl(0 To 4)
    ReDim G_Item_Tbl(0 To 2)
    ReDim D_Item_Tbl(0 To 49)
    
    
    
    
    k = -1
    g = -1
    d = -1
    
    K_Index = ptxK_HIN_GAI01
    G_Index = ptxG_HIN_GAI01
    DT_Index = ptxD_HIN_GAI01
    DC_Index = pcmbD_SYUBETSU01
    
            
        
    
    Call UniCode_Conv(K0_P_SSHIJI_K.SHIJI_No, Text1(ptxSHIJI_NO).text)
    Call UniCode_Conv(K0_P_SSHIJI_K.DATA_KBN, "")
    Call UniCode_Conv(K0_P_SSHIJI_K.SEQNO, "")
    
    com = BtOpGetGreaterEqual
    
    
    Do
        sts = BTRV(com, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
        Select Case sts
            Case BtNoErr
            
                If StrConv(P_SSHIJI_K_REC.SHIJI_No, vbUnicode) <> Text1(ptxSHIJI_NO).text Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���i���w�}�[�ް�(�e)")
                Exit Function
        
        End Select
        
        Select Case StrConv(P_SSHIJI_K_REC.DATA_KBN, vbUnicode)
        
            Case P_KOSOU    '������
            
                k = k + 1
                K_Item_Tbl(k).JGYOBU = StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode)
                K_Item_Tbl(k).NAIGAI = StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode)
                            '�i��
                Text1(K_Index).text = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
                
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
'                If K_Item_Tbl(k).JGYOBU = SHIZAI Then
'                    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
'                Else
'                    Call UniCode_Conv(K0_ITEM.JGYOBU, K_Item_Tbl(k).JGYOBU)
'                End If


                Select Case K_Item_Tbl(k).JGYOBU
                    Case SHIZAI
                        Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)
                
                    Case SETSUBI
                        Call UniCode_Conv(K0_ITEM.JGYOBU, YUKO_JGYOBU)
                    Case Else
                        Call UniCode_Conv(K0_ITEM.JGYOBU, K_Item_Tbl(k).JGYOBU)
                End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
                Call UniCode_Conv(K0_ITEM.NAIGAI, K_Item_Tbl(k).NAIGAI)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(K_Index).text)
                
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        '�i��
                        Text1(K_Index + 1) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                        '�W���I��
                        If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                            Text1(K_Index + 4) = ""
                        Else
                            Text1(K_Index + 4) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_DAN, vbUnicode)
                        End If
                    
                    Case BtErrKeyNotFound
'>>>>>>>>>>>>>>>>>>>>>>>>>> �i�Ԗ��o�^�\���̑Ή�    2012.12.21
                        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(K_Index).text)
    
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            
                                '�i��
                                Text1(K_Index + 1) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                                '�W���I��
                                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                                    Text1(K_Index + 4) = ""
                                Else
                                    Text1(K_Index + 4) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                        StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                        StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                        StrConv(ITEMREC.ST_DAN, vbUnicode)
                                End If
                            
                            Case BtErrKeyNotFound
    
                                Text1(K_Index + 1) = "���o�^�i��"
                                Text1(K_Index + 4) = ""
                            Case Else
                                Call Input_UnLock             '2008.01.15
                                Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                                Exit Function
    
                        End Select
'                        Text1(K_Index + 1) = "���o�^�i��"
'                        Text1(K_Index + 4) = ""
'>>>>>>>>>>>>>>>>>>>>>>>>>> �i�Ԗ��o�^�\���̑Ή�    2012.12.21
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "")
                        Exit Function
                
                End Select
            
            
                Text1(K_Index + 2).text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0.00")
                Text1(K_Index + 3).text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0.00")
                            
                K_Index = K_Index + 5
            
            
            
            Case P_GAISOU   '�O������
                g = g + 1
                G_Item_Tbl(g).JGYOBU = StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode)
                G_Item_Tbl(g).NAIGAI = StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode)
                            '�i��
                Text1(G_Index).text = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
                
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
'                If G_Item_Tbl(g).JGYOBU = SHIZAI Then
'                    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
'                Else
'                    Call UniCode_Conv(K0_ITEM.JGYOBU, G_Item_Tbl(g).JGYOBU)
'                End If

                Select Case G_Item_Tbl(g).JGYOBU
                    Case SHIZAI
                        Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)
                    Case SETSUBI
                        Call UniCode_Conv(K0_ITEM.JGYOBU, YUKO_JGYOBU)
                    Case Else
                        Call UniCode_Conv(K0_ITEM.JGYOBU, G_Item_Tbl(g).JGYOBU)
                End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
                Call UniCode_Conv(K0_ITEM.NAIGAI, G_Item_Tbl(g).NAIGAI)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(G_Index).text)
                
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        '�i��
                        Text1(G_Index + 1) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                        '�W���I��
                        If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                            Text1(G_Index + 4) = ""
                        Else
                            Text1(G_Index + 4) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_DAN, vbUnicode)
                        End If
                    
                    Case BtErrKeyNotFound
'>>>>>>>>>>>>>>>>>>>>>>>>>> �i�Ԗ��o�^�\���̑Ή�    2012.12.21
                        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(G_Index).text)
    
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            
                                '�i��
                                Text1(G_Index + 1) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                                '�W���I��
                                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                                    Text1(G_Index + 4) = ""
                                Else
                                    Text1(G_Index + 4) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                        StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                        StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                        StrConv(ITEMREC.ST_DAN, vbUnicode)
                                End If
                            
                            Case BtErrKeyNotFound
    
                                Text1(G_Index + 1) = "���o�^�i��"
                                Text1(G_Index + 4) = ""
                            Case Else
                                Call Input_UnLock             '2008.01.15
                                Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                                Exit Function
    
                        End Select
'                        Text1(G_Index + 1) = "���o�^�i��"
'                        Text1(G_Index + 4) = ""
'>>>>>>>>>>>>>>>>>>>>>>>>>> �i�Ԗ��o�^�\���̑Ή�    2012.12.21
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "")
                        Exit Function
                
                End Select
            
            
                Text1(G_Index + 2).text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0.00")
                Text1(G_Index + 3).text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0.00")
                            
                G_Index = G_Index + 5
            
            
            Case P_DOUKON   '�����^�\��
            
                d = d + 1
' 2013.01.07 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                D_Item_Tbl(d).JGYOBU = StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode)
'                D_Item_Tbl(d).JGYOBU = BUZAI            '����/�\���̎��ƕ����u���ށv�Œ�ɕύX
' 2013.01.07 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                D_Item_Tbl(d).NAIGAI = StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode)
                            
                D_Item_Tbl(d).SYUBETSU = StrConv(P_SSHIJI_K_REC.KO_SYUBETSU, vbUnicode)
                D_Item_Tbl(d).HIN_GAI = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
                D_Item_Tbl(d).QTY = CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode))
                D_Item_Tbl(d).SHIJI_QTY = CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode))
                D_Item_Tbl(d).BIKOU = StrConv(P_SSHIJI_K_REC.KO_BIKOU, vbUnicode)
                            
                
                            
                If d < 6 Then
                            
                            
                                '���
                    Combo1(DC_Index).ListIndex = -1
                    For i = 0 To Combo1(DC_Index).ListCount - 1
                    
                        If StrConv(P_SSHIJI_K_REC.KO_SYUBETSU, vbUnicode) = Right(Combo1(DC_Index).List(i), 2) Then
                            Combo1(DC_Index).ListIndex = i
                            Exit For
                        End If
                    
                    Next i
                                
                                
                                '���ƕ� 2016.01.27
                    Combo2(DC_Index).ListIndex = -1
                    For i = 0 To Combo2(DC_Index).ListCount - 1
                    
                        If StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode) = Right(Combo2(DC_Index).List(i), 1) Then
                            Combo2(DC_Index).ListIndex = i
                            Exit For
                        End If
                    
                    Next i
                                
                                
                                
                    DC_Index = DC_Index + 1
                                
                                '�i��
                    Text1(DT_Index).text = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
                    
' 2013.01.07 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'                    If D_Item_Tbl(d).JGYOBU = SHIZAI Then
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)
'                    Else
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, D_Item_Tbl(d).JGYOBU)
'                    End If

'2016.01.27                    Select Case D_Item_Tbl(d).JGYOBU
'2016.01.27                        Case SHIZAI
'2016.01.27                            Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)
'2016.01.27                        Case SETSUBI
'2016.01.27                            Call UniCode_Conv(K0_ITEM.JGYOBU, YUKO_JGYOBU)
'2016.01.27                        Case Else
'2016.01.27                            Call UniCode_Conv(K0_ITEM.JGYOBU, D_Item_Tbl(d).JGYOBU)
'2016.01.27                    End Select
' 2013.01.07 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                    Call UniCode_Conv(K0_ITEM.NAIGAI, D_Item_Tbl(d).NAIGAI)
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(DT_Index).text)
                    
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                            '�i��
                            Text1(DT_Index + 1) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                            '�W���I��
                            If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                                Text1(DT_Index + 4) = ""
                            Else
                                Text1(DT_Index + 4) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                            End If
                        
                            '>>>>>>>>>>>>>>>>>>>    2013.01.07
                            'If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIZAI Then
                            '    wkJgyobu = BUZAI
                            'Else
                            '    'wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)  2012.04.04
                            '    wkJgyobu = YUKO_JGYOBU                          '2012.04.04
                            'End If
                            
                            Select Case StrConv(ITEMREC.JGYOBU, vbUnicode)
                                Case SHIZAI
                                    wkJgyobu = BUZAI
                                Case BUZAI
                                    wkJgyobu = YUKO_JGYOBU
                                Case Else
                                    wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)
                            End Select
                            '>>>>>>>>>>>>>>>>>>>    2013.01.07
                        
                            '2013.01.11 �W���I�Ԃ�ݒ�
                            If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, wkJgyobu, _
                                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                                                    (StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)), _
                                                                    , , Jyogai_Soko_umu) Then
                                Exit Function
                            
                            End If
                        
                            Text1(DT_Index + 5) = Format(Sumi_Qty + Mi_Qty, "#0")
                        
                        
                        Case BtErrKeyNotFound
                            
'>>>>>>>>>>>>>>>>>>>>>>>>>> �i�Ԗ��o�^�\���̑Ή�    2012.12.21
                            Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(DT_Index).text)
        
                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                            Select Case sts
                                Case BtNoErr
                                
                                    '�i��
                                    Text1(DT_Index + 1) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                                    '�W���I��
                                    If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                                        Text1(DT_Index + 4) = ""
                                    Else
                                        Text1(DT_Index + 4) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                            StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                            StrConv(ITEMREC.ST_DAN, vbUnicode)
                                    End If
                                
                                
                                
                                    Wk_LOC = StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                             StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)
        
                                    If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                                            StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                                                            Wk_LOC) Then
                                        Exit Function
        
                                    End If
                                
                                    Text1(DT_Index + 5) = Format(Sumi_Qty + Mi_Qty, "#0")
                                
                                Case BtErrKeyNotFound
                                    Text1(DT_Index + 1) = "���o�^�i��"
                                    Text1(DT_Index + 4) = ""
                                    Text1(DT_Index + 5) = ""
                                Case Else
                                    Call Input_UnLock             '2008.01.15
                                    Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                                    Exit Function
        
                            End Select
'                            Text1(DT_Index + 1) = "���o�^�i��"
'                            Text1(DT_Index + 4) = ""
'                            Text1(DT_Index + 5) = ""
'>>>>>>>>>>>>>>>>>>>>>>>>>> �i�Ԗ��o�^�\���̑Ή�    2012.12.21
                                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "")
                            Exit Function
                    
                    End Select
                
                
                    Text1(DT_Index + 2).text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0.00")
                    Text1(DT_Index + 3).text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0.00")
                    Text1(DT_Index + 6).text = Trim(StrConv(P_SSHIJI_K_REC.KO_BIKOU, vbUnicode))
                                
                    DT_Index = DT_Index + 7
                End If
        
        End Select
        
        
        
        
        com = BtOpGetNext
    
    Loop
    
    Item_Disp_Proc = False

End Function

Private Function Update_Proc(Mode As Integer, Optional HAKKO_F As Integer = 0, Optional MSG As Integer = 0) As Integer
'----------------------------------------------------------------------------
'                   �\���}�X�^�����i���w���ް��o��
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim ans         As Integer
Dim com         As Integer

Dim SEQNO       As Integer

Dim i           As Integer
Dim j           As Integer

Dim k           As Integer              '2012.03.09

Dim SHIJINO     As Long


Dim ORDER_DT    As String * 10          '2012.03.27

    Update_Proc = True
                                        
    Call Input_Lock
                                        
                                        '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If

    
    If Text1(ptxSHIJI_NO).text = "" Then
                                        
                                            '�Ǘ��t�@�C�����w�}�[�ԍ��̊l��
        Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)
        
        Do
            sts = BTRV(BtOpGetEqual + BtSNoWait, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound
                
                    If P_KANRI_MAKE_Proc() Then
                        GoTo Abort_Tran
                    End If
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_KANRI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Update_Proc = True
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�Ǘ��}�X�^")
                    GoTo Abort_Tran
            
            End Select
        
        
        Loop
    
        '�w�}�[���{�P
    
        If CLng(StrConv(P_KANRIREC.SASHIZU_NO, vbUnicode)) = 99999999 Then
            Call UniCode_Conv(P_KANRIREC.SASHIZU_NO, "00000001")
        Else
            Call UniCode_Conv(P_KANRIREC.SASHIZU_NO, Format(CLng(StrConv(P_KANRIREC.SASHIZU_NO, vbUnicode)) + 1, "00000000"))
        End If
    
    
        Do
            
            DoEvents
            
            sts = BTRV(BtOpUpdate, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_KANRI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        sts = BTRV(BtOpUnlock, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "�Ǘ��}�X�^")
                        End If
                        GoTo Abort_Tran
                    End If
                Case Else
                    Call File_Error(sts, BtOpUpdate, "�Ǘ��}�X�^")
                    GoTo Abort_Tran
            End Select
        Loop

        SHIJINO = CLng(StrConv(P_KANRIREC.SASHIZU_NO, vbUnicode))
    Else
        
        SHIJINO = CLng(Text1(ptxSHIJI_NO).text)
    
    End If
    '---------------------------------------------------    '���P�^�S���җL��͕i�ڃ}�X�^�X�V
        
'    If PRI_S_TANTO Then
        
        Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
        Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).text)
                
        
        Do
            sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound
                    MsgBox "�i�ڃ}�X�^�����[���ŕύX����Ă��܂��B�X�V�����𒆎~���܂��B"
                    GoTo Abort_Tran
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Update_Proc = True
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
                    GoTo Abort_Tran
            
            End Select
        
        
        Loop
                                                                                '���P�^�S���Ҹ׽
        Call UniCode_Conv(ITEMREC.S_TANTO, Right(Combo1(pcmbS_TANTO).text, 2))
    
        If IsNumeric(StrConv(ITEMREC.L_LABEL, vbUnicode)) Then
            G_Kisyu_F = CInt(StrConv(ITEMREC.L_LABEL, vbUnicode))
            
            If IsNumeric(StrConv(ITEMREC.L_URIKIN1, vbUnicode)) Then
                
                L_URIKIN1 = CDbl(StrConv(ITEMREC.L_URIKIN1, vbUnicode))
            Else
                L_URIKIN1 = 0
            End If
        
            If IsNumeric(StrConv(ITEMREC.L_URIKIN2, vbUnicode)) Then
                
                L_URIKIN2 = CDbl(StrConv(ITEMREC.L_URIKIN2, vbUnicode))
            Else
                L_URIKIN2 = 0
            End If
        
            If IsNumeric(StrConv(ITEMREC.L_URIKIN3, vbUnicode)) Then
                
                L_URIKIN3 = CDbl(StrConv(ITEMREC.L_URIKIN3, vbUnicode))
            Else
                L_URIKIN3 = 0
            End If
        
        
        Else
            G_Kisyu_F = 0
            L_URIKIN1 = 0
            L_URIKIN2 = 0
            L_URIKIN3 = 0
        
        End If
        Do
            
            DoEvents
            
            sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        sts = BTRV(BtOpUnlock, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "�i�ڃ}�X�^")
                        End If
                        GoTo Abort_Tran
                    End If
                Case Else
                    Call File_Error(sts, BtOpUpdate, "�i�ڃ}�X�^")
                    GoTo Abort_Tran
            End Select
        Loop
    
'    End If
    '---------------------------------------------------    '�\���}�X�^�X�V
        
        
        
If PI00015_LOG <> "" Then
    Call LOG_OUT(PI00015_LOG, "ITEM_CODE=" & Text1(ptxHIN_GAI).text & " DEL START" & " Mode =" & Mode)
End If
        
    '�Y���f�[�^�S���폜
    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).text)
       
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, "")
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
       
    com = BtOpGetGreater
       
    Do
        
        DoEvents
        
        Do
        
            sts = BTRV(com + BtSNoWait, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                
            Select Case sts
                Case BtNoErr
                
                    If StrConv(P_COMPO_O_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Or _
                        StrConv(P_COMPO_O_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                        StrConv(P_COMPO_O_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Or _
                        Trim(StrConv(P_COMPO_O_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxHIN_GAI).text) Then
                        sts = BTRV(BtOpUnlock, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "�\���}�X�^")
                            GoTo Abort_Tran
                        End If
                        sts = BtErrEOF
                    End If
                    Exit Do
                Case BtErrEOF
                    Exit Do
                
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_CMOPO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        GoTo Abort_Tran
                    End If
                
                
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "�\���}�X�^")
                    GoTo Abort_Tran
            End Select
    
        Loop
            
        If sts = BtErrEOF Then
If PI00015_LOG <> "" Then
    Call LOG_OUT(PI00015_LOG, "ITEM_CODE=" & Text1(ptxHIN_GAI).text & " DEL NORMAL END" & " Mode =" & Mode)
End If
            Exit Do
        End If


        Do
            sts = BTRV(BtOpDelete, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_CMOPO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        sts = BTRV(BtOpUnlock, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "�\���}�X�^")
                        End If
                        GoTo Abort_Tran
                    End If
                
                
                Case Else
                    Call File_Error(sts, BtOpDelete, "�\���}�X�^")
                    GoTo Abort_Tran
            End Select
        Loop
    
        com = BtOpGetNext
    
    Loop
        
    '�\���}�X�^(ͯ�ް)�o��
If PI00015_LOG <> "" Then
    Call LOG_OUT(PI00015_LOG, "ITEM_CODE=" & Text1(ptxHIN_GAI).text & " HEAD INSERT" & " Mode =" & Mode)
End If
                                                                                '�d�����溰��
    Call UniCode_Conv(P_COMPO_O_REC.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 1, 2))
                                                                                '���ƕ�
    Call UniCode_Conv(P_COMPO_O_REC.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1))
                                                                                '�����O
    Call UniCode_Conv(P_COMPO_O_REC.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1))
    Call UniCode_Conv(P_COMPO_O_REC.HIN_GAI, Text1(ptxHIN_GAI).text)
    Call UniCode_Conv(P_COMPO_O_REC.DATA_KBN, P_HEAD)
    Call UniCode_Conv(P_COMPO_O_REC.SEQNO, "000")

    Call UniCode_Conv(P_COMPO_O_REC.CLASS_CODE, Text1(ptxS_CLASS_CODE).text)    '�׽����
    Call UniCode_Conv(P_COMPO_O_REC.BIKOU, RichTextBox1(prchBIKOU).text)        '���l
    
    Call UniCode_Conv(P_COMPO_O_REC.F_CLASS_CODE, Text1(ptxF_CLASS_CODE).text)  '�t������
    
    Call UniCode_Conv(P_COMPO_O_REC.N_CLASS_CODE, Text1(ptxN_CLASS_CODE).text)  '���E����


    Call UniCode_Conv(P_COMPO_O_REC.FILLER, "")

    Call UniCode_Conv(P_COMPO_O_REC.UPD_TANTO, Text1(ptxTANTO_CODE))            '�X�V�S���Һ���
                                                                                '�X�V����
    Call UniCode_Conv(P_COMPO_O_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))


    Do
        
        DoEvents
        
        sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, BtOpInsert, "�\���}�X�^")
                GoTo Abort_Tran
        End Select
    
    Loop
If PI00015_LOG <> "" Then
    Call LOG_OUT(PI00015_LOG, "ITEM_CODE=" & Text1(ptxHIN_GAI).text & " HEAD INSERT" & " Mode =" & Mode)
End If

    '�\���}�X�^(���ި)�o��


    '�����ޕ�
    SEQNO = 0
    j = 0
    For i = ptxK_HIN_GAI01 To ptxK_HIN_GAI05 Step 5
    
        If Trim(Text1(i).text) = "" Then
        Else
                                                                                        
            SEQNO = SEQNO + 10
                                                                                        
                                                                                        '�d�����溰��
            Call UniCode_Conv(P_COMPO_K_REC.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 1, 2))
                                                                                        '���ƕ�
            Call UniCode_Conv(P_COMPO_K_REC.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1))
                                                                                        '�����O
            Call UniCode_Conv(P_COMPO_K_REC.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1))
            Call UniCode_Conv(P_COMPO_K_REC.HIN_GAI, Text1(ptxHIN_GAI).text)
            Call UniCode_Conv(P_COMPO_K_REC.DATA_KBN, P_KOSOU)                          '�f�[�^�敪
            Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(SEQNO, "000"))                '�ǔ�
                        
            Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, "")                            '���
            Call UniCode_Conv(P_COMPO_K_REC.KO_JGYOBU, K_Item_Tbl(j).JGYOBU)            '���ƕ�
            Call UniCode_Conv(P_COMPO_K_REC.KO_NAIGAI, K_Item_Tbl(j).NAIGAI)            '�����O
            Call UniCode_Conv(P_COMPO_K_REC.KO_HIN_GAI, Text1(i).text)                  '�i��
                                                                                        '����
            Call UniCode_Conv(P_COMPO_K_REC.KO_QTY, Format(CDbl(Text1(i + 2).text), "000.00"))
            Call UniCode_Conv(P_COMPO_K_REC.KO_BIKOU, "")                               '���l
        
            Call UniCode_Conv(P_COMPO_K_REC.FILLER, "")
        
            Call UniCode_Conv(P_COMPO_K_REC.UPD_TANTO, Text1(ptxTANTO_CODE).text)       '�X�V�S���Һ���
                                                                                        '�X�V����
            Call UniCode_Conv(P_COMPO_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
        
        
            Do
                
                DoEvents
                
                sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            GoTo Abort_Tran
                        End If
                    Case Else
                        Call File_Error(sts, BtOpInsert, "�\���}�X�^")
                        GoTo Abort_Tran
                End Select
            
            Loop
        
        
        End If
        
        j = j + 1
    
    
    Next i

    '�O�����ޕ�
    SEQNO = 0
    j = 0
    For i = ptxG_HIN_GAI01 To ptxG_HIN_GAI03 Step 5
    
        If Trim(Text1(i).text) = "" Then
        Else
            
            SEQNO = SEQNO + 10
                                                                                        
                                                                                        '�d�����溰��
            Call UniCode_Conv(P_COMPO_K_REC.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 1, 2))
                                                                                        '���ƕ�
            Call UniCode_Conv(P_COMPO_K_REC.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1))
                                                                                        '�����O
            Call UniCode_Conv(P_COMPO_K_REC.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1))
            Call UniCode_Conv(P_COMPO_K_REC.HIN_GAI, Text1(ptxHIN_GAI).text)
            Call UniCode_Conv(P_COMPO_K_REC.DATA_KBN, P_GAISOU)                         '�f�[�^�敪
            Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(SEQNO, "000"))                '�ǔ�
                        
            Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, "")                            '���
            Call UniCode_Conv(P_COMPO_K_REC.KO_JGYOBU, G_Item_Tbl(j).JGYOBU)            '���ƕ�
            Call UniCode_Conv(P_COMPO_K_REC.KO_NAIGAI, G_Item_Tbl(j).NAIGAI)            '�����O
            Call UniCode_Conv(P_COMPO_K_REC.KO_HIN_GAI, Text1(i).text)                  '�i��
                                                                                        '����
            Call UniCode_Conv(P_COMPO_K_REC.KO_QTY, Format(CDbl(Text1(i + 2).text), "000.00"))
            Call UniCode_Conv(P_COMPO_K_REC.KO_BIKOU, "")                               '���l
        
            Call UniCode_Conv(P_COMPO_K_REC.FILLER, "")
        
            Call UniCode_Conv(P_COMPO_K_REC.UPD_TANTO, Text1(ptxTANTO_CODE).text)       '�X�V�S���Һ���
                                                                                        '�X�V����
            Call UniCode_Conv(P_COMPO_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
        
            Do
                
                DoEvents
                
                sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            GoTo Abort_Tran
                        End If
                    Case Else
                        Call File_Error(sts, BtOpInsert, "�\���}�X�^")
                        GoTo Abort_Tran
                End Select
            
            Loop
        
        
        
        End If
        
        j = j + 1
    
    
    Next i


    '�����^�\����
If PI00015_LOG <> "" Then
    Call LOG_OUT(PI00015_LOG, "ITEM_CODE=" & Text1(ptxHIN_GAI).text & " BODY INSERT START" & " Mode =" & Mode)
End If
    SEQNO = 0
    For i = 0 To 49
    
        If D_Item_Tbl(i).JGYOBU = vbNullChar Or _
            Trim(D_Item_Tbl(i).JGYOBU) = "" Then
        Else
            SEQNO = SEQNO + 10
                                                                                        
                                                                                        '�d�����溰��
            Call UniCode_Conv(P_COMPO_K_REC.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 1, 2))
                                                                                        '���ƕ�
            Call UniCode_Conv(P_COMPO_K_REC.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1))
                                                                                        '�����O
            Call UniCode_Conv(P_COMPO_K_REC.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1))
            Call UniCode_Conv(P_COMPO_K_REC.HIN_GAI, Text1(ptxHIN_GAI).text)
            Call UniCode_Conv(P_COMPO_K_REC.DATA_KBN, P_DOUKON)                         '�f�[�^�敪
            Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(SEQNO, "000"))                '�ǔ�
                        
            Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, D_Item_Tbl(i).SYUBETSU)        '���
            Call UniCode_Conv(P_COMPO_K_REC.KO_JGYOBU, D_Item_Tbl(i).JGYOBU)            '���ƕ�
            Call UniCode_Conv(P_COMPO_K_REC.KO_NAIGAI, D_Item_Tbl(i).NAIGAI)            '�����O
            Call UniCode_Conv(P_COMPO_K_REC.KO_HIN_GAI, D_Item_Tbl(i).HIN_GAI)          '�i��
                                                                                        '����
            Call UniCode_Conv(P_COMPO_K_REC.KO_QTY, Format(D_Item_Tbl(i).QTY, "000.00"))
            Call UniCode_Conv(P_COMPO_K_REC.KO_BIKOU, D_Item_Tbl(i).BIKOU)              '���l
        
            Call UniCode_Conv(P_COMPO_K_REC.FILLER, "")
        
            Call UniCode_Conv(P_COMPO_K_REC.UPD_TANTO, Text1(ptxTANTO_CODE).text)       '�X�V�S���Һ���
                                                                                        '�X�V����
            Call UniCode_Conv(P_COMPO_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
        
        
            Do
                
                DoEvents
                
                sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            GoTo Abort_Tran
                        End If
                    Case Else
                        Call File_Error(sts, BtOpInsert, "�\���}�X�^")
                        GoTo Abort_Tran
                End Select
            
            Loop
        
If PI00015_LOG <> "" Then
    Call LOG_OUT(PI00015_LOG, "ITEM_CODE=" & Text1(ptxHIN_GAI).text & " BODY INSERT" & " Mode =" & Mode & "KO_ITEM_CODE = " & D_Item_Tbl(i).HIN_GAI)
End If
        
        
        
        End If
    
    Next i
If PI00015_LOG <> "" Then
    Call LOG_OUT(PI00015_LOG, "ITEM_CODE=" & Text1(ptxHIN_GAI).text & " BODY INSERT NORMAL END" & " Mode =" & Mode)
End If


    If Mode = 1 Then
        GoTo End_Tran
    End If
    
    '---------------------------------------------------    '�w�}�[�f�[�^�X�V
    
    '�w�}�[�f�[�^(ͯ�ް)����
    
    
    Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_No, Format(SHIJINO, "00000000"))
    
    Do
    
        sts = BTRV(BtOpGetEqual + BtSNoWait, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
            
        Select Case sts
            Case BtNoErr
            
                com = BtOpUpdate
                
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_SSHIJI_O.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    GoTo Abort_Tran
                End If
            
            
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "���i���w�}�[�ް�(�e)")
                GoTo Abort_Tran
        End Select

    Loop
    
    
    If com = BtOpInsert Then
        '�V�K�쐬
        Call UniCode_Conv(P_SSHIJI_O_REC.SHIJI_No, Format(SHIJINO, "00000000"))    '�w�}�[��
        
        Call UniCode_Conv(P_SSHIJI_O_REC.HAKKO_DT, "")
        
        
        Call UniCode_Conv(P_SSHIJI_O_REC.Print_datetime, "")                    '���s����
        
        Call UniCode_Conv(P_SSHIJI_O_REC.KAN_F, P_KAN_OFF)                      '����F
        Call UniCode_Conv(P_SSHIJI_O_REC.KAN_DT, "")                            '������
        Call UniCode_Conv(P_SSHIJI_O_REC.BUNNOU_CNT, "00")                      '���[��
        Call UniCode_Conv(P_SSHIJI_O_REC.UKEIRE_QTY, "00000000")                '�����
    
        For i = 0 To 9                                                          '�����Ǘ�
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(i).NIN, "0.0")           '�l��
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(i).TIMES, "000.00")      '����
        Next i
    
            
        Call UniCode_Conv(P_SSHIJI_O_REC.JISEKI_NAME, "")                       '���ӗv����
        Call UniCode_Conv(P_SSHIJI_O_REC.JISEKI_NIN, "0.0")                     '        �l
        Call UniCode_Conv(P_SSHIJI_O_REC.JISEKI_TIMES, "000.00")                '        ��
            
        Call UniCode_Conv(P_SSHIJI_O_REC.TASEKI_NAME, "")                       '���ӗv����
        Call UniCode_Conv(P_SSHIJI_O_REC.TASEKI_NIN, "0.0")                     '        �l
        Call UniCode_Conv(P_SSHIJI_O_REC.TASEKI_TIMES, "000.00")                '        ��
            
            
    
    
        Call UniCode_Conv(P_SSHIJI_O_REC.CANCEL_F, P_CANCEL_OFF)                '��ݾ��׸�
        Call UniCode_Conv(P_SSHIJI_O_REC.CANCEL_DATETIME, "")                   '��ݾٓ���
        
        Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_TANTO, "")                   '�i�������S���Һ��� 2013.08.21
        Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_DATETIME, "")                '�i����������       2013.08.21
        
        
        
        Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_LABEL_CNT, "000")            '�i���������ٌ���   2010.09.03
        Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_GENPIN_CNT, "000")           '�i���������ٌ���   2010.09.03
        
        Call UniCode_Conv(P_SSHIJI_O_REC.ORDER_DT_SEQ, "")                      '�󒍓�(������)�}�� 2012.03.27
        Call UniCode_Conv(P_SSHIJI_O_REC.COMPO_END_F, "")                       '�\����������F(���PC) 9:���� 2012.04.13
            
            
        
'        Call UniCode_Conv(P_SSHIJI_O_REC.FILLER, "")                           '2016.01.13
        Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_GAISOU_CNT, "")              '2016.01.13

    End If
                                                                                '���s��
    If HAKKO_F = 1 Then
        Call UniCode_Conv(P_SSHIJI_O_REC.HAKKO_DT, Format(Now, "YYYYMMDD"))
        Call UniCode_Conv(P_SSHIJI_O_REC.Print_datetime, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    End If
                                                                                '�S���Һ���
    Call UniCode_Conv(P_SSHIJI_O_REC.TANTO_CODE, Text1(ptxTANTO_CODE).text)
                                                                                '���F�Һ���
    Call UniCode_Conv(P_SSHIJI_O_REC.SHONIN_CODE, Text1(ptxSHONIN_CODE).text)
                                                                                '�d�����溰��
    Call UniCode_Conv(P_SSHIJI_O_REC.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 1, 2))
                                                                                '���ƕ�
    Call UniCode_Conv(P_SSHIJI_O_REC.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1))
                                                                                '�����O
    Call UniCode_Conv(P_SSHIJI_O_REC.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1))
                                                                                '�i��
    Call UniCode_Conv(P_SSHIJI_O_REC.HIN_GAI, Text1(ptxHIN_GAI).text)
                                                                                '����
    Call UniCode_Conv(P_SSHIJI_O_REC.SHIJI_QTY, Format(CDbl(Text1(ptxSHIJI_QTY).text), "00000000.00"))
                                                                                '��z��
    Call UniCode_Conv(P_SSHIJI_O_REC.UKEHARAI_CODE, Text1(ptxUKEHARAI_CODE).text)
                                                                                '�����敪
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, Text1(ptxUKEHARAI_CODE).text)
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrEOF
            MsgBox "��z���񂪑��ŕύX����܂����B�X�V�����𒆎~���܂��B"
            GoTo Abort_Tran
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�󕥐�}�X�^")
            Exit Function
    
    End Select
    Call UniCode_Conv(P_SSHIJI_O_REC.TORI_KBN, StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode))
                                                                                '���i���׽
    Call UniCode_Conv(P_SSHIJI_O_REC.S_CLASS_CODE, Text1(ptxS_CLASS_CODE).text)
                                                                                '�t���׽
    Call UniCode_Conv(P_SSHIJI_O_REC.F_CLASS_CODE, Text1(ptxF_CLASS_CODE).text)
                                                                                '���E�׽
    Call UniCode_Conv(P_SSHIJI_O_REC.N_CLASS_CODE, Text1(ptxN_CLASS_CODE).text)
                                                                                '���P�^�S���Ҹ׽
    Call UniCode_Conv(P_SSHIJI_O_REC.S_TANTO, Right(Combo1(pcmbS_TANTO).text, 2))
                                                                                                                                                        
    If Check1(pchkSAMPLE_F).Value = vbChecked Then                              '���{�쐬
        Call UniCode_Conv(P_SSHIJI_O_REC.SAMPLE_F, P_SAMPLE_F_ON)
    Else
        Call UniCode_Conv(P_SSHIJI_O_REC.SAMPLE_F, P_SAMPLE_F_OFF)
    End If

    If Option1(poptSHIJI_NORMAL).Value Then
        Call UniCode_Conv(P_SSHIJI_O_REC.SHIJI_F, P_SHIJI_F_NORMAL)             '�ʏ�
    Else
        If Option1(poptSHIJI_SPOT).Value Then
            Call UniCode_Conv(P_SSHIJI_O_REC.SHIJI_F, P_SHIJI_F_SPOT)           '��߯�
        Else
            If Option1(poptSHIJI_KEPPIN).Value Then
                Call UniCode_Conv(P_SSHIJI_O_REC.SHIJI_F, P_SHIJI_F_KEPPIN)     '���i����
            
            Else
                If Option1(poptSHIJI_SAIKON).Value Then
                    Call UniCode_Conv(P_SSHIJI_O_REC.SHIJI_F, P_SHIJI_F_SAIKON) '�č��� 2007.11.09
            
                End If
            End If
        End If
    End If
    

    If Check1(pchkPRI_SHIJI).Value = vbChecked Then                             '�o�͑Ώہ@�w�}�[
        Call UniCode_Conv(P_SSHIJI_O_REC.PRI_SHIJI, P_PRI_SHIJI_ON)
'''        Call UniCode_Conv(P_SSHIJI_O_REC.Print_datetime, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    Else
        Call UniCode_Conv(P_SSHIJI_O_REC.PRI_SHIJI, P_PRI_SHIJI_OFF)
    End If

    If Check1(pchkPRI_PARTS).Value = vbChecked Then                             '�o�͑Ώہ@�߰�����
        Call UniCode_Conv(P_SSHIJI_O_REC.PRI_PARTS, P_PRI_PARTS_ON)
    Else
        Call UniCode_Conv(P_SSHIJI_O_REC.PRI_PARTS, P_PRI_PARTS_OFF)
    End If

    If Check1(pchkPRI_GAISOU).Value = vbChecked Then                            '�o�͑Ώہ@�O������
        Call UniCode_Conv(P_SSHIJI_O_REC.PRI_GAISOU, P_PRI_GAISOU_ON)
    Else
        Call UniCode_Conv(P_SSHIJI_O_REC.PRI_GAISOU, P_PRI_GAISOU_OFF)
    End If

    If Check1(pchkPRI_KISHU).Value = vbChecked Then                             '�o�͑Ώہ@�O������
        Call UniCode_Conv(P_SSHIJI_O_REC.PRI_KISHU, P_PRI_KISHU_ON)
    Else
        Call UniCode_Conv(P_SSHIJI_O_REC.PRI_KISHU, P_PRI_KISHU_OFF)
    End If

    Call UniCode_Conv(P_SSHIJI_O_REC.BIKOU, RichTextBox1(prchBIKOU).text)       '���l
    
                                                                                '�X�V����
    Call UniCode_Conv(P_SSHIJI_O_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                                                                                
                                                                                
'-----------------------------------------------------------------------------  2012.03.18�@�󒍓�--��������
                                                                                '�󒍓�
    If Trim(Text1(ptxORDER_NO).text) = "" Then
        Call UniCode_Conv(P_SSHIJI_O_REC.ORDER_DT, "")
        Call UniCode_Conv(P_SSHIJI_O_REC.ORDER_DT_SEQ, "")
    Else
'2012.03.17        Call UniCode_Conv(P_SSHIJI_O_REC.ORDER_DT, Format(Text1(ptxORDER_DT).text, "YYYYMMDD"))
                
        ORDER_DT = Text1(ptxORDER_NO).text
        Call UniCode_Conv(P_SSHIJI_O_REC.ORDER_DT, Mid(ORDER_DT, 1, 8))
        Call UniCode_Conv(P_SSHIJI_O_REC.ORDER_DT_SEQ, Mid(ORDER_DT, 9, 2))
    End If
'-----------------------------------------------------------------------------  2012.03.18�@�󒍓�--��������
    
    Do
        
        DoEvents
        
        sts = BTRV(com, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    If com = BtOpUpdate Then
                        sts = BTRV(BtOpUnlock, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "���i���w�}�ް�(�e)")
                        End If
                    End If
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, com, "���i���w�}�ް�(�e)")
                GoTo Abort_Tran
        End Select
    
    Loop
    
    If com = BtOpUpdate Then
        
        
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.09
        
                
        
        For k = 0 To UBound(ZAIKO_FUSOKU)
        
            ZAIKO_FUSOKU(k).IDO_SUMI = ""
            ZAIKO_FUSOKU(k).HIKIATE_QTY = 0
            ZAIKO_FUSOKU(k).IDO_SUMI_QTY = 0        '2012.04.13
        
        Next k
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.09
        
        
        '�Ώۂ̎q���폜����
        Call UniCode_Conv(K0_P_SSHIJI_K.SHIJI_No, Format(SHIJINO, "00000000"))
        Call UniCode_Conv(K0_P_SSHIJI_K.DATA_KBN, "")
        Call UniCode_Conv(K0_P_SSHIJI_K.SEQNO, "")
    
        com = BtOpGetGreater
    
        Do
            
            DoEvents
            
            Do
            
                sts = BTRV(com + BtSNoWait, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
                    
                Select Case sts
                    Case BtNoErr
                    
                        If StrConv(P_SSHIJI_K_REC.SHIJI_No, vbUnicode) <> Format(SHIJINO, "00000000") Then
                            sts = BTRV(BtOpUnlock, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
                            If sts Then
                                Call File_Error(sts, BtOpUnlock, "���i���w�}�ް�(�q)")
                                GoTo Abort_Tran
                            End If
                            sts = BtErrEOF
                        End If
                        Exit Do
                    Case BtErrEOF
                        Exit Do
                    
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_SSHIJI_K.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            GoTo Abort_Tran
                        End If
                    
                    
                    Case Else
                        Call File_Error(sts, com + BtSNoWait, "���i���w�}�ް�(�q)")
                        GoTo Abort_Tran
                End Select
        
            Loop
                
            If sts = BtErrEOF Then
                Exit Do
            End If
    
    
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.09
            
            
            Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI, "")                  '2013.02.13
            Call UniCode_Conv(P_SSHIJI_K_REC.HIKIATE_QTY, "")               '2013.02.13
            Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI_QTY, "")              '2013.02.13
            
            
            
            If Trim(Text1(ptxORDER_NO).text) <> "" Then
                For k = 0 To UBound(ZAIKO_FUSOKU)
                
                
                    If ZAIKO_FUSOKU(k).JGYOBU = StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode) And _
                        ZAIKO_FUSOKU(k).NAIGAI = StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode) And _
                        ZAIKO_FUSOKU(k).HIN_GAI = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode) Then
                
                        ZAIKO_FUSOKU(k).IDO_SUMI = StrConv(P_SSHIJI_K_REC.IDO_SUMI, vbUnicode)
                        ZAIKO_FUSOKU(k).HIKIATE_QTY = Val(StrConv(P_SSHIJI_K_REC.HIKIATE_QTY, vbUnicode))
                        ZAIKO_FUSOKU(k).IDO_SUMI_QTY = Val(StrConv(P_SSHIJI_K_REC.IDO_SUMI_QTY, vbUnicode))     '2012.04.13
                        
                
                        Exit For
                
                    End If
                Next k
            End If
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.09
            Do
                sts = BTRV(BtOpDelete, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
                Select Case sts
                    Case BtNoErr
                    
                        Exit Do
                    Case BtErrEOF
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_CMOPO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            sts = BTRV(BtOpUnlock, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
                            If sts Then
                                Call File_Error(sts, BtOpUnlock, "���i���w�}�ް�(�q)")
                                GoTo Abort_Tran
                            End If
                            GoTo Abort_Tran
                        End If
                    
                    
                    Case Else
                        Call File_Error(sts, BtOpDelete, "���i���w�}�ް�(�q)")
                        GoTo Abort_Tran
                End Select
            Loop
        
            com = BtOpGetNext
        
        Loop
    End If
    
    
    '���i���w�}�[�ް�(���ި)�o��


    '�����ޕ�
    SEQNO = 0
    j = 0
    For i = ptxK_HIN_GAI01 To ptxK_HIN_GAI05 Step 5
    
        If Trim(Text1(i).text) = "" Then
        Else
            Call UniCode_Conv(P_SSHIJI_K_REC.SHIJI_No, Format(SHIJINO, "00000000"))        '�w�}�[��
                                                                                        
                                                                                        
                                                                                        
            SEQNO = SEQNO + 10
            Call UniCode_Conv(P_SSHIJI_K_REC.DATA_KBN, P_KOSOU)                         '�f�[�^�敪
            
            Call UniCode_Conv(P_SSHIJI_K_REC.SEQNO, Format(SEQNO, "000"))               '�ǔ�
                        
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_SYUBETSU, "")                           '���
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_JGYOBU, K_Item_Tbl(j).JGYOBU)           '���ƕ�
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_NAIGAI, K_Item_Tbl(j).NAIGAI)           '�����O
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_HIN_GAI, Text1(i).text)                 '�i��
                                                                                        '����
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_QTY, Format(CDbl(Text1(i + 2).text), "000.00"))
                                                                                        '����
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_SHIJI_QTY, Format(CDbl(Text1(i + 3).text), "000000000.00"))
            
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_BIKOU, "")                              '���l
        
            Call UniCode_Conv(P_SSHIJI_K_REC.CALCEL_F, P_CANCEL_OFF)                    '��ݾ��׸�
            Call UniCode_Conv(P_SSHIJI_K_REC.CANCEL_DATETIME, "")                       '��ݾٓ���
        
        
        
 '--------------------------------------------------- ���  ���ޑΉ��@2012.03.18
            If K_Item_Tbl(j).JGYOBU = SHIZAI Then           '2013.03.31
            
            
                Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)    '2013.03.31
            Else                                            '2013.03.31
                Call UniCode_Conv(K0_ITEM.JGYOBU, K_Item_Tbl(j).JGYOBU)
            
            End If                                          '2013.03.31
            
            
            Call UniCode_Conv(K0_ITEM.NAIGAI, K_Item_Tbl(j).NAIGAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(i).text)
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Call UniCode_Conv(P_SSHIJI_K_REC.ST_TANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_DAN, vbUnicode))
                Case BtErrKeyNotFound
                    Call UniCode_Conv(P_SSHIJI_K_REC.ST_TANABAN, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                    GoTo Abort_Tran
            End Select
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.18
       
        
        
        
            Call UniCode_Conv(P_SSHIJI_K_REC.FILLER, "")
                                                                                        '�X�V����
            Call UniCode_Conv(P_SSHIJI_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
        
                                                                                        '�o�ח\��h�c
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_ID_NO, "")
        
        
                    
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.09
            Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI, "")                  '2013.02.13
            Call UniCode_Conv(P_SSHIJI_K_REC.HIKIATE_QTY, "")               '2013.02.13
            Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI_QTY, "")              '2013.02.13
            
            For k = 0 To UBound(ZAIKO_FUSOKU)
            
            
                If ZAIKO_FUSOKU(k).JGYOBU = StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode) And _
                    ZAIKO_FUSOKU(k).NAIGAI = StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode) And _
                    ZAIKO_FUSOKU(k).HIN_GAI = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode) Then
            
                    Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI, ZAIKO_FUSOKU(k).IDO_SUMI)
                    Call UniCode_Conv(P_SSHIJI_K_REC.HIKIATE_QTY, Format(ZAIKO_FUSOKU(k).HIKIATE_QTY, "00000000.00"))
                    
                    Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI_QTY, Format(ZAIKO_FUSOKU(k).IDO_SUMI_QTY, "00000000.00")) '2012.04.13
            
                    Exit For
            
                End If
            Next k
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.09
        
        
        
        
            Do
                
                DoEvents
                
                sts = BTRV(BtOpInsert, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_SSHIJI_K.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            GoTo Abort_Tran
                        End If
                    Case Else
                        Call File_Error(sts, BtOpInsert, "���i���w�}�ް�(�q)")
                        GoTo Abort_Tran
                End Select
            
            Loop
        
        
        End If
        
        j = j + 1
    
    
    Next i

    '�O�����ޕ�
    SEQNO = 0
    j = 0
    For i = ptxG_HIN_GAI01 To ptxG_HIN_GAI03 Step 5
    
        If Trim(Text1(i).text) = "" Then
        Else
            
            Call UniCode_Conv(P_SSHIJI_K_REC.SHIJI_No, Format(SHIJINO, "00000000"))        '�w�}�[��
            SEQNO = SEQNO + 10
            Call UniCode_Conv(P_SSHIJI_K_REC.DATA_KBN, P_GAISOU)                        '�f�[�^�敪
            Call UniCode_Conv(P_SSHIJI_K_REC.SEQNO, Format(SEQNO, "000"))               '�ǔ�
                        
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_SYUBETSU, "")                           '���
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_JGYOBU, G_Item_Tbl(j).JGYOBU)           '���ƕ�
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_NAIGAI, G_Item_Tbl(j).NAIGAI)           '�����O
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_HIN_GAI, Text1(i).text)                 '�i��
                                                                                        '����
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_QTY, Format(CDbl(Text1(i + 2).text), "000.00"))
                                                                                        '����
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_SHIJI_QTY, Format(CDbl(Text1(i + 3).text), "00000000.00"))
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_BIKOU, "")                               '���l
            
                        
            
            Call UniCode_Conv(P_SSHIJI_K_REC.CALCEL_F, P_CANCEL_OFF)                    '��ݾ��׸�
            Call UniCode_Conv(P_SSHIJI_K_REC.CANCEL_DATETIME, "")                       '��ݾٓ���
        
        
        
 '--------------------------------------------------- ���  ���ޑΉ��@2012.03.18
            
            
            If G_Item_Tbl(j).JGYOBU = SHIZAI Then                       '2013.03.31
                Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)                '2013.03.31
            Else                                                        '2013.03.31
                Call UniCode_Conv(K0_ITEM.JGYOBU, G_Item_Tbl(j).JGYOBU)
            End If                                                      '2013.03.31
            
            Call UniCode_Conv(K0_ITEM.NAIGAI, G_Item_Tbl(j).NAIGAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(i).text)
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Call UniCode_Conv(P_SSHIJI_K_REC.ST_TANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_DAN, vbUnicode))
                Case BtErrKeyNotFound
                    Call UniCode_Conv(P_SSHIJI_K_REC.ST_TANABAN, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                    GoTo Abort_Tran
            End Select
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.18
        
        
        
            Call UniCode_Conv(P_SSHIJI_K_REC.FILLER, "")
                                                                                        '�X�V����
            Call UniCode_Conv(P_SSHIJI_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
        
        
                                                                                        '�o�ח\��h�c
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_ID_NO, "")
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.09
            For k = 0 To UBound(ZAIKO_FUSOKU)
            
            
                If ZAIKO_FUSOKU(k).JGYOBU = StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode) And _
                    ZAIKO_FUSOKU(k).NAIGAI = StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode) And _
                    ZAIKO_FUSOKU(k).HIN_GAI = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode) Then
            
                    Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI, ZAIKO_FUSOKU(k).IDO_SUMI)
                    Call UniCode_Conv(P_SSHIJI_K_REC.HIKIATE_QTY, Format(ZAIKO_FUSOKU(k).HIKIATE_QTY, "00000000.00"))
                    
                    Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI_QTY, Format(ZAIKO_FUSOKU(k).IDO_SUMI_QTY, "00000000.00")) '2012.04.13
            
                    Exit For
            
                End If
            Next k
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.09
        
            Do
                
                DoEvents
                
                sts = BTRV(BtOpInsert, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_SSHIJI_K.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            GoTo Abort_Tran
                        End If
                    Case Else
                        Call File_Error(sts, BtOpInsert, "���i���w�}�ް�(�q)")
                        GoTo Abort_Tran
                End Select
            
            Loop
        
        End If
        
        j = j + 1
    
    
    Next i


    '�����^�\����
    SEQNO = 0
    For i = 0 To 49
    
        If D_Item_Tbl(i).JGYOBU = vbNullChar Or Trim(D_Item_Tbl(i).JGYOBU) = "" Then
        Else
            Call UniCode_Conv(P_SSHIJI_K_REC.SHIJI_No, Format(SHIJINO, "00000000"))        '�w�}�[��
            
            SEQNO = SEQNO + 10
                                                                                        
            Call UniCode_Conv(P_SSHIJI_K_REC.DATA_KBN, P_DOUKON)                        '�f�[�^�敪
            Call UniCode_Conv(P_SSHIJI_K_REC.SEQNO, Format(SEQNO, "000"))               '�ǔ�
                        
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_SYUBETSU, D_Item_Tbl(i).SYUBETSU)       '���
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_JGYOBU, D_Item_Tbl(i).JGYOBU)           '���ƕ�
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_NAIGAI, D_Item_Tbl(i).NAIGAI)           '�����O
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_HIN_GAI, D_Item_Tbl(i).HIN_GAI)         '�i��
                                                                                        '����
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_QTY, Format(D_Item_Tbl(i).QTY, "000.00"))
                                                                                        '����
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_SHIJI_QTY, Format(D_Item_Tbl(i).SHIJI_QTY, "00000000.00"))
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_BIKOU, D_Item_Tbl(i).BIKOU)             '���l
        
            Call UniCode_Conv(P_SSHIJI_K_REC.CALCEL_F, P_CANCEL_OFF)                    '��ݾ��׸�
            Call UniCode_Conv(P_SSHIJI_K_REC.CANCEL_DATETIME, "")                       '��ݾٓ���
            
            
 '--------------------------------------------------- ���  ���ޑΉ��@2012.03.18
            If D_Item_Tbl(i).JGYOBU = SHIZAI Then                       '2013.03.31
                Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)                '2013.03.31
            Else                                                        '2013.03.31
                Call UniCode_Conv(K0_ITEM.JGYOBU, D_Item_Tbl(i).JGYOBU)
            End If                                                      '2013.03.31
            Call UniCode_Conv(K0_ITEM.NAIGAI, D_Item_Tbl(i).NAIGAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, D_Item_Tbl(i).HIN_GAI)
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Call UniCode_Conv(P_SSHIJI_K_REC.ST_TANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_DAN, vbUnicode))
                Case BtErrKeyNotFound
                    Call UniCode_Conv(P_SSHIJI_K_REC.ST_TANABAN, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                    GoTo Abort_Tran
            End Select
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.18
            
            
            Call UniCode_Conv(P_SSHIJI_K_REC.FILLER, "")
                                                                                        '�X�V����
            Call UniCode_Conv(P_SSHIJI_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
        
        
            If POS_UMU Then
                '�o�׎w���̍쐬
'''2007.03.08                If Y_SYUKA_Make_Proc(i) Then
'''2007.03.08                    GoTo Abort_Tran
'''2007.03.08                End If
            End If
        
                                                                                        '�o�ח\��h�c
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_ID_NO, D_Item_Tbl(i).ID_NO)
        
        
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.09
            Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI, "")                  '2013.02.13
            Call UniCode_Conv(P_SSHIJI_K_REC.HIKIATE_QTY, "")               '2013.02.13
            Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI_QTY, "")              '2013.02.13
            
            
            If Trim(Text1(ptxORDER_NO).text) <> "" Then
                For k = 0 To UBound(ZAIKO_FUSOKU)
                
                
                    If ZAIKO_FUSOKU(k).JGYOBU = StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode) And _
                        ZAIKO_FUSOKU(k).NAIGAI = StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode) And _
                        ZAIKO_FUSOKU(k).HIN_GAI = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode) Then
                
                        Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI, ZAIKO_FUSOKU(k).IDO_SUMI)
                        Call UniCode_Conv(P_SSHIJI_K_REC.HIKIATE_QTY, Format(ZAIKO_FUSOKU(k).HIKIATE_QTY, "00000000.00"))
                        
                        Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI_QTY, Format(ZAIKO_FUSOKU(k).IDO_SUMI_QTY, "00000000.00")) '2012.04.13
                
                        Exit For
                
                    End If
                Next k
            End If
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.09
        
        
        
            Do
                
                DoEvents
                
                sts = BTRV(BtOpInsert, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_SSHIJI_K.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            GoTo Abort_Tran
                        End If
                    Case Else
                        Call File_Error(sts, BtOpInsert, "���i���w�}�ް�(�q)")
                        GoTo Abort_Tran
                End Select
            
            Loop
        
        End If
    
    Next i
    
    
'--------------------------------------------------- ���  ���ޑΉ��@2012.04.13
    'If Trim(Text1(ptxORDER_NO).text) = "" Then         '2013.05.22 DEL
    If Trim(Text1(ptxORDER_NO).text) <> "" Then         '2013.05.22 INS
        Call UniCode_Conv(K6_ODR_ORDER.ORDER_NO, Text1(ptxORDER_NO).text)

        sts = BTRV(BtOpGetEqual, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K6_ODR_ORDER, Len(K6_ODR_ORDER), 6)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Beep
                MsgBox "�����f�[�^���ύX����Ă��܂��B����������ʂŊm�F���Ă��������B"
                GoTo Abort_Tran
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�����f�[�^")
                GoTo Abort_Tran
        End Select

        Call UniCode_Conv(ODR_ORDER_REC.PRT_FLG, "F")
        sts = BTRV(BtOpUpdate, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K6_ODR_ORDER, Len(K6_ODR_ORDER), 6)
        Select Case sts
            Case BtNoErr
            Case Else
                Call File_Error(sts, BtOpUpdate, "�����f�[�^")
                GoTo Abort_Tran
        End Select
    End If
'--------------------------------------------------- ���  ���ޑΉ��@2012.04.13



End_Tran:
                                        '�g�����U�N�V�����I��
    sts = BTRV(BtOpEndTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
'2007.11.21    If Mode = 0 Then
    If MSG = 0 Then     '2007.11.21
        If Text1(ptxSHIJI_NO).text = "" Then
            MsgBox "�w�}�[���F" & Format(SHIJINO, "00000000") & "���쐬���܂����B"
        End If
    End If
    
    Call Input_UnLock
                                        '����ɑΏێw�}�[����ʒm
    Taget_Key = Format(SHIJINO, "00000000")
    
    Update_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If
    
    Call Input_UnLock

End Function

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If

    Select Case Index
        Case pcmbSHIMUKE        '�d������


        Case pcmbUKEHARAI       '��z��
            Text1(ptxUKEHARAI_CODE).text = Trim(Right(Combo1(pcmbUKEHARAI).text, 5))

        Case pcmbS_TANTO        '���P�^�S����

                                '�����^�\���@���
        Case pcmbD_SYUBETSU01, pcmbD_SYUBETSU02, pcmbD_SYUBETSU03, pcmbD_SYUBETSU04, pcmbD_SYUBETSU05, pcmbD_SYUBETSU06

            D_Item_Tbl(Index - pcmbD_SYUBETSU01).SYUBETSU = Right(Combo1(Index).text, 2)

    End Select

    Call Tab_Ctrl(Shift)        '�ړ�

End Sub

Private Sub Combo1_LostFocus(Index As Integer)

    Select Case Index
        Case pcmbSHIMUKE        '�d������

        Case pcmbUKEHARAI       '��z��
            Text1(ptxUKEHARAI_CODE).text = Trim(Right(Combo1(pcmbUKEHARAI).text, 5))

        Case pcmbS_TANTO        '���P�^�S����

                                '�����^�\���@���
        Case pcmbD_SYUBETSU01, pcmbD_SYUBETSU02, pcmbD_SYUBETSU03, pcmbD_SYUBETSU04, pcmbD_SYUBETSU05, pcmbD_SYUBETSU06

            D_Item_Tbl(Index - pcmbD_SYUBETSU01).SYUBETSU = Right(Combo1(Index).text, 2)

    End Select

End Sub

Private Sub Combo2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Call Tab_Ctrl(Shift)        '�ړ�


End Sub

Private Sub Command1_Click(Index As Integer)

Dim ans             As Integer
Dim i               As Integer

Dim rpt             As New PI00015F1
Dim f               As New PI000153

Dim com             As Integer
Dim sts             As Integer


Dim Parts_F         As Integer
Dim Gaisou_F        As Integer
Dim Kishu_F         As Integer

Dim objAccess       As Access.Application
Dim strAccessPath   As String

Dim GAISOU_QTY          As Long
Dim GAISOU_SHIJI_QYU    As Long

Dim L_print_Flg     As Boolean

'--------------------------------------------------- ���  ���ޑΉ��@2012.03.08
Dim Order_QTY       As Long
Dim SHIJI_QTY       As Long

Dim ZAIKO_F         As Boolean
Dim wkMSG           As String
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.08


    Select Case Index
        Case P_CMD_Upd        '�X�V
            
            
            For i = ptxSHIJI_NO To ptxD_BIKOU06
            
                If Error_Check_Proc(i, 0, 1) Then   '�G���[�`�F�b�N
                    Exit Sub
                End If
            
            Next i
            
            
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.08
'            If ORDER_Check_Proc(Order_QTY, SHIJI_QTY) Then
'                Unload Me
'            End If
'            If Order_QTY < SHIJI_QTY Then
'
'                wkMSG = "�e�i�ԁ@�������ƈقȂ�܂��B�������p�����܂����H" & Chr(13) & Chr(10)
'                wkMSG = wkMSG & "�e�i�ԁ@������:" & Format(Order_QTY, "#0") & Chr(13) & Chr(10)
'                wkMSG = wkMSG & "�@�@�@�@�@ �w����:" & Format(SHIJI_QTY, "#0")
'
'
'                ans = MsgBox(wkMSG, vbYesNo + vbDefaultButton2, "�m�F����")
'                If ans = vbNo Then
'                    Exit Sub
'                End If
'            End If

            
            
            If Trim(Text1(ptxORDER_NO).text) <> "" Then
                If Zaiko_Check_Proc(ZAIKO_F) Then
                    Unload Me
                End If
    
    
                If ZAIKO_F Then
                    wkMSG = "�݌ɕs�����������Ă��܂��B�������p�����܂����H" & Chr(13) & Chr(10)
                    For i = 0 To UBound(ZAIKO_FUSOKU)
                        If ZAIKO_FUSOKU(i).SAI_QTY < 0 Then
                            wkMSG = wkMSG & RTrim(ZAIKO_FUSOKU(i).HIN_GAI) & Chr(13) & Chr(10)
                        End If
                    Next i
                    ans = MsgBox(wkMSG, vbYesNo + vbDefaultButton2, "�m�F����")
                    If ans = vbNo Then
                        Exit Sub
                    End If
                End If
            End If
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.08
            
            Beep
            ans = MsgBox("�X�V���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Update_Proc(0, , 0) Then  '2007.11.21 �����ύX
                    Unload Me
                End If
                
                If Init_Proc() Then
                    Unload Me
                End If
            
                Text1(ptxSHIJI_NO).SetFocus
            
            
            Else
                Text1(ptxORDER_NO).SetFocus
            End If


'        Case P_CMD_DEL                      '�폜
        Case cmdMUPDATE                     'Ͻ��X�V
        
            For i = ptxSHIJI_NO To ptxD_BIKOU06
            
                If Error_Check_Proc(i, 1, 1) Then   '�G���[�`�F�b�N
                    Exit Sub
                End If
            
            Next i
            
            Beep
            ans = MsgBox("�X�V���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Update_Proc(1, , 1) Then  '2007.11.21�����ύX
                    Unload Me
                End If
                
'                If Init_Proc() Then
'                    Unload Me
'                End If
            
                Call UniCode_Conv(ITEMREC.JGYOBU, "")
                Call UniCode_Conv(P_COMPO_O_REC.JGYOBU, "")
            
            
                If Text1(ptxSHIJI_NO).Locked Then
                    Text1(ptxORDER_NO).SetFocus
                Else
                    Text1(ptxSHIJI_NO).SetFocus
                End If
            
            Else
                Text1(ptxORDER_NO).SetFocus
            End If
        
        
        Case P_CMD_DSP                      '����/�\��
        Case cmdNext                        '�\�����i��ʂ�
        
            Doukon_Start = 1
            PI000152.Show vbModal           '���i�ڍ׃t�H�[���\��
            If G_SCREEN_FLG = SYS_ERR Then
                Unload Me
            End If
        
            'ð��ق��\���^������\��
            If Tbl_To_Disp_Proc() Then
                Unload Me
            End If
        
        
        
        Case P_CMD_OUT                      '�ް��o��
        Case P_CMD_PRT                      '���
            
            
            
            For i = ptxSHIJI_NO To ptxD_BIKOU06
            
                If Error_Check_Proc(i, 0, 1) Then   '�G���[�`�F�b�N
                    Exit Sub
                End If
            
            Next i
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.08
'            If ORDER_Check_Proc(Order_QTY, SHIJI_QTY) Then
'                Unload Me
'            End If
'            If Order_QTY < SHIJI_QTY Then
'
'                wkMSG = "�e�i�ԁ@�������ƈقȂ�܂��B�������p�����܂����H" & Chr(13) & Chr(10)
'                wkMSG = wkMSG & "�e�i�ԁ@������:" & Format(Order_QTY, "#0") & Chr(13) & Chr(10)
'                wkMSG = wkMSG & "�@�@�@�@�@ �w����:" & Format(SHIJI_QTY, "#0")
'
'
'                ans = MsgBox(wkMSG, vbYesNo + vbDefaultButton2, "�m�F����")
'                If ans = vbNo Then
'                    Exit Sub
'                End If
'            End If

            
            
            If Trim(Text1(ptxORDER_NO).text) <> "" Then
                If Zaiko_Check_Proc(ZAIKO_F) Then
                    Unload Me
                End If
    
    
                If ZAIKO_F Then
                    wkMSG = "�݌ɕs�����������Ă��܂��B�������p�����܂����H" & Chr(13) & Chr(10)
                    For i = 0 To UBound(ZAIKO_FUSOKU)
                        If ZAIKO_FUSOKU(i).SAI_QTY < 0 Then
                            wkMSG = wkMSG & RTrim(ZAIKO_FUSOKU(i).HIN_GAI) & Chr(13) & Chr(10)
                        End If
                    Next i
                    ans = MsgBox(wkMSG, vbYesNo + vbDefaultButton2, "�m�F����")
                    If ans = vbNo Then
                        Exit Sub
                    End If
                End If
            Else                                    '2012.12.20
                Erase ZAIKO_FUSOKU                  '2012.12.20
                ReDim ZAIKO_FUSOKU(0 To 0)          '2012.12.20
                ZAIKO_FUSOKU(0).JGYOBU = ""         '2012.12.20
                ZAIKO_FUSOKU(0).NAIGAI = ""         '2012.12.20
                ZAIKO_FUSOKU(0).HIN_GAI = ""        '2012.12.20
                
            
            End If
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.08
            
            Beep
            ans = MsgBox("����^�X�V���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Update_Proc(0, 1, 1) Then
                    Unload Me
                End If
                
                
                If Check1(pchkPRI_SHIJI).Value = vbChecked Then
                
                    Set rpt = New PI00015F1
                
                    '���|�[�g��������܂��B�itrue�F����_�C�A���O���� false�F�Ȃ��j
                    rpt.PrintReport False
                
                    Set rpt = Nothing


'                    f.RunReport rpt
'                    f.Show
                
                End If
                
                
                
                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2013.11.20

                '���ټ��ш���v��




                If Check1(pchkPRI_PARTS).Value = vbChecked Or _
                    Check1(pchkPRI_GAISOU).Value = vbChecked Then

                    L_print_Flg = True

                    
'>>>>>>>>>>>>>>>>   2016.01.13
'                    If L_URIKIN1 = 0 And L_URIKIN2 = 0 And L_URIKIN3 = 0 Then
'
'                        Beep
'                        ans = MsgBox("�P�����ݒ�ł��B���x��������܂����H", vbYesNo + vbQuestion, "�m�F����")
'                        If ans = vbYes Then
'                        Else
'                            L_print_Flg = False
'                        End If
'                    Else
'                    End If
'>>>>>>>>>>>>>>>>   2016.01.13
                    
                    If L_print_Flg Then


                        On Error Resume Next
                        Set objAccess = GetObject(, "Access.Application")
                        If Err().Number <> 0 Then
    '2016.01.13                        MsgBox "���̒[���ł͏��i���x�����s�͍s���܂���B"
    '                        MsgBox "GetObject(Access.Application)" & Err().Number & " " & Err().Description
                        Else
    '                        MsgBox Err.Number

                            strAccessPath = App.Path
                            If Right(strAccessPath, 1) <> "\" Then
                                strAccessPath = strAccessPath & "\"
                            End If

                            strAccessPath = strAccessPath & "litem.mdb"
                            Set objAccess = GetObject(strAccessPath)



                            If Check1(pchkPRI_PARTS).Value = vbChecked Then
                                Parts_F = 1
                            Else
                                Parts_F = 0
                            End If


                            If Check1(pchkPRI_GAISOU).Value = vbChecked Then
                                Gaisou_F = 1
                            Else
                                Gaisou_F = 0
                            End If

                            If G_Kisyu_F = L_LABEL_ON Then
                                Kishu_F = 1
                                Parts_F = 0
                            Else
                                Kishu_F = 0
                            End If

                            If IsNumeric(Text1(ptxG_QTY01).text) Then
                                GAISOU_QTY = CLng(Text1(ptxG_QTY01).text)
                            Else
                                GAISOU_QTY = 0
                            End If
                            If IsNumeric(Text1(ptxG_SHIJI_QTY01).text) Then
                                GAISOU_SHIJI_QYU = CLng(Text1(ptxG_SHIJI_QTY01).text)
                            Else
                                GAISOU_SHIJI_QYU = 0
                            End If

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

                            Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
                            Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
                            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).text)


                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                            Select Case sts
                                Case BtNoErr


                                    Call UniCode_Conv(ITEMREC.L_IRI_QTY, Format(GAISOU_QTY, "00000000"))



                                    '�č���ϰ��ǉ�  2007.11.09

                                    If Option1(poptSHIJI_SAIKON).Value Then
                                        Call UniCode_Conv(ITEMREC.L_MARK, SAIKON_F)

                                    End If





                                    sts = BTRV(BtOpInsert, L_ITEM_POS, ITEMREC, Len(ITEMREC), K0_L_ITEM, Len(K0_L_ITEM), 0)

                                    Select Case sts
                                        Case BtNoErr

                                            objAccess.Run "PosPrintLabel", Trim(Text1(ptxHIN_GAI).text), CLng(Text1(ptxSHIJI_QTY).text), Parts_F, Gaisou_F, Kishu_F, GAISOU_QTY, GAISOU_SHIJI_QYU, 0



                                        Case Else
                                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                            Exit Sub


                                    End Select

                                Case BtErrKeyNotFound

                                Case Else
                                   Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                    Exit Sub

                            End Select





                            Set objAccess = Nothing
                        End If



                    End If
                End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2013.11.20
                
                If Init_Proc() Then
                    Unload Me
                End If
            
'2007.11.21                Text1(ptxSHIJI_NO).SetFocus
                
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.18
'                Text1(ptxHIN_GAI).SetFocus  '2007.11.21
                Text1(ptxORDER_NO).SetFocus
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.18
            Else
                Text1(ptxORDER_NO).SetFocus
            End If
            
            
            
        Case cmdCen                         '������
            If Init_Proc() Then
                Unload Me
            End If
            Text1(ptxSHIJI_NO).SetFocus
        Case P_CMD_End                      '�I��
            Unload Me
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
            Command1(KeyCode - vbKeyF1).Value = True
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()

Dim c           As String * 128
Dim sts         As Integer
Dim i           As Integer

Dim MUKE_CODE   As Variant

Dim YOMI_JGYOBU_w As Variant    '2014.03.24
Dim j           As Integer      '2014.03.24


    If App.PrevInstance Then
        MsgBox "����v���O�������s���ł��B"
        End
    End If
                                
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
        End
    End If
    LOG_F = RTrim(c)
                                '�o�׃��O�t�@�C������荞��
    If SYUKA_LOGF_GET_PROC() Then
        Beep
        MsgBox "�o�׃��O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
        End
    End If
                                

    PI000151.Caption = PI000151.Caption & LAST_UPDATE_DAY       '2017.10.17

                                
                                
                                
                                
                                
                                
                                
                                '�����p���O�t�@�C������荞��   2016.03.30
    If GetIni(App.EXEName, "PI00015_LOG", App.EXEName, c) Then
        PI00015_LOG = ""
    Else
        PI00015_LOG = Trim(c)
    End If
                                
                                
                                
                                
                                '���ƕ��̊l��       2016.01.27
    If JGYOB_TB_Set() Then
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B"
        End
    End If
                                
                                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    P_SYS.INI�@--���@PI00015.INI    2016.01.13
                                
                                
                                '��z���荞��
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "TEHAI", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "TEHAI", App.EXEName, c) Then
    Else
        TEHAI = RTrim(c)
    End If
                                
                                'POS���їL���̎�荞��
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "POS_UMU", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "POS_UMU", App.EXEName, c) Then
        POS_UMU = False
    Else
        If RTrim(c) = "0" Then
            POS_UMU = False
        Else
            POS_UMU = True
        End If
    End If
                                '�ް���ވ�
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "BCR", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "BCR", App.EXEName, c) Then
        PRI_MAIN_BCR = False
    Else
        If RTrim(c) = "0" Then
            PRI_MAIN_BCR = False
        Else
            If Not POS_UMU Then
                MsgBox "�o�n�r���т����ݒ�ł��B�����𒆎~���܂��B"
                End
            End If
            PRI_MAIN_BCR = True
        End If
    End If
                                    '���ה��l�󎚓��e
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "DET_BIKOU", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "DET_BIKOU", App.EXEName, c) Then
        PRI_BIKOU_BCR = 0
    Else
        
        If Not IsNumeric(RTrim(c)) Then
            PRI_BIKOU_BCR = 0
        Else
            Select Case RTrim(c)
                Case "0", "2"
                    PRI_BIKOU_BCR = CInt(RTrim(c))
                Case "1"
                    If Not POS_UMU Then
                        MsgBox "�o�n�r���т����ݒ�ł��B�����𒆎~���܂��B"
                        End
                    Else
                        PRI_BIKOU_BCR = CInt(RTrim(c))
                    End If
                Case Else
                    PRI_BIKOU_BCR = 0
            End Select
        
        End If
    End If
                                '���P�^�S���҂̎�荞��
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "S_TANTO", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "S_TANTO", App.EXEName, c) Then
        PRI_S_TANTO = False
    Else
        If RTrim(c) = "0" Then
            PRI_S_TANTO = False
        Else
            PRI_S_TANTO = True
        End If
    End If
                                
                                '��Ɠ��^���ʁ^�S�� 2007.05.22
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "SAGYO_DAY", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "SAGYO_DAY", App.EXEName, c) Then
        PRI_SAGYO_DAY = False
    Else
        If RTrim(c) = "0" Then
            PRI_SAGYO_DAY = False
        Else
            PRI_SAGYO_DAY = True
        End If
    End If
                                
                                
                                
                                '���i�����@�����̎�荞��
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "DOUKON", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "DOUKON", App.EXEName, c) Then
        PRI_DOUKON = False
    Else
        If RTrim(c) = "0" Then
            PRI_DOUKON = False
        Else
            PRI_DOUKON = True
        End If
    End If
                                '���Ɋ�����̎�荞��
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "NYUKO_IN", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "NYUKO_IN", App.EXEName, c) Then
        PRI_NYUKO_IN = False
    Else
        If RTrim(c) = "0" Then
            PRI_NYUKO_IN = False
        Else
            PRI_NYUKO_IN = True
        End If
    End If
                                '���͊�����̎�荞��
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "INPUT_IN", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "INPUT_IN", App.EXEName, c) Then
        PRI_INPUT_IN = False
    Else
        If RTrim(c) = "0" Then
            PRI_INPUT_IN = False
        Else
            PRI_INPUT_IN = True
        End If
    End If
    
    '�����@�i�ԁ^���^����   2007.05.22
    If PRI_NYUKO_IN Or PRI_INPUT_IN Then
    Else
'        If GetIni(StrConv(App.EXEName, vbUpperCase), "HINBAN_BIKOU", "P_SYS", c) Then
        If GetIni(StrConv(App.EXEName, vbUpperCase), "HINBAN_BIKOU", App.EXEName, c) Then
            PRI_HINBAN_BIKOU = False
        Else
            If RTrim(c) = "0" Then
                PRI_HINBAN_BIKOU = False
            Else
                PRI_HINBAN_BIKOU = True
            End If
        End If
    End If
    
    
                                '����
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "JISEKI", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "JISEKI", App.EXEName, c) Then
        JISEKI_TITLE = ""
    Else
        JISEKI_TITLE = Split(Trim(c), ",", -1)
    End If
    
                                '����
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "TASEKI", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "TASEKI", App.EXEName, c) Then
        TASEKI_TITLE = ""
    Else
        TASEKI_TITLE = Split(Trim(c), ",", -1)
    End If
    
                                '���o�^�i�Ԃ̉�
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "HIN_INV", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "HIN_INV", App.EXEName, c) Then
        HIN_INV = False
    Else
        If Trim(c) = "0" Then
            HIN_INV = False
        Else
            HIN_INV = True
        End If
    
    End If
    
    
    
    If PRI_BIKOU_BCR = 1 Then
                                    '������
'        If GetIni(StrConv(App.EXEName, vbUpperCase), "MTSSS", "P_SYS", c) Then
        If GetIni(StrConv(App.EXEName, vbUpperCase), "MTSSS", App.EXEName, c) Then
            MTS_CODE = ""
            SS_CODE = ""
        Else
        
            MUKE_CODE = Split(Trim(c), ",", -1)
        
            Select Case UBound(MUKE_CODE)
                Case 0
                
                    MTS_CODE = CStr(MUKE_CODE(0))
                    SS_CODE = ""
                Case 1
                    MTS_CODE = CStr(MUKE_CODE(0))
                    SS_CODE = CStr(MUKE_CODE(1))
                Case Else
                    MTS_CODE = ""
                    SS_CODE = ""
            End Select
        End If
                                    '������Ǘ��}�X�^�̃`�F�b�N
        If MTS_Open(BtOpenNomal) Then
            Unload Me
        End If
                                    
        Call UniCode_Conv(K0_MTS.MUKE_CODE, MTS_CODE)
        Call UniCode_Conv(K0_MTS.SS_CODE, SS_CODE)
                                            
        sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                MsgBox "�����悪���ݒ�ł��B�����𒆎~���܂��B"
                                            '������Ǘ��}�X�^�b�k�n�r�d
                sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
                If sts Then
                    If sts <> BtErrNoOpen Then
                        Call File_Error(sts, BtOpClose, "������Ǘ��}�X�^")
                    End If
                End If
                End
            
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "������Ǘ��}�X�^")
                                            '������Ǘ��}�X�^�b�k�n�r�d
                sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
                If sts Then
                    If sts <> BtErrNoOpen Then
                        Call File_Error(sts, BtOpClose, "������Ǘ��}�X�^")
                    End If
                End If
                End
        End Select
                                            '������Ǘ��}�X�^�b�k�n�r�d
        sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
        If sts Then
            If sts <> BtErrNoOpen Then
                Call File_Error(sts, BtOpClose, "������Ǘ��}�X�^")
            End If
        End If
                                            '�����敪�̊l��
'        If GetIni(StrConv(App.EXEName, vbUpperCase), "CYU_KBN", "P_SYS", c) Then
        If GetIni(StrConv(App.EXEName, vbUpperCase), "CYU_KBN", App.EXEName, c) Then
            CYU_KBN = ""
        Else
            CYU_KBN = Trim(c)
        End If
        
        
        
        Select Case CYU_KBN
            Case CYU_KBN_TUK            '����
                CYU_KBN_N = CYU_KBN_1
            Case CYU_KBN_SPO
                CYU_KBN_N = CYU_KBN_2   '�ً}
            Case CYU_KBN_HJU
                CYU_KBN_N = CYU_KBN_3   '��[
            Case CYU_KBN_TOK
                CYU_KBN_N = CYU_KBN_4   '����
            Case CYU_KBN_BOU
                CYU_KBN_N = CYU_KBN_E   '�f��
            Case Else
            MsgBox "�����敪�����ݒ�ł��B�����𒆎~���܂��B"
            End
        End Select
    Else
        MTS_CODE = ""
        SS_CODE = ""
        CYU_KBN = ""
    End If
                                
                                
                                        '�č���̊l��   2007.11.09
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "SAIKON_F", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "SAIKON_F", App.EXEName, c) Then
        SAIKON_F = ""
    Else
        SAIKON_F = Trim(c)
    End If
                                
                                
                                
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.20
    Jyogai_Soko_umu = False
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "JYOGAI_SOKO", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "JYOGAI_SOKO", App.EXEName, c) Then
    Else
        Jyogai_Soko_umu = True
        Zaiko_Syukei_Jyogai_Soko_No = Split(Trim(c), ",", -1)
    End If



'    If GetIni(StrConv(App.EXEName, vbUpperCase), "B", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "B", App.EXEName, c) Then
        YUKO_JGYOBU = "B"
    Else
        YUKO_JGYOBU = Trim(c)
    End If



'--------------------------------------------------- ���  ���ޑΉ��@2012.03.20
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "YOMI_JGYOBU", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "YOMI_JGYOBU", App.EXEName, c) Then
        Call DEF_JGYOBU_PROC
    Else
        YOMI_JGYOBU_w = Split(Trim(c), ",", -1)
        For j = 0 To UBound(YOMI_JGYOBU_w)
            ReDim Preserve YOMI_JGYOBU(j)
            YOMI_JGYOBU(j) = YOMI_JGYOBU_w(j)
        Next j
    End If
'--------------------------------------------------- ���ƕ��ǂݑւ��� 2014.03.24




                                
                                
                                
                                '���ԃ}�X�^�n�o�d�m
    If HATUBAN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '���i���ٗp�i�ڃ}�X�^�n�o�d�m
    If L_ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '�N���X�}�X�^�n�o�d�m
    If P_Class_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�R�[�h�}�X�^�n�o�d�m
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�\���}�X�^�n�o�d�m
    If P_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�Ǘ��}�X�^�n�o�d�m
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���i���w�}�i�q�j�ް��n�o�d�m
    If P_SSHIJI_K_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���i���w�}�i�e�j�ް��n�o�d�m
    If P_SSHIJI_O_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�S���҃}�X�^�n�o�d�m
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�o�ח\���ް��n�o�d�m
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�󕥐�}�X�^�n�o�d�m
    If P_UKEHARAI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌��ް��n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '���i���w�}�i�e�jܰ��n�o�d�m
    If wP_SSHIJI_O_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
                                '���������@�e�i�Ԓ���̧�قn�o�d�m   2012.03.08
    If ODR_ORDER_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
    
    '����Ͻ���`
    Call P_CODE_TBL_Proc
    
    
    
    
    
    
    
    
    Load PI000152
    Load PI000153
    
    
    
    '�Ǘ��}�X�^�̓ǂݍ���
    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)

    sts = BTRV(BtOpGetEqual, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            If P_KANRI_MAKE_Proc() Then
                Unload Me
            End If
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�Ǘ��}�X�^")
            Unload Me
    End Select
        
    
    
    '�d������̃Z�b�g
    If Code_Set_Proc(pcmbSHIMUKE, P_KBN04_CD, 0) Then
        Unload Me
    End If
    '���P�^�S���҂̃Z�b�g
    If Code_Set_Proc(pcmbS_TANTO, P_KBN05_CD, 0) Then
        Unload Me
    End If
    
    '�󕥐�
    If Ukeharai_Set_Proc() Then
        Unload Me
    End If
    
    
    Doukon_Tbl_No(0) = "�@"
    Doukon_Tbl_No(1) = "�A"
    Doukon_Tbl_No(2) = "�B"
    Doukon_Tbl_No(3) = "�C"
    Doukon_Tbl_No(4) = "�D"
    Doukon_Tbl_No(5) = "�E"
    Doukon_Tbl_No(6) = "�F"
    Doukon_Tbl_No(7) = "�G"
    Doukon_Tbl_No(8) = "�H"
    Doukon_Tbl_No(9) = "�I"
    Doukon_Tbl_No(10) = "�J"
    Doukon_Tbl_No(11) = "�K"
    Doukon_Tbl_No(12) = "�L"
    Doukon_Tbl_No(13) = "�M"
    Doukon_Tbl_No(14) = "�N"
    Doukon_Tbl_No(15) = "�O"
    Doukon_Tbl_No(16) = "�P"
    Doukon_Tbl_No(17) = "�Q"
    Doukon_Tbl_No(18) = "�R"
    Doukon_Tbl_No(19) = "�S"
    
    
    
    '��ʂ̃Z�b�g
    For i = pcmbD_SYUBETSU01 To pcmbD_SYUBETSU06
        If Code_Set_Proc(i, P_KBN06_CD, 1) Then
            Unload Me
        End If
    Next i
    
    '��ʏ����ݒ�
    If Init_Proc() Then
        Unload Me
    End If
    Combo1(pcmbSHIMUKE).ListIndex = 0                   '2007.11.01


    '�w���`��       2007.11.01
    Option1(poptSHIJI_NORMAL).Value = True
    Option1(poptSHIJI_SPOT).Value = False
    Option1(poptSHIJI_KEPPIN).Value = False




    




End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts     As Integer
Dim ans     As Integer      '2012.03.09


    ans = MsgBox("�������I�����܂����H", vbYesNo + vbDefaultButton1, "�m�F����")
    If ans = vbNo Then
        Cancel = True
        Exit Sub
    End If
                                            
                                            '���ԃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), K0_HATUBAN, Len(K0_HATUBAN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ԃ}�X�^")
        End If
    End If
                                            
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '���i���ٗp�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, L_ITEM_POS, L_ITEMREC, Len(L_ITEMREC), K0_L_ITEM, Len(K0_L_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���i���ٗp�i�ڃ}�X�^")
        End If
    End If
    
                                            '�N���X�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�N���X�}�X�^")
        End If
    End If
    
                                            '�R�[�h�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�R�[�h�}�X�^")
        End If
    End If
    
                                            '�\���}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�\���}�X�^")
        End If
    End If
                                            '�Ǘ��}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�Ǘ��}�X�^")
        End If
    End If
                                            '���i���w�}�ް�(�e)�b�k�n�r�d
    sts = BTRV(BtOpClose, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���i���w�}�ް�(�e)")
        End If
    End If
                                            '���i���w�}�ް�(�q)�b�k�n�r�d
    sts = BTRV(BtOpClose, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���i���w�}�ް�(�q)")
        End If
    End If
    
                                            '�S���҃}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�S���҃}�X�^")
        End If
    End If
    
                                            '�o�ח\��f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�o�ח\��f�[�^")
        End If
    End If
                                            '�󕥐�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�󕥐�}�X�^")
        End If
    End If
                                            '�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        End If
    End If
    
                                            '���i���w�}ܰ�(�e)�b�k�n�r�d
    sts = BTRV(BtOpClose, wP_SSHIJI_O_POS, wP_SSHIJI_O_REC, Len(wP_SSHIJI_O_REC), K0_wP_SSHIJI_O, Len(K0_wP_SSHIJI_O), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���i���w�}(�e)ܰ�")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PI000151 = Nothing
    Set PI000152 = Nothing
    Set PI000153 = Nothing

    End
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).text)
    End If

    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2013.12.28
    Select Case Index
        Case ptxHIN_GAI
            svHin_Gai = Text1(Index).text
    End Select
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2013.12.28

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
        
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.09
    Select Case Index
        Case ptxHIN_GAI, ptxK_HIN_GAI01, ptxK_HIN_GAI02, ptxK_HIN_GAI03, ptxK_HIN_GAI04, ptxK_HIN_GAI05, _
                ptxG_HIN_GAI01, ptxG_HIN_GAI02, ptxG_HIN_GAI03, _
                ptxD_HIN_GAI01, ptxD_HIN_GAI02, ptxD_HIN_GAI03, ptxD_HIN_GAI04, ptxD_HIN_GAI05, ptxD_HIN_GAI06 _

            Text1(Index).text = StrConv(Text1(Index).text, vbUpperCase)
    
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2013.12.28
            If Index = ptxHIN_GAI Then
                svHin_Gai = Text1(Index).text
            End If
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2013.12.28
    
    End Select
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.09
        
        
        
        
        
        
    If Error_Check_Proc(Index, 0, 0) Then   '�G���[�`�F�b�N
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '�ړ�
End Sub

Private Function Init_Proc() As Integer
'----------------------------------------------------------------------------
'                   ���͉�ʂ̏����ݒ�
'----------------------------------------------------------------------------
Dim i           As Integer
Dim j           As Integer      '2016.01.27
Dim sts         As Integer

Dim TANTO_CODE  As String
Dim TANTO_NAME  As String

Dim SHONIN_CODE     As String   '2007.11.21
Dim SHONIN_NAME     As String   '2007.11.21


    Init_Proc = True
    
    Text1(ptxSHIJI_NO).BackColor = G_INPUT_OK
    Text1(ptxSHIJI_NO).Locked = False
    Text1(ptxSHIJI_NO).TabStop = True
    
    
    Combo1(pcmbS_TANTO).Enabled = PRI_S_TANTO
    
    TANTO_CODE = Text1(ptxTANTO_CODE).text
    TANTO_NAME = Text1(ptxTANTO_NAME).text
    
    SHONIN_CODE = Text1(ptxSHONIN_CODE).text    '2007.11.21
    SHONIN_NAME = Text1(ptxSHONIN_NAME).text    '2007.11.21
    
    
    
    For i = ptxSHIJI_NO To ptxD_BIKOU06
        
        If i = ptxUKEHARAI_CODE Then '2007.11.01
        Else
            Text1(i).text = ""
        End If
    Next i
    Text1(ptxTANTO_CODE).text = TANTO_CODE
    Text1(ptxTANTO_NAME).text = TANTO_NAME




    RichTextBox1(prchBIKOU).text = ""

    For i = pchkSAMPLE_F To pchkPRI_KISHU              '2013.11.20
        Check1(i).Value = vbChecked                    '2013.11.20
    Next i                                             '2013.11.20
    
'    For i = pchkSAMPLE_F To pchkPRI_SHIJI               '2013.11.20
'        Check1(i).Value = vbChecked                     '2013.11.20
'    Next i                                              '2013.11.20
    
    
    
    
    

    
    Check1(pchkSAMPLE_F).Value = vbUnchecked
    Check1(pchkPRI_KISHU).Value = vbUnchecked

    For i = pcmbSHIMUKE To pcmbD_SYUBETSU06
        
        If i = pcmbSHIMUKE Or i = pcmbUKEHARAI Then     '2007.11.01
        Else
            Combo1(i).ListIndex = -1
        End If
    Next i
'    Combo1(pcmbSHIMUKE).ListIndex = 0                  2007.11.01
    

'--------------------------------------------------- ���  ���ޑΉ��@2012.03.18
    '�󒍓�
'    Text1(ptxORDER_NO).text = Format(Now, "YYYY/MM/DD")
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.18


    '���s��
'''    Text1(ptxHAKKO_DT).text = Format(Now, "YYYY/MM/DD")
    Text1(ptxHAKKO_DT).text = ""


    '���F�Ґݒ�
    If Trim(SHONIN_CODE) = "" Then      '2007.11.21
    
    
        Text1(ptxSHONIN_CODE).text = StrConv(P_KANRIREC.SHONIN_CODE, vbUnicode)
    
        Call UniCode_Conv(K0_TANTO.TANTO_CODE, StrConv(P_KANRIREC.SHONIN_CODE, vbUnicode))
    
        sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
        Select Case sts
            Case BtNoErr
                Text1(ptxSHONIN_NAME).text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
            Case BtErrKeyNotFound
                Text1(ptxSHONIN_NAME).text = ""
        
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
                Exit Function
        End Select
    Else
        Text1(ptxSHONIN_CODE).text = SHONIN_CODE    '2007.11.21
        Text1(ptxSHONIN_NAME).text = SHONIN_NAME    '2007.11.21
    End If
    '��z��
    If Trim(Text1(ptxUKEHARAI_CODE).text) = "" Then '2007.11.01
        Text1(ptxUKEHARAI_CODE).text = TEHAI
    End If

    '�w���`��
'2007.11.01    Option1(poptSHIJI_NORMAL).Value = True
'2007.11.01    Option1(poptSHIJI_SPOT).Value = False
'2007.11.01    Option1(poptSHIJI_KEPPIN).Value = False

    '�o�͑Ώ�
    Erase K_Item_Tbl
    Erase G_Item_Tbl
    Erase D_Item_Tbl
    
    ReDim K_Item_Tbl(0 To 4)
    ReDim G_Item_Tbl(0 To 2)
    ReDim D_Item_Tbl(0 To 49)

    For i = 0 To UBound(K_Item_Tbl)
        K_Item_Tbl(i).JGYOBU = ""
        K_Item_Tbl(i).NAIGAI = ""
    Next i

    For i = 0 To UBound(G_Item_Tbl)
        G_Item_Tbl(i).JGYOBU = ""
        G_Item_Tbl(i).NAIGAI = ""
    Next i

    For i = 0 To UBound(D_Item_Tbl)
        D_Item_Tbl(i).JGYOBU = ""
        D_Item_Tbl(i).NAIGAI = ""
        D_Item_Tbl(i).HIN_GAI = ""
    
    Next i


    Call UniCode_Conv(ITEMREC.JGYOBU, "")
    Call UniCode_Conv(P_COMPO_O_REC.JGYOBU, "")



'--------------------------------------------------- ���ƕ��@�Z�b�g 2016.01.27
    For i = 3 To 8
    
        Combo2(i).Clear
    
        For j = 0 To UBound(JGYOBU_T)
        
            Combo2(i).AddItem JGYOBU_T(j).NAME & Space(10) & JGYOBU_T(j).CODE
        
        Next j
    
    
    Next i
'--------------------------------------------------- ���ƕ��@�Z�b�g 2016.01.27

'--------------------------------------------------- ���  ���ޑΉ��@2012.03.18
    Call Disp_Lock_Proc(False)
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.18

    Init_Proc = False

End Function
Private Function Code_Set_Proc(Index As Integer, KBN As String, Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   �R�[�h�}�X�^���R���{�ɃZ�b�g����B
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim Key_Len     As Integer
Dim Option1     As Integer
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
    
    If Mode = 1 Then
        Combo1(Index).AddItem Space(Key_Len)
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
            wkOption = Trim(StrConv(P_CODEREC.Option1, vbUnicode))
        End If
        If P_KBN_TBL(i).KBN_OP2 Then
            wkOption = wkOption & Trim(StrConv(P_CODEREC.OPTION2, vbUnicode))
        End If
        
        
        
        Combo1(Index).AddItem StrConv(P_CODEREC.C_RNAME, vbUnicode) & "            " & _
                                Left(StrConv(P_CODEREC.C_Code, vbUnicode), Key_Len) & wkOption
        
        
        com = BtOpGetNext
    
    Loop

    Code_Set_Proc = False
    



End Function

Private Function Ukeharai_Set_Proc() As Integer
'----------------------------------------------------------------------------
'                   �󕥐�}�X�^���R���{�ɃZ�b�g����B
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer




Dim i           As Integer
    
    Ukeharai_Set_Proc = True
    
    Combo1(pcmbUKEHARAI).Clear
    
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, "")

    com = BtOpGetGreater

    Do
        DoEvents
    
        sts = BTRV(com, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
            
        Select Case sts
            Case BtNoErr
            
                                
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�󕥐�}�X�^")
                Exit Function
        
        End Select

        
        
        Combo1(pcmbUKEHARAI).AddItem StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode) & "            " & _
                                StrConv(P_UKEHARAIREC.UKEHARAI_CODE, vbUnicode)
        
        com = BtOpGetNext
    
    Loop

    Ukeharai_Set_Proc = False
    



End Function



Private Function P_SSHIJI_Read_Proc() As Integer
'----------------------------------------------------------------------------
'                   �w�}�f�[�^�̓ǂݍ���
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
    
    
    P_SSHIJI_Read_Proc = True
    
    
    '�w�}�[�ް��i�e�j
    Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_No, Text1(ptxSHIJI_NO))
    sts = BTRV(BtOpGetEqual, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
        
    Select Case sts
        Case BtNoErr
        
        Case Else
            P_SSHIJI_Read_Proc = sts
            Exit Function
    
    End Select
    
    
    If Item_Disp_Proc() Then
        Exit Function
    End If
    
    P_SSHIJI_Read_Proc = False
        
    

End Function


Private Function P_COMPO_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   �\���}�X�^�̓ǂݍ��݁��\��
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim k           As Integer
Dim g           As Integer
Dim d           As Integer

Dim K_Index     As Integer
Dim G_Index     As Integer
Dim DT_Index    As Integer
Dim DC_Index    As Integer


Dim Mi_Qty      As Long
Dim Sumi_Qty    As Long
    
Dim wkJgyobu    As String * 1
    
    
    P_COMPO_Disp_Proc = True
    
    For i = ptxK_HIN_GAI01 To ptxD_BIKOU06
        Text1(i).text = ""
    Next i
            
            
    For i = pcmbD_JGYOBU01 To pcmbD_JGYOBU06        '2016.04.07
        Combo2(i).ListIndex = -1                    '2016.04.07
    Next                                            '2016.04.07
            
            
    '�o�͑Ώ�
    Erase K_Item_Tbl
    Erase G_Item_Tbl
    Erase D_Item_Tbl
    
    ReDim K_Item_Tbl(0 To 4)
    ReDim G_Item_Tbl(0 To 2)
    ReDim D_Item_Tbl(0 To 49)

    For i = 0 To UBound(K_Item_Tbl)
        K_Item_Tbl(i).JGYOBU = ""
        K_Item_Tbl(i).NAIGAI = ""
    Next i

    For i = 0 To UBound(G_Item_Tbl)
        G_Item_Tbl(i).JGYOBU = ""
        G_Item_Tbl(i).NAIGAI = ""
    Next i

    For i = 0 To UBound(D_Item_Tbl)
        D_Item_Tbl(i).JGYOBU = ""
        D_Item_Tbl(i).NAIGAI = ""
        D_Item_Tbl(i).HIN_GAI = ""
    
    Next i
    
    
    
    
    
    Text1(ptxS_CLASS_CODE).text = ""
    Text1(ptxF_CLASS_CODE).text = ""
    Text1(ptxN_CLASS_CODE).text = ""

    
    
    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).text)
       
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_HEAD)
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")
       
    sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        
    Select Case sts
        Case BtNoErr
        
        Case Else
            
            For i = ptxK_HIN_GAI01 To ptxD_BIKOU06
                Text1(i).text = ""
            Next i
            
            '�o�͑Ώ�
            Erase K_Item_Tbl
            Erase G_Item_Tbl
            Erase D_Item_Tbl
            
            ReDim K_Item_Tbl(0 To 4)
            ReDim G_Item_Tbl(0 To 2)
            ReDim D_Item_Tbl(0 To 49)
        
            For i = 0 To UBound(K_Item_Tbl)
                K_Item_Tbl(i).JGYOBU = ""
                K_Item_Tbl(i).NAIGAI = ""
            Next i
        
            For i = 0 To UBound(G_Item_Tbl)
                G_Item_Tbl(i).JGYOBU = ""
                G_Item_Tbl(i).NAIGAI = ""
            Next i
        
            For i = 0 To UBound(D_Item_Tbl)
                D_Item_Tbl(i).JGYOBU = ""
                D_Item_Tbl(i).NAIGAI = ""
                D_Item_Tbl(i).HIN_GAI = ""
            
            Next i
            
            Text1(ptxS_CLASS_CODE).text = ""
            Text1(ptxF_CLASS_CODE).text = ""
            Text1(ptxN_CLASS_CODE).text = ""
            
            
            P_COMPO_Disp_Proc = sts
            Exit Function
    End Select

    '���i�׽
    Text1(ptxS_CLASS_CODE).text = Trim(StrConv(P_COMPO_O_REC.CLASS_CODE, vbUnicode))
    '�t���׽
    Text1(ptxF_CLASS_CODE).text = Trim(StrConv(P_COMPO_O_REC.F_CLASS_CODE, vbUnicode))
    '���E�׽
    Text1(ptxN_CLASS_CODE).text = Trim(StrConv(P_COMPO_O_REC.N_CLASS_CODE, vbUnicode))
    RichTextBox1(prchBIKOU) = Trim(StrConv(P_COMPO_O_REC.BIKOU, vbUnicode))
    '--------------------------------   �u�q�v���
    Erase K_Item_Tbl
    Erase G_Item_Tbl
    Erase D_Item_Tbl
    
    ReDim K_Item_Tbl(0 To 4)
    ReDim G_Item_Tbl(0 To 2)
    ReDim D_Item_Tbl(0 To 49)
    
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.06
    For i = 0 To UBound(K_Item_Tbl)
        K_Item_Tbl(i).JGYOBU = ""
        K_Item_Tbl(i).NAIGAI = ""
    Next i

    For i = 0 To UBound(G_Item_Tbl)
        G_Item_Tbl(i).JGYOBU = ""
        G_Item_Tbl(i).NAIGAI = ""
    Next i

    For i = 0 To UBound(D_Item_Tbl)
        D_Item_Tbl(i).JGYOBU = ""
        D_Item_Tbl(i).NAIGAI = ""
        D_Item_Tbl(i).HIN_GAI = ""
    
    Next i
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.06
    
    
    
    k = -1
    g = -1
    d = -1
    
    K_Index = ptxK_HIN_GAI01
    G_Index = ptxG_HIN_GAI01
    DT_Index = ptxD_HIN_GAI01
    DC_Index = pcmbD_SYUBETSU01
    
    Do
        
        
        sts = BTRV(BtOpGetNext, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        Select Case sts
            Case BtNoErr
            
                            
                If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Or _
                    StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                    StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Or _
                    Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxHIN_GAI).text) Then
                
                    Exit Do
            
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetNext, "�\���}�X�^")
                Exit Function
        
        
        End Select
        
        Select Case StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode)
        
            Case P_KOSOU    '������
            
                k = k + 1
                K_Item_Tbl(k).JGYOBU = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode)
                K_Item_Tbl(k).NAIGAI = StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode)
                            '�i��
                Text1(K_Index).text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
                
                
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
                Select Case K_Item_Tbl(k).JGYOBU
                    Case SHIZAI
                        Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)
                    Case SETSUBI
                        Call UniCode_Conv(K0_ITEM.JGYOBU, YUKO_JGYOBU)
                    Case Else
                        Call UniCode_Conv(K0_ITEM.JGYOBU, K_Item_Tbl(k).JGYOBU)
                End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
                Call UniCode_Conv(K0_ITEM.NAIGAI, K_Item_Tbl(k).NAIGAI)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                
                
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        '�i��
                        Text1(K_Index + 1) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                        '�W���I��
                        If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                            Text1(K_Index + 4) = ""
                        Else
                            Text1(K_Index + 4) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_DAN, vbUnicode)
                        End If
                    
                    Case BtErrKeyNotFound
                        
                        
'>>>>>>>>>>>>>>>>>>>>>>>>>> �i�Ԗ��o�^�\���̑Ή�    2012.12.21
                        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))

                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                                '�i��
                                Text1(K_Index + 1) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                                '�W���I��
                                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                                    Text1(K_Index + 4) = ""
                                Else
                                    Text1(K_Index + 4) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                        StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                        StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                        StrConv(ITEMREC.ST_DAN, vbUnicode)
                                End If
                            
                            Case BtErrKeyNotFound


                                Text1(K_Index + 1) = "���o�^�i��"
                                Text1(K_Index + 4) = ""
                            Case Else
                                Call Input_UnLock             '2008.01.15
                                Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                                Exit Function

                        End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>> �i�Ԗ��o�^�\���̑Ή�    2012.12.21
                        
                        
                        
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "")
                        Exit Function
                
                End Select
            
            
                Text1(K_Index + 2).text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
                
                If IsNumeric(Text1(ptxSHIJI_QTY).text) Then
                    Text1(K_Index + 3).text = Format(CDbl(CLng(Text1(ptxSHIJI_QTY).text) * CDbl(Text1(K_Index + 2).text)), "#0.00")
                Else
                    Text1(K_Index + 3).text = ""
                End If
                            
                K_Index = K_Index + 5
            
            
            
            Case P_GAISOU   '�O������
                g = g + 1
                G_Item_Tbl(g).JGYOBU = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode)
                G_Item_Tbl(g).NAIGAI = StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode)
                            '�i��
                Text1(G_Index).text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
                Select Case G_Item_Tbl(g).JGYOBU
                    Case SHIZAI
                        Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)
                    Case SETSUBI
                        Call UniCode_Conv(K0_ITEM.JGYOBU, YUKO_JGYOBU)
                    Case Else
                        Call UniCode_Conv(K0_ITEM.JGYOBU, G_Item_Tbl(g).JGYOBU)
                End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
                Call UniCode_Conv(K0_ITEM.NAIGAI, G_Item_Tbl(g).NAIGAI)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(G_Index).text)
                
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        '�i��
                        Text1(G_Index + 1) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                        '�W���I��
                        If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                            Text1(G_Index + 4) = ""
                        Else
                            Text1(G_Index + 4) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_DAN, vbUnicode)
                        End If
                    
                    Case BtErrKeyNotFound
                        
                        
'>>>>>>>>>>>>>>>>>>>>>>>>>> �i�Ԗ��o�^�\���̑Ή�    2012.12.21
                            Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
    
                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                            Select Case sts
                                Case BtNoErr
                                    '�i��
                                    Text1(G_Index + 1) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                                    '�W���I��
                                    If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                                        Text1(G_Index + 4) = ""
                                    Else
                                        Text1(G_Index + 4) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                            StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                            StrConv(ITEMREC.ST_DAN, vbUnicode)
                                    End If
                                
                                Case BtErrKeyNotFound
    
                                    Text1(G_Index + 1) = "���o�^�i��"
                                    Text1(G_Index + 4) = ""
                                Case Else
                                    Call Input_UnLock             '2008.01.15
                                    Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                                    Exit Function
    
                            End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>> �i�Ԗ��o�^�\���̑Ή�    2012.12.21
                        
                        
                        
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "")
                        Exit Function
                
                End Select
            
            
                Text1(G_Index + 2).text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
                If IsNumeric(Text1(ptxSHIJI_QTY).text) Then
                    Text1(G_Index + 3).text = Format(CDbl(CLng(Text1(ptxSHIJI_QTY).text) * CDbl(Text1(G_Index + 2).text)), "#0.00")
                Else
                    Text1(G_Index + 3).text = ""
                End If
                            
                G_Index = G_Index + 5
            
            
            Case P_DOUKON   '�����^�\��
            
                d = d + 1
                D_Item_Tbl(d).SYUBETSU = StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode)
' 2013.01.07 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                D_Item_Tbl(d).JGYOBU = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode)
'                D_Item_Tbl(d).JGYOBU = BUZAI            '����/�\���̎��ƕ����u���ށv�Œ�ɕύX
' 2013.01.07 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                D_Item_Tbl(d).NAIGAI = StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode)
                D_Item_Tbl(d).HIN_GAI = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
                D_Item_Tbl(d).QTY = CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode))
                D_Item_Tbl(d).BIKOU = StrConv(P_COMPO_K_REC.KO_BIKOU, vbUnicode)
                            
                If d > 5 Then
                Else
                            '���
                    For i = 0 To Combo1(DC_Index).ListCount - 1
                    
                        If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = Right(Combo1(DC_Index).List(i), 2) Then
                            Combo1(DC_Index).ListIndex = i
                            Exit For
                        End If
                    
                    Next i
                                
                                
                                
                                '���ƕ� 2016.01.27
                    Combo2(DC_Index).ListIndex = -1
                    For i = 0 To Combo2(DC_Index).ListCount - 1

                    
                        If D_Item_Tbl(d).JGYOBU = Right(Combo2(DC_Index).List(i), 1) Then
                            Combo2(DC_Index).ListIndex = i
                            Exit For
                        End If
                    
                    Next i
                                
                                
                                
                                
                    DC_Index = DC_Index + 1
                                
                                '�i��
                    Text1(DT_Index).text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2016.01.27
'                    Select Case D_Item_Tbl(d).JGYOBU
'                        Case SHIZAI
'                            Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)
'                            wkJgyobu = BUZAI
'                        Case SETSUBI
'                            Call UniCode_Conv(K0_ITEM.JGYOBU, YUKO_JGYOBU)
'                            wkJgyobu = YUKO_JGYOBU
'                        Case Else
'                            Call UniCode_Conv(K0_ITEM.JGYOBU, D_Item_Tbl(d).JGYOBU)
'                            wkJgyobu = D_Item_Tbl(d).JGYOBU
'                    End Select

                     wkJgyobu = D_Item_Tbl(d).JGYOBU

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2016.01.27
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
                    Call UniCode_Conv(K0_ITEM.JGYOBU, D_Item_Tbl(d).JGYOBU)
                    Call UniCode_Conv(K0_ITEM.NAIGAI, D_Item_Tbl(d).NAIGAI)
                    
                    
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(DT_Index).text)
                    
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                            '�i��
                            Text1(DT_Index + 1) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                            '�W���I��
                            If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                                Text1(DT_Index + 4) = ""
                            Else
                                Text1(DT_Index + 4) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                            End If
                        
                        
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.18
                            
                            '>>>>>>>>>>>>>>>>   2013.01.07  DEL
                            'If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIZAI Then
                            '    wkJgyobu = BUZAI
                            'Else
                            '    wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)  '2012.04.04
                            '    'wkJgyobu = YUKO_JGYOBU                          '2012.04.04
                            'End If
                            '>>>>>>>>>>>>>>>>   2013.01.07  DEL
                            
                            '�W���I�Ԃ�ݒ� 2013.01.11
                            If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, wkJgyobu, _
                                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                                                    (StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)), _
                                                                    , , Jyogai_Soko_umu) Then
                                Exit Function
                            
                            End If
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.18
                        
                            Text1(DT_Index + 5) = Format(Sumi_Qty + Mi_Qty, "#0")
                        
                        
                        Case BtErrKeyNotFound
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2016.01.27 �ǂݒ����p�~
                            
'>>>>>>>>>>>>>>>>>>>>>>>>>> �i�Ԗ��o�^�\���̑Ή�    2012.12.21
'                            Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
'                            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
'                            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
'
'                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                            Select Case sts
'                                Case BtNoErr
'                                    '�i��
'                                    Text1(DT_Index + 1) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
'                                    '�W���I��
'                                    If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
'                                        Text1(DT_Index + 4) = ""
'                                    Else
'                                        Text1(DT_Index + 4) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
'                                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
'                                                            StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
'                                                            StrConv(ITEMREC.ST_DAN, vbUnicode)
'                                    End If
'
'
'                                    '�W���I�ԕ���ݒ�   2013.01.11
'                                    If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
'                                                                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
'                                                                            StrConv(ITEMREC.HIN_GAI, vbUnicode), _
'                                                                            (StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode))) Then
'                                        Exit Function
'
'                                    End If
'
'                                    Text1(DT_Index + 5) = Format(Sumi_Qty + Mi_Qty, "#0")
'
'                                Case BtErrKeyNotFound
'
'
'                                    Text1(DT_Index + 1) = "���o�^�i��"
'                                    Text1(DT_Index + 4) = ""
'                                    Text1(DT_Index + 5) = ""
'                                Case Else
'                                    Call Input_UnLock             '2008.01.15
'                                    Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
'                                    Exit Function
'
'                            End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>> �i�Ԗ��o�^�\���̑Ή�    2012.12.21
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2016.01.27
                                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "")
                            Exit Function
                    
                    End Select
                
                
                    Text1(DT_Index + 2).text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
                    If IsNumeric(Text1(ptxSHIJI_QTY).text) Then
                        Text1(DT_Index + 3).text = Format(CDbl(CLng(Text1(ptxSHIJI_QTY).text) * CDbl(Text1(DT_Index + 2).text)), "#0.00")
                    Else
                        Text1(DT_Index + 3).text = ""
                    End If
                    Text1(DT_Index + 6).text = Trim(StrConv(P_COMPO_K_REC.KO_BIKOU, vbUnicode))
                                
                    DT_Index = DT_Index + 7
                End If
        
        End Select
        
        
        
        
        com = BtOpGetNext
    
    Loop

        
    
    
    P_COMPO_Disp_Proc = False

End Function

Private Function Tbl_To_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   ð��ق�蓯���^�\���̕\��
'----------------------------------------------------------------------------
Dim sts         As Integer


Dim i           As Integer
Dim j           As Integer

Dim DC_Index    As Integer
Dim DT_Index    As Integer

Dim Sumi_Qty    As Long
Dim Mi_Qty      As Long

Dim wkJgyobu    As String * 1

    Tbl_To_Disp_Proc = True

    
    DT_Index = ptxD_HIN_GAI01
    DC_Index = pcmbD_SYUBETSU01
                
                
    For i = 0 To 5          '�ŏ��̂U�s��\��
                    
                    '���
        Combo1(DC_Index).ListIndex = -1
        For j = 0 To Combo1(DC_Index).ListCount - 1
        
            If D_Item_Tbl(i).SYUBETSU = Right(Combo1(DC_Index).List(j), 2) Then
                Combo1(DC_Index).ListIndex = j
                Exit For
            End If
        
        Next j
    
    
                    '���ƕ�
'        Combo1(DC_Index).ListIndex = -1            '2017.10.17
        Combo2(DC_Index).ListIndex = -1             '2017.10.17
        For j = 0 To Combo2(DC_Index).ListCount - 1
        
            If D_Item_Tbl(i).JGYOBU = Right(Combo2(DC_Index).List(j), 1) Then
                Combo2(DC_Index).ListIndex = j
                Exit For
            End If
        
        Next j
    
    
        DC_Index = DC_Index + 1
    
    
        If Trim(D_Item_Tbl(i).HIN_GAI) = "" Then
            Text1(DT_Index).text = ""
            Text1(DT_Index + 1).text = ""
            Text1(DT_Index + 2).text = ""
            Text1(DT_Index + 3).text = ""
            Text1(DT_Index + 4).text = ""
            Text1(DT_Index + 5).text = ""
            Text1(DT_Index + 6).text = ""
        Else
        
            Text1(DT_Index).text = D_Item_Tbl(i).HIN_GAI    '�i��
                    
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
            
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2016.01.27
'            Select Case D_Item_Tbl(i).JGYOBU
'                Case SHIZAI
'                    Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)
'                Case SETSUBI
'                    Call UniCode_Conv(K0_ITEM.JGYOBU, YUKO_JGYOBU)
'                Case Else
'                    Call UniCode_Conv(K0_ITEM.JGYOBU, D_Item_Tbl(i).JGYOBU)
'            End Select

            Call UniCode_Conv(K0_ITEM.JGYOBU, D_Item_Tbl(i).JGYOBU)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2016.01.27

'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
            Call UniCode_Conv(K0_ITEM.NAIGAI, D_Item_Tbl(i).NAIGAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(DT_Index).text)
                    
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    '�i��
                    Text1(DT_Index + 1) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                    '�W���I��
                    If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                        Text1(DT_Index + 4) = ""
                    Else
                        Text1(DT_Index + 4) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_DAN, vbUnicode)
                    End If
                        
                
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.18
                    
                    
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
                    'If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIZAI Then
                    '    wkJgyobu = BUZAI
                    'Else
                    '    'wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)  '2012.04.04
                    '    wkJgyobu = YUKO_JGYOBU                          '2012.04.04
                    'End If

'>>>>>>>>>  2016.01.27
'                    Select Case StrConv(ITEMREC.JGYOBU, vbUnicode)
'                        Case SHIZAI
'                            wkJgyobu = SHIZAI
'                        Case SETSUBI
'                            wkJgyobu = YUKO_JGYOBU
'                        Case Else
'                            wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)
'                    End Select



'                   Select Case StrConv(ITEMREC.JGYOBU, vbUnicode)
'                       Case SHIZAI
'                            wkJgyobu = BUZAI
'                        Case SETSUBI
'                            wkJgyobu = YUKO_JGYOBU
'                        Case Else
'                            wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)
'                    End Select

                    wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)
'>>>>>>>>>  2016.01.27

'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
                    
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07 �����ɕW���I�Ԃ�ǉ�
                    If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, wkJgyobu, _
                                                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                            StrConv(ITEMREC.HIN_GAI, vbUnicode), StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                                                StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                                                StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                                                StrConv(ITEMREC.ST_DAN, vbUnicode), , , Jyogai_Soko_umu) Then
                        Exit Function
                    
                    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07 �����ɕW���I�Ԃ�ǉ�
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.18
                    '�݌ɐ�
                    Text1(DT_Index + 5) = Format(Sumi_Qty + Mi_Qty, "#0")
                        
                Case BtErrKeyNotFound
                    
                    Text1(DT_Index + 1) = "���o�^�i��"
                    Text1(DT_Index + 4) = ""
                    Text1(DT_Index + 5) = ""
                                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                    Exit Function
            
            End Select
        
            '����
            Text1(DT_Index + 2).text = Format(D_Item_Tbl(i).QTY, "#0.00")
            '����
            Text1(DT_Index + 3).text = Format(D_Item_Tbl(i).SHIJI_QTY, "#0.00")
            '���l
            Text1(DT_Index + 6).text = D_Item_Tbl(i).BIKOU
        
        End If
    
        DT_Index = DT_Index + 7
    
    Next i
                
    Tbl_To_Disp_Proc = False


End Function

Private Function Y_SYUKA_Make_Proc(i As Integer) As Integer
'----------------------------------------------------------------------------
'                   �o�׎w���̍쐬
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer

Dim ID_NO   As String * 12
Dim DEN_NO  As String * 6


    Y_SYUKA_Make_Proc = True

    '�i�ڃ}�X�^�ǂݍ��݁i�݌ɗL���̔���j
    Call UniCode_Conv(K0_ITEM.JGYOBU, D_Item_Tbl(i).JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, D_Item_Tbl(i).NAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, D_Item_Tbl(i).HIN_GAI)

    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            
            If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIZAI Then
            
                If StrConv(ITEMREC.ZAIKO_F, vbUnicode) <> P_ZAIKO_F_ON Then
                    '�݌ɑΏۊO�i��
                    Y_SYUKA_Make_Proc = False
                    Exit Function
                End If
            End If
        Case BtErrKeyNotFound
            D_Item_Tbl(i).ID_NO = ""
            Y_SYUKA_Make_Proc = False
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Function
    
    End Select

    '-------------------------------------------------- �o�ח\��ҏW
    Call UniCode_Conv(Y_SYUREC.WEL_ID, "")                                  '�g�p�q�@�h�c
    Call UniCode_Conv(Y_SYUREC.PRG_ID, "")                                  '�g�p���v���O����
    Call UniCode_Conv(Y_SYUREC.KAN_KBN, "0")                                '�����敪
    Call UniCode_Conv(Y_SYUREC.DT_SYU, "R")                                 '�f�[�^���
    Call UniCode_Conv(Y_SYUREC.JGYOBU, D_Item_Tbl(i).JGYOBU)                '���ƕ��敪
    Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN)                        '�����敪
    Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN)

    If Den_No_Set_Proc(21, D_Item_Tbl(i).JGYOBU, ID_NO) Then                'IDNO
        Exit Function
    End If
    
    Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, ID_NO)
    Call UniCode_Conv(Y_SYUREC.ID_NO, ID_NO)

    Call UniCode_Conv(Y_SYUREC.NAIGAI, D_Item_Tbl(i).NAIGAI)                '�����O
                                                                    
    Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, D_Item_Tbl(i).HIN_GAI)           '�i�ڔԍ�
    Call UniCode_Conv(Y_SYUREC.HIN_NO, D_Item_Tbl(i).HIN_GAI)               '�i�ڔԍ�

                                                                            '���Ӑ�R�[�h
    Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, MTS_CODE)
    Call UniCode_Conv(Y_SYUREC.MUKE_CODE, MTS_CODE)
                                                                            '������R�[�h
    Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, SS_CODE)
    Call UniCode_Conv(Y_SYUREC.SS_CODE, SS_CODE)
                                                                            '�o�ד�
    Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode))
    Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode))

    Call UniCode_Conv(Y_SYUREC.JGYOBA, "")                                  '���Ə�
    Call UniCode_Conv(Y_SYUREC.DATA_KBN, "")                                '�f�[�^�敪
    Call UniCode_Conv(Y_SYUREC.TORI_KBN, "")                                '����敪
                                                                            '�`�[��
    If Den_No_Set_Proc(20, D_Item_Tbl(i).JGYOBU, DEN_NO) Then
        Exit Function
    End If
    Call UniCode_Conv(Y_SYUREC.DEN_NO, DEN_NO)
                                                                            '�o�ɐ���
    Call UniCode_Conv(Y_SYUREC.SURYO, Format(Int(D_Item_Tbl(i).SHIJI_QTY + 0.9), "0000000"))
        
    Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, "")                             '�o�Ɏ��x
    Call UniCode_Conv(Y_SYUREC.ODER_NO, "")                                 '�I�[�_�[�ԍ�
    Call UniCode_Conv(Y_SYUREC.ITEM_NO, "")                                 '�A�C�e���ԍ�
    Call UniCode_Conv(Y_SYUREC.ODER_NO_R, "")                               '�I�[�_�[�ԍ�����
                                                                            '���Ӑ於��
    Call UniCode_Conv(Y_SYUREC.MUKE_NAME, StrConv(MTSREC.MUKE_NAME, vbUnicode))
                                                                            '�����敪����
    Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, CYU_KBN_N)
                                                                            '�i��
    Call UniCode_Conv(Y_SYUREC.HIN_NAME, StrConv(ITEMREC.HIN_NAME, vbUnicode))
        
        
    
    Call UniCode_Conv(Y_SYUREC.TANABAN1, "")
    Call UniCode_Conv(Y_SYUREC.TANABAN2, "")
    Call UniCode_Conv(Y_SYUREC.TANABAN3, "")
        
        
    Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, "")
    Call UniCode_Conv(Y_SYUREC.ORIGIN1, "")
    Call UniCode_Conv(Y_SYUREC.ORIGIN2, "")
    Call UniCode_Conv(Y_SYUREC.BIKOU2, "")
    Call UniCode_Conv(Y_SYUREC.HAN_KBN, "")
    Call UniCode_Conv(Y_SYUREC.UNIT_ID_NO, "")
    Call UniCode_Conv(Y_SYUREC.ZAIKO_HIKIATE, "")
    Call UniCode_Conv(Y_SYUREC.GOKON_KANRI_NO, "")
    Call UniCode_Conv(Y_SYUREC.JYUCHU_ZAN, "")
    Call UniCode_Conv(Y_SYUREC.KYOKYU_KBN, "")
    Call UniCode_Conv(Y_SYUREC.SHOHIN_SYUSI, "")
    Call UniCode_Conv(Y_SYUREC.S_SHISAN_SYUSI, "")
    Call UniCode_Conv(Y_SYUREC.S_HOJYO_SYUSI, "")
    Call UniCode_Conv(Y_SYUREC.BIKOU1, "")
    Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, "")
    Call UniCode_Conv(Y_SYUREC.JYU_HIN_NO, "")
    Call UniCode_Conv(Y_SYUREC.HIN_NAME, StrConv(ITEMREC.HIN_NAME, vbUnicode))
    Call UniCode_Conv(Y_SYUREC.HIN_CHANGE_KBN, "")
    Call UniCode_Conv(Y_SYUREC.MODULE_EXCHANGE, "")
    Call UniCode_Conv(Y_SYUREC.ZAIKO_SYUSI, "")
    Call UniCode_Conv(Y_SYUREC.ZAN_SHISAN_SYUSI, "")
    Call UniCode_Conv(Y_SYUREC.ZAN_HOJYO_SYUSI, "")
    Call UniCode_Conv(Y_SYUREC.NOUKI_YMD, "")
    Call UniCode_Conv(Y_SYUREC.SERVICE_KANRI_NO, "")
    Call UniCode_Conv(Y_SYUREC.KISHU_CODE, "")
    Call UniCode_Conv(Y_SYUREC.ENVIRONMENT_KBN, "")
    Call UniCode_Conv(Y_SYUREC.SS_CODE, "")
    Call UniCode_Conv(Y_SYUREC.KEPIN_KAIJYO, "")
                                                                            '�z�X�g�I��
    Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                            StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                            StrConv(ITEMREC.ST_DAN, vbUnicode))
    
    Call UniCode_Conv(Y_SYUREC.PRINT_YMD, "")                               '������t
    Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")                                 '�������t
    Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")                              '���i���t
    Call UniCode_Conv(Y_SYUREC.TOK_KBN, "")                                 '������敪
    
    Call UniCode_Conv(Y_SYUREC.JITU_SURYO, "00000000")                      '���ѐ���
                                                                            '�X�V����
    Call UniCode_Conv(Y_SYUREC.INS_NOW, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    Call UniCode_Conv(Y_SYUREC.FILLER, "")

    
    Do
        sts = BTRV(BtOpInsert, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case BtErrDuplicates
                                        '�������ԃf�[�^�d���͍Ĕ��s
                sts = Den_No_Set_Proc(21, D_Item_Tbl(i).JGYOBU, ID_NO)
                If sts Then
                    Exit Function
                End If

                Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, ID_NO)
                Call UniCode_Conv(Y_SYUREC.ID_NO, ID_NO)
                
            Case Else
                Call File_Error(sts, BtOpInsert, "�o�ח\��f�[�^")
                Exit Function
        End Select
    Loop
    
    
        
    
    If SYUKA_LOG_ON Then
        Call SYUKA_LOG_OUT_PROC("INS", "AFT")
    End If
    
    D_Item_Tbl(i).ID_NO = ID_NO
    
    Y_SYUKA_Make_Proc = False
    
End Function
Public Function wP_SSHIJI_O_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���i���w�}(�e)���[�N  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    wP_SSHIJI_O_Open = True
                                            '���i���w�}(�e)�ް��t���p�X�捞��
    sts = GetIni("FILE", P_SSHIJI_O_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_SSHIJI_O]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, wP_SSHIJI_O_POS, wP_SSHIJI_O_REC, Len(wP_SSHIJI_O_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "���i���w�}(�e)ܰ�")
                Exit Function
        End Select
    Loop
    
    wP_SSHIJI_O_Open = False

End Function


Private Function ORDER_Check_Proc(Order_QTY As Long, SHIJI_QTY As Long) As Integer

'----------------------------------------------------------------------------
'                   ���������@�e�i�Ԓ����e�Ƃ̂��荇�킹
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim com     As Integer
    
    
    ORDER_Check_Proc = True
    
    
    Order_QTY = 0
    SHIJI_QTY = Val(Text1(ptxSHIJI_QTY).text)
    
    
    
    
    
    
    '------------------------------------   ���������@�e�i�Ԓ����e�@�`�F�b�N
    Call UniCode_Conv(K6_ODR_ORDER.ORDER_NO, Text1(ptxORDER_NO).text)
    
    
        
    sts = BTRV(com, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K6_ODR_ORDER, Len(K6_ODR_ORDER), 6)
    Select Case sts
        Case BtNoErr
                    
            Order_QTY = Val(Text1(ptxSHIJI_QTY).text)
        
        Case BtErrKeyNotFound
        
        
        Case Else
            Call File_Error(sts, com, "���������F�e�i�Ԓ����e")
            Exit Function
    End Select
    
    
    '------------------------------------   ���i���w�}�f�[�^�i�e�j�@�`�F�b�N
    Call UniCode_Conv(K5_P_SSHIJI_O.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
    Call UniCode_Conv(K5_P_SSHIJI_O.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K5_P_SSHIJI_O.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K5_P_SSHIJI_O.HIN_GAI, Text1(ptxHIN_GAI).text)
    Call UniCode_Conv(K5_P_SSHIJI_O.KAN_F, "")
    com = BtOpGetGreaterEqual
    Do
        
       DoEvents
        
        sts = BTRV(com, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K5_P_SSHIJI_O, Len(K5_P_SSHIJI_O), 5)
        Select Case sts
            Case BtNoErr
                If StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Or _
                    StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                    StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Or _
                    StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode) <> Text1(ptxHIN_GAI).text Then
                    
                    Exit Do
                
                End If
            
                If Trim(StrConv(P_SSHIJI_O_REC.KAN_F, vbUnicode)) <> "" Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���i���w�}�f�[�^�i�e�j")
                Exit Function
        End Select
    
        If Trim(StrConv(ODR_ORDER_REC.BUN_NO, vbUnicode)) <> "" Then
        Else
            Order_QTY = Order_QTY + Val(StrConv(ODR_ORDER_REC.ODR_QTY, vbUnicode))
        End If
        
        com = BtOpGetNext
    
    Loop
        
'    Order_QTY = Val(Text1(ptxSHIJI_QTY).text)
    
    
    
    
    
    ORDER_Check_Proc = False
End Function

Private Function Zaiko_Check_Proc(ZAIKO_F As Boolean) As Integer
'----------------------------------------------------------------------------
'                   ���݌ɂƂ̂��荇�킹
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer

Dim i               As Integer
Dim j               As Integer
 
Dim Sumi_Qty        As Long
Dim Mi_Qty          As Long
 
Dim wkJgyobu        As String * 1
 
 
    Zaiko_Check_Proc = True
    
       
    Erase ZAIKO_FUSOKU
    
    
    
    ZAIKO_F = False
    
    j = -1
    
'---------------------------------------------------------------------- �����ނ̎g�p�W�v
    For i = ptxK_HIN_GAI01 To ptxK_HIN_GAI05 Step 5
    
    
        If Trim(Text1(i).text) = "" Then
        Else
            
            
            If Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1) = SHIZAI Then
                wkJgyobu = BUZAI
            Else
                'wkJgyobu = Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1)   '2012.04.04
                wkJgyobu = YUKO_JGYOBU                                      '2012.04.04
            End If
                        
            Call UniCode_Conv(K0_ITEM.JGYOBU, wkJgyobu)
           
            
            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(i).text)
            
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(ITEMREC.JGYOBU, wkJgyobu)
                    Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
                    Call UniCode_Conv(ITEMREC.HIN_GAI, Text1(i).text)
            
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                    Exit Function
            
            End Select
    
    
            If j = -1 Then
                j = 0
                ReDim Preserve ZAIKO_FUSOKU(0 To j)
                ZAIKO_FUSOKU(j).JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
                ZAIKO_FUSOKU(j).NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
                ZAIKO_FUSOKU(j).HIN_GAI = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                ZAIKO_FUSOKU(j).USE_QTY = Val(Text1(i + 3).text)
                ZAIKO_FUSOKU(j).ZAIKO_QTY = 0
            Else
                For j = 0 To UBound(ZAIKO_FUSOKU)
                
                    If ZAIKO_FUSOKU(j).JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode) Or _
                        ZAIKO_FUSOKU(j).NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode) Or _
                        RTrim(ZAIKO_FUSOKU(j).HIN_GAI) = RTrim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                
                        Exit For
                    End If
                
                Next j
                
                If j > UBound(ZAIKO_FUSOKU) Then
                    ReDim Preserve ZAIKO_FUSOKU(0 To j)
                    ZAIKO_FUSOKU(j).JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
                    ZAIKO_FUSOKU(j).NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
                    ZAIKO_FUSOKU(j).HIN_GAI = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                    ZAIKO_FUSOKU(j).USE_QTY = Val(Text1(i + 3).text)
                    ZAIKO_FUSOKU(j).ZAIKO_QTY = 0
                Else
                    ZAIKO_FUSOKU(j).USE_QTY = ZAIKO_FUSOKU(j).USE_QTY + Val(Text1(i + 3).text)
                End If
            End If
        End If
    Next i
'---------------------------------------------------------------------- �����ނ̎g�p�W�v
    
    
'---------------------------------------------------------------------- �O�����ނ̎g�p�W�v
    For i = ptxG_HIN_GAI01 To ptxK_HIN_GAI03 Step 5
    
    
        If Trim(Text1(i).text) = "" Then
        Else
    
            If Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1) = SHIZAI Then
                wkJgyobu = BUZAI
            Else
                'wkJgyobu = Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1)   '2012.04.04
                wkJgyobu = YUKO_JGYOBU                                      '2012.04.04
            End If
                        
            Call UniCode_Conv(K0_ITEM.JGYOBU, wkJgyobu)
            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(i).text)
            
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(ITEMREC.JGYOBU, wkJgyobu)
                    Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
                    Call UniCode_Conv(ITEMREC.HIN_GAI, Text1(i).text)
            
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                    Exit Function
            
            End Select
    
    
            If j = -1 Then
                j = 0
                ReDim Preserve ZAIKO_FUSOKU(0 To j)
                ZAIKO_FUSOKU(j).JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
                ZAIKO_FUSOKU(j).NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
                ZAIKO_FUSOKU(j).HIN_GAI = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                ZAIKO_FUSOKU(j).USE_QTY = Val(Text1(i + 3).text)
                ZAIKO_FUSOKU(j).ZAIKO_QTY = 0
            
            
            Else
                For j = 0 To UBound(ZAIKO_FUSOKU)
                
                    If ZAIKO_FUSOKU(j).JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode) Or _
                        ZAIKO_FUSOKU(j).NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode) Or _
                        RTrim(ZAIKO_FUSOKU(j).HIN_GAI) = RTrim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                
                        Exit For
                
                    End If
                Next j
                
                If j > UBound(ZAIKO_FUSOKU) Then
                    ReDim Preserve ZAIKO_FUSOKU(0 To j)
                    ZAIKO_FUSOKU(j).JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
                    ZAIKO_FUSOKU(j).NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
                    ZAIKO_FUSOKU(j).HIN_GAI = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                    ZAIKO_FUSOKU(j).USE_QTY = Val(Text1(i + 3).text)
                    ZAIKO_FUSOKU(j).ZAIKO_QTY = 0
                Else
                    ZAIKO_FUSOKU(j).USE_QTY = ZAIKO_FUSOKU(j).USE_QTY + Val(Text1(i + 3).text)
                End If
            End If
        End If
    Next i
'---------------------------------------------------------------------- �O�����ނ̎g�p�W�v
    
    
    
'---------------------------------------------------------------------- �������̎g�p�W�v
    For i = 0 To UBound(D_Item_Tbl)
        If Trim(D_Item_Tbl(i).HIN_GAI) = "" Then
        Else
            '>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
            'If D_Item_Tbl(i).JGYOBU = SHIZAI Then
            '    wkJgyobu = BUZAI
            'Else
            '    'wkJgyobu = D_Item_Tbl(i).JGYOBU        '2012.04.04
            '    wkJgyobu = YUKO_JGYOBU                  '2012.04.04
            'End If
            
            
            Select Case D_Item_Tbl(i).JGYOBU
'2016.05.31                Case SHIZAI
'2016.05.31                    wkJgyobu = BUZAI
'2016.05.31                Case SETSUBI
'2016.05.31                    wkJgyobu = YUKO_JGYOBU
                Case Else
                    wkJgyobu = D_Item_Tbl(i).JGYOBU
            End Select
            '>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
                        
            Call UniCode_Conv(K0_ITEM.JGYOBU, wkJgyobu)
            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, D_Item_Tbl(i).HIN_GAI)
            
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(ITEMREC.JGYOBU, wkJgyobu)
                    Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
                    Call UniCode_Conv(ITEMREC.HIN_GAI, D_Item_Tbl(i).HIN_GAI)
            
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                    Exit Function
            
            End Select
    
    
            If j = -1 Then
                j = 0
                ReDim Preserve ZAIKO_FUSOKU(0 To j)
                ZAIKO_FUSOKU(j).JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
                ZAIKO_FUSOKU(j).NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
                ZAIKO_FUSOKU(j).HIN_GAI = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                ZAIKO_FUSOKU(j).USE_QTY = D_Item_Tbl(i).SHIJI_QTY
                ZAIKO_FUSOKU(j).ZAIKO_QTY = 0
                
            
            Else
                For j = 0 To UBound(ZAIKO_FUSOKU)
                
                    If ZAIKO_FUSOKU(j).JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode) And _
                        ZAIKO_FUSOKU(j).NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode) And _
                        RTrim(ZAIKO_FUSOKU(j).HIN_GAI) = RTrim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                
                        Exit For
                
                    End If
                
                Next j
                
                
                If j > UBound(ZAIKO_FUSOKU) Then
                    ReDim Preserve ZAIKO_FUSOKU(0 To j)
                    ZAIKO_FUSOKU(j).JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
                    ZAIKO_FUSOKU(j).NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
                    ZAIKO_FUSOKU(j).HIN_GAI = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                    ZAIKO_FUSOKU(j).USE_QTY = D_Item_Tbl(i).SHIJI_QTY
                Else
                    ZAIKO_FUSOKU(j).USE_QTY = ZAIKO_FUSOKU(j).USE_QTY + D_Item_Tbl(i).SHIJI_QTY
                End If
            End If
        End If
    
    
    
    Next i
'---------------------------------------------------------------------- �������̎g�p�W�v
    
    If j = -1 Then
        Zaiko_Check_Proc = False
        Exit Function
    End If
    
'---------------------------------------------------------------------- ���݌ɂ̏W�v
    For i = 0 To UBound(ZAIKO_FUSOKU)
        
'        If ZAIKO_FUSOKU(i).JGYOBU = SHIZAI Then
'            wkJgyobu = BUZAI
'        Else
'            'wkJgyobu = ZAIKO_FUSOKU(i).JGYOBU  '2012.04.04
'            wkJgyobu = YUKO_JGYOBU              '2012.04.04
'        End If
        
        
If ZAIKO_FUSOKU(i).JGYOBU = "B" Then
Debug.Print
End If
        
        
        
        If Zaiko_Syukei_Proc(Sumi_Qty, _
                                Mi_Qty, _
                                ZAIKO_FUSOKU(i).JGYOBU, _
                                ZAIKO_FUSOKU(i).NAIGAI, _
                                ZAIKO_FUSOKU(i).HIN_GAI, , , , Jyogai_Soko_umu) Then
            Exit Function
        End If
        ZAIKO_FUSOKU(i).ZAIKO_QTY = Sumi_Qty + Mi_Qty
    Next i
'---------------------------------------------------------------------- ���݌ɂ̏W�v
    
'---------------------------------------------------------------------- �g�p�\�񕪂̏W�v
    For i = 0 To UBound(ZAIKO_FUSOKU)
        
'        If ZAIKO_FUSOKU(i).JGYOBU = SHIZAI Then
'            wkJgyobu = BUZAI
'        Else
'            'wkJgyobu = ZAIKO_FUSOKU(i).JGYOBU  '2012.04.04
'            wkJgyobu = YUKO_JGYOBU              '2012.04.04
'        End If
        
        
        Call UniCode_Conv(K2_P_SSHIJI_K.KO_JGYOBU, ZAIKO_FUSOKU(i).JGYOBU)
        Call UniCode_Conv(K2_P_SSHIJI_K.KO_NAIGAI, ZAIKO_FUSOKU(i).NAIGAI)
        Call UniCode_Conv(K2_P_SSHIJI_K.KO_HIN_GAI, ZAIKO_FUSOKU(i).HIN_GAI)
        Call UniCode_Conv(K2_P_SSHIJI_K.IDO_SUMI, "")
    
        com = BtOpGetGreaterEqual
    
        Do
            DoEvents
            
            sts = BTRV(com, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K2_P_SSHIJI_K, Len(K2_P_SSHIJI_K), 2)
            Select Case sts
                Case BtNoErr
                    
                    If StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode) <> ZAIKO_FUSOKU(i).JGYOBU Or _
                        StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode) <> ZAIKO_FUSOKU(i).NAIGAI Or _
                        StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode) <> ZAIKO_FUSOKU(i).HIN_GAI Then
                        Exit Do
                    End If
                            


                Case BtErrEOF
                    Exit Do


                Case Else
                    Call File_Error(sts, com, "���i���w�}�f�[�^�i�q�j")
                    Exit Function


            
            End Select
        
        
            If StrConv(P_SSHIJI_K_REC.CALCEL_F, vbUnicode) <> P_CANCEL_ON Then
            
                ZAIKO_FUSOKU(i).ZAIKO_QTY = ZAIKO_FUSOKU(i).ZAIKO_QTY - Val(StrConv(P_SSHIJI_K_REC.HIKIATE_QTY, vbUnicode))
            
            End If
        
        
            com = BtOpGetNext
        
        Loop
    Next i
'---------------------------------------------------------------------- �g�p�\�񕪂̏W�v
    
'---------------------------------------------------------------------- ���ِ��̏W�v
    For i = 0 To UBound(ZAIKO_FUSOKU)
        ZAIKO_FUSOKU(i).SAI_QTY = ZAIKO_FUSOKU(i).ZAIKO_QTY - ZAIKO_FUSOKU(i).USE_QTY
        If ZAIKO_FUSOKU(i).SAI_QTY < 0 Then
            ZAIKO_F = True
        End If
    Next i
'---------------------------------------------------------------------- ���ِ��̏W�v
    
    Zaiko_Check_Proc = False

End Function

Private Sub Text1_LostFocus(Index As Integer)



'--------------------------------------------------- ���  ���ޑΉ��@2012.03.09
    Select Case Index
        
        
        
        Case ptxORDER_NO                                        '2016.06.22
        
            If Error_Check_Proc(Index, 0, 0) Then   '�G���[�`�F�b�N
                Exit Sub
            End If
        
        
        
        Case ptxHIN_GAI
            Text1(Index).text = StrConv(Text1(Index).text, vbUpperCase)
    
            '>>>>>>>>>>>>>> 2013.12.28
'            If svHin_Gai <> Text1(Index).text Then             '2016.01.27
            If Trim(svHin_Gai) <> Text1(Index).text Then        '2016.01.27
            
                        
                If Error_Check_Proc(Index, 0, 0) Then   '�G���[�`�F�b�N
                    Exit Sub
                End If
            
            End If
            '>>>>>>>>>>>>>> 2013.12.28
    
    End Select
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.09




End Sub

Private Sub Disp_Lock_Proc(Mode As Integer)
'----------------------------------------------------------------------------
'                   ���Lock/UnLock
'           2012.03.18
'----------------------------------------------------------------------------
Dim i   As Integer


    Text1(ptxHIN_GAI).Locked = Mode
    Text1(ptxSHIJI_QTY).Locked = Mode
    
    For i = ptxK_HIN_GAI01 To ptxD_BIKOU06
'        Text1(ptxSHIJI_QTY).Locked = Mode      '2016.05.18
        Text1(i).Locked = Mode                  '2016.05.18
    Next i


    For i = pcmbD_SYUBETSU01 To pcmbD_SYUBETSU06
        Combo1(i).Locked = Mode
    Next i

        
    For i = pcmbD_JGYOBU01 To pcmbD_JGYOBU06    '2016.05.18
        Combo2(i).Locked = Mode                 '2016.05.18
    Next i                                      '2016.05.18
        
        
        
    For i = 0 To 24
        PI000152.Combo1(i).Locked = Mode
    Next i
    
    
    For i = 0 To 24                             '2016.05.18
        PI000152.Combo2(i).Locked = Mode        '2016.05.18
    Next i                                      '2016.05.18
    
    
    For i = 0 To 174
        PI000152.Text1(i).Locked = Mode
    Next i




End Sub


Private Sub DEF_JGYOBU_PROC()
'-------------------------------------------------------------------------
'
'   �ǂݑւ��f�t�H���g���ƕ��@�Z�b�g
'
'       2014.03.24
'
'
'-------------------------------------------------------------------------
Dim i   As Integer
Dim j   As Integer
    
Dim c   As String * 128
    
    
    i = 0
    j = 0
    Do
        If GetIni("JIGYOBU", "code" & RTrim(Format$(i + 1, "#0")), "SYS", c) Then
            Exit Sub
        End If
        If RTrim(c) = "0" Then
            Exit Do
        End If

        ReDim Preserve YOMI_JGYOBU(j)

        YOMI_JGYOBU(j) = RTrim(c)
        j = j + 1
        i = i + 1
    Loop

End Sub
