VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form PI000101 
   Caption         =   "���i���w�}�[���s "
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   ClipControls    =   0   'False
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
   ScaleHeight     =   10155
   ScaleWidth      =   15240
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CommandButton Command1 
      Caption         =   "�L�����Z��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   8040
      TabIndex        =   193
      TabStop         =   0   'False
      Top             =   9360
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "��"
      Height          =   255
      Index           =   5
      Left            =   2280
      TabIndex        =   27
      Top             =   3720
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "�v��"
      Height          =   255
      Index           =   6
      Left            =   2880
      TabIndex        =   28
      Top             =   3720
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "�K�p�@�탉�x��"
      Height          =   255
      Index           =   7
      Left            =   3720
      TabIndex        =   29
      Top             =   3720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ListBox lstGensankoku 
      Height          =   780
      Left            =   12735
      Sorted          =   -1  'True
      TabIndex        =   181
      Top             =   2220
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.TextBox txGensankoku 
      Height          =   375
      Left            =   12180
      TabIndex        =   177
      Top             =   240
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   98
      Left            =   13230
      MaxLength       =   8
      TabIndex        =   119
      Top             =   9000
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   11
      Left            =   14160
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   10
      Left            =   13200
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1080
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   1800
      TabIndex        =   173
      Top             =   2160
      Width           =   4455
      Begin VB.OptionButton Option1 
         Caption         =   "���i����"
         Height          =   375
         Index           =   2
         Left            =   2880
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "���O"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   175
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "�X�|�b�g"
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   97
      Left            =   10200
      TabIndex        =   118
      Top             =   8520
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   96
      Left            =   9600
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   117
      TabStop         =   0   'False
      Top             =   8520
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   95
      Left            =   8160
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   116
      TabStop         =   0   'False
      Top             =   8520
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   94
      Left            =   7080
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   115
      TabStop         =   0   'False
      Top             =   8520
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   93
      Left            =   6240
      MaxLength       =   6
      TabIndex        =   114
      Top             =   8520
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   92
      Left            =   3120
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   113
      TabStop         =   0   'False
      Top             =   8520
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   91
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   112
      Top             =   8520
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   90
      Left            =   10200
      TabIndex        =   110
      Top             =   8160
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   89
      Left            =   9600
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   109
      TabStop         =   0   'False
      Top             =   8160
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   88
      Left            =   8160
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   108
      TabStop         =   0   'False
      Top             =   8160
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   87
      Left            =   7080
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   107
      TabStop         =   0   'False
      Top             =   8160
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   86
      Left            =   6240
      MaxLength       =   6
      TabIndex        =   106
      Top             =   8160
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   85
      Left            =   3120
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   8160
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   84
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   104
      Top             =   8160
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   83
      Left            =   10200
      TabIndex        =   102
      Top             =   7800
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   82
      Left            =   9600
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   101
      TabStop         =   0   'False
      Top             =   7800
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   81
      Left            =   8160
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   100
      TabStop         =   0   'False
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   80
      Left            =   7080
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   7800
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   79
      Left            =   6240
      MaxLength       =   6
      TabIndex        =   98
      Top             =   7800
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   78
      Left            =   3120
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   97
      TabStop         =   0   'False
      Top             =   7800
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   77
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   96
      Top             =   7800
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   76
      Left            =   10200
      TabIndex        =   94
      Top             =   7440
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   75
      Left            =   9600
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   7440
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   74
      Left            =   8160
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   7440
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   73
      Left            =   7080
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   7440
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   72
      Left            =   6240
      MaxLength       =   6
      TabIndex        =   90
      Top             =   7440
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   71
      Left            =   3120
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   7440
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   70
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   88
      Top             =   7440
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   69
      Left            =   10200
      TabIndex        =   86
      Top             =   7080
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   68
      Left            =   9600
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   7080
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   67
      Left            =   8160
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   7080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   66
      Left            =   7080
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   7080
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   65
      Left            =   6240
      MaxLength       =   6
      TabIndex        =   82
      Top             =   7080
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   64
      Left            =   3120
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   7080
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   63
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   80
      Top             =   7080
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   55
      Left            =   13320
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   54
      Left            =   12240
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   53
      Left            =   11400
      MaxLength       =   6
      TabIndex        =   68
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   52
      Left            =   9240
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   5280
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   51
      Left            =   7920
      MaxLength       =   20
      TabIndex        =   66
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   50
      Left            =   13320
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   49
      Left            =   12240
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   48
      Left            =   11400
      MaxLength       =   6
      TabIndex        =   63
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   47
      Left            =   9240
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   4920
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   46
      Left            =   7920
      MaxLength       =   20
      TabIndex        =   61
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   45
      Left            =   13320
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   44
      Left            =   12240
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   43
      Left            =   11400
      MaxLength       =   6
      TabIndex        =   58
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   42
      Left            =   9240
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   41
      Left            =   7920
      MaxLength       =   20
      TabIndex        =   56
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   40
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   39
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   38
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   53
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   37
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   6000
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   36
      Left            =   600
      MaxLength       =   20
      TabIndex        =   51
      Top             =   6000
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   35
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   34
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   33
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   48
      Top             =   5640
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   32
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   5640
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   31
      Left            =   600
      MaxLength       =   20
      TabIndex        =   46
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   30
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   29
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   28
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   43
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   27
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   5280
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   26
      Left            =   600
      MaxLength       =   20
      TabIndex        =   41
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   25
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   24
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   23
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   38
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   15
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   18
      Top             =   1800
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1095
      Index           =   0
      Left            =   8040
      TabIndex        =   30
      Top             =   2640
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   1931
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"PI000101.frx":0000
   End
   Begin VB.ComboBox Combo1 
      Height          =   336
      Index           =   8
      Left            =   240
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   111
      Top             =   8520
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   336
      Index           =   7
      Left            =   240
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   103
      Top             =   8160
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   336
      Index           =   6
      Left            =   240
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   95
      Top             =   7800
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   336
      Index           =   5
      Left            =   240
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   87
      Top             =   7440
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   336
      Index           =   4
      Left            =   240
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   79
      Top             =   7080
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   336
      Index           =   3
      Left            =   240
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   71
      Top             =   6720
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "�o�͑Ώ�"
      Height          =   732
      Left            =   240
      TabIndex        =   155
      Top             =   2880
      Width           =   6630
      Begin VB.ComboBox Combo2 
         Height          =   336
         Index           =   0
         Left            =   1440
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   184
         Top             =   240
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Caption         =   "�@�탉�x��"
         Height          =   375
         Index           =   4
         Left            =   6240
         TabIndex        =   26
         Top             =   240
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.CheckBox Check1 
         Caption         =   "�O�����x��"
         Height          =   375
         Index           =   3
         Left            =   4320
         TabIndex        =   25
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "�p�[�c���x��"
         Height          =   375
         Index           =   2
         Left            =   2040
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "�w�}�["
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   22
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   4920
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   21
      Left            =   600
      MaxLength       =   20
      TabIndex        =   36
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   20
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   19
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   18
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   33
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   17
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   16
      Left            =   600
      MaxLength       =   20
      TabIndex        =   31
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   5
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   12
      Left            =   240
      MaxLength       =   5
      TabIndex        =   14
      Top             =   1800
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "���{�쐬"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   20
      Top             =   2400
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   2
      Left            =   9840
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   19
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   14
      Left            =   5640
      MaxLength       =   10
      TabIndex        =   17
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   13
      Left            =   4200
      MaxLength       =   10
      TabIndex        =   16
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   336
      Index           =   1
      Left            =   960
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   15
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   9
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   8
      Left            =   9960
      MaxLength       =   8
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   7
      Left            =   5640
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   3075
      MaxLength       =   20
      TabIndex        =   7
      Top             =   1080
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   240
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   6
      Top             =   1080
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   6240
      MaxLength       =   5
      TabIndex        =   4
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   3
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   3340
      MaxLength       =   5
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   1
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
      Width           =   1050
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
      TabIndex        =   131
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
      TabIndex        =   130
      Top             =   9000
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��߰���ēo�^"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   8040
      TabIndex        =   129
      TabStop         =   0   'False
      Top             =   9000
      Width           =   1575
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
      Left            =   7200
      TabIndex        =   128
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
      Left            =   6240
      TabIndex        =   127
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
      Left            =   5400
      TabIndex        =   126
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
      Left            =   4560
      TabIndex        =   125
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
      Left            =   3720
      TabIndex        =   124
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
      TabIndex        =   123
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
      TabIndex        =   122
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
      TabIndex        =   121
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
      TabIndex        =   120
      Top             =   9000
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   62
      Left            =   10200
      TabIndex        =   78
      Top             =   6720
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   61
      Left            =   9600
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   6720
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   60
      Left            =   8160
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   6720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   59
      Left            =   7080
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   6720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   58
      Left            =   6240
      MaxLength       =   6
      TabIndex        =   74
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   57
      Left            =   3120
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   6720
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   56
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   72
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Label lblL_URIKIN3 
      Height          =   135
      Left            =   9840
      TabIndex        =   192
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblL_URIKIN2 
      Height          =   375
      Left            =   9240
      TabIndex        =   191
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblGAI_BUHIN 
      Height          =   135
      Left            =   8760
      TabIndex        =   190
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblL_Hin_Name_E 
      Height          =   255
      Left            =   8760
      TabIndex        =   189
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblL_JGYOBU_N 
      Height          =   315
      Left            =   11280
      TabIndex        =   188
      Top             =   9600
      Width           =   690
   End
   Begin VB.Label lblL_KAISHA_N 
      Height          =   315
      Left            =   9960
      TabIndex        =   187
      Top             =   9600
      Width           =   690
   End
   Begin VB.Label lblKISHU2 
      Height          =   255
      Left            =   11280
      TabIndex        =   186
      Top             =   5880
      Width           =   3375
   End
   Begin VB.Label lblKISHU1 
      Height          =   255
      Left            =   7560
      TabIndex        =   185
      Top             =   5880
      Width           =   3375
   End
   Begin VB.Label lblL_JGYOBU 
      Height          =   315
      Left            =   14490
      TabIndex        =   183
      Top             =   3300
      Width           =   690
   End
   Begin VB.Label lblL_KAISHA 
      Height          =   315
      Left            =   14490
      TabIndex        =   182
      Top             =   2880
      Width           =   690
   End
   Begin VB.Label lblGensankoku 
      BorderStyle     =   1  '����
      Height          =   375
      Index           =   1
      Left            =   12735
      TabIndex        =   180
      Top             =   1800
      Width           =   2445
   End
   Begin VB.Label lblGensankoku 
      Height          =   255
      Index           =   0
      Left            =   13590
      TabIndex        =   179
      Top             =   1560
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���Y��"
      Height          =   255
      Index           =   25
      Left            =   12780
      TabIndex        =   178
      Top             =   1560
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���x�����s����"
      Height          =   255
      Index           =   24
      Left            =   11445
      TabIndex        =   176
      Top             =   9120
      Width           =   1710
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���i����"
      Height          =   255
      Index           =   23
      Left            =   14040
      TabIndex        =   13
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�����i"
      Height          =   255
      Index           =   17
      Left            =   13200
      TabIndex        =   174
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�I��"
      Height          =   252
      Index           =   16
      Left            =   8520
      TabIndex        =   172
      Top             =   6480
      Width           =   492
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "����"
      Height          =   252
      Index           =   22
      Left            =   7560
      TabIndex        =   171
      Top             =   6480
      Width           =   492
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�i��"
      Height          =   252
      Index           =   21
      Left            =   3720
      TabIndex        =   170
      Top             =   6480
      Width           =   492
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���E�׽"
      Height          =   255
      Index           =   20
      Left            =   7320
      TabIndex        =   169
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���l"
      Height          =   252
      Index           =   19
      Left            =   10920
      TabIndex        =   168
      Top             =   6480
      Width           =   492
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�݌�"
      Height          =   252
      Index           =   18
      Left            =   9720
      TabIndex        =   167
      Top             =   6480
      Width           =   492
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "����"
      Height          =   252
      Index           =   15
      Left            =   6480
      TabIndex        =   166
      Top             =   6480
      Width           =   492
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�i��"
      Height          =   252
      Index           =   14
      Left            =   1440
      TabIndex        =   165
      Top             =   6480
      Width           =   492
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���"
      Height          =   252
      Index           =   12
      Left            =   240
      TabIndex        =   164
      Top             =   6480
      Width           =   492
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�O�����އ�"
      Height          =   372
      Index           =   17
      Left            =   7560
      TabIndex        =   163
      Top             =   4200
      Width           =   1692
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�@"
      Height          =   372
      Index           =   16
      Left            =   7560
      TabIndex        =   162
      Top             =   4560
      Width           =   372
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�A"
      Height          =   372
      Index           =   15
      Left            =   7560
      TabIndex        =   161
      Top             =   4920
      Width           =   372
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�B"
      Height          =   372
      Index           =   14
      Left            =   7560
      TabIndex        =   160
      Top             =   5280
      Width           =   372
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�i��"
      Height          =   372
      Index           =   13
      Left            =   9240
      TabIndex        =   159
      Top             =   4200
      Width           =   2172
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "����"
      Height          =   372
      Index           =   12
      Left            =   11400
      TabIndex        =   158
      Top             =   4200
      Width           =   852
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "����"
      Height          =   372
      Index           =   11
      Left            =   12240
      TabIndex        =   157
      Top             =   4200
      Width           =   1092
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�I��"
      Height          =   372
      Index           =   10
      Left            =   13320
      TabIndex        =   156
      Top             =   4200
      Width           =   1452
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�I��"
      Height          =   372
      Index           =   9
      Left            =   6000
      TabIndex        =   154
      Top             =   4200
      Width           =   1452
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "����"
      Height          =   372
      Index           =   8
      Left            =   4920
      TabIndex        =   153
      Top             =   4200
      Width           =   1092
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "����"
      Height          =   372
      Index           =   7
      Left            =   4080
      TabIndex        =   152
      Top             =   4200
      Width           =   852
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�i��"
      Height          =   372
      Index           =   6
      Left            =   1920
      TabIndex        =   151
      Top             =   4200
      Width           =   2172
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�D"
      Height          =   372
      Index           =   5
      Left            =   240
      TabIndex        =   150
      Top             =   6000
      Width           =   372
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�C"
      Height          =   372
      Index           =   4
      Left            =   240
      TabIndex        =   149
      Top             =   5640
      Width           =   372
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�B"
      Height          =   372
      Index           =   3
      Left            =   240
      TabIndex        =   148
      Top             =   5280
      Width           =   372
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�A"
      Height          =   372
      Index           =   2
      Left            =   240
      TabIndex        =   147
      Top             =   4920
      Width           =   372
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�@"
      Height          =   372
      Index           =   1
      Left            =   240
      TabIndex        =   146
      Top             =   4560
      Width           =   372
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�����އ�"
      Height          =   372
      Index           =   0
      Left            =   240
      TabIndex        =   145
      Top             =   4200
      Width           =   1692
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���l"
      Height          =   255
      Index           =   13
      Left            =   8040
      TabIndex        =   144
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���P/�S����"
      Height          =   255
      Index           =   11
      Left            =   9960
      TabIndex        =   143
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�t���׽"
      Height          =   255
      Index           =   10
      Left            =   5880
      TabIndex        =   142
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
      TabIndex        =   141
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
      TabIndex        =   140
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�W���I��"
      Height          =   255
      Index           =   7
      Left            =   11520
      TabIndex        =   139
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
      TabIndex        =   138
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
      TabIndex        =   137
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
      TabIndex        =   136
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���F"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   6240
      TabIndex        =   135
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�S����"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   134
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���s��"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   133
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�w�}�[��"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   132
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "PI000101"
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


Private TEHAI       As String


'�e�L�X�g�p�Y��
Private Const ptxSHIJI_NO% = 0              '�w�}�[��
Private Const ptxHAKKO_DT% = 1              '���s��
Private Const ptxTANTO_CODE% = 2            '�S���Һ���
Private Const ptxTANTO_NAME% = 3            '�S���Җ���
Private Const ptxSHONIN_CODE% = 4           '���F�Һ���
Private Const ptxSHONIN_NAME% = 5           '���F�Җ���
Private Const ptxHIN_GAI% = 6               '�i��
Private Const ptxHIN_NAME% = 7              '�i��
Private Const ptxSHIJI_QTY% = 8             '����
Private Const ptxST_LOCATION% = 9           '�W���I��
Private Const ptxMI_QTY% = 10               '�����i
Private Const ptxSUMI_QTY% = 11             '���i����
Private Const ptxUKEHARAI_CODE% = 12        '��z�溰��
Private Const ptxS_CLASS_CODE% = 13         '���i���׽
Private Const ptxF_CLASS_CODE% = 14         '�t���׽
Private Const ptxN_CLASS_CODE% = 15         '���E�׽


Private Const ptxK_HIN_GAI01% = 16          '�@�@�����އ�
Private Const ptxK_HIN_NAME01% = 17         '�@�@�����ޖ���
Private Const ptxK_QTY01% = 18              '�@�@����
Private Const ptxK_SHIJI_QTY01% = 19        '�@�@����
Private Const ptxK_ST_LOCATION01% = 20      '�@�@�I��

Private Const ptxK_HIN_GAI02% = 21          '�A�@�����އ�
Private Const ptxK_HIN_NAME02% = 22         '�A�@�����ޖ���
Private Const ptxK_QTY02% = 23              '�A�@����
Private Const ptxK_SHIJI_QTY02% = 24        '�A�@����
Private Const ptxK_ST_LOCATION02% = 25      '�A�@�I��

Private Const ptxK_HIN_GAI03% = 26          '�B�@�����އ�
Private Const ptxK_HIN_NAME03% = 27         '�B�@�����ޖ���
Private Const ptxK_QTY03% = 28              '�B�@����
Private Const ptxK_SHIJI_QTY03% = 29        '�B�@����
Private Const ptxK_ST_LOCATION03% = 30      '�B
Private Const ptxK_HIN_GAI04% = 31          '�C�@�����އ�
Private Const ptxK_HIN_NAME04% = 32         '�C�@�����ޖ���
Private Const ptxK_QTY04% = 33              '�C�@����
Private Const ptxK_SHIJI_QTY04% = 34        '�C�@����
Private Const ptxK_ST_LOCATION04% = 35      '�C�@�I��

Private Const ptxK_HIN_GAI05% = 36          '�D�@�����އ�
Private Const ptxK_HIN_NAME05% = 37         '�D�@�����ޖ���
Private Const ptxK_QTY05% = 38              '�D�@����
Private Const ptxK_SHIJI_QTY05% = 39        '�D�@����
Private Const ptxK_ST_LOCATION05% = 40      '�D�@�I��


Private Const ptxG_HIN_GAI01% = 41          '�@�@�O�����އ�
Private Const ptxG_HIN_NAME01% = 42         '�@�@�O�����ޖ���
Private Const ptxG_QTY01% = 43              '�@�@����
Private Const ptxG_SHIJI_QTY01% = 44        '�@�@����
Private Const ptxG_ST_LOCATION01% = 45      '�@�@�I��

Private Const ptxG_HIN_GAI02% = 46          '�A�@�O�����އ�
Private Const ptxG_HIN_NAME02% = 47         '�A�@�O�����ޖ���
Private Const ptxG_QTY02% = 48              '�A�@����
Private Const ptxG_SHIJI_QTY02% = 49        '�A�@����
Private Const ptxG_ST_LOCATION02% = 50      '�A�@�I��

Private Const ptxG_HIN_GAI03% = 51          '�B�@�O�����އ�
Private Const ptxG_HIN_NAME03% = 52         '�B�@�O�����ޖ���
Private Const ptxG_QTY03% = 53              '�B�@����
Private Const ptxG_SHIJI_QTY03% = 54        '�B�@����
Private Const ptxG_ST_LOCATION03% = 55      '�B�@�I��

Private Const ptxD_HIN_GAI01% = 56          '�@�@�����^�\���i��
Private Const ptxD_HIN_NAME01% = 57         '�@�@�����^�\���i��
Private Const ptxD_QTY01% = 58              '�@�@����
Private Const ptxD_SHIJI_QTY01% = 59        '�@�@����
Private Const ptxD_ST_LOCATION01% = 60      '�@�@�I��
Private Const ptxD_ZAIKO_QTY01% = 61        '�@�@�݌ɐ�
Private Const ptxD_BIKOU01% = 62            '�@�@���l

Private Const ptxD_HIN_GAI02% = 63          '�A�@�����^�\���i��
Private Const ptxD_HIN_NAME02% = 64         '�A�@�����^�\���i��
Private Const ptxD_QTY02% = 65              '�A�@����
Private Const ptxD_SHIJI_QTY02% = 66        '�A�@����
Private Const ptxD_ST_LOCATION02% = 67      '�A�@�I��
Private Const ptxD_ZAIKO_QTY02% = 68        '�A�@�݌ɐ�
Private Const ptxD_BIKOU02% = 69            '�A�@���l

Private Const ptxD_HIN_GAI03% = 70          '�B�@�����^�\���i��
Private Const ptxD_HIN_NAME03% = 71         '�B�@�����^�\���i��
Private Const ptxD_QTY03% = 72              '�B�@����
Private Const ptxD_SHIJI_QTY03% = 73        '�B�@����
Private Const ptxD_ST_LOCATION03% = 74      '�B�@�I��
Private Const ptxD_ZAIKO_QTY03% = 75        '�B�@�݌ɐ�
Private Const ptxD_BIKOU03% = 76            '�B�@���l

Private Const ptxD_HIN_GAI04% = 77          '�C�@�����^�\���i��
Private Const ptxD_HIN_NAME04% = 78         '�C�@�����^�\���i��
Private Const ptxD_QTY04% = 79              '�C�@����
Private Const ptxD_SHIJI_QTY04% = 80        '�C�@����
Private Const ptxD_ST_LOCATION04% = 81      '�C�@�I��
Private Const ptxD_ZAIKO_QTY04% = 82        '�C�@�݌ɐ�
Private Const ptxD_BIKOU04% = 83            '�C�@���l

Private Const ptxD_HIN_GAI05% = 84          '�D�@�����^�\���i��
Private Const ptxD_HIN_NAME05% = 85         '�D�@�����^�\���i��
Private Const ptxD_QTY05% = 86              '�D�@����
Private Const ptxD_SHIJI_QTY05% = 87        '�D�@����
Private Const ptxD_ST_LOCATION05% = 88      '�D�@�I��
Private Const ptxD_ZAIKO_QTY05% = 89        '�D�@�݌ɐ�
Private Const ptxD_BIKOU05% = 90            '�D�@���l

Private Const ptxD_HIN_GAI06% = 91          '�E�@�����^�\���i��
Private Const ptxD_HIN_NAME06% = 92         '�E�@�����^�\���i��
Private Const ptxD_QTY06% = 93              '�E�@����
Private Const ptxD_SHIJI_QTY06% = 94        '�E�@����
Private Const ptxD_ST_LOCATION06% = 95      '�E�@�I��
Private Const ptxD_ZAIKO_QTY06% = 96        '�E�@�݌ɐ�
Private Const ptxD_BIKOU06% = 97            '�E�@���l


Private Const ptxLabel_QTY% = 98            '���x�����s���� 2007.12.11




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

'�`�F�b�N�p�Y��
Private Const pchkSAMPLE_F% = 0             '���{�쐬
Private Const pchkPRI_SHIJI% = 1            '�o�͑Ώہ@�w�}�[
Private Const pchkPRI_PARTS% = 2            '�o�͑Ώہ@�߰�����
Private Const pchkPRI_GAISOU% = 3           '�o�͑Ώہ@�O������
Private Const pchkPRI_KISHU% = 4            '�o�͑Ώہ@�@������

Private Const pchkL_PAPER% = 5              '��             2010.11.12
Private Const pchkL_PLASTIC% = 6            '��׽���        2010.11.12
Private Const pchkL_LABEL% = 7              '�K�p�@������   2010.11.12

'��߼�����ݗp�Y��
Private Const poptSHIJI_NORMAL% = 0         '�ʏ�
Private Const poptSHIJI_SPOT% = 1           '�X�|�b�g
Private Const poptSHIJI_KEPPIN% = 2         '���i����


'���b�`�e�L�X�g�p�Y��
Private Const prchBIKOU% = 0                '���l



'�R�}���h�{�^���ŗL����
Private Const cmdMUPDATE% = 3               'Ͻ��X�V

Private Const cmdNext% = 5                  '�\�����i��ʂ�
Private Const cmdCen% = 10                  '������

Private GENSANKOKU_FLG  As String * 1       '���Y���@�󎚗L��   2008.06.13


Private wkSURYO         As Long             '208
Private chenge_F        As Boolean          '2008.07.30
Private svJGYOBU        As String * 1       '2008.07.30
Private svNAIGAI        As String * 1       '2008.07.30

Private svSHIMUKE       As String * 4       '2019.06.11 �ǉ�


Private GENSANKOKU_CHECK_TBL _
                        As Variant          '���Y�������L���i���ƕ��j 2009.03.28

Private L_GENSANKOKU    As String           '2009.03.28

Private KAISYA_CHK_F    As Boolean          '��Ё^���ƕ��̃G���[�����L�� 2010.07.20

Private KISHU_CHECK     As Boolean          '��\�@������� 2012.09.03

Private GAI_BUHIN_CHECK As Boolean          '�C�O�����敪�����L��   2016.02.01

Private TANKA_SPACE_F   As String           '2016.02.01

Private KAISYA_RESTRICT_F   As String


Private SHIMUKE_CHK_TBL As Variant          '�����i�@�d������   2013.08.29
Private svSHIMUKE_CODE  As String * 2       '2013.08.29

Private LABEL_PRINT_F       As Integer      '���x������f�t�H���g�\��   2019.03.07
Private GA_LABEL_PRINT_F    As Integer      '�O�����x������f�t�H���g�\��   2019.03.07

Dim L_print_Flg     As Boolean

'Private Const Last_Update_day$ = "���i���w�}�[���s (PI00010 2019.04.18 10:00)"
'Private Const Last_Update_day$ = "���i���w�}�[���s (PI00010 2019.04.18 11:45)"
'Private Const Last_Update_day$ = "���i���w�}�[���s (PI00010 2019.05.27 18:05)" '����
'Private Const Last_Update_day$ = "���i���w�}�[���s (PI00010 2019.05.28 13:50)" '����
'Private Const Last_Update_day$ = "���i���w�}�[���s (PI00010 2019.05.28 16:15)"  '����
'Private Const Last_Update_day$ = "���i���w�}�[���s (PI00010 2019.06.02 17:15)"  '����
'Private Const Last_Update_day$ = "���i���w�}�[���s (PI00010 2019.06.04 17:15)"  '����
'Private Const Last_Update_day$ = "���i���w�}�[���s (PI00010 2019.06.04 11:00)"  '����
'Private Const Last_Update_day$ = "���i���w�}�[���s (PI00010 2019.06.05 11:30)"  '����
'Private Const Last_Update_day$ = "���i���w�}�[���s (PI00010 2019.06.10 18:30)"  '����
'Private Const Last_Update_day$ = "���i���w�}�[���s (PI00010 2019.06.11 11:30)"  '����
'Private Const Last_Update_day$ = "���i���w�}�[���s (PI00010 2019.06.11 16:20)"  '����
'Private Const Last_Update_day$ = "���i���w�}�[���s (PI00010 2019.06.12 17:35)"  '����
'Private Const Last_Update_day$ = "���i���w�}�[���s (PI00010 2019.06.18 10:55)"  '����
'Private Const Last_Update_day$ = "���i���w�}�[���s (PI00010 2019.06.30 20:55)"  '����
'Private Const Last_Update_day$ = "���i���w�}�[���s (PI00010 2019.08.27 11:50)�e�X�g��"  '����   �i�ړ��͌�̏����ňꕔ�ύX
'Private Const Last_Update_day$ = "���i���w�}�[���s (PI00010 2019.08.28 10:25)"  '����   �i�ړ��͌�̏����ňꕔ�ύX
'Private Const Last_Update_day$ = "���i���w�}�[���s (PI00010 2019.09.24 13:35)"  '����   Init_Proc2�Ɉꕔ�ǉ�
'Private Const Last_Update_day$ = "���i���w�}�[���s (PI00010 2019.11.07 15:30) �o�ɗ\��o�[�R�[�h�Ή�"
'Private Const Last_Update_day$ = "���i���w�}�[���s (PI00010 2019.12.18 12:00) �o�͑ΏۑO��\�����c�錏���C��(�K�p�@�탉�x��)"
Private Const Last_Update_day$ = "���i���w�}�[���s (PI00010 2019.12.18 16:30) �����i���p�[�c���x���Ȃ��֑ؑΉ�"

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    PI000101.MousePointer = vbHourglass

    Call Ctrl_Lock(PI000101)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PI000101)


    PI000101.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer, chk As Integer, flg, Optional opt As Integer = 0) As Integer
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

Dim com         As Integer

Dim wkTanto     As String

Dim L           As Integer  '2011.02.10

Dim m           As Integer  '2013.01.17


Dim Shimuke_flg    As Integer  '2013.09.04

Dim wkGENSANKOKU    As String * 20  '2015.10.09


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

                    If Text1(ptxSHIJI_NO).text = StrConv(P_SSHIJI_O_REC.SHIJI_No, vbUnicode) Then
                    Else

Start_Proc1:        '2015.03.13


                        chenge_F = False
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
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                If sts > 3000 Or sts = 3 Then
            
                
                                    Call File_Error(sts, BtOpGetEqual, "�w�}�[(�e)", 0)
                                    '>>>>>>>>>>>>>  2015.04.24
                                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                    'If sts Then
                                    '    Call File_Error(sts, BtOpReset, "")
                                    'End If
                                
                                
                                    'Call File_Open_Proc
                                    Do
                                        If Not File_Open_Proc() Then
                                            Exit Do
                                        End If
                                    Loop
                                    '>>>>>>>>>>>>>  2015.04.24
                    
                                
                                    GoTo Start_Proc1
                                End If
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                Exit Function
                        End Select
                        Text1(Mode).SetFocus        '2008.01.15

                    End If
                End If


                '=========================================== 2007/03/19 =====
''                Text1(ptxSHIJI_NO).BackColor = G_INPUT_NG
''                Text1(ptxSHIJI_NO).Locked = True
''                Text1(ptxSHIJI_NO).TabStop = False
                '============================================================



            End If

        Case ptxHAKKO_DT    '���s��

            If chk = 1 Then
            Else
                If Not IsDate(Text1(ptxHAKKO_DT).text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(���s��)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxHAKKO_DT).text = Format(CDate(Text1(ptxHAKKO_DT).text), "YYYY/MM/DD")
                End If
            End If

        Case ptxTANTO_CODE      '�S����

           If chk = 1 Then
            Else
                
Start_Proc2:    '2015.03.13
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
                        
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc2
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        
                        
                        
                        
                        
                        Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
                        Exit Function

                End Select
            End If

        Case ptxSHONIN_CODE     '���F��

            If chk = 1 Then
            Else
                
                
Start_Proc3:    '2015.03.13
                
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
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc3
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        
                        Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
                        Exit Function



                End Select
            End If
        Case ptxHIN_GAI         '�i��


            '========================================================= 2007/03/19 =====

Start_Proc4:    '2015.03.13

            chenge_F = False
            
            
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


                    lblL_Hin_Name_E.Caption = ""        '2016.02.10
                    lblGAI_BUHIN.Caption = ""           '2016.02.10
                    lblL_URIKIN2.Caption = ""           '2016.02.10
                    lblL_URIKIN3.Caption = ""           '2016.02.10

'2010.11.12 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                    Check1(pchkL_PAPER).Value = vbUnchecked         '��
                    Check1(pchkL_PLASTIC).Value = vbUnchecked       '�v��
                    Check1(pchkL_LABEL).Value = vbUnchecked         '�K�p�@�탉�x��
'2010.11.12 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> MT_2009.06.01
                    'MsgBox "���͂������ڂ̓G���[�ł��B(�i��)"
                    'Text1(Mode).SetFocus
                    'Exit Function

                    lblGensankoku(0).Caption = ""
                    lblGensankoku(1).Caption = ""


                    If Trim(Text1(ptxHIN_GAI).text) = "" Then
                        MsgBox "���͂������ڂ̓G���[�ł��B(�i��)"
                        Text1(Mode).SetFocus
                        Exit Function
                    End If

                    wkTanto = Text1(ptxTANTO_CODE)
                    If Trim(wkTanto) = "" Then
                        wkTanto = "PSHIJ"
                    End If

                    Last_JGYOBU = Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1)
                    If PN_CHK(Text1(Mode), "G", wkTanto, 1) Then          '�O���i��
                        Text1(Mode).SetFocus
                        Call Text1_GotFocus(Mode)
                        Exit Function
                    End If

                    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


                    lblL_KAISHA.Caption = ""
                    lblL_JGYOBU.Caption = ""
                    

                    lblKISHU1.Caption = ""          '2016.02.01
                    lblKISHU2.Caption = ""          '2016.02.01
                Case Else
                    
                    
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24

                        
                            GoTo Start_Proc4
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                    
                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                    Exit Function

            End Select

'''            yn = False
'''
'''            For i = 0 To Combo1(pcmbSHIMUKE).ListCount - 1
'''                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).List(i), 4), 3, 1))
'''                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).List(i), 4), 4, 1))
'''                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).text)
'''
'''
'''                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'''                Select Case sts
'''                    Case BtNoErr
'''                        Combo1(pcmbSHIMUKE).ListIndex = i
'''                        yn = True
'''                        Exit For
'''                    Case BtErrKeyNotFound
'''
'''                    Case Else
'''                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
'''                        Exit Function
'''
'''                End Select
'''            Next i
'''
'''            If yn = False Then
'''                Text1(ptxHIN_NAME).text = ""
'''                Text1(ptxST_LOCATION).text = ""
'''                Text1(ptxMI_QTY).text = ""
'''                Text1(ptxSUMI_QTY).text = ""
'''
'''                MsgBox "���͂������ڂ̓G���[�ł��B(�i��)"
'''                Text1(Mode).SetFocus
'''                Exit Function
'''            End If

            '==========================================================================

            Text1(ptxHIN_NAME).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)


            lblL_Hin_Name_E.Caption = StrConv(ITEMREC.L_HIN_NAME_E, vbUnicode)  '2016.02.10

            lblGAI_BUHIN.Caption = StrConv(ITEMREC.GAI_BUHIN, vbUnicode)        '2016.02.10
            lblL_URIKIN2.Caption = StrConv(ITEMREC.L_URIKIN2, vbUnicode)        '2016.02.10
            lblL_URIKIN3.Caption = StrConv(ITEMREC.L_URIKIN3, vbUnicode)        '2016.02.10

'2013.09.04 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            Shimuke_flg = False
'            If svSHIMUKE_CODE <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Then '2019/12/18 �R�����g�A�E�g
                For i = 0 To UBound(SHIMUKE_CHK_TBL)
                
                    If SHIMUKE_CHK_TBL(i) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Then
                        Shimuke_flg = True
                        Exit For
                    End If
                
                Next i
'            End If
'2013.09.04 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'2010.11.12 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            'If opt <> 9 Then                   '2013.09.04
            If opt <> 9 And Not Shimuke_flg Then      '2013.09.04
                If StrConv(ITEMREC.L_PAPER, vbUnicode) = L_PAPER_ON Then        '��
                    Check1(pchkL_PAPER).Value = vbChecked
                Else
                    Check1(pchkL_PAPER).Value = vbUnchecked
                End If
    
                If StrConv(ITEMREC.L_PLASTIC, vbUnicode) = L_PLASTIC_ON Then    '�v��
                    Check1(pchkL_PLASTIC).Value = vbChecked
                Else
                    Check1(pchkL_PLASTIC).Value = vbUnchecked
                End If
    
                If StrConv(ITEMREC.L_LABEL, vbUnicode) = L_LABEL_ON Then        '�K�p�@�탉�x��
                    Check1(pchkL_LABEL).Value = vbChecked
                Else
                    Check1(pchkL_LABEL).Value = vbUnchecked
                End If
            
'                If LABEL_PRINT_F = 1 Then          '2019/12/18 <> 1 �� = 1 �֕ύX
                    '2011.02.10
'                    If Trim(StrConv(ITEMREC.L_LABEL, vbUnicode)) = "" Then
'                        Combo2(0).ListIndex = 1
                    '2019.08.27 �݌����񂩂�̎w���ŉ��L�Ƃ����B�Ƃ������A�e�X�g��
'                    If Trim(StrConv(ITEMREC.L_LABEL, vbUnicode)) = "" Or Trim(StrConv(ITEMREC.L_LABEL, vbUnicode)) = "0" Then
'                        Combo2(0).ListIndex = 1
'                    Else
                        For L = 1 To Combo2(0).ListCount
                            If StrConv(ITEMREC.L_LABEL, vbUnicode) = Right(Combo2(0).List(L), 1) Then
                                Combo2(0).ListIndex = L
                                Exit For
                            
                            End If
                        Next L
'                    End If
'               End If                              '2019.03.07
                
                
                If Trim(StrConv(ITEMREC.L_LABEL, vbUnicode)) = "" Or Trim(StrConv(ITEMREC.L_LABEL, vbUnicode)) = "2" Then
                    Check1(pchkL_PAPER).Enabled = False
                    Check1(pchkL_PLASTIC).Enabled = False
                Else
                    Check1(pchkL_PAPER).Enabled = True
                    Check1(pchkL_PLASTIC).Enabled = True
                End If
                
                '2011.02.10
            
            
            
            
            End If
'2010.11.12 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<



            If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                Text1(ptxST_LOCATION).text = ""
            Else
                Text1(ptxST_LOCATION).text = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_DAN, vbUnicode)
            End If

            If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                Exit Function

            End If

            Text1(ptxMI_QTY).text = Format(Mi_Qty, "#0")
            Text1(ptxSUMI_QTY).text = Format(Sumi_Qty, "#0")



            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    TORI_GENSANKOKU�̗L���`�F�b�N����������   2012.01.31

'            chk_TORI_GENSANKOKU = StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)           '���Y���L�������p   2013.01.08

            
            '>>>>>>>>>>>>>>>>2015.10.09
            wkGENSANKOKU = StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)
            For m = 1 To 20
                If Mid(wkGENSANKOKU, m, 1) < " " Then
                    Mid(wkGENSANKOKU, m, 1) = " "
                End If
            Next m
            '>>>>>>>>>>>>>>>>2015.10.09
            
'            If Trim(StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)) = "" Then     '2015.10.09
            If Trim(wkGENSANKOKU) = "" Then                                     '2015.10.09
            Else
                
Start_Proc5:        '2015.03.13
                
                Call UniCode_Conv(K0_GENSAN.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_GENSAN.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_GENSAN.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                Call UniCode_Conv(K0_GENSAN.GENSANKOKU, StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode))
                
                sts = BTRV(BtOpGetEqual, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                    
                        Call UniCode_Conv(GENSANREC.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                        Call UniCode_Conv(GENSANREC.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                        Call UniCode_Conv(GENSANREC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                        Call UniCode_Conv(GENSANREC.GENSANKOKU, StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode))
                        Call UniCode_Conv(GENSANREC.FILLER, "")
                
                        Call UniCode_Conv(GENSANREC.INS_TANTO, "PI010")
                        Call UniCode_Conv(GENSANREC.Ins_DateTime, Format(Now, "YYYYMMDDHHMMSS"))
                
                        Call UniCode_Conv(GENSANREC.UPD_TANTO, "")
                        Call UniCode_Conv(GENSANREC.UPD_DATETIME, "")
                    
                    
                        sts = BTRV(BtOpInsert, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                        Select Case sts
                        
                            Case BtNoErr
                            Case BtErrDuplicates
                            Case Else
                                
                                
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                If sts > 3000 Or sts = 3 Then
            
                
                                    Call File_Error(sts, BtOpGetEqual, "���Y���}�X�^", 0)
                                    '>>>>>>>>>>>>>  2015.04.24
                                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                    'If sts Then
                                    '    Call File_Error(sts, BtOpReset, "")
                                    'End If
                                
                                
                                    'Call File_Open_Proc
                                    Do
                                        If Not File_Open_Proc() Then
                                            Exit Do
                                        End If
                                    Loop
                                    '>>>>>>>>>>>>>  2015.04.24
                    
                                
                                    GoTo Start_Proc5
                                End If
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                
                                
                                
                                Call File_Error(sts, com, "���Y���}�X�^")
                                Exit Function
                        End Select
                    
                    
                    
                    
                    Case Else
                        Exit Function
                End Select
            End If
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    TORI_GENSANKOKU�̗L���`�F�b�N����������   2012.01.31







            txGensankoku.text = Trim(StrConv(ITEMREC.GENSANKOKU, vbUnicode))            '2009.03.28
            chk_TORI_GENSANKOKU = Trim(StrConv(ITEMREC.GENSANKOKU, vbUnicode))          '���Y���L�������p   2013.01.08


            For m = 1 To Len(chk_TORI_GENSANKOKU)
                If Mid(chk_TORI_GENSANKOKU, m, 1) < " " Then
                    chk_TORI_GENSANKOKU = ""
                End If
            Next m
            

            '2010.07.20 ��
            
Start_Proc6:        '2013.03.13
            
            lstGensankoku.Clear

            Call UniCode_Conv(K0_GENSAN.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_GENSAN.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_GENSAN.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
            Call UniCode_Conv(K0_GENSAN.GENSANKOKU, "")

            com = BtOpGetGreater

            Do

                DoEvents

                sts = BTRV(com, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                Select Case sts
                    Case BtNoErr
                        If StrConv(ITEMREC.JGYOBU, vbUnicode) <> StrConv(GENSANREC.JGYOBU, vbUnicode) Or _
                            StrConv(ITEMREC.NAIGAI, vbUnicode) <> StrConv(GENSANREC.NAIGAI, vbUnicode) Or _
                            StrConv(ITEMREC.HIN_GAI, vbUnicode) <> StrConv(GENSANREC.HIN_GAI, vbUnicode) Then
                            Exit Do
                        End If
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "�w�}�[(�e)", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc6
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        
                        
                        
                        Exit Function
                End Select
                
                'lstGensankoku.AddItem StrConv(GENSANREC.Ins_DateTime, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)    2013.02.19

'                If Trim(StrConv(GENSANREC.UPD_DATETIME, vbUnicode)) = "" Then
'                    lstGensankoku.AddItem StrConv(GENSANREC.Ins_DateTime, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)     '2013.02.19    2014.02.18
'                Else                                                                                                                '2013.02.19    2014.02.18
'                    lstGensankoku.AddItem StrConv(GENSANREC.UPD_DATETIME, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)     '2013.02.19    2014.02.18
'                End If


                If StrConv(GENSANREC.UPD_DATETIME, vbUnicode) > StrConv(GENSANREC.Ins_DateTime, vbUnicode) Then                     '2014.02.18
                    lstGensankoku.AddItem StrConv(GENSANREC.UPD_DATETIME, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)     '2014.02.18
                Else                                                                                                                '2014.02.18
                    lstGensankoku.AddItem StrConv(GENSANREC.Ins_DateTime, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)     '2014.02.18
                End If



                com = BtOpGetNext
            Loop

            lblGensankoku(0).Caption = ""
            lblGensankoku(1).Caption = ""
            If lstGensankoku.ListCount < 1 Then
                lblGensankoku(1).Caption = Trim(StrConv(ITEMREC.GENSANKOKU, vbUnicode))
                If Trim(lblGensankoku(1).Caption) <> "" Then
                    lblGensankoku(0).Caption = ""
                End If
            Else

                lblGensankoku(1).Caption = Right(lstGensankoku.List(lstGensankoku.ListCount - 1), 20)
                lblGensankoku(0).Caption = "��" & StrConv(Format(lstGensankoku.ListCount, "#0"), vbWide)
            End If
            txGensankoku.text = Trim(lblGensankoku(1).Caption)


            lblL_KAISHA.Caption = Trim(StrConv(ITEMREC.L_KAISHA_CODE, vbUnicode))
            lblL_JGYOBU.Caption = Trim(StrConv(ITEMREC.L_JGYOBU_CODE, vbUnicode))



            '���̃Z�b�g2016.02.01
Start_Proc6_1:        '2016.02.01
            Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN07_CD)
            Call UniCode_Conv(K0_P_CODE.C_Code, lblL_KAISHA.Caption)
            
            


            sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            Select Case sts
                Case BtNoErr
                    lblL_KAISHA_N.Caption = StrConv(P_CODEREC.C_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    lblL_KAISHA_N.Caption = lblL_KAISHA.Caption
                Case Else
                    If sts > 3000 Or sts = 3 Then

    
                        Call File_Error(sts, BtOpGetEqual, "����Ͻ�", 0)
                        Do
                            If Not File_Open_Proc() Then
                                Exit Do
                            End If
                        Loop
                        GoTo Start_Proc6_1
                    End If
                    Exit Function
            End Select
            
            
Start_Proc6_2:        '2016.02.01
            Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN07_CD)
            Call UniCode_Conv(K0_P_CODE.C_Code, lblL_JGYOBU.Caption)
            
            
            sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            Select Case sts
                Case BtNoErr
                     lblL_JGYOBU_N.Caption = StrConv(P_CODEREC.C_NAME, vbUnicode)
                Case BtErrKeyNotFound
                     lblL_JGYOBU_N.Caption = lblL_JGYOBU.Caption
                Case Else
                    If sts > 3000 Or sts = 3 Then

    
                        Call File_Error(sts, BtOpGetEqual, "����Ͻ�", 0)
                        Do
                            If Not File_Open_Proc() Then
                                Exit Do
                            End If
                        Loop
                        GoTo Start_Proc6_2
                    End If
                    Exit Function
            End Select
            
            
            
            
            '���̃Z�b�g2016.02.01



            '2010.07.20 ��




            lblKISHU1.Caption = Trim(StrConv(ITEMREC.L_KISHU1, vbUnicode))          '2016.02.01
            lblKISHU2.Caption = Trim(StrConv(ITEMREC.L_KISHU2, vbUnicode))          '2016.02.01








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
                        Text1(Mode).SetFocus         '2008.01.15
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
                        Text1(Mode).SetFocus         '2008.01.15


                    End If
                End If
            End If
            
            '2019.06.10
            Text1(ptxSHIJI_QTY).SetFocus
            
        Case ptxSHIJI_QTY       '����

            If chk = 1 Then
            Else
                If Not IsNumeric(Text1(ptxSHIJI_QTY).text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(����)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxSHIJI_QTY).text = Format(CLng(Text1(ptxSHIJI_QTY).text), "#0")


                    If Trim(Text1(ptxLabel_QTY).text) = "" Then '2008.02.06
'                        Text1(ptxLabel_QTY).text = Format(CLng(Text1(ptxSHIJI_QTY).text) + 1, "#0")                '2015.04.02
                        Text1(ptxLabel_QTY).text = Format(CLng(Text1(ptxSHIJI_QTY).text) + LABEL_PLUS, "#0")        '2015.04.02
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

Start_Proc9:        '2015.03.13
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
                        
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "���i���׽", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc9
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        
                        
                        
                        Call File_Error(sts, BtOpGetEqual, "���i���׽")
                        Exit Function

                End Select
            End If
        Case ptxF_CLASS_CODE    '�t���׽

            If Trim(Text1(ptxF_CLASS_CODE).text) = "" Then
            Else
                
                
Start_Proc10:       '2015.03.13
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
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "���i���׽", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc10
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        
                        
                        Call File_Error(sts, BtOpGetEqual, "���i���׽")
                        Exit Function

                End Select
            End If

        Case ptxN_CLASS_CODE    '���E�׽

            If Trim(Text1(ptxN_CLASS_CODE).text) = "" Then
            Else
                
Start_Proc11:       '2015.03.13
                
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
                        
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "���i���׽", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc11
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        
                        
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
                
Start_Proc12:       '2015.03.13
                
                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1))
                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)

                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        '���ޕi�œǂݑւ�
Start_Proc12_2:       '2015.03.13

                        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)

                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound

                                If HIN_INV Then
                                    Call Rclr_ITEMREC                               '2019.06.02 �P�s�ǉ��i����j
                                    '���o�^�i�ԁ@�@���ނƂ��Ă���
                                    Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)
                                    Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
                                    Call UniCode_Conv(ITEMREC.HIN_NAME, "���o�^�i��")
                                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                                Else
                                    MsgBox "���͂������ڂ̓G���[�ł��B(�����ށ@�i��)"
                                    Text1(Mode).SetFocus
                                    Exit Function
                                End If
                            Case Else
                                
                                
                                
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                If sts > 3000 Or sts = 3 Then
            
                
                                    Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�", 0)
                                    '>>>>>>>>>>>>>  2015.04.24
                                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                    'If sts Then
                                    '    Call File_Error(sts, BtOpReset, "")
                                    'End If
                                
                                
                                    'Call File_Open_Proc
                                    Do
                                        If Not File_Open_Proc() Then
                                            Exit Do
                                        End If
                                    Loop
                                    '>>>>>>>>>>>>>  2015.04.24
                    
                                
                                    GoTo Start_Proc12_2
                                End If
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                
                                
                                Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                                Exit Function

                        End Select

                    Case Else
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc12
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        
                        
                        
                        Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                        Exit Function

                End Select

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
                If Trim(Text1(Mode - 2).text) <> "" Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(�����ށ@����)"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            Else
                If Trim(Text1(Mode - 2).text) = "" Then
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
                
Start_Proc13:   '2015.03.13
                
                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1))
                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)

                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        '���ޕi�œǂݑւ�
Start_Proc14:   '2015.03.13
                        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)

                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound

                                If HIN_INV Then
                                    Call Rclr_ITEMREC                               '2019.06.02 �P�s�ǉ��i����j
                                    '���o�^�i�ԁ@�@���ނƂ��Ă���
                                    Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)
                                    Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
                                    Call UniCode_Conv(ITEMREC.HIN_NAME, "���o�^�i��")
                                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                                Else

                                    MsgBox "���͂������ڂ̓G���[�ł��B(�O�����ށ@�i��)"
                                    Text1(Mode).SetFocus
                                    Exit Function
                                End If
                            Case Else
                                
                                
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                If sts > 3000 Or sts = 3 Then
            
                
                                    Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�", 0)
                                    '>>>>>>>>>>>>>  2015.04.24
                                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                    'If sts Then
                                    '    Call File_Error(sts, BtOpReset, "")
                                    'End If
                                    'Call File_Open_Proc
                                    Do
                                        If Not File_Open_Proc() Then
                                            Exit Do
                                        End If
                                    Loop
                                    '>>>>>>>>>>>>>  2015.04.24
                    
                                
                                    GoTo Start_Proc14
                                End If
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                
                                Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                                Exit Function

                        End Select

                    Case Else
                        
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc13
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        
                        
                        
                        Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                        Exit Function

                End Select

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
                
Start_Proc15:       '2015.03.13
                
                
                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1))
                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)

                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
Start_Proc16:       '2015.03.13

                        '�i�ԁi���j�œǂݑւ�
                        Call UniCode_Conv(K2_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1))
                        Call UniCode_Conv(K2_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1))
                        Call UniCode_Conv(K2_ITEM.HIN_NAI, Text1(Mode).text)

                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound


Start_Proc17:       '2015.03.13




                                '���ޕi�œǂݑւ�

                                Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)

                                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                Select Case sts
                                    Case BtNoErr
                                    Case BtErrKeyNotFound

                                        If HIN_INV Then
                                            Call Rclr_ITEMREC                               '2019.06.02 �P�s�ǉ��i����j
                                            '���o�^�i�ԁ@�@���ނƂ��Ă���
                                            Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)
                                            Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
                                            Call UniCode_Conv(ITEMREC.HIN_GAI, Text1(Mode).text)
                                            Call UniCode_Conv(ITEMREC.HIN_NAME, "���o�^�i��")
                                            Call UniCode_Conv(ITEMREC.ST_SOKO, "")

                                        Else

                                            MsgBox "���͂������ڂ̓G���[�ł��B(�����^�\���@�i��)"
                                            Text1(Mode).SetFocus
                                            Exit Function
                                        End If
                                    Case Else
                                        
                                        
                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                        If sts > 3000 Or sts = 3 Then
                    
                        
                                            Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�", 0)
                                            '>>>>>>>>>>>>>  2015.04.24
                                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                            'If sts Then
                                            '    Call File_Error(sts, BtOpReset, "")
                                            'End If
                                        
                                        
                                            'Call File_Open_Proc
                                            Do
                                                If Not File_Open_Proc() Then
                                                    Exit Do
                                                End If
                                            Loop
                                            '>>>>>>>>>>>>>  2015.04.24
                            
                                        
                                            GoTo Start_Proc17
                                        End If
                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                        
                                        
                                        Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                                        Exit Function

                                End Select

                            Case Else
                                
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                If sts > 3000 Or sts = 3 Then
            
                
                                    Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�", 0)
                                    '>>>>>>>>>>>>>  2015.04.24
                                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                    'If sts Then
                                    '    Call File_Error(sts, BtOpReset, "")
                                    'End If
                                
                                
                                    'Call File_Open_Proc
                                    Do
                                        If Not File_Open_Proc() Then
                                            Exit Do
                                        End If
                                    Loop
                                    '>>>>>>>>>>>>>  2015.04.24
                    
                                
                                    GoTo Start_Proc16
                                End If
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                
                                Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                                Exit Function
                       End Select

                    Case Else
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc15
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                        Exit Function

                End Select

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

                '�݌ɐ�
                If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                        StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                        StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
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

                D_Item_Tbl(i).JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
                D_Item_Tbl(i).NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
                D_Item_Tbl(i).HIN_GAI = StrConv(ITEMREC.HIN_GAI, vbUnicode)


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
                End If
            End If
                
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

        Case ptxLabel_QTY       '���x�����s���� 2007.12.11

            If chk = 1 Then
            Else
                If Not IsNumeric(Text1(ptxSHIJI_QTY).text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(���x�����s����)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxLabel_QTY).text = Format(CLng(Text1(ptxLabel_QTY).text), "#0")
                    If CLng(Text1(ptxLabel_QTY).text) <= 0 Then
                        MsgBox "���͂������ڂ̓G���[�ł��B(���x�����s����)"
                        Text1(Mode).SetFocus
                        Exit Function
                    End If
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


Dim L           As Integer  '2011.02.10

Dim Wk_LOC      As String   '2013.01.07


Dim Ret_sts     As Integer  '2015.03.13

Dim Zaiko_sts   As Integer  '2015.03.13


    Item_Disp_Proc = True

    Call Input_Lock         '2008.01.15

    For i = ptxK_HIN_GAI01 To ptxD_BIKOU06
        Text1(i).text = ""
    Next i

                                '2008.07.30
    For i = pcmbD_SYUBETSU01 To pcmbD_SYUBETSU06

            Combo1(i).ListIndex = -1

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

    Text1(ptxLabel_QTY).text = ""       '2008.02.27


    '--------------------------------   �u�e�v���


    Text1(ptxSHIJI_NO).text = StrConv(P_SSHIJI_O_REC.SHIJI_No, vbUnicode)           '�w�}�[��
                                                                                    '���s��
    Text1(ptxHAKKO_DT).text = Mid(StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 7, 2)

    Text1(ptxTANTO_CODE).text = StrConv(P_SSHIJI_O_REC.TANTO_CODE, vbUnicode)       '�S���Һ��ށ^����
    
    
Start_Proc1:    '2015.03.13
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxTANTO_CODE).text)

    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
            Text1(ptxTANTO_NAME).text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Text1(ptxTANTO_NAME).text = ""
        Case Else
            
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
            If sts > 3000 Or sts = 3 Then


                Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^", 0)
                '>>>>>>>>>>>>>  2015.04.24
                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                'If sts Then
                '    Call File_Error(sts, BtOpReset, "")
                'End If
            
            
                'Call File_Open_Proc
                Do
                    If Not File_Open_Proc() Then
                        Exit Do
                    End If
                Loop
                '>>>>>>>>>>>>>  2015.04.24

            
                GoTo Start_Proc1
            End If
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
            
            
            Call Input_UnLock         '2008.01.15
            Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
            Exit Function

    End Select

    Text1(ptxSHONIN_CODE).text = StrConv(P_SSHIJI_O_REC.SHONIN_CODE, vbUnicode)     '���F�Һ��ށ^����
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxSHONIN_CODE).text)

    
Start_Proc2:        '2015.03.13
    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
            Text1(ptxSHONIN_NAME).text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Text1(ptxSHONIN_NAME).text = ""
        Case Else
            
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
            If sts > 3000 Or sts = 3 Then


                Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^", 0)
                '>>>>>>>>>>>>>  2015.04.24
                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                'If sts Then
                '    Call File_Error(sts, BtOpReset, "")
                'End If
            
            
                'Call File_Open_Proc
                Do
                    If Not File_Open_Proc() Then
                        Exit Do
                    End If
                Loop
                '>>>>>>>>>>>>>  2015.04.24

            
                GoTo Start_Proc2
            End If
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
            
            
            Call Input_UnLock         '2008.01.15
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

Start_Proc3:    '2015.03.13

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





            lblKISHU1.Caption = Trim(StrConv(ITEMREC.L_KISHU1, vbUnicode))          '2012.10.26
            lblKISHU2.Caption = Trim(StrConv(ITEMREC.L_KISHU2, vbUnicode))          '2012.10.26


'>>>>>>>>>>>>>>>>>> 2015.03.13
'            If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
'                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
'                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode), , , , , , Ret_sts) Then
'                Exit Function
'
'            End If


Start_Proc4:        '2015.03.13
            Zaiko_sts = Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode), , , , , , Ret_sts, 0)
             If Zaiko_sts Then
                If Ret_sts > 3000 Or Ret_sts = 3 Then


                    Call File_Error(Ret_sts, BtOpGetEqual, "�݌��ް�", 0)
                    '>>>>>>>>>>>>>  2015.04.24
                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                    'If sts Then
                    '    Call File_Error(sts, BtOpReset, "")
                    'End If
                
                
                    'Call File_Open_Proc
                    Do
                        If Not File_Open_Proc() Then
                            Exit Do
                        End If
                    Loop
                    '>>>>>>>>>>>>>  2015.04.24
                
                    GoTo Start_Proc4
                End If
                Call File_Error(Ret_sts, BtOpGetEqual, "�݌��ް�")
                Exit Function
            End If
'>>>>>>>>>>>>>>>>>> 2015.03.13

            Text1(ptxMI_QTY).text = Format(Mi_Qty, "#0")
            Text1(ptxSUMI_QTY).text = Format(Sumi_Qty, "#0")


'2010.11.12 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            If StrConv(ITEMREC.L_PAPER, vbUnicode) = L_PAPER_ON Then        '��
                Check1(pchkL_PAPER).Value = vbChecked
            Else
                Check1(pchkL_PAPER).Value = vbUnchecked
            End If

            If StrConv(ITEMREC.L_PLASTIC, vbUnicode) = L_PLASTIC_ON Then    '�v��
                Check1(pchkL_PLASTIC).Value = vbChecked
            Else
                Check1(pchkL_PLASTIC).Value = vbUnchecked
            End If

            If StrConv(ITEMREC.L_LABEL, vbUnicode) = L_LABEL_ON Then        '�K�p�@�탉�x��
                Check1(pchkL_LABEL).Value = vbChecked
            Else
                Check1(pchkL_LABEL).Value = vbUnchecked
            End If
'2010.11.12 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

            '2011.02.10
'            If Trim(StrConv(ITEMREC.L_LABEL, vbUnicode)) = "" Then
'                Combo2(0).ListIndex = 1
            '2019.08.28 �������L�ɕύX
            If Trim(StrConv(ITEMREC.L_LABEL, vbUnicode)) = "" Or Trim(StrConv(ITEMREC.L_LABEL, vbUnicode)) = "0" Then
'''                Combo2(0).ListIndex = 1
            
            Else
                For L = 1 To Combo2(0).ListCount
                
                
                    If StrConv(ITEMREC.L_LABEL, vbUnicode) = Right(Combo2(0).List(L), 1) Then
                
                        Combo2(0).ListIndex = L
                        Exit For
                    
                    End If
                Next L
            
            
            
            
            
            End If
            
            
            
            
            
            
            '2011.02.10




        Case BtErrKeyNotFound
            Text1(ptxHIN_NAME).text = ""
            Text1(ptxST_LOCATION).text = ""
            Text1(ptxMI_QTY).text = ""
            Text1(ptxSUMI_QTY).text = ""

'2010.11.12 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            Check1(pchkL_PAPER).Value = vbUnchecked         '��
            Check1(pchkL_PLASTIC).Value = vbUnchecked       '�v��
            Check1(pchkL_LABEL).Value = vbUnchecked         '�K�p�@�탉�x��
'2010.11.12 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

            lblKISHU1.Caption = ""                      '2012.10.26
            lblKISHU2.Caption = ""                      '2012.10.26


        Case Else
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
          
            If sts > 3000 Or sts = 3 Then


                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                '>>>>>>>>>>>>>  2015.04.24
                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                'If sts Then
                '    Call File_Error(sts, BtOpReset, "")
                'End If
            
            
                'Call File_Open_Proc
                Do
                    If Not File_Open_Proc() Then
                        Exit Do
                    End If
                Loop
                '>>>>>>>>>>>>>  2015.04.24

            
                GoTo Start_Proc3
            End If
            
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
            
            
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
                                                                                    '�o�͑Ώہ@�߰����� 2010.07.20
    If StrConv(P_SSHIJI_O_REC.PRI_PARTS, vbUnicode) = P_PRI_PARTS_OFF Or Trim(StrConv(P_SSHIJI_O_REC.PRI_PARTS, vbUnicode)) = "" Then
        Check1(pchkPRI_PARTS).Value = vbUnchecked
    Else
        Check1(pchkPRI_PARTS).Value = vbChecked
    End If


    RichTextBox1(prchBIKOU).text = StrConv(P_SSHIJI_O_REC.BIKOU, vbUnicode)         '���l




    txGensankoku.text = Trim(StrConv(ITEMREC.GENSANKOKU, vbUnicode))                '2009.03.28



    '2010.07.20 ��
    
Start_Proc4_2:           '2015.03.13
    
    lstGensankoku.Clear

    Call UniCode_Conv(K0_GENSAN.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_GENSAN.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_GENSAN.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
    Call UniCode_Conv(K0_GENSAN.GENSANKOKU, "")

    com = BtOpGetGreater

    Do

        DoEvents

        sts = BTRV(com, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(ITEMREC.JGYOBU, vbUnicode) <> StrConv(GENSANREC.JGYOBU, vbUnicode) Or _
                    StrConv(ITEMREC.NAIGAI, vbUnicode) <> StrConv(GENSANREC.NAIGAI, vbUnicode) Or _
                    StrConv(ITEMREC.HIN_GAI, vbUnicode) <> StrConv(GENSANREC.HIN_GAI, vbUnicode) Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                
                
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                If sts > 3000 Or sts = 3 Then


                    Call File_Error(sts, BtOpGetEqual, "���Y��Ͻ�", 0)
                    '>>>>>>>>>>>>>  2015.04.24
                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                    'If sts Then
                    '    Call File_Error(sts, BtOpReset, "")
                    'End If
                
                
                    'Call File_Open_Proc
                    Do
                        If Not File_Open_Proc() Then
                            Exit Do
                        End If
                    Loop
                    '>>>>>>>>>>>>>  2015.04.24
    
                
                    GoTo Start_Proc4_2
                End If
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                
                
                
                Exit Function
        End Select


        'lstGensankoku.AddItem StrConv(GENSANREC.Ins_DateTime, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)        2013.01.28
'        If Trim(StrConv(GENSANREC.UPD_DATETIME, vbUnicode)) = "" Then                                                       '2013.01.28     2014.02.18
'            lstGensankoku.AddItem StrConv(GENSANREC.Ins_DateTime, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)     '2013.01.28     2014.02.18
'        Else                                                                                                                '2013.01.28     2014.02.18
'            lstGensankoku.AddItem StrConv(GENSANREC.UPD_DATETIME, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)     '2013.01.28     2014.02.18
'        End If

        If StrConv(GENSANREC.UPD_DATETIME, vbUnicode) > StrConv(GENSANREC.Ins_DateTime, vbUnicode) Then                     '2014.02.18
            lstGensankoku.AddItem StrConv(GENSANREC.UPD_DATETIME, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)     '2014.02.18
        Else                                                                                                                '2014.02.18
            lstGensankoku.AddItem StrConv(GENSANREC.Ins_DateTime, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)     '2014.02.18
        End If


        com = BtOpGetNext
    Loop

    lblGensankoku(0).Caption = ""
    lblGensankoku(1).Caption = ""
    If lstGensankoku.ListCount < 1 Then
        lblGensankoku(1).Caption = Trim(StrConv(ITEMREC.GENSANKOKU, vbUnicode))
        If Trim(lblGensankoku(1).Caption) <> "" Then
            lblGensankoku(0).Caption = ""
        End If
    Else

        lblGensankoku(1).Caption = Right(lstGensankoku.List(lstGensankoku.ListCount - 1), 20)
        lblGensankoku(0).Caption = "��" & StrConv(Format(lstGensankoku.ListCount, "#0"), vbWide)
    End If
    txGensankoku.text = Trim(lblGensankoku(1).Caption)




    '2010.07.20 ��



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


Start_Proc5:        '2013.03.13

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
                
                

                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                If sts > 3000 Or sts = 3 Then


                    Call File_Error(sts, BtOpGetEqual, "�w�}�[(�q)", 0)
                    '>>>>>>>>>>>>>  2015.04.24
                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                    'If sts Then
                    '    Call File_Error(sts, BtOpReset, "")
                    'End If
                
                
                    'Call File_Open_Proc
                    Do
                        If Not File_Open_Proc() Then
                            Exit Do
                        End If
                    Loop
                    '>>>>>>>>>>>>>  2015.04.24
    
                
                    GoTo Start_Proc5
                End If
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                
                
                
                
                Call Input_UnLock         '2008.01.15
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

Start_Proc6:        '2015.03.13

                Call UniCode_Conv(K0_ITEM.JGYOBU, K_Item_Tbl(k).JGYOBU)
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
                        
Start_Proc7:        '2015.03.13
                        
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
                                
                                
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                If sts > 3000 Or sts = 3 Then
            
                
                                    Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�", 0)
                                    '>>>>>>>>>>>>>  2015.04.24
                                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                    'If sts Then
                                    '    Call File_Error(sts, BtOpReset, "")
                                    'End If
                                
                                
                                    'Call File_Open_Proc
                                    Do
                                        If Not File_Open_Proc() Then
                                            Exit Do
                                        End If
                                    Loop
                                    '>>>>>>>>>>>>>  2015.04.24
                    
                                
                                    GoTo Start_Proc7
                                End If
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                
                                
                                
                                
                                
                                Call Input_UnLock             '2008.01.15
                                Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                                Exit Function
    
                        End Select
'                        Text1(K_Index + 1) = "���o�^�i��"
'                        Text1(K_Index + 4) = ""
'>>>>>>>>>>>>>>>>>>>>>>>>>> �i�Ԗ��o�^�\���̑Ή�    2012.12.21
                    Case Else
                        
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc6
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        
                        
                        
                        Call Input_UnLock         '2008.01.15
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


Start_Proc8:        '2015.03.13

                Call UniCode_Conv(K0_ITEM.JGYOBU, G_Item_Tbl(g).JGYOBU)
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
Start_Proc9:        '2015.03.13
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
                                
                                
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                If sts > 3000 Or sts = 3 Then
            
                
                                    Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�", 0)
                                    '>>>>>>>>>>>>>  2015.04.24
                                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                    'If sts Then
                                    '    Call File_Error(sts, BtOpReset, "")
                                    'End If
                                
                                
                                    'Call File_Open_Proc
                                    Do
                                        If Not File_Open_Proc() Then
                                            Exit Do
                                        End If
                                    Loop
                                    '>>>>>>>>>>>>>  2015.04.24
                    
                                
                                    GoTo Start_Proc9
                                End If
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                
                                
                                Call Input_UnLock             '2008.01.15
                                Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                                Exit Function
    
                        End Select
'                        Text1(G_Index + 1) = "���o�^�i��"
'                        Text1(G_Index + 4) = ""
'>>>>>>>>>>>>>>>>>>>>>>>>>> �i�Ԗ��o�^�\���̑Ή�    2012.12.21
                        
                    Case Else
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc8
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        Call Input_UnLock         '2008.01.15
                        Call File_Error(sts, BtOpGetEqual, "")
                        Exit Function

                End Select


                Text1(G_Index + 2).text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0.00")
                Text1(G_Index + 3).text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0.00")

                G_Index = G_Index + 5


            Case P_DOUKON   '�����^�\��

                d = d + 1
                D_Item_Tbl(d).JGYOBU = StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode)
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

                    DC_Index = DC_Index + 1

                                '�i��
                    Text1(DT_Index).text = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)


Start_Proc10:   '2015.03.13

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


' 2013.01.07 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'             �������i�͕W���I�Ԃ̍݌ɐ��̂ݕ\������l�ɕύX ���o�b�ȊO�͌���̂܂�

Start_Proc11:   '2015.03.13
                            
                            
'>>>>>>>>>>>>>>>    2015.03.13
'                            If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
'                                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
'                                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode), , , , , , Ret_sts, 0) Then
'                                Exit Function
'                            End If


                            Zaiko_sts = Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode), , , , , , Ret_sts, 0)
                                
                            If Zaiko_sts Then
                                If Ret_sts > 3000 Or Ret_sts = 3 Then
                                    Call File_Error(sts, BtOpGetEqual, "�݌��ް�", 0)
                                    '>>>>>>>>>>>>>  2015.04.24
                                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                    'If sts Then
                                    '    Call File_Error(sts, BtOpReset, "")
                                    'End If
                                
                                
                                    'Call File_Open_Proc
                                    Do
                                        If Not File_Open_Proc() Then
                                            Exit Do
                                        End If
                                    Loop
                                    '>>>>>>>>>>>>>  2015.04.24
                    
                                
                                    GoTo Start_Proc11
                                End If
                                Call File_Error(sts, BtOpGetEqual, "�݌��ް�")
                                Exit Function
                            End If
'>>>>>>>>>>>>>>>    2015.03.13

'                            Wk_LOC = StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & _
'                                     StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)
'
'                            If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
'                                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
'                                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode), _
'                                                                    Wk_LOC) Then
'                                Exit Function
'
'                            End If
' 2013.01.07 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

                            Text1(DT_Index + 5) = Format(Sumi_Qty + Mi_Qty, "#0")


                        Case BtErrKeyNotFound




'>>>>>>>>>>>>>>>>>>>>>>>>>> �i�Ԗ��o�^�\���̑Ή�    2012.12.21
                            
                            
Start_Proc12:   '2015.03.13
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
                                    
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                If sts > 3000 Or sts = 3 Then
            
                
                                    Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�", 0)
                                    '>>>>>>>>>>>>>  2015.04.24
                                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                    'If sts Then
                                    '    Call File_Error(sts, BtOpReset, "")
                                    'End If
                                
                                
                                    'Call File_Open_Proc
                                    Do
                                        If Not File_Open_Proc() Then
                                            Exit Do
                                        End If
                                    Loop
                                    '>>>>>>>>>>>>>  2015.04.24
                    
                                
                                    GoTo Start_Proc12
                                End If
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                    
                                    
                                    Call Input_UnLock             '2008.01.15
                                    Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                                    Exit Function
        
                            End Select
'                            Text1(DT_Index + 1) = "���o�^�i��"
'                            Text1(DT_Index + 4) = ""
'                            Text1(DT_Index + 5) = ""
'>>>>>>>>>>>>>>>>>>>>>>>>>> �i�Ԗ��o�^�\���̑Ή�    2012.12.21
                    
                        Case Else
                            
                            
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                            If sts > 3000 Or sts = 3 Then
        
            
                                Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�", 0)
                                '>>>>>>>>>>>>>  2015.04.24
                                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                'If sts Then
                                '    Call File_Error(sts, BtOpReset, "")
                                'End If
                            
                            
                                'Call File_Open_Proc
                                Do
                                    If Not File_Open_Proc() Then
                                        Exit Do
                                    End If
                                Loop
                                '>>>>>>>>>>>>>  2015.04.24
                
                            
                                GoTo Start_Proc10
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                            
                            
                            Call Input_UnLock         '2008.01.15
                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
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

    Call Input_UnLock         '2008.01.15



    If Right(Combo2(0).text, 1) = " " Or Right(Combo2(0).text, 1) = "2" Then
        Check1(pchkL_PAPER).Enabled = False
        Check1(pchkL_PLASTIC).Enabled = False
    Else
        Check1(pchkL_PAPER).Enabled = True
        Check1(pchkL_PLASTIC).Enabled = True
    End If



    Item_Disp_Proc = False

End Function

Private Function Update_Proc(Mode As Integer, MSG As Integer) As Integer
'----------------------------------------------------------------------------
'                   �\���}�X�^�����i���w���ް��o��
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim ans         As Integer
Dim com         As Integer

Dim SEQNO       As Integer

Dim i           As Integer
Dim j           As Integer

Dim SHIJINO     As Long


Dim NEW_F       As Boolean          '2008.05.19

    Update_Proc = True

    Call Input_Lock

                                        
Start_Proc0:        '2015.03.26
                                        
                                        '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If

    NEW_F = False       '2008.05.19

    If Text1(ptxSHIJI_NO).text = "" Then
        NEW_F = True    '2008.05.19


        Do                              '<----------------------    2013.10.04

            DoEvents                                                '2013.10.04


            '2008.02.02
            'If IsNumeric(Text1(ptxSHIJI_QTY).text) And CDbl(Text1(ptxSHIJI_QTY).text) = 0 Then
            '    SHIJINO = "00000"
            'Else
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
                            
                            
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            If sts > 3000 Or sts = 3 Then
        
            
                                Call File_Error(sts, BtOpGetEqual, "�Ǘ��}�X�^", 0)
    
            
                                sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                                If sts <> BtNoErr Then
                                    Call File_Error(sts, BtOpAbortTransaction, "")
                                End If
    
                                '>>>>>>>>>>>>>  2015.04.24
                                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                'If sts Then
                                '    Call File_Error(sts, BtOpReset, "")
                                'End If
                            
                                'Call File_Open_Proc
                                Do
                                    If Not File_Open_Proc() Then
                                        Exit Do
                                    End If
                                Loop
                                '>>>>>>>>>>>>>  2015.04.24
                
                            
                                GoTo Start_Proc0
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            
                            
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
                            
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            If sts > 3000 Or sts = 3 Then
        
            
                                Call File_Error(sts, BtOpGetEqual, "�Ǘ��}�X�^", 0)
    
            
                                sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                                If sts <> BtNoErr Then
                                    Call File_Error(sts, BtOpAbortTransaction, "")
                                End If
                                '>>>>>>>>>>>>>  2015.04.24
                                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                'If sts Then
                                '    Call File_Error(sts, BtOpReset, "")
                                'End If
                            
                                'Call File_Open_Proc
                                Do
                                    If Not File_Open_Proc() Then
                                        Exit Do
                                    End If
                                Loop
                                '>>>>>>>>>>>>>  2015.04.24
                
                            
                                GoTo Start_Proc0
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            
                            
                            Call File_Error(sts, BtOpUpdate, "�Ǘ��}�X�^")
                            GoTo Abort_Tran
                    End Select
                Loop
    
                SHIJINO = CLng(StrConv(P_KANRIREC.SASHIZU_NO, vbUnicode))
    
    
                Text1(ptxSHIJI_NO).text = Format(SHIJINO, "00000000")
            'End If
        
                                        '                           2013.10.04  �w�}�f�[�^�̑��݃`�F�b�N��ǉ�
                Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_No, Format(SHIJINO, "00000000"))
                sts = BTRV(BtOpGetEqual, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        Exit Do
                    Case Else
                        
                        
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            If sts > 3000 Or sts = 3 Then
        
            
                                Call File_Error(sts, BtOpGetEqual, "�w�}�[�ް�(�e)", 0)
    
            
                                sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                                If sts <> BtNoErr Then
                                    Call File_Error(sts, BtOpAbortTransaction, "")
                                End If
    
                                '>>>>>>>>>>>>>  2015.04.24
                                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                'If sts Then
                                '    Call File_Error(sts, BtOpReset, "")
                                'End If
                            
                                'Call File_Open_Proc
                                Do
                                    If Not File_Open_Proc() Then
                                        Exit Do
                                    End If
                                Loop
                                '>>>>>>>>>>>>>  2015.04.24
                
                            
                                GoTo Start_Proc0
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        
                        
                        Call File_Error(sts, BtOpUpdate, "�w�}�[�ް�(�e)")
                        GoTo Abort_Tran
                End Select
                                        '                           2013.10.04  �w�}�f�[�^�̑��݃`�F�b�N��ǉ�
        Loop                            '<----------------------    2013.10.04
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
                    
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            If sts > 3000 Or sts = 3 Then
        
            
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
    
            
                                sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                                If sts <> BtNoErr Then
                                    Call File_Error(sts, BtOpAbortTransaction, "")
                                End If
    
                                '>>>>>>>>>>>>>  2015.04.24
                                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                'If sts Then
                                '    Call File_Error(sts, BtOpReset, "")
                                'End If
                            
                                'Call File_Open_Proc
                                Do
                                    If Not File_Open_Proc() Then
                                        Exit Do
                                    End If
                                Loop
                                '>>>>>>>>>>>>>  2015.04.24
                
                            
                                GoTo Start_Proc0
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                    
                    
                    
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
                    GoTo Abort_Tran

            End Select


        Loop


'2010.11.12 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        
    If Trim(Right(Combo2(0).text, 1)) = "0" Or Trim(Right(Combo2(0).text, 1)) = "1" Then
        If Check1(pchkL_PAPER).Value = vbChecked Then                           '��
            Call UniCode_Conv(ITEMREC.L_PAPER, L_PAPER_ON)
        Else
            Call UniCode_Conv(ITEMREC.L_PAPER, L_PAPER_OFF)
        End If

        If Check1(pchkL_PLASTIC).Value = vbChecked Then                         '�v���X�`�b�N
            Call UniCode_Conv(ITEMREC.L_PLASTIC, L_PLASTIC_ON)
        Else
            Call UniCode_Conv(ITEMREC.L_PLASTIC, L_PLASTIC_OFF)
        End If

        If Check1(pchkL_LABEL).Value = vbChecked Then                           '�K�p�@�탉�x��
            Call UniCode_Conv(ITEMREC.L_LABEL, L_LABEL_ON)
        Else
            Call UniCode_Conv(ITEMREC.L_LABEL, L_LABEL_OFF)
        End If
            

    End If

'2010.11.12 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


'2011.02.10
    If Trim(Right(Combo2(0).text, 1)) <> "" Then
        Call UniCode_Conv(ITEMREC.L_LABEL, Right(Combo2(0).text, 1))
    End If
'2011.02.10

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


        '2011.02.16
        Call UniCode_Conv(ITEMREC.UPD_TANTO, "PI010")
        Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
        


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
                    
                    
                    
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            If sts > 3000 Or sts = 3 Then
        
            
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
    
            
                                sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                                If sts <> BtNoErr Then
                                    Call File_Error(sts, BtOpAbortTransaction, "")
                                End If
    
                                '>>>>>>>>>>>>>  2015.04.24
                                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                'If sts Then
                                '    Call File_Error(sts, BtOpReset, "")
                                'End If
                            
                                'Call File_Open_Proc
                                Do
                                    If Not File_Open_Proc() Then
                                        Exit Do
                                    End If
                                Loop
                                '>>>>>>>>>>>>>  2015.04.24
                
                            
                                GoTo Start_Proc0
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                    
                    
                    Call File_Error(sts, BtOpUpdate, "�i�ڃ}�X�^")
                    GoTo Abort_Tran
            End Select
        Loop

'    End If
    '---------------------------------------------------    '�\���}�X�^�X�V

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
                    
                    
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            If sts > 3000 Or sts = 3 Then
        
            
                                Call File_Error(sts, BtOpGetEqual, "�\���}�X�^", 0)
    
            
                                sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                                If sts <> BtNoErr Then
                                    Call File_Error(sts, BtOpAbortTransaction, "")
                                End If
    
                                '>>>>>>>>>>>>>  2015.04.24
                                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                'If sts Then
                                '    Call File_Error(sts, BtOpReset, "")
                                'End If
                            
                                'Call File_Open_Proc
                                Do
                                    If Not File_Open_Proc() Then
                                        Exit Do
                                    End If
                                Loop
                                '>>>>>>>>>>>>>  2015.04.24
                
                            
                                GoTo Start_Proc0
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                    
                    
                    Call File_Error(sts, com + BtSNoWait, "�\���}�X�^")
                    GoTo Abort_Tran
            End Select

        Loop

        If sts = BtErrEOF Then
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
                    
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                    If sts > 3000 Or sts = 3 Then

    
                        Call File_Error(sts, BtOpGetEqual, "�\���}�X�^", 0)

    
                        sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                        If sts <> BtNoErr Then
                            Call File_Error(sts, BtOpAbortTransaction, "")
                        End If

                        '>>>>>>>>>>>>>  2015.04.24
                        'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                        'If sts Then
                        '    Call File_Error(sts, BtOpReset, "")
                        'End If
                    
                        'Call File_Open_Proc
                        Do
                            If Not File_Open_Proc() Then
                                Exit Do
                            End If
                        Loop
                        '>>>>>>>>>>>>>  2015.04.24
        
                    
                        GoTo Start_Proc0
                    End If
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                    
                    
                    
                    Call File_Error(sts, BtOpDelete, "�\���}�X�^")
                    GoTo Abort_Tran
            End Select
        Loop

        com = BtOpGetNext

    Loop

    '�\���}�X�^(ͯ�ް)�o��
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
        
    Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_TANTO, "")                   '�i�������S���Һ���     2013.08.21
    Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_DATETIME, "")                '�i����������           2013.08.21
    Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_LABEL_CNT, "")               '�i���������ٌ���       2013.08.21
    Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_GENPIN_CNT, "")              '�i���������i�[����     2013.08.21


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
                
                
                
                
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                If sts > 3000 Or sts = 3 Then


                    Call File_Error(sts, BtOpGetEqual, "�\���}�X�^", 0)


                    sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                    If sts <> BtNoErr Then
                        Call File_Error(sts, BtOpAbortTransaction, "")
                    End If

                    '>>>>>>>>>>>>>  2015.04.24
                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                    'If sts Then
                    '    Call File_Error(sts, BtOpReset, "")
                    'End If
                
                    'Call File_Open_Proc
                    Do
                        If Not File_Open_Proc() Then
                            Exit Do
                        End If
                    Loop
                    '>>>>>>>>>>>>>  2015.04.24
    
                
                    GoTo Start_Proc0
                End If
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                
                
                
                
                Call File_Error(sts, BtOpInsert, "�\���}�X�^")
                GoTo Abort_Tran
        End Select

    Loop

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
                        
                        
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            If sts > 3000 Or sts = 3 Then
        
            
                                Call File_Error(sts, BtOpGetEqual, "�\���}�X�^", 0)
    
            
                                sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                                If sts <> BtNoErr Then
                                    Call File_Error(sts, BtOpAbortTransaction, "")
                                End If
    
                                '>>>>>>>>>>>>>  2015.04.24
                                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                'If sts Then
                                '    Call File_Error(sts, BtOpReset, "")
                                'End If
                                
                                'Call File_Open_Proc
                                Do
                                    If Not File_Open_Proc() Then
                                        Exit Do
                                    End If
                                Loop
                                '>>>>>>>>>>>>>  2015.04.24
                
                            
                                GoTo Start_Proc0
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        
                        
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
                        
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            If sts > 3000 Or sts = 3 Then
        
            
                                Call File_Error(sts, BtOpGetEqual, "�\���}�X�^", 0)
    
            
                                sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                                If sts <> BtNoErr Then
                                    Call File_Error(sts, BtOpAbortTransaction, "")
                                End If
                                
                                '>>>>>>>>>>>>>  2015.04.24
                                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                'If sts Then
                                '    Call File_Error(sts, BtOpReset, "")
                                'End If
                            
                                'Call File_Open_Proc
                                Do
                                    If Not File_Open_Proc() Then
                                        Exit Do
                                    End If
                                Loop
                                '>>>>>>>>>>>>>  2015.04.24
                
                            
                                GoTo Start_Proc0
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        
                        
                        Call File_Error(sts, BtOpInsert, "�\���}�X�^")
                        GoTo Abort_Tran
                End Select

            Loop



        End If

        j = j + 1


    Next i


    '�����^�\����
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
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "�\���}�X�^", 0)

        
                            sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                            If sts <> BtNoErr Then
                                Call File_Error(sts, BtOpAbortTransaction, "")
                            End If

                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc0
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        
                        
                        Call File_Error(sts, BtOpInsert, "�\���}�X�^")
                        GoTo Abort_Tran
                End Select

            Loop


        End If

    Next i



    If Mode = 1 Then
        GoTo End_Tran
    End If

    If CDbl(Text1(ptxSHIJI_QTY).text) = 0 Then      '2008.02.02
        GoTo End_Tran
    End If
    '---------------------------------------------------    '�w�}�[�f�[�^�X�V

    '�w�}�[�f�[�^(ͯ�ް)����


    Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_No, Format(SHIJINO, "00000000"))  '2008.02.13

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
                
                
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            If sts > 3000 Or sts = 3 Then
        
            
                                Call File_Error(sts, BtOpGetEqual, "���i���w�}�[�ް�(�e)", 0)
    
            
                                sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                                If sts <> BtNoErr Then
                                    Call File_Error(sts, BtOpAbortTransaction, "")
                                End If
                                
                                '>>>>>>>>>>>>>  2015.04.24
                                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                'If sts Then
                                '    Call File_Error(sts, BtOpReset, "")
                                'End If
                            
                                'Call File_Open_Proc
                                Do
                                    If Not File_Open_Proc() Then
                                        Exit Do
                                    End If
                                Loop
                                '>>>>>>>>>>>>>  2015.04.24
                
                            
                                GoTo Start_Proc0
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "���i���w�}�[�ް�(�e)")
                GoTo Abort_Tran
        End Select

    Loop


    If com = BtOpInsert Then
        '�V�K�쐬
        Call UniCode_Conv(P_SSHIJI_O_REC.SHIJI_No, Format(SHIJINO, "00000000")) '�w�}�[��   2008.02.13
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

'        Call UniCode_Conv(P_SSHIJI_O_REC.FILLER, "")                           '2016.02.01
        Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_GAISOU_CNT, "")              '2016.02.01 �i�������O���i�Ԍ���

    End If
                                                                                '���s��
    Call UniCode_Conv(P_SSHIJI_O_REC.HAKKO_DT, Format(Text1(ptxHAKKO_DT).text, "YYYYMMDD"))
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
            
            
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
            If sts > 3000 Or sts = 3 Then


                Call File_Error(sts, BtOpGetEqual, "�󕥐�}�X�^", 0)


                sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                If sts <> BtNoErr Then
                    Call File_Error(sts, BtOpAbortTransaction, "")
                End If

                '>>>>>>>>>>>>>  2015.04.24
                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                'If sts Then
                '    Call File_Error(sts, BtOpReset, "")
                'End If
            
                'Call File_Open_Proc
                Do
                    If Not File_Open_Proc() Then
                        Exit Do
                    End If
                Loop
                '>>>>>>>>>>>>>  2015.04.24

            
                GoTo Start_Proc0
            End If
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
            
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
            End If
        End If
    End If


    If Check1(pchkPRI_SHIJI).Value = vbChecked Then                             '�o�͑Ώہ@�w�}�[
        Call UniCode_Conv(P_SSHIJI_O_REC.PRI_SHIJI, P_PRI_SHIJI_ON)
        Call UniCode_Conv(P_SSHIJI_O_REC.Print_datetime, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
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
                
                
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                If sts > 3000 Or sts = 3 Then


                    Call File_Error(sts, BtOpGetEqual, "���i���w�}�ް�(�e)", 0)


                    sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                    If sts <> BtNoErr Then
                        Call File_Error(sts, BtOpAbortTransaction, "")
                    End If

                    '>>>>>>>>>>>>>  2015.04.24
                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                    'If sts Then
                    '    Call File_Error(sts, BtOpReset, "")
                    'End If
                
                    'Call File_Open_Proc
                    Do
                        If Not File_Open_Proc() Then
                            Exit Do
                        End If
                    Loop
                    '>>>>>>>>>>>>>  2015.04.24
    
                
                    GoTo Start_Proc0
                End If
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                
                
                Call File_Error(sts, com, "���i���w�}�ް�(�e)")
                GoTo Abort_Tran
        End Select

    Loop

    If com = BtOpUpdate Then
        '�Ώۂ̎q���폜����
        Call UniCode_Conv(K0_P_SSHIJI_K.SHIJI_No, Format(SHIJINO, "00000000"))  '2008.02.13
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
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "���i���w�}�ް�(�q)", 0)

        
                            sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                            If sts <> BtNoErr Then
                                Call File_Error(sts, BtOpAbortTransaction, "")
                            End If

                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc0
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        
                        Call File_Error(sts, com + BtSNoWait, "���i���w�}�ް�(�q)")
                        GoTo Abort_Tran
                End Select

            Loop

            If sts = BtErrEOF Then
                Exit Do
            End If


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
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "���i���w�}�ް�(�q)", 0)

        
                            sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                            If sts <> BtNoErr Then
                                Call File_Error(sts, BtOpAbortTransaction, "")
                            End If

                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc0
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        
                        
                        
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
            Call UniCode_Conv(P_SSHIJI_K_REC.SHIJI_No, Format(SHIJINO, "00000000"))     '�w�}�[�� 2008.02.13



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


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.04.20
            Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI_QTY, "")                          '�ړ��ςݐ��� 2012.04.20
            Call UniCode_Conv(P_SSHIJI_K_REC.COMPO_TANTO, "")                           '�\������   �S����          2012.04.20
            Call UniCode_Conv(P_SSHIJI_K_REC.COMPO_Sumi_Cnt, "")                        '           �����ςݐ�      2012.04.20
            Call UniCode_Conv(P_SSHIJI_K_REC.COMPO_ALL_Cnt, "")                         '           �\����          2012.04.20
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.04.20
            
            Call UniCode_Conv(P_SSHIJI_K_REC.FILLER, "")
                                                                                        '�X�V����
            Call UniCode_Conv(P_SSHIJI_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))

                                                                                        '�o�ח\��h�c
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_ID_NO, "")

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
                        
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "���i���w�}�ް�(�q)", 0)

        
                            sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                            If sts <> BtNoErr Then
                                Call File_Error(sts, BtOpAbortTransaction, "")
                            End If

                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc0
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        
                        
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

            Call UniCode_Conv(P_SSHIJI_K_REC.SHIJI_No, Format(SHIJINO, "00000000"))     '�w�}�[��   2008.02.13
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

            Call UniCode_Conv(P_SSHIJI_K_REC.FILLER, "")
                                                                                        
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.04.20
            Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI_QTY, "")                          '�ړ��ςݐ��� 2012.04.20
            Call UniCode_Conv(P_SSHIJI_K_REC.COMPO_TANTO, "")                           '�\������   �S����          2012.04.20
            Call UniCode_Conv(P_SSHIJI_K_REC.COMPO_Sumi_Cnt, "")                        '           �����ςݐ�      2012.04.20
            Call UniCode_Conv(P_SSHIJI_K_REC.COMPO_ALL_Cnt, "")                         '           �\����          2012.04.20
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.04.20
                                                                                        
                                                                                        
                                                                                        
                                                                                        '�X�V����
            Call UniCode_Conv(P_SSHIJI_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))


                                                                                        '�o�ח\��h�c
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_ID_NO, "")

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
                        
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "���i���w�}�ް�(�q)", 0)

        
                            sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                            If sts <> BtNoErr Then
                                Call File_Error(sts, BtOpAbortTransaction, "")
                            End If
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
           
                        
                            GoTo Start_Proc0
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        
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
            Call UniCode_Conv(P_SSHIJI_K_REC.SHIJI_No, Format(SHIJINO, "00000000"))     '�w�}�[��   2008.02.13

            SEQNO = SEQNO + 10

            Call UniCode_Conv(P_SSHIJI_K_REC.DATA_KBN, P_DOUKON)                        '�f�[�^�敪
            Call UniCode_Conv(P_SSHIJI_K_REC.SEQNO, Format(SEQNO, "000"))               '�ǔ�

            Call UniCode_Conv(P_SSHIJI_K_REC.KO_SYUBETSU, D_Item_Tbl(i).SYUBETSU)       '���
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_JGYOBU, D_Item_Tbl(i).JGYOBU)           '���ƕ�
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_NAIGAI, D_Item_Tbl(i).NAIGAI)           '�����O
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_HIN_GAI, D_Item_Tbl(i).HIN_GAI)         '�i��
                                                                                        '����
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_QTY, Format(D_Item_Tbl(i).QTY, "000.00"))
                                                                                        '��
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_SHIJI_QTY, Format(D_Item_Tbl(i).SHIJI_QTY, "00000000.00"))
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_BIKOU, D_Item_Tbl(i).BIKOU)             '���l

            Call UniCode_Conv(P_SSHIJI_K_REC.CALCEL_F, P_CANCEL_OFF)                    '��ݾ��׸�
            Call UniCode_Conv(P_SSHIJI_K_REC.CANCEL_DATETIME, "")                       '��ݾٓ���

            
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.04.20
            Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI_QTY, "")                          '�ړ��ςݐ��� 2012.04.20
            Call UniCode_Conv(P_SSHIJI_K_REC.COMPO_TANTO, "")                           '�\������   �S����          2012.04.20
            Call UniCode_Conv(P_SSHIJI_K_REC.COMPO_Sumi_Cnt, "")                        '           �����ςݐ�      2012.04.20
            Call UniCode_Conv(P_SSHIJI_K_REC.COMPO_ALL_Cnt, "")                         '           �\����          2012.04.20
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.04.20
            
            
            
            Call UniCode_Conv(P_SSHIJI_K_REC.FILLER, "")
                                                                                        '�X�V����
            Call UniCode_Conv(P_SSHIJI_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))


'''Y_SYUKA �̏o�͂�߂�@2010.09.17
            If POS_UMU Then
                '�o�׎w���̍쐬
                If Y_SYUKA_Make_Proc(i) Then
                    GoTo Abort_Tran
                End If
            End If

                                                                                        '�o�ח\��h�c
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_ID_NO, D_Item_Tbl(i).ID_NO)


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
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "���i���w�}�ް�(�q)", 0)

        
                            sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                            If sts <> BtNoErr Then
                                Call File_Error(sts, BtOpAbortTransaction, "")
                            End If
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc0
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        
                        
                        Call File_Error(sts, BtOpInsert, "���i���w�}�ް�(�q)")
                        GoTo Abort_Tran
                End Select

            Loop

        End If

    Next i





End_Tran:
                                        '�g�����U�N�V�����I��
    sts = BTRV(BtOpEndTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If

'    If Mode = 0 Then       '2007.11.21
    If MSG = 0 Then         '2007.11.21

'2008.05.19        If Text1(ptxSHIJI_NO).text = "" Then
        If NEW_F Then       '2008.05.19
            MsgBox "�w�}�[���F" & Format(SHIJINO, "00000000") & "���쐬���܂����B"  '2008.02.13
        End If
    End If

    Call Input_UnLock
                                        '����ɑΏێw�}�[����ʒm
    Taget_Key = Format(SHIJINO, "00000000") '2008.02.13

    Update_Proc = False

    Exit Function

Abort_Tran:

    sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If

    Call Input_UnLock

End Function

Private Sub Check1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'
'2010.11.12  �߰����ٔ��s(PM0040)�̓����ɍ��킹��ׁA�������ۼ��ެ��ǉ�

    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If

    Call Tab_Ctrl(Shift)        '�ړ�

End Sub


Private Sub Combo1_Click(Index As Integer)

Dim sts         As Integer

Dim Sumi_Qty    As Long
Dim Mi_Qty      As Long

Dim TABLCTRL_SW As Integer      '2019.06.11 �ǉ�
    
    TABLCTRL_SW = 0
    
    Select Case Index
        Case pcmbSHIMUKE        '�d������
'            svJGYOBU = Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1)
'            svNAIGAI = Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1)
            '2019.06.11 �������ŃZ�b�g����ƁA���L��If�������Ӗ��I
            '           �R�����g�ɂ����B
'            If Trim(Text1(ptxHIN_GAI)) <> "" Then
'                If svJGYOBU <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
'                    svNAIGAI <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Then
'
'                    chenge_F = True
'                End If
'            End If
            
            '2019.06.11 �d������̂S���Ŕ���
            If Trim(Text1(ptxHIN_GAI)) <> "" Then
                If svSHIMUKE <> Right(Combo1(pcmbSHIMUKE), 4) Then
                    chenge_F = True
                End If
            End If
            
            svSHIMUKE_CODE = Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2)   '2013.08.29

            svSHIMUKE = Right(Combo1(pcmbSHIMUKE), 4)





            If chenge_F Then
                TABLCTRL_SW = 1
                
                If Error_Check_Proc(ptxHIN_GAI, 0, 0) Then   '�G���[�`�F�b�N
                    chenge_F = True
                    Text1(ptxHIN_GAI).SelLength = Len(Text1(ptxHIN_GAI))
                    Text1(ptxHIN_GAI).SelStart = 0
                    Text1(ptxHIN_GAI).SetFocus
                    Exit Sub
                End If

Start_Proc1:    '2015.03.26
                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).text)


                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                sts = BTRV(BtOpGetGreaterEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        Text1(ptxHIN_GAI) = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
                    Case BtErrKeyNotFound, BtErrEOF

                        Text1(ptxHIN_NAME).text = ""
                        Text1(ptxST_LOCATION).text = ""
                        Text1(ptxMI_QTY).text = ""
                        Text1(ptxSUMI_QTY).text = ""

                        
'2010.11.12 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                        Check1(pchkL_PAPER).Value = vbUnchecked         '��
                        Check1(pchkL_PLASTIC).Value = vbUnchecked       '�v��
                        Check1(pchkL_LABEL).Value = vbUnchecked         '�K�p�@�탉�x��
'2010.11.12 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                        
                        MsgBox "���͂������ڂ̓G���[�ł��B(�i��)"
                        Text1(ptxHIN_GAI).SetFocus
                        Exit Sub                        '2019.06.10 ����
'                        Exit Sub    '2010.11.17
                    Case Else
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                           GoTo Start_Proc1
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        
                        
                        
                        
                        
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                        Unload Me

                End Select



                Text1(ptxHIN_NAME).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)

'2010.11.12 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                If StrConv(ITEMREC.L_PAPER, vbUnicode) = L_PAPER_ON Then        '��
                    Check1(pchkL_PAPER).Value = vbChecked
                Else
                    Check1(pchkL_PAPER).Value = vbUnchecked
                End If
    
                If StrConv(ITEMREC.L_PLASTIC, vbUnicode) = L_PLASTIC_ON Then    '�v��
                    Check1(pchkL_PLASTIC).Value = vbChecked
                Else
                    Check1(pchkL_PLASTIC).Value = vbUnchecked
                End If
    
                If StrConv(ITEMREC.L_LABEL, vbUnicode) = L_LABEL_ON Then        '�K�p�@�탉�x��
                    Check1(pchkL_LABEL).Value = vbChecked
                Else
                    Check1(pchkL_LABEL).Value = vbUnchecked
                End If
'2010.11.12 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Text1(ptxST_LOCATION).text = ""
                Else
                    Text1(ptxST_LOCATION).text = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If

                If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                        StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                        StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then

                    Unload Me
                End If

                Text1(ptxMI_QTY).text = Format(Mi_Qty, "#0")
                Text1(ptxSUMI_QTY).text = Format(Sumi_Qty, "#0")



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
                            Case Else
                                
                                
                                
                                
                                
                                Call File_Error(sts, BtOpGetEqual, "�\���}�X�^")
                                Unload Me
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
                            Case Else
                                
                                
                                
                                
                                
                                
                                Call File_Error(sts, BtOpGetEqual, "�\���}�X�^")
                                Unload Me
                        End Select

                    End If
                End If

                chenge_F = False
                
                DoEvents
                
'                Combo1(pcmbSHIMUKE).SetFocus
                
                Text1(ptxHIN_GAI).SelLength = Len(Text1(ptxHIN_GAI))
                Text1(ptxHIN_GAI).SelStart = 0
                Text1(ptxHIN_GAI).SetFocus
                Exit Sub
            End If
        
        svJGYOBU = Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1)
        svNAIGAI = Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1)
        
        svSHIMUKE = Right(Combo1(pcmbSHIMUKE), 4)
        
        '2019.06.10 ����
        If TABLCTRL_SW = 1 Then
            Text1(ptxHIN_GAI).SetFocus
        Else
            Call Tab_Ctrl(0)        '�ړ�
        End If
        
        Exit Sub
    Case Else
    
    End Select

End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'
'
'Dim sts         As Integer
'
'Dim Sumi_Qty    As Long
'Dim Mi_Qty      As Long
'
'
'
'
'    If KeyCode <> vbKeyReturn Then
'        Exit Sub
'    End If
'
'    Select Case Index
'        Case pcmbSHIMUKE        '�d������
'
'            If svJGYOBU <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
'                svNAIGAI <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Then
'
'                chenge_F = True
'
'
'            End If
'
'
'
'            If chenge_F Then
'
'Start_Proc1:    '2015.03.13
'                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
'                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).text)
'
'
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                Select Case sts
'                    Case BtNoErr
'                    Case BtErrKeyNotFound
'
'                        Text1(ptxHIN_NAME).text = ""
'                        Text1(ptxST_LOCATION).text = ""
'                        Text1(ptxMI_QTY).text = ""
'                        Text1(ptxSUMI_QTY).text = ""
'
''2010.11.12 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                        Check1(pchkL_PAPER).Value = vbUnchecked         '��
'                        Check1(pchkL_PLASTIC).Value = vbUnchecked       '�v��
'                        Check1(pchkL_LABEL).Value = vbUnchecked         '�K�p�@�탉�x��
''2010.11.12 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'
'                        MsgBox "���͂������ڂ̓G���[�ł��B(�i��)"
'                        Text1(ptxHIN_GAI).SetFocus
'
'                        Exit Sub    '2010.11.17
'                    Case Else
'
'                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
'                        If sts > 3000 Or sts = 3 Then
'
'
'                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
'                            '>>>>>>>>>>>>>  2015.04.24
'                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
'                            'If sts Then
'                            '    Call File_Error(sts, BtOpReset, "")
'                            'End If
'
'                            'Call File_Open_Proc
'                            Do
'                                If Not File_Open_Proc() Then
'                                    Exit Do
'                                End If
'                            Loop
'                            '>>>>>>>>>>>>>  2015.04.24
'
'
'                            GoTo Start_Proc1
'                        End If
'                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
'
'
'
'                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
'                        Unload Me
'
'                End Select
'
'
'
'                Text1(ptxHIN_NAME).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
'
''2010.11.12 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                If StrConv(ITEMREC.L_PAPER, vbUnicode) = L_PAPER_ON Then        '��
'                    Check1(pchkL_PAPER).Value = vbChecked
'                Else
'                    Check1(pchkL_PAPER).Value = vbUnchecked
'                End If
'
'                If StrConv(ITEMREC.L_PLASTIC, vbUnicode) = L_PLASTIC_ON Then    '�v��
'                    Check1(pchkL_PLASTIC).Value = vbChecked
'                Else
'                    Check1(pchkL_PLASTIC).Value = vbUnchecked
'                End If
'
'                If StrConv(ITEMREC.L_LABEL, vbUnicode) = L_LABEL_ON Then        '�K�p�@�탉�x��
'                    Check1(pchkL_LABEL).Value = vbChecked
'                Else
'                    Check1(pchkL_LABEL).Value = vbUnchecked
'                End If
''2010.11.12 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'
'
'                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
'                    Text1(ptxST_LOCATION).text = ""
'                Else
'                    Text1(ptxST_LOCATION).text = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
'                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
'                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
'                                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
'                End If
'
'                If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
'                                                        StrConv(ITEMREC.NAIGAI, vbUnicode), _
'                                                        StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
'
'                    Unload Me
'                End If
'
'                Text1(ptxMI_QTY).text = Format(Mi_Qty, "#0")
'                Text1(ptxSUMI_QTY).text = Format(Sumi_Qty, "#0")
'
'
'
'                If Trim(Text1(ptxSHIJI_NO).text) = "" Then
'                    If StrConv(P_COMPO_O_REC.SHIMUKE_CODE, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) And _
'                        StrConv(P_COMPO_O_REC.JGYOBU, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) And _
'                        StrConv(P_COMPO_O_REC.NAIGAI, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) And _
'                        Trim(StrConv(P_COMPO_O_REC.HIN_GAI, vbUnicode)) = Trim(Text1(ptxHIN_GAI).text) Then
'                    Else
'
'
'
'                        sts = P_COMPO_Disp_Proc()
'                        Select Case sts
'                            Case BtNoErr
'                            Case BtErrKeyNotFound
'                            Case Else
'
'
'
'                                Call File_Error(sts, BtOpGetEqual, "�\���}�X�^")
'                                Unload Me
'                        End Select
'                    End If
'                Else
'                    If StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) And _
'                        StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) And _
'                        StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) And _
'                        Trim(StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode)) = Trim(Text1(ptxHIN_GAI).text) Then
'                    Else
'
'
'                        sts = P_COMPO_Disp_Proc()
'                        Select Case sts
'                            Case BtNoErr
'                            Case BtErrKeyNotFound
'                            Case Else
'
'
'
'                                Call File_Error(sts, BtOpGetEqual, "�\���}�X�^")
'                                Unload Me
'                        End Select
'
'                    End If
'                End If
'                Text1(ptxSHIJI_QTY).SetFocus
'
'                chenge_F = False
'                Exit Sub
'            End If
'
'            '2019.06.10 ���L��ǉ�
'            Text1(ptxHIN_GAI).SetFocus
'
'            Exit Sub
'
'        Case pcmbUKEHARAI       '��z��
'            Text1(ptxUKEHARAI_CODE).text = Trim(Right(Combo1(pcmbUKEHARAI).text, 5))
'        Case pcmbS_TANTO        '���P�^�S����
'
'                                '�����^�\���@���
'        Case pcmbD_SYUBETSU01, pcmbD_SYUBETSU02, pcmbD_SYUBETSU03, pcmbD_SYUBETSU04, pcmbD_SYUBETSU05, pcmbD_SYUBETSU06
'
'            D_Item_Tbl(Index - pcmbD_SYUBETSU01).SYUBETSU = Right(Combo1(Index).text, 2)
'    End Select
'
    Call Tab_Ctrl(Shift)        '�ړ�

End Sub


Private Sub Combo1_LostFocus(Index As Integer)
                                               '2019.06.10�@�S�āA�R�����g�ɂ��Ă݂��B
'                                               '2019.06.11�@�S�āA���A���Ă݂��B
Dim i   As Integer  '2013.08.29

    Select Case Index
        Case pcmbSHIMUKE        '�d������

'            If svJGYOBU <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
'                svNAIGAI <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Then
'
'                chenge_F = True
'
'            End If
            '2019.06.11 �S���Ŕ���ɕύX
            If svSHIMUKE <> Right(Combo1(pcmbSHIMUKE), 4) Then
                chenge_F = True
            End If


'2013.08.29 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'2013.11.21            If svSHIMUKE_CODE <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Then
                For i = 0 To UBound(SHIMUKE_CHK_TBL)

                    If SHIMUKE_CHK_TBL(i) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Then
                        Combo2(0).ListIndex = 0
                        Check1(pchkPRI_GAISOU).Value = vbUnchecked
                        Exit For
                    End If

                Next i
'2013.11.21            End If
'2013.08.29 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>



        Case pcmbUKEHARAI       '��z��
            Text1(ptxUKEHARAI_CODE).text = Trim(Right(Combo1(pcmbUKEHARAI).text, 5))
        Case pcmbS_TANTO        '���P�^�S����

                                '�����^�\���@���
        Case pcmbD_SYUBETSU01, pcmbD_SYUBETSU02, pcmbD_SYUBETSU03, pcmbD_SYUBETSU04, pcmbD_SYUBETSU05, pcmbD_SYUBETSU06

            D_Item_Tbl(Index - pcmbD_SYUBETSU01).SYUBETSU = Right(Combo1(Index).text, 2)
    End Select

End Sub


Private Sub Combo2_Click(Index As Integer)
    
    If Right(Combo2(0).text, 1) = " " Or Right(Combo2(0).text, 1) = "2" Then
        Check1(pchkL_PAPER).Enabled = False
        Check1(pchkL_PLASTIC).Enabled = False
    Else
        Check1(pchkL_PAPER).Enabled = True
        Check1(pchkL_PLASTIC).Enabled = True
    End If


End Sub

Private Sub Combo2_GotFocus(Index As Integer)

    '2011.11.19
    If Right(Combo2(0).text, 1) = "2" Then
        Check1(pchkL_PAPER).Enabled = False
        Check1(pchkL_PLASTIC).Enabled = False
    Else
        Check1(pchkL_PAPER).Enabled = True
        Check1(pchkL_PLASTIC).Enabled = True
    End If
    '2011.11.19


End Sub

Private Sub Combo2_LostFocus(Index As Integer)

    If Right(Combo2(0).text, 1) = " " Or Right(Combo2(0).text, 1) = "2" Then
        Check1(pchkL_PAPER).Enabled = False
        Check1(pchkL_PLASTIC).Enabled = False
    Else
        Check1(pchkL_PAPER).Enabled = True
        Check1(pchkL_PLASTIC).Enabled = True
    End If

End Sub

Private Sub Command1_Click(Index As Integer)

Dim ans             As Integer
Dim i               As Integer

Dim rpt             As New PI00010F1
Dim f               As New PI000103

Dim rpt2            As New PI00010F2


Dim com             As Integer
Dim sts             As Integer


Dim Parts_F         As Integer
Dim Gaisou_F        As Integer
Dim Kishu_F         As Integer

Dim objAccess       As Access.Application
Dim strAccessPath   As String

Dim GAISOU_QTY          As Long
Dim GAISOU_SHIJI_QYU    As Long

'Dim L_print_Flg     As Boolean

Dim FileNo      As Long         '2008.05.30


'=============================== 2007/03/19 =====
Dim wk_SHIJI_NORMAL As Integer
Dim wk_SHIJI_SPOT   As Integer
Dim wk_SHIJI_KEPPIN As Integer
'================================================


'=============================== 2011.02.16 =====
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

Dim L_QTY       As Long         '2008.10.03

'=============================== 2011.02.16 =====

Dim KISHU1      As String       '2012.09.03
Dim KISHU2      As String       '2012.09.03


Dim LABEL_CHECK_F   As Boolean  '2013.11.05


Dim GYO_SU      As Long         '2016.01.05


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �G���[�`�F�b�N 2016.01.29
Dim GEN_NG_F        As Integer      '���Y����
Dim GEN_AT_GAI_F    As Integer      '���Y������(�C�O)
Dim GEN_AT_PLU_F    As Integer      '���Y������(����)
Dim TANKA_SP_F      As Integer      '�P����(�P���Q,�P���R)
Dim KISHU_NG_F      As Integer      '�@���
Dim KAISYA_NG_F     As Integer      '��Ё^���ƕ���

Dim yn              As Integer
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �G���[�`�F�b�N 2016.01.29


    For i = 0 To 97




        Select Case i
            Case ptxHIN_GAI, ptxK_HIN_GAI01, ptxK_HIN_GAI02, ptxK_HIN_GAI03, ptxK_HIN_GAI04, ptxK_HIN_GAI05, _
                    ptxG_HIN_GAI01, ptxG_HIN_GAI02, ptxG_HIN_GAI03, _
                    ptxD_HIN_GAI01, ptxD_HIN_GAI02, ptxD_HIN_GAI03, ptxD_HIN_GAI04, ptxD_HIN_GAI05, ptxD_HIN_GAI06 _


                Text1(i).text = RTrim(StrConv(Text1(i).text, vbUpperCase))

        End Select

    Next i



        GYO_SU = SendMessage(RichTextBox1(prchBIKOU).hwnd, EM_GETLINECOUNT, 0&, 0&)  '2016.01.05



    Select Case Index
        Case P_CMD_Upd        '�X�V


            For i = ptxSHIJI_NO To ptxD_BIKOU06

                If chenge_F Then        '2008.07.30

                    If Error_Check_Proc(i, 0, 0, 9) Then '�G���[�`�F�b�N
                        Exit Sub
                    End If
                Else
                    If Error_Check_Proc(i, 0, 1, 9) Then '�G���[�`�F�b�N
                        Exit Sub
                    End If
                End If

            Next i

'>>>>>> 2016.02.10
            If LenB(StrConv(RTrim(RichTextBox1(prchBIKOU).text), vbFromUnicode)) > 120 Then
                yn = MsgBox("���l�������I�[�o�[���Ă��܂�(�ő�120����)�A�I�[�o���������͐؂�̂Ă��܂��B", vbYesNo, "�m�F����")
                If yn = vbNo Then
                    RichTextBox1(prchBIKOU).SetFocus
                    Exit Sub
                End If
            End If
'>>>>>> 2016.02.10


            If GYO_SU > 5 Then                                                      '2016.01.05
                MsgBox "���l�ő�󎚍s���͂T�s�ł��B���e���m�F���ĉ������B"         '2016.01.05
                RichTextBox1(prchBIKOU).SetFocus                                    '2016.01.05
                Exit Sub                                                            '2016.01.05
            End If                                                                  '2016.01.05


            Beep
            ans = MsgBox("�X�V���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Update_Proc(0, 0) Then       '�����ǉ�   2007.11.21
                    Unload Me
                End If

                If Init_Proc() Then
                    Unload Me
                End If
                chenge_F = False                '2019.06.20 �N���A���Ȃ��ƕi�ԃ`�F�b�N���s���Ă��܂��I
                DoEvents
                Text1(ptxSHIJI_NO).SetFocus

                PI000104_OLD_HIN_GAI = ""       '2019.04.18
                
                
                
            Else
                chenge_F = False                '2019.06.20 �N���A���Ȃ��ƕi�ԃ`�F�b�N���s���Ă��܂��I
                DoEvents
                Text1(ptxHAKKO_DT).SetFocus
            End If


'        Case P_CMD_DEL                      '�폜
        Case cmdMUPDATE                     'Ͻ��X�V

            
            
            
            For i = ptxSHIJI_NO To ptxD_BIKOU06

                If chenge_F Then        '2008.07.30

                    If Error_Check_Proc(i, 1, 0, 9) Then '�G���[�`�F�b�N
                        Exit Sub
                    End If
                Else
                    If Error_Check_Proc(i, 1, 1, 9) Then '�G���[�`�F�b�N
                        Exit Sub
                    End If
                End If

            Next i
'>>>>>> 2016.02.10
            If LenB(StrConv(RTrim(RichTextBox1(prchBIKOU).text), vbFromUnicode)) > 120 Then
                yn = MsgBox("���l�������I�[�o�[���Ă��܂�(�ő�120����)�A�I�[�o���������͐؂�̂Ă��܂��B", vbYesNo, "�m�F����")
                If yn = vbNo Then
                    RichTextBox1(prchBIKOU).SetFocus
                    Exit Sub
                End If
            End If
'>>>>>> 2016.02.10

            If GYO_SU > 5 Then                                                      '2016.01.05
                MsgBox "���l�ő�󎚍s���͂T�s�ł��B���e���m�F���ĉ������B"         '2016.01.05
                RichTextBox1(prchBIKOU).SetFocus                                    '2016.01.05
                Exit Sub                                                            '2016.01.05
            End If                                                                  '2016.01.05


            Beep
            ans = MsgBox("�X�V���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Update_Proc(1, 1) Then       '�����ǉ�   2007.11.21
                    Unload Me
                End If

'                If Init_Proc() Then
'                    Unload Me
'                End If

                Call UniCode_Conv(ITEMREC.JGYOBU, "")
                Call UniCode_Conv(P_COMPO_O_REC.JGYOBU, "")
                Text1(ptxSHIJI_NO) = ""

                If Text1(ptxSHIJI_NO).Locked Then
                    Text1(ptxHAKKO_DT).SetFocus
                Else
                    Text1(ptxSHIJI_NO).SetFocus
                End If


                Text1(ptxHIN_GAI).Locked = False            '2019.03.18
    '            Text1(ptxHIN_GAI).Enabled = True           '2019.03.18
                Text1(ptxHIN_GAI).BackColor = &H80000005     '2019.03.18

                PI000104_OLD_HIN_GAI = ""       '2019.04.18

                chenge_F = False                '2019.06.20 �N���A���Ȃ��ƕi�ԃ`�F�b�N���s���Ă��܂��I
            Else
                chenge_F = False                '2019.06.20 �N���A���Ȃ��ƕi�ԃ`�F�b�N���s���Ă��܂��I
                Text1(ptxHAKKO_DT).SetFocus
            End If


        Case P_CMD_DSP                      '����/�\��
        Case cmdNext                        '�\�����i��ʂ�

            Doukon_Start = 1
            PI000102.Show vbModal           '���i�ڍ׃t�H�[���\��
            If G_SCREEN_FLG = SYS_ERR Then
                Unload Me
            End If

            'ð��ق��\���^������\��
            If Tbl_To_Disp_Proc() Then
                Unload Me
            End If

            chenge_F = False                '2019.06.20 �N���A���Ȃ��ƕi�ԃ`�F�b�N���s���Ă��܂��I

        Case P_CMD_OUT                      '�ް��o��
        
        
        
        Case P_CMD_PRT                      '���


            For i = ptxSHIJI_NO To ptxD_BIKOU06


                If chenge_F Then        '2008.07.30

                    If Error_Check_Proc(i, 0, 0, 9) Then '�G���[�`�F�b�N
                        Exit Sub
                    End If
                Else
                    If Error_Check_Proc(i, 0, 1, 9) Then '�G���[�`�F�b�N
                        Exit Sub
                    End If
                End If

            Next i
'>>>>>> 2016.02.10
            If LenB(StrConv(RTrim(RichTextBox1(prchBIKOU).text), vbFromUnicode)) > 120 Then
                yn = MsgBox("���l�������I�[�o�[���Ă��܂�(�ő�120����)�A�I�[�o���������͐؂�̂Ă��܂��B", vbYesNo, "�m�F����")
                If yn = vbNo Then
                    RichTextBox1(prchBIKOU).SetFocus
                    Exit Sub
                End If
            End If
'>>>>>> 2016.02.10

            If GYO_SU > 5 Then                                                      '2016.01.05
                MsgBox "���l�ő�󎚍s���͂T�s�ł��B���e���m�F���ĉ������B"         '2016.01.05
                RichTextBox1(prchBIKOU).SetFocus                                    '2016.01.05
                Exit Sub                                                            '2016.01.05
            End If                                                                  '2016.01.05



Debug.Print Combo1(0).text


            Beep
            
            
            
            
'>>>>>>>>   2016.02.10 �����ʒu�ύX


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �G���[�`�F�b�N 2016.01.29
            GEN_NG_F = 0        '���Y����
            GEN_AT_GAI_F = 0    '���Y������(�C�O)
            GEN_AT_PLU_F = 0    '���Y������(����)
            TANKA_SP_F = 0      '�P����(�P���Q,�P���R)
            KISHU_NG_F = 0      '�@���
            KAISYA_NG_F = 0     '��Ё^���ƕ���

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>���Y���󔒃`�F�b�N>>>>>>>>>>>>>>>>>>>>>>>>>>
            For i = 0 To UBound(GENSANKOKU_CHECK_TBL)
                If Last_JGYOBU = GENSANKOKU_CHECK_TBL(i) Then
                    Exit For
                End If
            Next i
            If i > UBound(GENSANKOKU_CHECK_TBL) Then
                GEN_NG_F = 9
            Else
                If GENSANKOKU_FLG <> "1" Then
                    GEN_NG_F = 9
                Else
                    '���Y���A�󔒂��H
                    If Trim(txGensankoku.text) = "" Then
                        GEN_NG_F = 1
                    Else
                    End If
                End If
            End If

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>�C�O�����敪�`�F�b�N>>>>>>>>>>>>>>>>>>>>>>>>>>
            If GAI_BUHIN_CHECK Then
                If Trim(lblGAI_BUHIN.Caption) = "1" Or _
                   Trim(lblGAI_BUHIN.Caption) = "2" Or _
                    Trim(lblGAI_BUHIN.Caption) = "3" Then
                    GEN_AT_GAI_F = 1
                End If
            End If
                    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>���Y���C�O�����`�F�b�N�����Y�������`�F�b�N>>>>
            If lstGensankoku.ListCount < 1 Then
                GEN_AT_PLU_F = 0
            Else
                GEN_AT_PLU_F = lstGensankoku.ListCount
            End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>�P���`�F�b�N>>>>>>>>>>>>>>>>>>>>>>>>>>
            If TANKA_SPACE_F = "1" Then
                If Not IsNumeric(lblL_URIKIN2) Or _
                     Not IsNumeric(lblL_URIKIN3) Then
                    TANKA_SP_F = 1
                End If
            End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>��\�@��`�F�b�N>>>>>>>>>>>>>>>>>>>>>>>>>>
            
            If KISHU_CHECK Then
                KISHU1 = ""
                KISHU2 = ""
                
                For i = 1 To Len(Trim(lblKISHU1.Caption))
                    If Mid(StrConv(lblKISHU1.Caption, vbUnicode), i, 1) <= " " Then
                    Else
                        KISHU1 = KISHU1 & Mid(lblKISHU1.Caption, i, 1)
                    End If
                Next i
                
                For i = 1 To Len(Trim(lblKISHU2.Caption))
                    If Mid(Len(lblKISHU2.Caption), i, 1) <= " " Then
                    Else
                        KISHU2 = KISHU2 & Mid(lblKISHU1.Caption, i, 1)
                    End If
                Next i
                
                If Trim(KISHU1) = "" And Trim(KISHU2) = "" Then
                    KISHU_NG_F = 1
                End If
            End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>��Ж��^���ƕ����`�F�b�N>>>>>>>>>>>>>>>>>>>>>>>>>>
            If KAISYA_RESTRICT_F Then
                KAISYA_NG_F = 9
            Else
                If KAISYA_CHK_F Then
                    If Trim(lblL_KAISHA.Caption) = "" Or Trim(lblL_JGYOBU.Caption) = "" Then
                        KAISYA_NG_F = 1
                    End If
                End If
            End If

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>���b�Z�[�W�쐬>>>>>>>>>>>>>>>>>>>>>>>>>>
            
            If Right(Combo2(0).text, 1) <> " " Then      '2019.03.07
            
                ans = Mesg_Set_Proc(GEN_NG_F, GEN_AT_GAI_F, GEN_AT_PLU_F, TANKA_SP_F, KISHU_NG_F, KAISYA_NG_F, KISHU1, KISHU2)
                If ans = vbCancel Then
                    GoTo Next_Step
                End If

            End If                                      '2019.03.07
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �G���[�`�F�b�N 2016.01.29




'>>>>>>>>   2016.02.10 �����ʒu�ύX
            
            
            
            
 '2016.02.10           ans = MsgBox("����^�X�V���܂����H", vbYesNo + vbQuestion, "�m�F����")
            ans = vbYes '2016.02.10
            If ans = vbYes Then
                
                If Update_Proc(0, 1) Then       '�����ǉ�   2007.11.21
                    Unload Me
                End If


                Taget_Key = Text1(ptxSHIJI_NO).text

                If Check1(pchkPRI_SHIJI).Value = vbChecked Then

                    '2008.06.26 ��
                    On Error Resume Next
                    Set objAccess = GetObject(, "Access.Application")
                    If Err().Number <> 0 Then
                        On Error GoTo 0
                    Else
                        On Error GoTo 0
                    '2008.06.26�@��


                        LABEL_CHECK_F = False '2013.11.05

                        '>>>>>>>>>> 2013.11.12
                        If Trim(Right(Combo2(0).text, 1)) <> "" Or _
                            Check1(pchkPRI_GAISOU).Value = vbChecked Then
                        
                        '>>>>>>>>>> 2013.11.12



                            '��2008.05.30
                            Do
                                
                                
                                
                                On Error Resume Next
    
    
                                FileNo = FreeFile
    
                                Open LabelPrint_F For Input As FileNo
    
                                Select Case Err.Number
                                    Case 0
    
                                        Close #FileNo
    
                                        ans = MsgBox("���x�����s���ł�", vbAbortRetryIgnore + vbDefaultButton3 + vbQuestion, "�m�F����")
    
                                        Select Case ans
    
                                            Case vbAbort    '���~
    
                                                Exit Sub
    
                                            Case vbIgnore   '����
    
                                                LABEL_CHECK_F = True    '2013.11.05
    
                                                Exit Do
    
                                        End Select
    
    
    
    
                                    Case 53
                                        Exit Do
    
    
                                    Case Else
    
    
                                        Exit Sub
    
    
                                End Select
    
                                On Error GoTo 0
    
                            Loop
    
                            'Open LabelPrint_F For Output As FileNo         2013.11.05
                            'Close #FileNo                                  2013.11.05
                            '��2008.05.30
                        End If                                  '2013.11.12
                    End If              '2008.06.26









                    If CDbl(Text1(ptxSHIJI_QTY).text) <> 0 Then '2008.02.02


'2013.01.08 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'2013.02.19                        If StrConv(ITEMREC.PRT_GENSANKOKU, vbUnicode) = "1" Then
'2013.02.19                            chk_TORI_GENSANKOKU = lblGensankoku(1).Caption
'2013.02.19                        ElseIf StrConv(ITEMREC.PRT_GENSANKOKU, vbUnicode) = "0" Then
'2013.02.19                            chk_TORI_GENSANKOKU = ""
'2013.02.19                        Else
'2013.02.19                                '�i��Ͻ����ڂ����ݒ��ini��`�l�ɂ�菉���\��
'2013.02.19                            If GENSANKOKU_FLG = "1" Then
'2013.02.19                                chk_TORI_GENSANKOKU = lblGensankoku(1).Caption
'2013.02.19                            Else
'2013.02.19                                chk_TORI_GENSANKOKU = ""
'2013.02.19                            End If
'2013.02.19                        End If

                    
                    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2016.02.01
'                        '>>>>�C�O���L�敪�̃`�F�b�N�@2015.07.23
'                        If GAI_BUHIN_CHK Then
'                            If StrConv(ITEMREC.GAI_BUHIN, vbUnicode) = "1" Or StrConv(ITEMREC.GAI_BUHIN, vbUnicode) = "2" Or StrConv(ITEMREC.GAI_BUHIN, vbUnicode) = "3" Then
'                                MsgBox "���Y������"
'                            End If
'                        Else
'                        '>>>>�C�O���L�敪�̃`�F�b�N�@2015.07.23
'
'                            If GENSANKOKU_MSG_F Then                        '2013.02.19
'                                If Trim(chk_TORI_GENSANKOKU) <> "" Then
'                                        MsgBox "���Y������"
'                                End If
'                            End If
'                        End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2016.02.01


'2013.01.08 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

                        PRINT_STOP_F = False        '2015.03.26
                        Set rpt = New PI00010F1

                        '���|�[�g��������܂��B�itrue�F����_�C�A���O���� false�F�Ȃ��j
                        
                        If Not PRINT_STOP_F Then    '2015.03.26
                            rpt.PrintReport False
                        End If                      '2015.03.26

                        Set rpt = Nothing


    '                    f.RunReport rpt
    '                    f.Show



''''2011.02.17
''''                        If Check1(pchkPRI_PARTS).Value = vbChecked Or _
''''                            Check1(pchkPRI_GAISOU).Value = vbChecked Then

                
                            If Trim(Right(Combo2(0).text, 1)) <> "" Or _
                                Check1(pchkPRI_GAISOU).Value = vbChecked Then

                            


''''2011.02.17
                            L_print_Flg = True


'>>>>>>>>   2016.02.10 �����ʒu�ύX
'
'
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �G���[�`�F�b�N 2016.01.29
'            GEN_NG_F = 0        '���Y����
'            GEN_AT_GAI_F = 0    '���Y������(�C�O)
'            GEN_AT_PLU_F = 0    '���Y������(����)
'            TANKA_SP_F = 0      '�P����(�P���Q,�P���R)
'            KISHU_NG_F = 0      '�@���
'            KAISYA_NG_F = 0     '��Ё^���ƕ���
'
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>���Y���󔒃`�F�b�N>>>>>>>>>>>>>>>>>>>>>>>>>>
'            For i = 0 To UBound(GENSANKOKU_CHECK_TBL)
'                If Last_JGYOBU = GENSANKOKU_CHECK_TBL(i) Then
'                    Exit For
'                End If
'            Next i
'            If i > UBound(GENSANKOKU_CHECK_TBL) Then
'                GEN_NG_F = 9
'            Else
'                If GENSANKOKU_FLG <> "1" Then
'                    GEN_NG_F = 9
'                Else
'                    '���Y���A�󔒂��H
'                    If Trim(txGensankoku.text) = "" Then
'                        GEN_NG_F = 1
'                    Else
'                    End If
'                End If
'            End If
'
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>�C�O�����敪�`�F�b�N>>>>>>>>>>>>>>>>>>>>>>>>>>
'            If GAI_BUHIN_CHECK Then
'                If StrConv(ITEMREC.GAI_BUHIN, vbUnicode) = "1" Or _
'                    StrConv(ITEMREC.GAI_BUHIN, vbUnicode) = "2" Or _
'                    StrConv(ITEMREC.GAI_BUHIN, vbUnicode) = "3" Then
'                    GEN_AT_GAI_F = 1
'                End If
'            End If
'
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>���Y���C�O�����`�F�b�N�����Y�������`�F�b�N>>>>
'            If lstGensankoku.ListCount < 1 Then
'                GEN_AT_PLU_F = 0
'            Else
'                GEN_AT_PLU_F = lstGensankoku.ListCount
'            End If
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>�P���`�F�b�N>>>>>>>>>>>>>>>>>>>>>>>>>>
'            If TANKA_SPACE_F = "1" Then
'                If Not IsNumeric(StrConv(ITEMREC.L_URIKIN2, vbUnicode)) Or _
'                     Not IsNumeric(StrConv(ITEMREC.L_URIKIN2, vbUnicode)) Then
'                    TANKA_SP_F = 1
'                End If
'            End If
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>��\�@��`�F�b�N>>>>>>>>>>>>>>>>>>>>>>>>>>
'
'            If KISHU_CHECK Then
'                KISHU1 = ""
'                KISHU2 = ""
'
'                For i = 1 To Len(Trim(lblKISHU1.Caption))
'                    If Mid(StrConv(lblKISHU1.Caption, vbUnicode), i, 1) <= " " Then
'                    Else
'                        KISHU1 = KISHU1 & Mid(lblKISHU1.Caption, i, 1)
'                    End If
'                Next i
'
'                For i = 1 To Len(Trim(lblKISHU2.Caption))
'                    If Mid(Len(lblKISHU2.Caption), i, 1) <= " " Then
'                    Else
'                        KISHU2 = KISHU2 & Mid(lblKISHU1.Caption, i, 1)
'                    End If
'                Next i
'
'                If Trim(KISHU1) = "" And Trim(KISHU2) = "" Then
'                    KISHU_NG_F = 1
'                End If
'            End If
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>��Ж��^���ƕ����`�F�b�N>>>>>>>>>>>>>>>>>>>>>>>>>>
'            If KAISYA_RESTRICT_F Then
'                KAISYA_NG_F = 9
'            Else
'                If KAISYA_CHK_F Then
'                    If Trim(lblL_KAISHA.Caption) = "" Or Trim(lblL_JGYOBU.Caption) = "" Then
'                        KAISYA_NG_F = 1
'                    End If
'                End If
'            End If
'
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>���b�Z�[�W�쐬>>>>>>>>>>>>>>>>>>>>>>>>>>
'            ans = Mesg_Set_Proc(GEN_NG_F, GEN_AT_GAI_F, GEN_AT_PLU_F, TANKA_SP_F, KISHU_NG_F, KAISYA_NG_F, KISHU1, KISHU2)
'            If ans = vbCancel Then
'                GoTo Next_Step
'            End If
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �G���[�`�F�b�N 2016.01.29




'>>>>>>>>   2016.02.10 �����ʒu�ύX












                            If L_URIKIN1 = 0 And L_URIKIN2 = 0 And L_URIKIN3 = 0 Then

                                Beep
'2016.02.01                                ans = MsgBox("�P�����ݒ�ł��B���x��������܂����H", vbYesNo + vbQuestion, "�m�F����")
                                ans = vbYes
                                If ans = vbYes Then
                                Else
'                                    L_print_Flg = False            '2013.08.29
                                    
                                    
                                    '>>>>> 2013.11.05
                                    'Text1(ptxHIN_GAI).SetFocus      '2013.08.29
                                    'Exit Sub                        '2013.08.29
                                    GoTo Next_Step
                                    '>>>>> 2013.11.05
                                End If
                            Else
                            End If


                            '��Ў��ƕ��G���[�����L�� 2010.07.20
                            If KAISYA_CHK_F Then

                                If Trim(lblL_KAISHA.Caption) = "" Or Trim(lblL_JGYOBU.Caption) = "" Then
'2016.0.01                                    ans = MsgBox("��Ж�/���ƕ� ���w�肳��Ă��܂���B(�n�j�����s�A��ݾ�=���s���Ȃ�)", vbOKCancel + vbQuestion + vbDefaultButton2, "�m�F����") '2013.08.29
'                                    ans = MsgBox("���/���ƕ����w�肳��Ă��܂���B���x��������܂����H", vbYesNo + vbQuestion + vbDefaultButton2, "�m�F����")                                                      '2013.08.29
                                    ans = vbYes
                                    'If ans = vbCancel Then     '2013.08.29
                                    If ans = vbNo Then          '2013.08.29
' 2013.01.05 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                        L_print_Flg = False
                                        
                                        
                                        '>>>>> 2013.11.05
                                        'Text1(ptxHIN_GAI).SetFocus
                                        'Exit Sub    '2013.08.29
                                        GoTo Next_Step
                                        '>>>>> 2013.11.05
                                    Else
'2013.08.27                                        sts = Shell(App.Path & "\PM00040.exe " & Right(Combo1(pcmbSHIMUKE).text, 2) & Trim(Text1(ptxHIN_GAI).text), vbNormalFocus)
                                    End If

'2013.08.27                                    Exit Sub
' 2013.01.05 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                                End If

                            End If
                            '��Ў��ƕ��G���[�����L�� 2010.07.20



                            '2009.03.28
                            For i = 0 To UBound(GENSANKOKU_CHECK_TBL)


                                If Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) = GENSANKOKU_CHECK_TBL(i) Then
                                    Exit For
                                End If

                            Next i



                            '2009.03.28
                            If i > UBound(GENSANKOKU_CHECK_TBL) Then
                            Else


                                If Trim(txGensankoku.text) = "" Then

                                    'ans = MsgBox("���Y�����󔒂ł��B(�n�j��������~�A��ݾ�=�p��)", vbOKCancel + vbQuestion, "�m�F����")    '2013.08.29
'2016.02.01                                    ans = MsgBox("���Y�����󔒂ł��B���x��������܂����H", vbYesNo + vbQuestion + vbDefaultButton2, "�m�F����")                    '2013.08.29
                                    ans = vbYes
                                    'If ans = vbCancel Then '2013.08.29
                                    If ans = vbYes Then     '2013.08.29
' 2013.01.05 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'2013.08.27                                        sts = Shell(App.Path & "\PM00040.exe " & Right(Combo1(pcmbSHIMUKE).text, 2) & Trim(Text1(ptxHIN_GAI).text), vbNormalFocus)
                                    Else
'                                        L_print_Flg = False
                                        '>>>>> 2013.11.05
                                        Text1(ptxHIN_GAI).SetFocus
                                        Exit Sub    '2013.08.29
                                        GoTo Next_Step
                                        '>>>>> 2013.11.05
                                    End If

'2013.08.27                                    Exit Sub
' 2013.01.05 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                                End If

                            End If



                            '2012.09.03     ��\�@������        2012.10.26 itemrec.L_KISHU1 -- > lblKISHU1,itemrec.L_KISHU2 -- > lblKISHU2
                            If KISHU_CHECK Then
                                KISHU1 = ""
                                KISHU2 = ""
                                
                                For i = 1 To Len(Trim(lblKISHU1.Caption))
                                    If Mid(StrConv(lblKISHU1.Caption, vbUnicode), i, 1) <= " " Then
                                    Else
                                        KISHU1 = KISHU1 & Mid(lblKISHU1.Caption, i, 1)
                                    End If
                                Next i
                                
                                For i = 1 To Len(Trim(lblKISHU2.Caption))
                                    If Mid(Len(lblKISHU2.Caption), i, 1) <= " " Then
                                    Else
                                        KISHU2 = KISHU2 & Mid(lblKISHU1.Caption, i, 1)
                                    End If
                                Next i
                                
                                If Trim(KISHU1) = "" And Trim(KISHU2) = "" Then
                                    'ans = MsgBox("��\�@�킪�󔒂ł��B(�n�j��������~�A��ݾ�=�p��)", vbOKCancel + vbQuestion, "�m�F����")  '2013.08.29
'2016.02.01                                    ans = MsgBox("��\�@�킪�󔒂ł��B���x��������܂���?", vbYesNo + vbQuestion + vbDefaultButton2, "�m�F����")                   '2013.08.29
                                    ans = vbYes
                                    If ans = vbYes Then
                                    Else
                                        L_print_Flg = False
                                    End If
                                End If
                            End If
                            
                            '2012.09.03     ��\�@������



                            If L_print_Flg Then


                                On Error Resume Next
                                Set objAccess = GetObject(, "Access.Application")
                                If Err().Number <> 0 Then
                                    MsgBox "���̒[���ł͏��i���x�����s�͍s���܂���B"
            '                        MsgBox "GetObject(Access.Application)" & Err().Number & " " & Err().Description
                                Else
            '                        MsgBox Err.Number


                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>   ���[���̏󋵍ă`�F�b�N  2013.11.05

                                    
                                    
                                    If Not LABEL_CHECK_F Then   '2013.11.05
                                        Do
                                            On Error Resume Next
                
                                            FileNo = FreeFile
                
                                            Open LabelPrint_F For Input As FileNo
                
                                            Select Case Err.Number
                                                Case 0
                
                                                    Close #FileNo
                
                                                    ans = MsgBox("���x�����s���ł�", vbAbortRetryIgnore + vbDefaultButton3 + vbQuestion, "�m�F����")
                
                                                    Select Case ans
                
                                                        Case vbAbort    '���~
                
                                                            Exit Sub
                
                                                        Case vbIgnore   '����
                
                                                            Exit Do
                
                                                    End Select
                
                
                
                
                                                Case 53
                                                    Exit Do
                
                
                                                Case Else
                
                
                                                    Exit Sub
                
                
                                            End Select
                
                                            On Error GoTo 0
                
                                        Loop
                                    End If                  '2013.11.05
                                    
                                    
                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>   ���[���̏󋵍ă`�F�b�N  2013.11.05
                                    
                                    Open LabelPrint_F For Output As FileNo          '2013.11.05
                                    Close #FileNo                                   '2013.11.05
                                    
                                    
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


Start_Proc1:        '2015.03.26

                                        sts = BTRV(com, L_ITEM_POS, L_ITEMREC, Len(L_ITEMREC), K0_L_ITEM, Len(K0_L_ITEM), 0)
                                        Select Case sts
                                            Case BtNoErr

Start_Proc2:        '2015.03.26
                                                
                                                sts = BTRV(BtOpDelete, L_ITEM_POS, L_ITEMREC, Len(L_ITEMREC), K0_L_ITEM, Len(K0_L_ITEM), 0)


                                                Select Case sts

                                                    Case BtNoErr
                                                    Case Else
                                                        
                                                        
                                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                                                        If sts > 3000 Or sts = 3 Then
                                    
                                        
                                                            Call File_Error(sts, BtOpGetEqual, "���ٗp�i��Ͻ�", 0)
                                
                                        
                                                            '>>>>>>>>>>>>>  2015.04.24
                                                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                                            'If sts Then
                                                            '    Call File_Error(sts, BtOpReset, "")
                                                            'End If
                                                        
                                                            'Call File_Open_Proc
                                                            Do
                                                                If Not File_Open_Proc() Then
                                                                    Exit Do
                                                                End If
                                                            Loop
                                                            '>>>>>>>>>>>>>  2015.04.24
                                            
                                                        
                                                            GoTo Start_Proc2
                                                        End If
                                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                                                        
                                                        
                                                        Call File_Error(sts, com, "���ٗp�i��Ͻ�")
                                                        Exit Sub
                                                End Select

                                            Case BtErrEOF
                                                Exit Do
                                            Case Else
                                                
                                                
                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                                                If sts > 3000 Or sts = 3 Then
                            
                                
                                                    Call File_Error(sts, BtOpGetEqual, "���ٗp�i��Ͻ�", 0)
                        
                                
                                                    '>>>>>>>>>>>>>  2015.04.24
                                                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                                    'If sts Then
                                                    '    Call File_Error(sts, BtOpReset, "")
                                                    'End If
                                                
                                                    'Call File_Open_Proc
                                                    Do
                                                        If Not File_Open_Proc() Then
                                                            Exit Do
                                                        End If
                                                    Loop
                                                    '>>>>>>>>>>>>>  2015.04.24
                                    
                                                
                                                    GoTo Start_Proc1
                                                End If
                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                                                
                                                
                                                Call File_Error(sts, com, "���ٗp�i��Ͻ�")
                                                Exit Sub
                                        End Select

                                        com = BtOpGetNext


                                    Loop

Start_Proc3:    '2015.03.26

                                    Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
                                    Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
                                    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).text)


                                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                    Select Case sts
                                        Case BtNoErr


                                            Call UniCode_Conv(ITEMREC.L_IRI_QTY, Format(GAISOU_QTY, "00000000"))

''2010.11.15                                            If GENSANKOKU_FLG = "0" Then        '���Y�� 2008.06.13
''2010.11.15                                                Call UniCode_Conv(ITEMREC.GENSANKOKU, "")
''2010.11.15
''2010.11.15                                            '2010.07.20 ��
''2010.11.15                                            Else
''2010.11.15                                                Call UniCode_Conv(ITEMREC.GENSANKOKU, lblGensankoku(1).Caption)
''2010.11.15                                            '2010.07.20 ��
''2010.11.15                                            End If
                                            
                                            
                                            
                                            
            If StrConv(ITEMREC.PRT_GENSANKOKU, vbUnicode) = "1" Then
                Call UniCode_Conv(ITEMREC.GENSANKOKU, lblGensankoku(1).Caption)
            ElseIf StrConv(ITEMREC.PRT_GENSANKOKU, vbUnicode) = "0" Then
                Call UniCode_Conv(ITEMREC.GENSANKOKU, "")
            
            Else
                    '�i��Ͻ����ڂ����ݒ��ini��`�l�ɂ�菉���\��
                If GENSANKOKU_FLG = "1" Then
                    Call UniCode_Conv(ITEMREC.GENSANKOKU, lblGensankoku(1).Caption)
                Else
                    Call UniCode_Conv(ITEMREC.GENSANKOKU, "")
                End If
            End If




Debug.Print StrConv(ITEMREC.GENSANKOKU, vbUnicode)

                                            '2008.10.29 �I��(1)�ɕW���I�Ԃ��Z�b�g
                                            Call UniCode_Conv(ITEMREC.L_TANA1, StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                                                StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                                                StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                                                StrConv(ITEMREC.ST_DAN, vbUnicode))

                                            '2008.10.29


Start_Proc4:    '2015.03.26

                                            sts = BTRV(BtOpInsert, L_ITEM_POS, ITEMREC, Len(ITEMREC), K0_L_ITEM, Len(K0_L_ITEM), 0)
                                            Select Case sts
                                                Case BtNoErr

                                                    '2007.12.11objAccess.Run "PosPrintLabel", Trim(Text1(ptxHIN_GAI).text), CLng(Text1(ptxSHIJI_QTY).text), Parts_F, Gaisou_F, Kishu_F, GAISOU_QTY, GAISOU_SHIJI_QYU, 0
''2011.02.17 PosPrintLabel-->NewPosPrintLabel
''                                                    objAccess.Run "PosPrintLabel", _
''                                                                    Trim(Text1(ptxHIN_GAI).text), _
''                                                                    CLng(Text1(ptxLabel_QTY).text), _
''                                                                    Parts_F, _
''                                                                    Gaisou_F, _
''                                                                    Kishu_F, _
''                                                                    GAISOU_QTY, _
''                                                                    GAISOU_SHIJI_QYU, _
''                                                                    0




                                                    PartsLabel = 0
                                                    KisyuLabel = 0
                                                    JanLabel = 0
                                                    GLabel = 0
                                                    ItemLabel = 0


                                                    '�i�ڃR�[�h
                                                    Parts = Text1(ptxHIN_GAI).text
                                                    '�p�[�c���x��
                                                    If Right(Combo2(0).text, 1) = "0" Then
                                                        PartsLabel = CLng(Text1(ptxLabel_QTY).text)
                                                    End If
                                                    '�@�탉�x��
                                                    If Right(Combo2(0).text, 1) = "1" Then
                                                        KisyuLabel = CLng(Text1(ptxLabel_QTY).text)
                                                    End If
                                                    'Jan���x��
                                                    If Right(Combo2(0).text, 1) = "2" Then
                                                        JanLabel = CLng(Text1(ptxLabel_QTY).text)
                                                    End If

                                                    '�O������
                                                    If Check1(pchkPRI_GAISOU).Value = vbChecked Then
                                                        GLabel = CLng(Text1(ptxG_SHIJI_QTY01).text)
                                                    End If
                                                    'ID
                                                    ID = 0
                                                    '�A�C�e�����x��
                                                    ItemLabel = 0
                                                    '�I�[�_�[��
                                                    OrderNo = ""
                                                    '�A�C�e����
                                                    ItemNo = ""
                                                    '������t
                                                    Pri_Date = Format(Now, "yyyy/mm/dd")
                                                    '����
                                                    L_QTY = 1
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



                                                Case Else
                                                    
                                                    
                                                    
                                                    
                                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                                                    If sts > 3000 Or sts = 3 Then
                                
                                    
                                                        Call File_Error(sts, BtOpGetEqual, "���ٗp�i�ڃ}�X�^", 0)
                            
                                    
                            
                                                        '>>>>>>>>>>>>>  2015.04.24
                                                        'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                                        'If sts Then
                                                        '    Call File_Error(sts, BtOpReset, "")
                                                        'End If
                                                    
                                                        'Call File_Open_Proc
                                        
                                                        Do
                                                            If Not File_Open_Proc() Then
                                                                Exit Do
                                                            End If
                                                        Loop
                                                        '>>>>>>>>>>>>>  2015.04.24
                                                    
                                                        GoTo Start_Proc4
                                                    End If
                                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                                                    
                                                    
                                                    Call File_Error(sts, BtOpInsert, "���ٗp�i�ڃ}�X�^")
                                                    Exit Sub


                                            End Select

                                        Case BtErrKeyNotFound

                                        Case Else
                                            
                                            
                                            
                                            
                                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                                            If sts > 3000 Or sts = 3 Then
                        
                            
                                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                    
                            
                                                '>>>>>>>>>>>>>  2015.04.24
                                                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                                'If sts Then
                                                '    Call File_Error(sts, BtOpReset, "")
                                                'End If
                                            
                                                'Call File_Open_Proc
                                                    
                                                Do
                                                    If Not File_Open_Proc() Then
                                                        Exit Do
                                                    End If
                                                Loop
                                                '>>>>>>>>>>>>>  2015.04.24
                                            
                                                GoTo Start_Proc3
                                            End If
                                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                                            
                                            
                                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                            Exit Sub

                                    End Select





                                    Set objAccess = Nothing
                                End If



                            End If
                        End If





                    Else                                                '2008.02.02
                        Taget_SHIMUKE_CODE_KEY = Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2)
                        Taget_JGYOBU_key = Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1)
                        Taget_NAIGAI_key = Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1)
                        Taget_Hin_key = Trim(Text1(ptxHIN_GAI).text)    '2008.02.02
                        Set rpt2 = New PI00010F2                        '2008.02.02
                        '���|�[�g��������܂��B�itrue�F����_�C�A���O���� false�F�Ȃ��j
                        rpt2.PrintReport False                          '2008.02.02
                        Set rpt2 = Nothing                              '2008.02.02
                    End If                                              '2008.02.02

                End If

                '���ټ��ш���v��




Next_Step:



                '=============================== 2007/03/19 =====
                wk_SHIJI_NORMAL = Option1(poptSHIJI_NORMAL).Value
                wk_SHIJI_SPOT = Option1(poptSHIJI_SPOT).Value
                wk_SHIJI_KEPPIN = Option1(poptSHIJI_KEPPIN).Value
                '================================================

                If Init_Proc() Then
                    Unload Me
                End If
                                '2019.06.18 ������鎞�́Achenge_F��False�ɂ����B�Q�s�ǉ�
                DoEvents
                chenge_F = False

                '=============================== 2007/03/19 =====
                Option1(poptSHIJI_NORMAL).Value = wk_SHIJI_NORMAL
                Option1(poptSHIJI_SPOT).Value = wk_SHIJI_SPOT
                Option1(poptSHIJI_KEPPIN).Value = wk_SHIJI_KEPPIN

'                Text1(ptxSHIJI_NO).SetFocus
                Text1(ptxHIN_GAI).SetFocus
                '================================================


            Else
                                '2019.06.18 ������鎞�́Achenge_F��False�ɂ����B�Q�s�ǉ�
                DoEvents
                chenge_F = False
                
                
                Text1(ptxHAKKO_DT).SetFocus
            End If


        Case 9                              'COPY���   2019.03.14
            
            PI000104.Show vbModal
        
                    
            If PI000104_CANCEL_F = 1 Then
                Exit Sub
            End If
        
        
        
            If PI000104_Error_F = 1 Then
                Unload Me
            End If
            PI000104_OLD_HIN_GAI = ""
            If Trim(PI000104_HIN_GAI) <> "" Then
                PI000104_OLD_HIN_GAI = Text1(ptxHIN_GAI).text
            End If
                    
            Text1(ptxHIN_GAI).text = PI000104_HIN_GAI
            
                    
            
            
            '2019.05.27 �@�@�@�@�@�@�@�@�@�@�������̂Q�Ԗڂ́A�P�ł́H
'            If Error_Check_Proc(ptxHIN_GAI, 0, 0) Then
'                Exit Sub
'            End If
            '2019.05.27                     ������ύX���Ă݂��B    ����
            If Error_Check_Proc(ptxHIN_GAI, 1, 0) Then
                Exit Sub
            End If
            
            
            
            Command1(cmdMUPDATE).SetFocus
        
        
            Text1(ptxHIN_GAI).Locked = True             '2019.03.18
            Text1(ptxHIN_GAI).BackColor = &H8000000F    '2019.03.18
'            Text1(ptxHIN_GAI).Enabled = False           '2019.03.18
        
        
        Case cmdCen                         '������
            If Init_Proc() Then
                Unload Me
            End If
            
            
            Text1(ptxHIN_GAI).Locked = False            '2019.03.18
'            Text1(ptxHIN_GAI).Enabled = True           '2019.03.18
            Text1(ptxHIN_GAI).BackColor = &H80000005     '2019.03.18
            
            '2019.06.12 ���L�Q�s�ǉ�
            DoEvents
            chenge_F = False
            
            
            Text1(ptxSHIJI_NO).SetFocus
        Case P_CMD_End                      '�I��
            Unload Me
    
    
        Case 12                                         '2019.03.18
            
            PI000104_OLD_HIN_GAI = ""       '2019.04.18

            
            Text1(ptxHIN_GAI).Locked = False            '2019.03.18
'            Text1(ptxHIN_GAI).Enabled = True           '2019.03.18
            Text1(ptxHIN_GAI).BackColor = &H80000005     '2019.03.18
            
            If Init_Proc() Then
                Unload Me
            End If
            Text1(ptxSHIJI_NO).SetFocus

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


    If App.PrevInstance Then
        MsgBox "����v���O�������s���ł��B"
        End
    End If

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.10.09
    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "���i���w�}�[���s�@�u�N���������v", Me.hwnd, 0)
    '�Ō�̗v�f��-1�ɂ����
    '�e�E�B���h�E�̑S�̂̕��̎c��̕���
    '�����I�Ɋ��蓖�Ă�
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)
                                
                                
    Me.Enabled = False
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.10.09
                                
                                
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
        End
    End If
    LOG_F = RTrim(c)


                                '���x������p�R���g���[���e�l��2008.05.30
    If GetIni("FILE", "labelprint", "SYS", c) Then
        Beep
        MsgBox "���x������p�R���g���[���e�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LabelPrint_F = RTrim(c)

Show    '2015.03.26
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2016.01.29  PI00010.INI --> P_SYS.INI[PLABEL]
'------------------------------------------ P_SYS.INI--> PI00010.INI 2011.08.04

                                '���Y���󎚗L�� 2008.06.12
    If GetIni("PLABEL", "GENSANKOKU_DEF_F", "P_SYS", c) Then
        GENSANKOKU_FLG = "0"
    Else
        GENSANKOKU_FLG = RTrim(c)
    End If


                                '���Y�������� 2009.03.28
    If GetIni("PLABEL", "GENSANKOKU_CHECK", "P_SYS", c) Then
        ReDim GENSANKOKU_CHECK_TBL(0 To 0)
        GENSANKOKU_CHECK_TBL(0) = "*"
    Else
        GENSANKOKU_CHECK_TBL = Split(Trim(c), ",", -1)
    End If


                                '��\�@������   2012.09.03
    If GetIni("PLABEL", "KISHU_CHECK", "P_SYS", c) Then
        KISHU_CHECK = False
    Else
        If Trim(c) = "1" Then
            KISHU_CHECK = True
        Else
            KISHU_CHECK = False
        End If
    End If


                                '��Ў��ƕ��G���[�����L�� 2010.07.20
'    If GetIni(App.EXEName, "KAISYA_CHECK", App.EXEName, c) Then
    If GetIni("PLABEL", "KAISYA_CHECK", "P_SYS", c) Then
        KAISYA_CHK_F = False
    Else

        If Trim(c) = "1" Then
            KAISYA_CHK_F = True
        Else
            KAISYA_CHK_F = False
        End If
    End If
'

                                '���Y���C�O�����敪���� 2016.02.01
    If GetIni("PLABEL", "GAI_BUHIN_CHK", "P_SYS", c) Then
        GAI_BUHIN_CHK = False
    Else

        If Trim(c) = "1" Then
            GAI_BUHIN_CHK = True
        Else
            GAI_BUHIN_CHK = False
        End If
    End If

                                '���Y���C�O�����敪���� 2016.02.01
    If GetIni("PLABEL", "TANKA_SPACE_F", "P_SYS", c) Then
        TANKA_SPACE_F = "0"
    Else
        If Trim(c) = "1" Then
            TANKA_SPACE_F = "1"
        Else
            TANKA_SPACE_F = "0"
        End If
    End If



'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2016.01.29  PI00010.INI --> P_SYS.INI[PLABEL]

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> ���Y��ү���ނ̕\�� 2013.02.19
    If GetIni(App.EXEName, "GENSANKOKU_MSG_F", App.EXEName, c) Then
        GENSANKOKU_MSG_F = False
    Else

        If Trim(c) = "1" Then
            GENSANKOKU_MSG_F = True
        Else
            GENSANKOKU_MSG_F = False
        End If
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> ���Y��ү���ނ̕\�� 2013.02.19
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>> ��Ж�/���ƕ�����\���ݒ� 1=��\���ݒ�L�� 2016.02.01
    If GetIni(App.EXEName, "KAISYA_RESTRICT_F", App.EXEName, c) Then
        KAISYA_RESTRICT_F = False
    Else
        If Trim(c) = "1" Then
            KAISYA_RESTRICT_F = True
        Else
            KAISYA_RESTRICT_F = False
        End If
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>> ��Ж�/���ƕ�����\���ݒ� 1=��\���ݒ�L�� 2016.02.01




                                '�o�׃��O�t�@�C������荞��
    If SYUKA_LOGF_GET_PROC() Then
        Beep
        MsgBox "�o�׃��O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
        End
    End If

                                '��z���荞��
    If GetIni(StrConv(App.EXEName, vbUpperCase), "TEHAI", App.EXEName, c) Then
    Else
        TEHAI = RTrim(c)
    End If

                                'POS���їL���̎�荞��
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
    If GetIni(StrConv(App.EXEName, vbUpperCase), "DET_BIKOU", App.EXEName, c) Then
        PRI_BIKOU_BCR = False
    Else
        If RTrim(c) = "0" Then
            PRI_BIKOU_BCR = False
        Else
            If Not POS_UMU Then
                MsgBox "�o�n�r���т����ݒ�ł��B�����𒆎~���܂��B"
                End
            End If
            PRI_BIKOU_BCR = True
        End If
    End If

                                '���P�^�S���҂̎�荞��
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
    If GetIni(StrConv(App.EXEName, vbUpperCase), "DOUKON", App.EXEName, c) Then
        PRI_DOUKON = False
    Else
        
        '2011.08.04
'        If RTrim(c) = "0" Then
'            PRI_DOUKON = False
'        Else
'            PRI_DOUKON = True
'        End If
        
        Select Case RTrim(c)
            Case "0"
                PRI_DOUKON = 0
            Case "1"
                PRI_DOUKON = 1
            Case "2"
                PRI_DOUKON = 2
            Case Else
                PRI_DOUKON = 0
        End Select
        '2011.08.04
    End If
                                '���Ɋ�����̎�荞��
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
    If GetIni(StrConv(App.EXEName, vbUpperCase), "JISEKI", App.EXEName, c) Then
        JISEKI_TITLE = ""
    Else
        JISEKI_TITLE = Split(Trim(c), ",", -1)
    End If

                                '����
    If GetIni(StrConv(App.EXEName, vbUpperCase), "TASEKI", App.EXEName, c) Then
        TASEKI_TITLE = ""
    Else
        TASEKI_TITLE = Split(Trim(c), ",", -1)
    End If

                                '���o�^�i�Ԃ̉�
    If GetIni(StrConv(App.EXEName, vbUpperCase), "HIN_INV", App.EXEName, c) Then
        HIN_INV = False
    Else
        If Trim(c) = "0" Then
            HIN_INV = False
        Else
            HIN_INV = True
        End If

    End If



    If PRI_BIKOU_BCR Then
                                    '������
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

    '�b�^����荞�� 2008.08.19
    If GetIni(StrConv(App.EXEName, vbUpperCase), "JISSEKI_DSP", App.EXEName, c) Then
        JISSEKI_DSP = "m"
    Else
        JISSEKI_DSP = Trim(c)
        If JISSEKI_DSP <> "m" And JISSEKI_DSP <> "s" Then
            JISSEKI_DSP = "m"
        End If
    End If




'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> �J���E���[�h�����E�S���h�~�E�̕\�� 2013.01.16
    If GetIni(App.EXEName, "KAIKON_PRI", App.EXEName, c) Then
        KAIKON_PRI = False
    Else

        If Trim(c) = "1" Then
            KAIKON_PRI = True
        Else
            KAIKON_PRI = False
        End If
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> �J���E���[�h�����E�S���h�~�E�̕\�� 2013.01.16







'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    �����i�d������  2013.08.29
    If GetIni(App.EXEName, "SHIMUKE_CHK", App.EXEName, c) Then
        
        
        ReDim SHIMUKE_CHK_TBL(0 To 0)
        SHIMUKE_CHK_TBL(0) = "**"
    Else
    
        SHIMUKE_CHK_TBL = Split(Trim(c), ",", -1)
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    �����i�d������  2013.08.29





'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> ���׎��ɏՍނ̕\�� 2013.11.05
    If GetIni(App.EXEName, "NYUKA_KANSYOZAI", App.EXEName, c) Then
        NYUKA_KANSYOZAI = False
    Else

        If Trim(c) = "1" Then
            NYUKA_KANSYOZAI = True
        Else
            NYUKA_KANSYOZAI = False
        End If
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> ���׎��ɏՍނ̕\�� 2013.11.05



'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> ���x�����s�̕\�� 2019.03.07
    If GetIni(App.EXEName, "LABEL_PRINT_F", App.EXEName, c) Then
        LABEL_PRINT_F = 0
    Else

        If Trim(c) = "1" Then
            LABEL_PRINT_F = 1
        Else
            LABEL_PRINT_F = 0
        End If
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> ���x�����s�̕\�� 2019.03.07


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> �O�����x�����s�̕\�� 2019.03.07
    If GetIni(App.EXEName, "GA_LABEL_PRINT_F", App.EXEName, c) Then
        GA_LABEL_PRINT_F = 0
    Else

        If Trim(c) = "1" Then
            GA_LABEL_PRINT_F = 1
        Else
            GA_LABEL_PRINT_F = 0
        End If
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> �O�����x�����s�̕\�� 2019.03.07



'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> ���x�����s�����̎w�� 2015.04.02
    If GetIni(App.EXEName, "LABEL_PLUS", App.EXEName, c) Then
        LABEL_PLUS = 1
    Else
        If Not IsNumeric(Trim(c)) Then
            LABEL_PLUS = 1
        Else
            LABEL_PLUS = Val(Trim(c))
        End If
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> ���x�����s�����̎w�� 2015.04.02


    PI000101.Caption = Last_Update_day      '2016.02.10



'���x���I��     2011.02.10
    Combo2(0).Clear
    Combo2(0).AddItem "���x���Ȃ��@�@" & "          " & " "
    Combo2(0).AddItem "�p�[�c���x���@" & "          " & "0"
    Combo2(0).AddItem "�K�p�@�탉�x��" & "          " & "1"
    Combo2(0).AddItem "JAN���x���@�@ " & "          " & "2"





'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.04.24
'
'                                '���ԃ}�X�^�n�o�d�m
'    If HATUBAN_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'                                '�i�ڃ}�X�^�n�o�d�m
'    If ITEM_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'
'                                '���i���ٗp�i�ڃ}�X�^�n�o�d�m
'    If L_ITEM_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'
'                                '�N���X�}�X�^�n�o�d�m
'    If P_Class_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'                                '�R�[�h�}�X�^�n�o�d�m
'    If P_CODE_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'                                '�\���}�X�^�n�o�d�m
'    If P_COMPO_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'                                '�Ǘ��}�X�^�n�o�d�m
'    If P_KANRI_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'                                '���i���w�}�i�q�j�ް��n�o�d�m
'    If P_SSHIJI_K_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'                                '���i���w�}�i�e�j�ް��n�o�d�m
'    If P_SSHIJI_O_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'                                '�S���҃}�X�^�n�o�d�m
'    If TANTO_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'                                '�o�ח\���ް��n�o�d�m
'    If Y_SYU_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'                                '�󕥐�}�X�^�n�o�d�m
'    If P_UKEHARAI_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'
'
'    '2010.07.20 ��
'                                '���Y���}�X�^�n�o�d�m
'    If GENSAN_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'    '2010.07.20 ��
'                                '�݌��ް��n�o�d�m
'    If ZAIKO_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'
'                                '���i���w�}�i�e�jܰ��n�o�d�m
'    If wP_SSHIJI_O_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'
'                                '���o�ɒP���ݒ�}�X�^�n�o�d�m   2008.09.20
'    If SE_LOC_TANKA_M_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'
'
'    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  MT_2009.06.01
'                                'PN�}�X�^�n�o�d�m
'    If PN_M_Open(0) Then
'        Beep
'        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
'        Unload Me
'    End If
'    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                            
    Do
        If Not File_Open_Proc() Then
            Exit Do
        End If
    Loop
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.04.24


    '����Ͻ���`
    Call P_CODE_TBL_Proc



    Load PI000102
    Load PI000103



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

'2009.03.25
    Combo1(pcmbSHIMUKE).ListIndex = 0
    Last_JGYOBU = Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1)      '2016.02.01
    
    PI000104_OLD_HIN_GAI = ""       '2019.03.14
    


    chenge_F = False


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.10.09
    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "���i���w�}�[���s�@�u���������v", Me.hwnd, 0)
    '�Ō�̗v�f��-1�ɂ����
    '�e�E�B���h�E�̑S�̂̕��̎c��̕���
    '�����I�Ɋ��蓖�Ă�
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)
                                
                                
    Me.Enabled = True
    DoEvents
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.10.09
    Text1(ptxSHIJI_NO).SetFocus
    

End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer



    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  MT_2009.06.01
                                            'PN�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, PN_M_POS, PN_MREC, Len(PN_MREC), K0_PN_M, Len(K0_PN_M), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
'            Call File_Error(sts, BtOpClose, "PN�}�X�^")             2015.05.14
            Call File_Error(sts, BtOpClose, "PN�}�X�^", 0)          '2015.05.14
'2015.03.26            Beep
'2015.03.26            MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly
        End If
    End If
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<



                                            '���ԃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), K0_HATUBAN, Len(K0_HATUBAN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ԃ}�X�^", 0)
        End If
    End If

                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^", 0)
        End If
    End If
                                            '���i���ٗp�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, L_ITEM_POS, L_ITEMREC, Len(L_ITEMREC), K0_L_ITEM, Len(K0_L_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���i���ٗp�i�ڃ}�X�^", 0)
        End If
    End If

                                            '�N���X�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�N���X�}�X�^", 0)
        End If
    End If

                                            '�R�[�h�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�R�[�h�}�X�^", 0)
        End If
    End If

                                            '�\���}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�\���}�X�^", 0)
        End If
    End If
                                            '�Ǘ��}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�Ǘ��}�X�^", 0)
        End If
    End If
                                            '���i���w�}�ް�(�e)�b�k�n�r�d
    sts = BTRV(BtOpClose, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���i���w�}�ް�(�e)", 0)
        End If
    End If
                                            '���i���w�}�ް�(�q)�b�k�n�r�d
    sts = BTRV(BtOpClose, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���i���w�}�ް�(�q)", 0)
        End If
    End If

                                            '�S���҃}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�S���҃}�X�^", 0)
        End If
    End If

                                            '�o�ח\��f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�o�ח\��f�[�^", 0)
        End If
    End If
                                            '�󕥐�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�󕥐�}�X�^", 0)
        End If
    End If
                                            '�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^", 0)
        End If
    End If

                                            '���i���w�}ܰ�(�e)�b�k�n�r�d
    sts = BTRV(BtOpClose, wP_SSHIJI_O_POS, wP_SSHIJI_O_REC, Len(wP_SSHIJI_O_REC), K0_wP_SSHIJI_O, Len(K0_wP_SSHIJI_O), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���i���w�}(�e)ܰ�", 0)
        End If
    End If


    sts = BTRV(BtOpReset, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "", 0)
    End If
    Set PI000101 = Nothing
    Set PI000102 = Nothing
    Set PI000103 = Nothing

    End
End Sub

Private Sub RichTextBox1_GotFocus(Index As Integer)
Dim sts         As Integer

Dim Sumi_Qty    As Long
Dim Mi_Qty      As Long


    Select Case Index
'        Case ptxHIN_GAI
        Case Else
            If chenge_F Then

Start_Proc1:        '2015.03.26
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

'2010.11.12 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                        Check1(pchkL_PAPER).Value = vbUnchecked         '��
                        Check1(pchkL_PLASTIC).Value = vbUnchecked       '�v��
                        Check1(pchkL_LABEL).Value = vbUnchecked         '�K�p�@�탉�x��
'2010.11.12 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<



                        MsgBox "���͂������ڂ̓G���[�ł��B(�i��)"

                        chenge_F = False

                        Text1(ptxHIN_GAI).SetFocus
                        Exit Sub
                    Case Else
                        
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        If sts > 3000 Or sts = 3 Then
    
                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                            Do                                  '2015.04.24
                                If Not File_Open_Proc() Then    '2015.04.24
                                    Exit Do                     '2015.04.24
                                End If                          '2015.04.24
                            Loop                                '2015.04.24
                            '>>>>>>>>>>>>>  2015.04.24
                            
                            GoTo Start_Proc1
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        
                        
                        
                        
                        
                        
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                        Unload Me

                End Select



                Text1(ptxHIN_NAME).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)

'2010.11.12 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                If StrConv(ITEMREC.L_PAPER, vbUnicode) = L_PAPER_ON Then        '��
                    Check1(pchkL_PAPER).Value = vbChecked
                Else
                    Check1(pchkL_PAPER).Value = vbUnchecked
                End If
    
                If StrConv(ITEMREC.L_PLASTIC, vbUnicode) = L_PLASTIC_ON Then    '�v��
                    Check1(pchkL_PLASTIC).Value = vbChecked
                Else
                    Check1(pchkL_PLASTIC).Value = vbUnchecked
                End If
    
                If StrConv(ITEMREC.L_LABEL, vbUnicode) = L_LABEL_ON Then        '�K�p�@�탉�x��
                    Check1(pchkL_LABEL).Value = vbChecked
                Else
                    Check1(pchkL_LABEL).Value = vbUnchecked
                End If
'2010.11.12 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<








                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Text1(ptxST_LOCATION).text = ""
                Else
                    Text1(ptxST_LOCATION).text = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If

                If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                        StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                        StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then

                    Unload Me
                End If

                Text1(ptxMI_QTY).text = Format(Mi_Qty, "#0")
                Text1(ptxSUMI_QTY).text = Format(Sumi_Qty, "#0")



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
                            Case Else
                                
                                
                                
                                
                                
                                
                                
                                
                                Call File_Error(sts, BtOpGetEqual, "�\���}�X�^")
                                Unload Me
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
                            Case Else
                                
                                
                                

                                
                                
                                
                                Call File_Error(sts, BtOpGetEqual, "�\���}�X�^")
                                Unload Me
                        End Select

                    End If
                End If
'                Text1(ptxSHIJI_QTY).SetFocus

                chenge_F = False

                RichTextBox1(Index).SetFocus

            End If
    End Select

End Sub

Private Sub Text1_Change(Index As Integer)
    If Index = ptxSHIJI_QTY Then        '2008.02.27
        Text1(ptxLabel_QTY) = ""
    End If

    If Index = ptxHIN_GAI Then          '2008.07.10
'        If Trim(Text1(ptxHIN_GAI)) <> "" Then
            chenge_F = True
'        End If
    End If

End Sub

Private Sub Text1_GotFocus(Index As Integer)


Dim sts         As Integer

Dim Sumi_Qty    As Long
Dim Mi_Qty      As Long
Dim wkINDEX     As Integer

    If Index = ptxHIN_GAI Then
        If Trim(Text1(ptxHIN_GAI).text) = "" Then
        
        
            If Text1(Index).TabStop = True Then
                Text1(Index) = Trim(Text1(Index).text)
                Text1(Index).SelStart = 0
                Text1(Index).SelLength = Len(Text1(Index).text)
            End If
    
    
            Exit Sub
        End If
    End If
    
'    MsgBox "Index=" & Index
    
    Select Case Index
        
        Case ptxHIN_GAI
            Text1(Index) = Trim(Text1(Index).text)
            Text1(Index).SelStart = 0
            Text1(Index).SelLength = Len(Text1(Index).text)
            Exit Sub
        Case Else
        
'            If chenge_F Then
            '2019.06.20 �����������ŕi�ԁ��󔒂ł̃`�F�b�N�Ƃ����B
            If chenge_F = True And Trim(Text1(ptxHIN_GAI)) <> "" Then
'                MsgBox "chenge_F=True"
                
Start_Proc1:        '2015.03.26
                wkINDEX = Index
                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).text)


                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                sts = BTRV(BtOpGetGreaterEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        Text1(ptxHIN_GAI) = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
                    Case BtErrKeyNotFound, BtErrEOF

                        Text1(ptxHIN_NAME).text = ""
                        Text1(ptxST_LOCATION).text = ""
                        Text1(ptxMI_QTY).text = ""
                        Text1(ptxSUMI_QTY).text = ""

'2010.11.12 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                        Check1(pchkL_PAPER).Value = vbUnchecked         '��
                        Check1(pchkL_PLASTIC).Value = vbUnchecked       '�v��
                        Check1(pchkL_LABEL).Value = vbUnchecked         '�K�p�@�탉�x��
'2010.11.12 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


                        MsgBox "���͂������ڂ̓G���[�ł��B(�i��)"


                        Text1(ptxHIN_GAI).SetFocus

                        Text1(ptxHIN_GAI) = Trim(Text1(ptxHIN_GAI).text)
                        Text1(ptxHIN_GAI).SelStart = 0
                        Text1(ptxHIN_GAI).SelLength = Len(Text1(ptxHIN_GAI).text)


                        Exit Sub
                    Case Else
                        
                        
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        If sts > 3000 Or sts = 3 Then
    
                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
                           
            
            
            
                        
                            GoTo Start_Proc1
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        
                        
                        
                        
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                        Unload Me

                End Select



                Text1(ptxHIN_NAME).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)



'2010.11.12 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                If StrConv(ITEMREC.L_PAPER, vbUnicode) = L_PAPER_ON Then        '��
                    Check1(pchkL_PAPER).Value = vbChecked
                Else
                    Check1(pchkL_PAPER).Value = vbUnchecked
                End If
    
                If StrConv(ITEMREC.L_PLASTIC, vbUnicode) = L_PLASTIC_ON Then    '�v��
                    Check1(pchkL_PLASTIC).Value = vbChecked
                Else
                    Check1(pchkL_PLASTIC).Value = vbUnchecked
                End If
    
                If StrConv(ITEMREC.L_LABEL, vbUnicode) = L_LABEL_ON Then        '�K�p�@�탉�x��
                    Check1(pchkL_LABEL).Value = vbChecked
                Else
                    Check1(pchkL_LABEL).Value = vbUnchecked
                End If
'2010.11.12 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<




                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Text1(ptxST_LOCATION).text = ""
                Else
                    Text1(ptxST_LOCATION).text = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If

                If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                        StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                        StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then

                    Unload Me
                End If

                Text1(ptxMI_QTY).text = Format(Mi_Qty, "#0")
                Text1(ptxSUMI_QTY).text = Format(Sumi_Qty, "#0")



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
                            Case Else
                                
                                
                                
                                Call File_Error(sts, BtOpGetEqual, "�\���}�X�^")
                                Unload Me
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
                            Case Else
                                
                                
                                
                                Call File_Error(sts, BtOpGetEqual, "�\���}�X�^")
                                Unload Me
                        End Select

                    End If
                End If
    '                Text1(ptxSHIJI_QTY).SetFocus


                chenge_F = False
                Text1(wkINDEX).SetFocus
                Exit Sub
            End If

    End Select


    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).text)
    End If

    
    

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim WK_STR  As String


    If KeyCode <> vbKeyReturn Then Exit Sub
    


    Select Case Index
        Case ptxHIN_GAI, ptxK_HIN_GAI01, ptxK_HIN_GAI02, ptxK_HIN_GAI03, ptxK_HIN_GAI04, ptxK_HIN_GAI05, _
                ptxG_HIN_GAI01, ptxG_HIN_GAI02, ptxG_HIN_GAI03, _
                ptxD_HIN_GAI01, ptxD_HIN_GAI02, ptxD_HIN_GAI03, ptxD_HIN_GAI04, ptxD_HIN_GAI05, ptxD_HIN_GAI06 _


            Text1(Index).text = RTrim(StrConv(Text1(Index).text, vbUpperCase))

    End Select
    
    '2019.06.04 �ǉ�                �v�]�̉��߂Łu����v�����Ɠ��ꏈ���Ƃ��Ēǉ��I
    If Index = ptxHIN_GAI Then
        WK_STR = Trim(Text1(Index))
'        Call Command1_Click(10)             '�������L�[����
        
        '2019.06.05 ��L��Call�ł͂Ȃ��A���̓��e���R�s�[�����B
        '           �w�}�[����SetFocus���Ă����ׁI
'        If Init_Proc() Then
        '2019.06.10 ����ύX�F�u���O�A�X�|�b�g�A���i�����v�̃N���A���Ȃ��I
        
        '2019.08.23 ���ꕔ�A�C���E�E�E�w�}���̃R���{�����������Ȃ��I
        If Init_Proc_2() Then
            Unload Me
        End If
            
        Text1(ptxHIN_GAI).Locked = False            '2019.03.18
'            Text1(ptxHIN_GAI).Enabled = True           '2019.03.18
        Text1(ptxHIN_GAI).BackColor = &H80000005     '2019.03.18

        
        Text1(Index) = WK_STR
        DoEvents
    End If
    '2019.06.04 �����܂Œǉ�
    

    If Error_Check_Proc(Index, 0, 0) Then   '�G���[�`�F�b�N
        Exit Sub
    End If
    
    DoEvents
    
    If Index = ptxHIN_GAI Then
        Text1(ptxSHIJI_QTY) = Trim(Text1(ptxSHIJI_QTY).text)
        Text1(ptxSHIJI_QTY).SelStart = 0
        Text1(ptxSHIJI_QTY).SelLength = Len(Text1(ptxSHIJI_QTY).text)
        DoEvents
        Exit Sub
    Else
        Call Tab_Ctrl(Shift)        '�ړ�
    End If
End Sub

Private Function Init_Proc() As Integer
'----------------------------------------------------------------------------
'                   ���͉�ʂ̏����ݒ�
'----------------------------------------------------------------------------
Dim i           As Integer
Dim sts         As Integer

Dim TANTO_CODE  As String
Dim TANTO_NAME  As String

    Init_Proc = True

    Text1(ptxSHIJI_NO).BackColor = G_INPUT_OK
    Text1(ptxSHIJI_NO).Locked = False
    Text1(ptxSHIJI_NO).TabStop = True


    Combo1(pcmbS_TANTO).Enabled = PRI_S_TANTO

    TANTO_CODE = Text1(ptxTANTO_CODE).text
    TANTO_NAME = Text1(ptxTANTO_NAME).text

    For i = ptxSHIJI_NO To ptxLabel_QTY
        Text1(i).text = ""
    Next i
    Text1(ptxTANTO_CODE).text = TANTO_CODE
    Text1(ptxTANTO_NAME).text = TANTO_NAME



    RichTextBox1(prchBIKOU).text = ""

    For i = pchkSAMPLE_F To pchkPRI_KISHU
        Check1(i).Value = vbChecked
    Next i
    Check1(pchkSAMPLE_F).Value = vbUnchecked    '���{�쐬
    Check1(pchkPRI_KISHU).Value = vbUnchecked   '�o�͑Ώہ@�@������

    Check1(pchkL_PAPER).Value = vbUnchecked     '��           2010.11.12
    Check1(pchkL_PLASTIC).Value = vbUnchecked   '��׽���      2010.11.12
    Check1(pchkL_LABEL).Value = vbUnchecked     '�K�p�@������ 2010.11.12


'2009.03.25
'    For i = pcmbSHIMUKE To pcmbD_SYUBETSU06
'
'            Combo1(i).ListIndex = -1
'
'    Next i
'    Combo1(pcmbSHIMUKE).ListIndex = 0


    For i = pcmbUKEHARAI To pcmbD_SYUBETSU06

            Combo1(i).ListIndex = -1

    Next i
'    Combo1(pcmbUKEHARAI).ListIndex = 0
'2009.03.25



    '���s��
    Text1(ptxHAKKO_DT).text = Format(Now, "YYYY/MM/DD")


    '���F�Ґݒ�
    Text1(ptxSHONIN_CODE).text = StrConv(P_KANRIREC.SHONIN_CODE, vbUnicode)


    PI000104_OLD_HIN_GAI = ""       '2019.03.14

Start_Proc1:        '2015.03.26ok

    Call UniCode_Conv(K0_TANTO.TANTO_CODE, StrConv(P_KANRIREC.SHONIN_CODE, vbUnicode))

    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
            Text1(ptxSHONIN_NAME).text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Text1(ptxSHONIN_NAME).text = ""

        Case Else
            
            
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        If sts > 3000 Or sts = 3 Then
    
                            
                            Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc1
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
            
            
            Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
            Exit Function
    End Select
    '��z��
    Text1(ptxUKEHARAI_CODE).text = TEHAI
    txGensankoku.text = ""                  '2009.03.28



    lblGensankoku(0).Caption = ""
    lblGensankoku(1).Caption = ""

    '�w���`��
    Option1(poptSHIJI_NORMAL).Value = True
    Option1(poptSHIJI_SPOT).Value = False
    Option1(poptSHIJI_KEPPIN).Value = False


    '2011.02.10
    Combo2(0).ListIndex = 1
    '2011.02.10

    If LABEL_PRINT_F = 1 Then                       '2019.03.07
        Combo2(0).ListIndex = 0                     '2019.03.07
    End If                                          '2019.03.07



    '>>>>>>>>>>>>>>>>>>>>>> 2013.09.12
    For i = 0 To UBound(SHIMUKE_CHK_TBL)
    
        If SHIMUKE_CHK_TBL(i) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Then
            Combo2(0).ListIndex = 0
            Check1(pchkPRI_GAISOU).Value = vbUnchecked
            Exit For
        End If
    
    Next i
    '>>>>>>>>>>>>>>>>>>>>>> 2013.09.12

    
    If GA_LABEL_PRINT_F = 1 Then                    '2019.03.07
        Check1(pchkPRI_GAISOU).Value = vbUnchecked  '2019.03.07
    End If                                          '2019.03.07
    '2019.09.24 ���L�t���O���N���A�iFalse�j
    '                                       ����́ACombo2(0)�̉E�[<>"" �������͊O�����x���̎w����True�ɂ��Ă���B
    L_print_Flg = False


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

    Call UniCode_Conv(P_SSHIJI_O_REC.SHIJI_No, "")




    Text1(ptxHIN_GAI).Locked = False            '2019.03.18
'            Text1(ptxHIN_GAI).Enabled = True           '2019.03.18
    Text1(ptxHIN_GAI).BackColor = &H80000005     '2019.03.18



    Init_Proc = False

End Function


Private Function Init_Proc_2() As Integer
'----------------------------------------------------------------------------
'                   ���͉�ʂ̏����ݒ�
'----------------------------------------------------------------------------
Dim i           As Integer
Dim sts         As Integer

Dim TANTO_CODE  As String
Dim TANTO_NAME  As String

    Init_Proc_2 = True

    Text1(ptxSHIJI_NO).BackColor = G_INPUT_OK
    Text1(ptxSHIJI_NO).Locked = False
    Text1(ptxSHIJI_NO).TabStop = True


    Combo1(pcmbS_TANTO).Enabled = PRI_S_TANTO

    TANTO_CODE = Text1(ptxTANTO_CODE).text
    TANTO_NAME = Text1(ptxTANTO_NAME).text

    For i = ptxSHIJI_NO To ptxLabel_QTY
        Text1(i).text = ""
    Next i
    Text1(ptxTANTO_CODE).text = TANTO_CODE
    Text1(ptxTANTO_NAME).text = TANTO_NAME



    RichTextBox1(prchBIKOU).text = ""

    For i = pchkSAMPLE_F To pchkPRI_KISHU
        Check1(i).Value = vbChecked
    Next i
    Check1(pchkSAMPLE_F).Value = vbUnchecked    '���{�쐬
    Check1(pchkPRI_KISHU).Value = vbUnchecked   '�o�͑Ώہ@�@������

    Check1(pchkL_PAPER).Value = vbUnchecked     '��           2010.11.12
    Check1(pchkL_PLASTIC).Value = vbUnchecked   '��׽���      2010.11.12
    Check1(pchkL_LABEL).Value = vbUnchecked     '�K�p�@������ 2010.11.12


'2009.03.25
'    For i = pcmbSHIMUKE To pcmbD_SYUBETSU06
'
'            Combo1(i).ListIndex = -1
'
'    Next i
'    Combo1(pcmbSHIMUKE).ListIndex = 0


    For i = pcmbUKEHARAI To pcmbD_SYUBETSU06

            Combo1(i).ListIndex = -1

    Next i
'    Combo1(pcmbUKEHARAI).ListIndex = 0
'2009.03.25



    '���s��
    Text1(ptxHAKKO_DT).text = Format(Now, "YYYY/MM/DD")


    '���F�Ґݒ�
    Text1(ptxSHONIN_CODE).text = StrConv(P_KANRIREC.SHONIN_CODE, vbUnicode)


    PI000104_OLD_HIN_GAI = ""       '2019.03.14

Start_Proc1:        '2015.03.26ok

    Call UniCode_Conv(K0_TANTO.TANTO_CODE, StrConv(P_KANRIREC.SHONIN_CODE, vbUnicode))

    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
            Text1(ptxSHONIN_NAME).text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Text1(ptxSHONIN_NAME).text = ""

        Case Else
            
            
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        If sts > 3000 Or sts = 3 Then
    
                            
                            Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc1
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
            
            
            Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
            Exit Function
    End Select
    '��z��
    Text1(ptxUKEHARAI_CODE).text = TEHAI
    txGensankoku.text = ""                  '2009.03.28



    lblGensankoku(0).Caption = ""
    lblGensankoku(1).Caption = ""
    
    '2019.06.10 �uInit_Proc�v�Ƃ̈Ⴂ�́A���L�̂R�s���L�����ۂ��I�̂݁A
'    '�w���`��
'    Option1(poptSHIJI_NORMAL).Value = True
'    Option1(poptSHIJI_SPOT).Value = False
'    Option1(poptSHIJI_KEPPIN).Value = False

    '2019.08.23 ���L���R�����g�ɂ����B
    '           �� �d����ɂ���āu���x���Ȃ��v�ƃZ�b�g���Ă����Ă��AClear����Ă��܂��I
    ''2019.08.28 ���L�𕜋A���Ă݂��B
    '                   2019.09.24 �ēx�A�R�����g�ɂ����I
    '2011.02.10
'    Combo2(0).ListIndex = 1
    '2011.02.10

    If LABEL_PRINT_F = 1 Then                       '2019.03.07
        Combo2(0).ListIndex = 0                     '2019.03.07
    End If                                          '2019.03.07
    ''2019.08.23 �����܂�
    ''2019.08.28 �����܂�

    '>>>>>>>>>>>>>>>>>>>>>> 2013.09.12
    For i = 0 To UBound(SHIMUKE_CHK_TBL)
    
        If SHIMUKE_CHK_TBL(i) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Then
            Combo2(0).ListIndex = 0
            Check1(pchkPRI_GAISOU).Value = vbUnchecked
            Exit For
        End If
    
    Next i
    '>>>>>>>>>>>>>>>>>>>>>> 2013.09.12

    '2019.08.23 ���L���R�����g�ɂ����B
    '           �� �d����ɂ���āu���x���Ȃ��v�ƃZ�b�g���Ă����Ă��ACkear����Ă��܂��I
'    If GA_LABEL_PRINT_F = 1 Then                    '2019.03.07
'        Check1(pchkPRI_GAISOU).Value = vbUnchecked  '2019.03.07
'    End If                                          '2019.03.07

    '2019.09.24 ��L�𕜋A
    If GA_LABEL_PRINT_F = 1 Then                    '2019.03.07
        Check1(pchkPRI_GAISOU).Value = vbUnchecked  '2019.03.07
    End If                                          '2019.03.07
    
    '2019.09.24 ���L�t���O���N���A�iFalse�j
    '                                       ����́ACombo2(0)�̉E�[<>"" �������͊O�����x���̎w����True�ɂ��Ă���B
    L_print_Flg = False
    
    
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

    Call UniCode_Conv(P_SSHIJI_O_REC.SHIJI_No, "")




    Text1(ptxHIN_GAI).Locked = False            '2019.03.18
'            Text1(ptxHIN_GAI).Enabled = True           '2019.03.18
    Text1(ptxHIN_GAI).BackColor = &H80000005     '2019.03.18



    Init_Proc_2 = False

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

Start_Proc0:        '2015.03.26ok

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
                
                
                
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "�R�[�h�}�X�^", 0)
                            
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
                        
                            GoTo Start_Proc0
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                
                
                
                
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

Start_Proc1:            '2015.03.26ok

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
                
                
                
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                If sts > 3000 Or sts = 3 Then


                    Call File_Error(sts, BtOpGetEqual, "�󕥐�}�X�^", 0)
                    
                    
                    
                    '>>>>>>>>>>>>>  2015.04.24
                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                    'If sts Then
                    '    Call File_Error(sts, BtOpReset, "")
                    'End If
                    'Call File_Open_Proc
                    Do
                        If Not File_Open_Proc() Then
                            Exit Do
                        End If
                    Loop
                    '>>>>>>>>>>>>>  2015.04.24
    
                
                    GoTo Start_Proc1
                End If
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                
                
                
                
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
    chenge_F = False

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

Dim wkHin_Gai   As String * 20          '2019.03.14


    P_COMPO_Disp_Proc = True
    Call Input_Lock             '2008.01.15

Start_Proc1:    '2013.05.26ok


    For i = ptxK_HIN_GAI01 To ptxD_BIKOU06
        Text1(i).text = ""
    Next i

                                '2008.07.30
    For i = pcmbD_SYUBETSU01 To pcmbD_SYUBETSU06

            Combo1(i).ListIndex = -1

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

    Text1(ptxLabel_QTY).text = ""       '2008.02.27



    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
        
    
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).text)

    '2019.03.14
    If Trim(PI000104_OLD_HIN_GAI) <> "" Then
        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, PI000104_OLD_HIN_GAI)
    End If
    '2019.03.14


    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_HEAD)
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")

    sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)

    Select Case sts
        Case BtNoErr

        Case Else


                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "�\���}�X�^", 0)
                                                        
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc1
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26


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


            '2012.03.21
            RichTextBox1(prchBIKOU) = ""


            Call Input_UnLock           '2008.01.15
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

                '2019.03.14
                wkHin_Gai = Text1(ptxHIN_GAI).text
                If Trim(PI000104_OLD_HIN_GAI) <> "" Then
                    wkHin_Gai = PI000104_OLD_HIN_GAI
                End If
                '2019.03.14

                If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Or _
                    StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                    StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Or _
                    Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(wkHin_Gai) Then

                    Exit Do

                End If

            Case BtErrEOF
                Exit Do
            Case Else
                
                
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "�\���}�X�^", 0)
                            
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc1
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                
                Call Input_UnLock             '2008.01.15
                Call File_Error(sts, BtOpGetNext, "�\���}�X�^")
                Exit Function


        End Select

        Select Case StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode)

            Case P_KOSOU    '������

                k = k + 1

                If k > 36 Then
                    MsgBox "�����ޓo�^�������I�[�o�[���Ă��܂��B�폜���Ă�������"
                Else
                    K_Item_Tbl(k).JGYOBU = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode)
                    K_Item_Tbl(k).NAIGAI = StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode)
                                '�i��
                    Text1(K_Index).text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)

                    Call UniCode_Conv(K0_ITEM.JGYOBU, K_Item_Tbl(k).JGYOBU)
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
                                    
                                    
                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                                    If sts > 3000 Or sts = 3 Then
                
                    
                                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                        '>>>>>>>>>>>>>  2015.04.24
                                        'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                        'If sts Then
                                        '    Call File_Error(sts, BtOpReset, "")
                                        'End If
                                        'Call File_Open_Proc
                                        Do
                                            If Not File_Open_Proc() Then
                                                Exit Do
                                            End If
                                        Loop
                                        '>>>>>>>>>>>>>  2015.04.24
                        
                                    
                                        GoTo Start_Proc1
                                    End If
                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                                    
                                    
                                    Call Input_UnLock             '2008.01.15
                                    Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                                    Exit Function
    
                            End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>> �i�Ԗ��o�^�\���̑Ή�    2012.12.21
                        Case Else
                            
                            
                            
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            If sts > 3000 Or sts = 3 Then
        
            
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                '>>>>>>>>>>>>>  2015.04.24
                                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                'If sts Then
                                '    Call File_Error(sts, BtOpReset, "")
                                'End If
                                'Call File_Open_Proc
                                Do
                                    If Not File_Open_Proc() Then
                                        Exit Do
                                    End If
                                Loop
                                '>>>>>>>>>>>>>  2015.04.24
                
                            
                                GoTo Start_Proc1
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            
                            
                            Call Input_UnLock             '2008.01.15
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
                End If


            Case P_GAISOU   '�O������
                g = g + 1


                If g > 51 Then
                    MsgBox "�O�����ޓo�^�������I�[�o�[���Ă��܂��B�폜���Ă�������"
                Else

                    G_Item_Tbl(g).JGYOBU = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode)
                    G_Item_Tbl(g).NAIGAI = StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode)
                                '�i��
                    Text1(G_Index).text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)

                    Call UniCode_Conv(K0_ITEM.JGYOBU, G_Item_Tbl(g).JGYOBU)
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
                                    
                                    
                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                                    If sts > 3000 Or sts = 3 Then
                
                    
                                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                        '>>>>>>>>>>>>>  2015.04.24
                                        'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                        'If sts Then
                                        '    Call File_Error(sts, BtOpReset, "")
                                        'End If
                                       'Call File_Open_Proc
                                        Do
                                            If Not File_Open_Proc() Then
                                                Exit Do
                                            End If
                                        Loop
                                        '>>>>>>>>>>>>>  2015.04.24
                        
                                    
                                        GoTo Start_Proc1
                                    End If
                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                                    
                                    
                                    
                                    Call Input_UnLock             '2008.01.15
                                    Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                                    Exit Function
    
                            End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>> �i�Ԗ��o�^�\���̑Ή�    2012.12.21
                            
                        
                        Case Else
                            
                            
                            
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            If sts > 3000 Or sts = 3 Then
        
            
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                '>>>>>>>>>>>>>  2015.04.24
                                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                'If sts Then
                                '    Call File_Error(sts, BtOpReset, "")
                                'End If
                                'Call File_Open_Proc
                                Do
                                    If Not File_Open_Proc() Then
                                        Exit Do
                                    End If
                                Loop
                                '>>>>>>>>>>>>>  2015.04.24
                
                            
                                GoTo Start_Proc1
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            
                            
                            
                            Call Input_UnLock             '2008.01.15
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
                End If

            Case P_DOUKON   '�����^�\��

                d = d + 1
                D_Item_Tbl(d).SYUBETSU = StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode)
                D_Item_Tbl(d).JGYOBU = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode)
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

                    DC_Index = DC_Index + 1

                                '�i��
                    Text1(DT_Index).text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)


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


                            If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                                Exit Function

                            End If

                            Text1(DT_Index + 5) = Format(Sumi_Qty + Mi_Qty, "#0")


                        Case BtErrKeyNotFound


'>>>>>>>>>>>>>>>>>>>>>>>>>> �i�Ԗ��o�^�\���̑Ή�    2012.12.21
                            
                            
                            Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
    
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
        
        
                                    If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                                            StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                                        Exit Function
        
                                    End If
        
                                    Text1(DT_Index + 5) = Format(Sumi_Qty + Mi_Qty, "#0")
                                
                                Case BtErrKeyNotFound
    

                                    Text1(DT_Index + 1) = "���o�^�i��"
                                    Text1(DT_Index + 4) = ""
                                    Text1(DT_Index + 5) = ""
                                Case Else
                                    
                                    
                                    
                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                                    If sts > 3000 Or sts = 3 Then
                
                    
                                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                        '>>>>>>>>>>>>>  2015.04.24
                                        'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                        'If sts Then
                                        '    Call File_Error(sts, BtOpReset, "")
                                        'End If
                                        'Call File_Open_Proc
                                        Do
                                            If Not File_Open_Proc() Then
                                                Exit Do
                                            End If
                                        Loop
                                        '>>>>>>>>>>>>>  2015.04.24
                        
                                    
                                        GoTo Start_Proc1
                                    End If
                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                                    
                                    
                                    
                                    
                                    
                                    
                                    Call Input_UnLock             '2008.01.15
                                    Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                                    Exit Function
    
                            End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>> �i�Ԗ��o�^�\���̑Ή�    2012.12.21








                        Case Else
                            
                            
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            If sts > 3000 Or sts = 3 Then
        
            
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                '>>>>>>>>>>>>>  2015.04.24
                                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                'If sts Then
                                '    Call File_Error(sts, BtOpReset, "")
                                'End If
                                'Call File_Open_Proc
                                Do
                                    If Not File_Open_Proc() Then
                                        Exit Do
                                    End If
                                Loop
                                '>>>>>>>>>>>>>  2015.04.24
                
                            
                                GoTo Start_Proc1
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            
                            
                            
                            
                            
                            
                            Call Input_UnLock             '2008.01.15
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

    Call Input_UnLock             '2008.01.15


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
Start_Proc1:        '2015.03.26
            Call UniCode_Conv(K0_ITEM.JGYOBU, D_Item_Tbl(i).JGYOBU)
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


                    If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                            StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                        Exit Function
                    End If
                    '�݌ɐ�
                    Text1(DT_Index + 5) = Format(Sumi_Qty + Mi_Qty, "#0")

                Case BtErrKeyNotFound

                    Text1(DT_Index + 1) = "���o�^�i��"
                    Text1(DT_Index + 4) = ""
                    Text1(DT_Index + 5) = ""

                Case Else
                    
                    
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                    If sts > 3000 Or sts = 3 Then

    
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                        '>>>>>>>>>>>>>  2015.04.24
                        'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                        'If sts Then
                        '    Call File_Error(sts, BtOpReset, "")
                        'End If
                        'Call File_Open_Proc
                        Do
                            If Not File_Open_Proc() Then
                                Exit Do
                            End If
                        Loop
                        '>>>>>>>>>>>>>  2015.04.24
        
                    
                        GoTo Start_Proc1
                    End If
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                    
                    
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

    If Den_No_Set_Proc(21, Last_JGYOBU, ID_NO) Then                         'IDNO
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
    If Den_No_Set_Proc(20, Last_JGYOBU, DEN_NO) Then
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
                sts = Den_No_Set_Proc(21, Last_JGYOBU, ID_NO)
                If sts Then
                    Exit Function
                End If

                Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, ID_NO)
                Call UniCode_Conv(Y_SYUREC.ID_NO, ID_NO)

Debug.Print StrConv(Y_SYUREC.KEY_ID_NO, vbUnicode)
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

'Private Sub Text1_LostFocus(Index As Integer)
'
'
'    Select Case Index
'        Case ptxHIN_GAI, ptxK_HIN_GAI01, ptxK_HIN_GAI02, ptxK_HIN_GAI03, ptxK_HIN_GAI04, ptxK_HIN_GAI05, _
'                ptxG_HIN_GAI01, ptxG_HIN_GAI02, ptxG_HIN_GAI03, _
'                ptxD_HIN_GAI01, ptxD_HIN_GAI02, ptxD_HIN_GAI03, ptxD_HIN_GAI04, ptxD_HIN_GAI05, ptxD_HIN_GAI06 _
'
'
'            Text1(Index).text = RTrim(StrConv(Text1(Index).text, vbUpperCase))
'
'    End Select
'
'End Sub

Private Function Mesg_Set_Proc(GEN_NG_F As Integer, GEN_AT_GAI_F As Integer, GEN_AT_PLU_F As Integer, TANKA_SP_F As Integer, KISHU_NG_F As Integer, KAISYA_NG_F As Integer, KISHU1 As String, KISHU2 As String) As Integer
'----------------------------------------------------------------------------
'               �G���[���b�Z�[�W�쐬
'           2016.01.29
'----------------------------------------------------------------------------
Dim Mesg        As String
Dim i           As Integer
    
    
Dim GENSANKOKU  As String * 20
    
Dim KAISYA_NAME As String * 20
Dim JGYOBU_NAME As String * 20
    
    
Dim TANKA       As String * 20
    
Dim KISHU       As String * 20
    
Dim Tanka2      As String * 9
Dim Tanka3      As String * 9
    
        Mesg = "�p�[�c���x���̓��e���m�F���ĉ�����" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
        Mesg = Mesg & "���i��   " & Text1(ptxHIN_GAI).text & Chr(13) & Chr(10)
        Mesg = Mesg & "���i��   " & Text1(ptxHIN_NAME).text & Chr(13) & Chr(10)
        Mesg = Mesg & "���i���d " & lblL_Hin_Name_E.Caption & Chr(13) & Chr(10)

    
        GENSANKOKU = lblGensankoku(1)




        Select Case GEN_NG_F
            Case 0
                If GEN_AT_PLU_F < 2 Then
                    If GEN_AT_GAI_F = 0 Then
                        Mesg = Mesg & "�����Y�� " & GENSANKOKU & Chr(13) & Chr(10)
                    Else
                        Mesg = Mesg & "�����Y�� " & GENSANKOKU & "�@�����Y�����Ӂi�C�O�����j" & Chr(13) & Chr(10)
                
                    End If
                Else
                        Mesg = Mesg & "�����Y�� "
                        For i = lstGensankoku.ListCount - 1 To 0 Step -1
                            GENSANKOKU = Right(lstGensankoku.List(i), 20)
                            
                            If i = lstGensankoku.ListCount - 1 Then
                                Mesg = Mesg & GENSANKOKU
                            Else
                                Mesg = Mesg & "�@�@�@�@�@   " & GENSANKOKU
                            End If
                            If i = 0 Then
                                Mesg = Mesg & "�@�����Y�����Ӂi�����j" & Chr(13) & Chr(10)
                            Else
                                Mesg = Mesg & Chr(13) & Chr(10)
                            End If
                        Next i
                End If
            
            Case 1
                Mesg = Mesg & "�~���Y�� " & GENSANKOKU & "�@���󔒂ł�" & Chr(13) & Chr(10)
            Case 9
                Mesg = Mesg & "�����Y�� " & Chr(13) & Chr(10)
        End Select
    
    
        Mesg = Mesg & "�������敪�C�O " & "     " & lblGAI_BUHIN.Caption & Chr(13) & Chr(10)
    
        If IsNumeric(lblL_URIKIN2) Then
            Tanka2 = Format(CDbl(lblL_URIKIN2), "#0")
        Else
            Tanka2 = ""
        End If
        If IsNumeric(lblL_URIKIN3) Then
            Tanka3 = Format(CDbl(lblL_URIKIN3), "#0")
        Else
            Tanka3 = ""
        End If
            
        
        
        TANKA = Tanka2 & " " & Tanka3
        If TANKA_SP_F = 1 Then
            Mesg = Mesg & "�~�P���@ " & TANKA & "   �@  ���󔒂ł�" & Chr(13) & Chr(10)
        Else
            Mesg = Mesg & "���P���@ " & TANKA & Chr(13) & Chr(10)
        End If
    
        KISHU = KISHU1
        If KISHU_NG_F = 1 Then
            Mesg = Mesg & "�~��\�@��@ " & KISHU & " ���󔒂ł�" & Chr(13) & Chr(10)
        Else
            Mesg = Mesg & "����\�@��@ " & KISHU & Chr(13) & Chr(10)
        End If
    
        If KAISYA_NG_F = 9 Then
        Else
            KAISYA_NAME = lblL_KAISHA_N.Caption
            JGYOBU_NAME = lblL_JGYOBU_N.Caption
            If KAISYA_NG_F = 1 Then
                If Trim(KAISYA_NAME) = "" Then
                    Mesg = Mesg & "�~��Ж� " & KAISYA_NAME & " " & "�@�@ ���󔒂ł�" & Chr(13) & Chr(10)
                Else
                    Mesg = Mesg & "����Ж� " & KAISYA_NAME & Chr(13) & Chr(10)
                End If
                If Trim(JGYOBU_NAME) = "" Then
                    Mesg = Mesg & "�~���ƕ��� " & JGYOBU_NAME & " " & "�@���󔒂ł�" & Chr(13) & Chr(10)
                Else
                    Mesg = Mesg & "�����ƕ��� " & JGYOBU_NAME & Chr(13) & Chr(10)
                End If
            Else
                    Mesg = Mesg & "����Ж� " & KAISYA_NAME & Chr(13) & Chr(10)
                    Mesg = Mesg & "�����ƕ��� " & JGYOBU_NAME & Chr(13) & Chr(10)
            End If
        End If
    
    
        Mesg = Mesg & Chr(13) & Chr(10)
    
    
'        Mesg = Mesg & "�@�@�@�@�y�n�j�z�p�[�c���x�������" & Chr(13) & Chr(10)     '2016.02.10
        Mesg = Mesg & "�@�@�@�@�y�n�j�z����^�X�V" & Chr(13) & Chr(10)              '2016.02.10
        Mesg = Mesg & " �y�L�����Z���z������~" & Chr(13) & Chr(10)
    
    
    
'        Mesg_Set_Proc = MsgBox(Mesg, vbOKCancel + vbDefaultButton2 + vbExclamation, "�p�[�c���x�����ڊm�F")    '2016.02.10
        Mesg_Set_Proc = MsgBox(Mesg, vbOKCancel + vbDefaultButton1 + vbExclamation, "�p�[�c���x�����ڊm�F")     '2016.02.10


End Function


