VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form SEI00181 
   Caption         =   "[�����V�X�e��]���Ϗ��ꊇ���s����"
   ClientHeight    =   13080
   ClientLeft      =   2025
   ClientTop       =   -3510
   ClientWidth     =   15945
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
   LockControls    =   -1  'True
   ScaleHeight     =   13080
   ScaleWidth      =   15945
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Frame Frame1 
      Height          =   12975
      Left            =   0
      TabIndex        =   222
      Top             =   -120
      Width           =   15855
      Begin VB.TextBox txtBUHIN 
         Appearance      =   0  '�ׯ�
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3480
         MaxLength       =   1
         TabIndex        =   255
         Top             =   1080
         Width           =   210
      End
      Begin VB.ComboBox cmbSHIMUKE 
         Appearance      =   0  '�ׯ�
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3495
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   238
         Top             =   600
         Width           =   2100
      End
      Begin VB.TextBox txtTanto_Name 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '�Ȃ�
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   237
         TabStop         =   0   'False
         Top             =   240
         Width           =   2325
      End
      Begin VB.TextBox txtTANTO_CODE 
         Appearance      =   0  '�ׯ�
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3480
         MaxLength       =   5
         TabIndex        =   236
         Top             =   240
         Width           =   750
      End
      Begin VB.TextBox Text2 
         Height          =   4335
         Left            =   2760
         MultiLine       =   -1  'True
         TabIndex        =   235
         ToolTipText     =   "�i�Ԃ��R�s�[���ĉ�����"
         Top             =   2880
         Width           =   5412
      End
      Begin VB.CommandButton Command2 
         Caption         =   "�ر"
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
         Left            =   3060
         TabIndex        =   234
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtNG_CNT 
         Alignment       =   1  '�E����
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   13800
         Locked          =   -1  'True
         TabIndex        =   233
         Top             =   12240
         Width           =   855
      End
      Begin VB.TextBox txtOK_CNT 
         Alignment       =   1  '�E����
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   13800
         Locked          =   -1  'True
         TabIndex        =   232
         Top             =   11760
         Width           =   855
      End
      Begin VB.TextBox txtIN_CNT 
         Alignment       =   1  '�E����
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   7260
         Locked          =   -1  'True
         TabIndex        =   231
         Top             =   12000
         Width           =   855
      End
      Begin VB.ListBox List2 
         Height          =   4380
         ItemData        =   "SEI00181.frx":0000
         Left            =   8550
         List            =   "SEI00181.frx":0002
         TabIndex        =   230
         Top             =   2880
         Width           =   6225
      End
      Begin VB.CommandButton Command2 
         Caption         =   "�I��"
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
         Left            =   13695
         TabIndex        =   229
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "���s"
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
         Left            =   12360
         TabIndex        =   228
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtKIN_NG_CNT 
         Alignment       =   1  '�E����
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   13800
         Locked          =   -1  'True
         TabIndex        =   227
         Top             =   12840
         Width           =   855
      End
      Begin VB.CommandButton Command2 
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
         Index           =   2
         Left            =   9240
         TabIndex        =   226
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtOUT_CNT 
         Alignment       =   1  '�E����
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   10260
         Locked          =   -1  'True
         TabIndex        =   225
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   4215
         Left            =   2760
         MultiLine       =   -1  'True
         TabIndex        =   224
         ToolTipText     =   "�i�Ԃ��R�s�[���ĉ�����"
         Top             =   7560
         Width           =   5412
      End
      Begin VB.ListBox List3 
         Height          =   4140
         ItemData        =   "SEI00181.frx":0004
         Left            =   8520
         List            =   "SEI00181.frx":0006
         TabIndex        =   223
         Top             =   7560
         Width           =   6225
      End
      Begin VB.Label Label3 
         Caption         =   "0:��Ώ�/1:�Ώ�/2:�Ő؈ē���/3:�Ő�/�󔒁F�S��"
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   256
         Top             =   1080
         Width           =   5535
      End
      Begin VB.Label Label3 
         Caption         =   "�����敪"
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   254
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Height          =   240
         Index           =   9
         Left            =   2640
         TabIndex        =   253
         Top             =   1080
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�d����"
         Height          =   240
         Index           =   110
         Left            =   2640
         TabIndex        =   252
         Top             =   660
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�S����"
         Height          =   240
         Index           =   111
         Left            =   2640
         TabIndex        =   251
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "�m�f����"
         Height          =   255
         Index           =   112
         Left            =   12855
         TabIndex        =   250
         Top             =   12360
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "�n�j����"
         Height          =   255
         Index           =   113
         Left            =   12855
         TabIndex        =   249
         Top             =   11880
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "�Ǎ��݌���"
         Height          =   255
         Index           =   114
         Left            =   6000
         TabIndex        =   248
         Top             =   12120
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "�X�V����"
         Height          =   255
         Index           =   115
         Left            =   13440
         TabIndex        =   247
         Top             =   2640
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "���z��ϯ�����"
         Height          =   255
         Index           =   116
         Left            =   12255
         TabIndex        =   246
         Top             =   12960
         Width           =   1665
      End
      Begin VB.Label Label1 
         Caption         =   "�q�i��"
         Height          =   255
         Index           =   117
         Left            =   8640
         TabIndex        =   245
         Top             =   2640
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "�e�i��"
         Height          =   255
         Index           =   118
         Left            =   11160
         TabIndex        =   244
         Top             =   2640
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "�q�i��"
         Height          =   255
         Index           =   119
         Left            =   2880
         TabIndex        =   243
         Top             =   2640
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "�o�͌���"
         Height          =   255
         Index           =   120
         Left            =   10200
         TabIndex        =   242
         Top             =   1800
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "�e�i��"
         Height          =   255
         Index           =   121
         Left            =   2880
         TabIndex        =   241
         Top             =   7320
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "�e�i��"
         Height          =   255
         Index           =   122
         Left            =   8640
         TabIndex        =   240
         Top             =   7320
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "�X�V����"
         Height          =   255
         Index           =   123
         Left            =   10920
         TabIndex        =   239
         Top             =   7320
         Width           =   960
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�P���Эڰ���"
      Enabled         =   0   'False
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
      Index           =   6
      Left            =   10560
      TabIndex        =   202
      ToolTipText     =   "���i���P�����v�Z���܂�(F9)"
      Top             =   0
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   2520
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   5
      Top             =   1200
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   72
      Left            =   1440
      MaxLength       =   8
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   25
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   15
      Left            =   9840
      TabIndex        =   29
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   20
      Left            =   2400
      MaxLength       =   5
      TabIndex        =   34
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   21
      Left            =   3120
      TabIndex        =   35
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   3120
      TabIndex        =   22
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   2400
      MaxLength       =   5
      TabIndex        =   21
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   71
      Left            =   3360
      TabIndex        =   89
      Top             =   9480
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   70
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   9480
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   69
      Left            =   3360
      TabIndex        =   87
      Top             =   9150
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   68
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   9150
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   57
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   7170
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   56
      Left            =   2520
      TabIndex        =   74
      Top             =   7170
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   67
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   8820
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   66
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   8820
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   65
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   8490
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   64
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   8490
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   63
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   8160
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   62
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   8160
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   61
      Left            =   3360
      TabIndex        =   79
      Top             =   7830
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   60
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   7830
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   59
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   7500
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   58
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   7500
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   110
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   114
      TabStop         =   0   'False
      Top             =   9480
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   111
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   115
      TabStop         =   0   'False
      Top             =   9480
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   170
      Left            =   14370
      MaxLength       =   1
      TabIndex        =   123
      Top             =   9960
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   168
      Left            =   11250
      MaxLength       =   1
      TabIndex        =   121
      Top             =   9960
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   167
      Left            =   11250
      MaxLength       =   10
      TabIndex        =   120
      Top             =   9570
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   166
      Left            =   6960
      TabIndex        =   39
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   165
      Left            =   6960
      TabIndex        =   26
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   164
      Left            =   8880
      TabIndex        =   41
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   162
      Left            =   8880
      TabIndex        =   28
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   163
      Left            =   7920
      TabIndex        =   40
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '�Ȃ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   159
      Left            =   12840
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '�Ȃ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   158
      Left            =   12480
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   157
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   126
      TabStop         =   0   'False
      Top             =   7920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   156
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   125
      TabStop         =   0   'False
      Top             =   7320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   155
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   124
      TabStop         =   0   'False
      Top             =   6960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   113
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   117
      TabStop         =   0   'False
      Top             =   9810
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   112
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   116
      TabStop         =   0   'False
      Top             =   9480
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   109
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   113
      TabStop         =   0   'False
      Top             =   9150
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   108
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   112
      TabStop         =   0   'False
      Top             =   9150
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   107
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   111
      TabStop         =   0   'False
      Top             =   9150
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   106
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   110
      TabStop         =   0   'False
      Top             =   8820
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   105
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   109
      TabStop         =   0   'False
      Top             =   8820
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   104
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   108
      TabStop         =   0   'False
      Top             =   8820
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   103
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   107
      TabStop         =   0   'False
      Top             =   8490
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   101
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   8490
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   102
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   106
      TabStop         =   0   'False
      Top             =   8490
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   100
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   104
      TabStop         =   0   'False
      Top             =   8160
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   99
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   103
      TabStop         =   0   'False
      Top             =   8160
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   98
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   102
      TabStop         =   0   'False
      Top             =   8160
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   97
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   101
      TabStop         =   0   'False
      Top             =   7830
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   96
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   100
      TabStop         =   0   'False
      Top             =   7830
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   95
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   7830
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   94
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   98
      TabStop         =   0   'False
      Top             =   7500
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   93
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   97
      TabStop         =   0   'False
      Top             =   7500
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   92
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   7500
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   91
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   7170
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   90
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   94
      TabStop         =   0   'False
      Top             =   7170
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   89
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   7170
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   86
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   6840
      Width           =   855
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1575
      Index           =   0
      Left            =   10680
      TabIndex        =   118
      Top             =   6750
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2778
      _Version        =   393217
      Appearance      =   0
      TextRTF         =   $"SEI00181.frx":0008
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   55
      Left            =   3360
      TabIndex        =   73
      Top             =   6840
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   54
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   6840
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   2
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '�Ȃ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   840
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '�Ȃ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '�Ȃ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '�Ȃ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '�Ȃ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   53
      Left            =   14040
      Locked          =   -1  'True
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   52
      Left            =   12960
      Locked          =   -1  'True
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   51
      Left            =   11880
      Locked          =   -1  'True
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   50
      Left            =   10800
      Locked          =   -1  'True
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   49
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   48
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   47
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   46
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   45
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   44
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   43
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   42
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   40
      Left            =   14040
      Locked          =   -1  'True
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   39
      Left            =   12960
      Locked          =   -1  'True
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   38
      Left            =   11880
      Locked          =   -1  'True
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   37
      Left            =   10800
      Locked          =   -1  'True
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   36
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   35
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   34
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   33
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   32
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   31
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   30
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   29
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   27
      Left            =   11520
      TabIndex        =   44
      Top             =   3000
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   26
      Left            =   10800
      TabIndex        =   43
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   24
      Left            =   6000
      TabIndex        =   38
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   23
      Left            =   5040
      TabIndex        =   37
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   22
      Left            =   4080
      TabIndex        =   36
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   19
      Left            =   1440
      MaxLength       =   8
      TabIndex        =   33
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   18
      Left            =   480
      MaxLength       =   8
      TabIndex        =   32
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   17
      Left            =   11520
      TabIndex        =   31
      Top             =   2760
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   16
      Left            =   10800
      TabIndex        =   30
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   14
      Left            =   6000
      TabIndex        =   25
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   13
      Left            =   5040
      TabIndex        =   24
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   4080
      TabIndex        =   23
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   1440
      MaxLength       =   8
      TabIndex        =   20
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   480
      MaxLength       =   8
      TabIndex        =   19
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�P���X�V"
      Enabled         =   0   'False
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
      Index           =   5
      Left            =   8760
      TabIndex        =   133
      ToolTipText     =   "���i���P����i�ڃ}�X�^�[�ɓo�^���܂�"
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���Ϗ����s"
      Enabled         =   0   'False
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
      Index           =   4
      Left            =   7080
      TabIndex        =   132
      ToolTipText     =   "���i���P�����Ϗ�(EXCEL)���쐬���܂�"
      Top             =   0
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '�Ȃ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1440
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   6
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�P���v�Z"
      Enabled         =   0   'False
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
      Index           =   3
      Left            =   5280
      TabIndex        =   131
      ToolTipText     =   "���i���P�����v�Z���܂�(F9)"
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�ۑ�"
      Enabled         =   0   'False
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
      Index           =   2
      Left            =   3480
      TabIndex        =   130
      ToolTipText     =   "���i���\����ۑ����܂�"
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   11040
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   128
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�Ǎ�"
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
      Index           =   1
      Left            =   1800
      TabIndex        =   129
      ToolTipText     =   "���i���\����ǂݍ��݂܂��i�e5�j"
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
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
      Index           =   0
      Left            =   480
      TabIndex        =   127
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   41
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   28
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   87
      Left            =   7560
      TabIndex        =   91
      Top             =   6840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   88
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   6840
      Width           =   855
   End
   Begin TrueDBGrid80.TDBDropDown TDBDropDown1 
      Height          =   2055
      Left            =   5640
      TabIndex        =   176
      Top             =   4320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   3625
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).ValueItems(0)._DefaultItem=   0
      Columns(0).ValueItems(0).Value=   "aaaa"
      Columns(0).ValueItems(0).Value.vt=   8
      Columns(0).ValueItems(0).DisplayValue=   "aaaa"
      Columns(0).ValueItems(0).DisplayValue.vt=   8
      Columns(0).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(0).ValueItems.Count=   1
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   1
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   688
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=1"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4366"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4233"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits.Count    =   1
      AllowRowSizing  =   0   'False
      Appearance      =   1
      BorderStyle     =   1
      ColumnHeaders   =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      Enabled         =   -1  'True
      HeadLines       =   1
      RowDividerStyle =   2
      LayoutName      =   ""
      LayoutFileName  =   ""
      LayoutURL       =   ""
      EmptyRows       =   0   'False
      ListField       =   ""
      DataField       =   "ub_grid2"
      IntegralHeight  =   0   'False
      FetchRowStyle   =   0   'False
      AlternatingRowStyle=   0   'False
      ColumnFooters   =   0   'False
      FootLines       =   1
      DeadAreaBackColor=   14215660
      ValueTranslate  =   0   'False
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(5)   =   ":id=0,.fontname=�l�r �S�V�b�N"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=�l�r �S�V�b�N"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=�l�r �S�V�b�N"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=�l�r �S�V�b�N"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Named:id=33:Normal"
      _StyleDefs(41)  =   ":id=33,.parent=0"
      _StyleDefs(42)  =   "Named:id=34:Heading"
      _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(44)  =   ":id=34,.wraptext=-1"
      _StyleDefs(45)  =   "Named:id=35:Footing"
      _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(47)  =   "Named:id=36:Selected"
      _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(49)  =   "Named:id=37:Caption"
      _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(51)  =   "Named:id=38:HighlightRow"
      _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=39:EvenRow"
      _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(55)  =   "Named:id=40:OddRow"
      _StyleDefs(56)  =   ":id=40,.parent=33"
      _StyleDefs(57)  =   "Named:id=41:RecordSelector"
      _StyleDefs(58)  =   ":id=41,.parent=34"
      _StyleDefs(59)  =   "Named:id=42:FilterBar"
      _StyleDefs(60)  =   ":id=42,.parent=33"
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   2295
      Index           =   0
      Left            =   0
      TabIndex        =   71
      Top             =   4200
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   4048
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "���ƕ�"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "�����O"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   1
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "���"
      Columns(2).DataField=   ""
      Columns(2).DropDown=   "TDBDropDown1"
      Columns(2).DropDown.vt=   8
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "�i��"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "�i��"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "����"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "�d����"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "�̔���"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "���ʒP����"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "�d�����z�v"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "�̔����z�v"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "��Ǝ��ԁi�b�j"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "�W������i�b�j"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "���l"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "�̔����z�@���×p"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   15
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=15"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1217"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1111"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=1032"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=926"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(12)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(13)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(15)=   "Column(2).Width=2196"
      Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=2090"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(19)=   "Column(2).Button=1"
      Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(21)=   "Column(3).Width=1905"
      Splits(0)._ColumnProps(22)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._WidthInPix=1799"
      Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(26)=   "Column(4).Width=3757"
      Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=3651"
      Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=8708"
      Splits(0)._ColumnProps(30)=   "Column(4).AllowFocus=0"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(32)=   "Column(5).Width=1164"
      Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1058"
      Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=514"
      Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(37)=   "Column(6).Width=1879"
      Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=1773"
      Splits(0)._ColumnProps(40)=   "Column(6)._ColStyle=8706"
      Splits(0)._ColumnProps(41)=   "Column(6).AllowFocus=0"
      Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(43)=   "Column(7).Width=2143"
      Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=2037"
      Splits(0)._ColumnProps(46)=   "Column(7)._ColStyle=8706"
      Splits(0)._ColumnProps(47)=   "Column(7).AllowFocus=0"
      Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(49)=   "Column(8).Width=2143"
      Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=2037"
      Splits(0)._ColumnProps(52)=   "Column(8)._ColStyle=514"
      Splits(0)._ColumnProps(53)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(54)=   "Column(9).Width=2117"
      Splits(0)._ColumnProps(55)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(9)._WidthInPix=2011"
      Splits(0)._ColumnProps(57)=   "Column(9)._ColStyle=8706"
      Splits(0)._ColumnProps(58)=   "Column(9).AllowFocus=0"
      Splits(0)._ColumnProps(59)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(60)=   "Column(10).Width=2249"
      Splits(0)._ColumnProps(61)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(62)=   "Column(10)._WidthInPix=2143"
      Splits(0)._ColumnProps(63)=   "Column(10)._ColStyle=8706"
      Splits(0)._ColumnProps(64)=   "Column(10).AllowFocus=0"
      Splits(0)._ColumnProps(65)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(66)=   "Column(11).Width=2858"
      Splits(0)._ColumnProps(67)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(68)=   "Column(11)._WidthInPix=2752"
      Splits(0)._ColumnProps(69)=   "Column(11)._ColStyle=8706"
      Splits(0)._ColumnProps(70)=   "Column(11).AllowFocus=0"
      Splits(0)._ColumnProps(71)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(72)=   "Column(12).Width=3201"
      Splits(0)._ColumnProps(73)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(74)=   "Column(12)._WidthInPix=3096"
      Splits(0)._ColumnProps(75)=   "Column(12)._ColStyle=8706"
      Splits(0)._ColumnProps(76)=   "Column(12).AllowFocus=0"
      Splits(0)._ColumnProps(77)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(78)=   "Column(13).Width=3810"
      Splits(0)._ColumnProps(79)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(80)=   "Column(13)._WidthInPix=3704"
      Splits(0)._ColumnProps(81)=   "Column(13)._ColStyle=516"
      Splits(0)._ColumnProps(82)=   "Column(13).Visible=0"
      Splits(0)._ColumnProps(83)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(84)=   "Column(14).Width=3810"
      Splits(0)._ColumnProps(85)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(86)=   "Column(14)._WidthInPix=3704"
      Splits(0)._ColumnProps(87)=   "Column(14)._ColStyle=516"
      Splits(0)._ColumnProps(88)=   "Column(14).Visible=0"
      Splits(0)._ColumnProps(89)=   "Column(14).Order=15"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=�l�r �S�V�b�N"
      PrintInfos(0).PageFooterFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=�l�r �S�V�b�N"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      AllowAddNew     =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   14215660
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(5)   =   ":id=0,.fontname=�l�r �S�V�b�N"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=�l�r �S�V�b�N"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bold=0,.fontsize=975"
      _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=�l�r �S�V�b�N"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=�l�r �S�V�b�N"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=82,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=78,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=75,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=76,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=77,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.bgcolor=&H80000016&,.locked=-1"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=1"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=54,.parent=13,.alignment=1,.bgcolor=&H8000000F&"
      _StyleDefs(61)  =   ":id=54,.locked=-1"
      _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
      _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
      _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
      _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=58,.parent=13,.alignment=1,.bgcolor=&H8000000F&"
      _StyleDefs(66)  =   ":id=58,.locked=-1"
      _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=94,.parent=13,.alignment=1"
      _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=91,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=92,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=93,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=62,.parent=13,.alignment=1,.bgcolor=&H8000000F&"
      _StyleDefs(75)  =   ":id=62,.locked=-1"
      _StyleDefs(76)  =   "Splits(0).Columns(9).HeadingStyle:id=59,.parent=14"
      _StyleDefs(77)  =   "Splits(0).Columns(9).FooterStyle:id=60,.parent=15"
      _StyleDefs(78)  =   "Splits(0).Columns(9).EditorStyle:id=61,.parent=17"
      _StyleDefs(79)  =   "Splits(0).Columns(10).Style:id=66,.parent=13,.alignment=1,.bgcolor=&H8000000F&"
      _StyleDefs(80)  =   ":id=66,.locked=-1"
      _StyleDefs(81)  =   "Splits(0).Columns(10).HeadingStyle:id=63,.parent=14"
      _StyleDefs(82)  =   "Splits(0).Columns(10).FooterStyle:id=64,.parent=15"
      _StyleDefs(83)  =   "Splits(0).Columns(10).EditorStyle:id=65,.parent=17"
      _StyleDefs(84)  =   "Splits(0).Columns(11).Style:id=70,.parent=13,.alignment=1,.bgcolor=&H8000000F&"
      _StyleDefs(85)  =   ":id=70,.locked=-1"
      _StyleDefs(86)  =   "Splits(0).Columns(11).HeadingStyle:id=67,.parent=14"
      _StyleDefs(87)  =   "Splits(0).Columns(11).FooterStyle:id=68,.parent=15"
      _StyleDefs(88)  =   "Splits(0).Columns(11).EditorStyle:id=69,.parent=17"
      _StyleDefs(89)  =   "Splits(0).Columns(12).Style:id=74,.parent=13,.alignment=1,.bgcolor=&H8000000F&"
      _StyleDefs(90)  =   ":id=74,.locked=-1"
      _StyleDefs(91)  =   "Splits(0).Columns(12).HeadingStyle:id=71,.parent=14"
      _StyleDefs(92)  =   "Splits(0).Columns(12).FooterStyle:id=72,.parent=15"
      _StyleDefs(93)  =   "Splits(0).Columns(12).EditorStyle:id=73,.parent=17"
      _StyleDefs(94)  =   "Splits(0).Columns(13).Style:id=86,.parent=13"
      _StyleDefs(95)  =   "Splits(0).Columns(13).HeadingStyle:id=83,.parent=14"
      _StyleDefs(96)  =   "Splits(0).Columns(13).FooterStyle:id=84,.parent=15"
      _StyleDefs(97)  =   "Splits(0).Columns(13).EditorStyle:id=85,.parent=17"
      _StyleDefs(98)  =   "Splits(0).Columns(14).Style:id=90,.parent=13"
      _StyleDefs(99)  =   "Splits(0).Columns(14).HeadingStyle:id=87,.parent=14"
      _StyleDefs(100) =   "Splits(0).Columns(14).FooterStyle:id=88,.parent=15"
      _StyleDefs(101) =   "Splits(0).Columns(14).EditorStyle:id=89,.parent=17"
      _StyleDefs(102) =   "Named:id=33:Normal"
      _StyleDefs(103) =   ":id=33,.parent=0"
      _StyleDefs(104) =   "Named:id=34:Heading"
      _StyleDefs(105) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(106) =   ":id=34,.wraptext=-1"
      _StyleDefs(107) =   "Named:id=35:Footing"
      _StyleDefs(108) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(109) =   "Named:id=36:Selected"
      _StyleDefs(110) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(111) =   "Named:id=37:Caption"
      _StyleDefs(112) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(113) =   "Named:id=38:HighlightRow"
      _StyleDefs(114) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(115) =   "Named:id=39:EvenRow"
      _StyleDefs(116) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(117) =   "Named:id=40:OddRow"
      _StyleDefs(118) =   ":id=40,.parent=33"
      _StyleDefs(119) =   "Named:id=41:RecordSelector"
      _StyleDefs(120) =   ":id=41,.parent=34"
      _StyleDefs(121) =   "Named:id=42:FilterBar"
      _StyleDefs(122) =   ":id=42,.parent=33"
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   161
      Left            =   7920
      TabIndex        =   27
      Top             =   2760
      Width           =   975
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   975
      Index           =   1
      Left            =   10680
      TabIndex        =   119
      Top             =   8550
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1720
      _Version        =   393217
      Appearance      =   0
      TextRTF         =   $"SEI00181.frx":00C6
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   171
      Left            =   4080
      TabIndex        =   13
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   172
      Left            =   6000
      TabIndex        =   14
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   173
      Left            =   6960
      TabIndex        =   15
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   174
      Left            =   7920
      TabIndex        =   16
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   175
      Left            =   8880
      TabIndex        =   17
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   176
      Left            =   9840
      TabIndex        =   18
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
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
      Index           =   169
      Left            =   10770
      MaxLength       =   8
      TabIndex        =   122
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "�W���I��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   8520
      TabIndex        =   221
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   11040
      TabIndex        =   220
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "�I�敪"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   11760
      TabIndex        =   219
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  '�ׯ�
      Caption         =   "���s"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   140
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   11520
      TabIndex        =   218
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      Caption         =   "�S����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   10800
      TabIndex        =   217
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      Caption         =   "�ݒ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   9840
      TabIndex        =   215
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      Caption         =   "�ؑ֓�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   9840
      TabIndex        =   216
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      Caption         =   "BU���H"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   8880
      TabIndex        =   214
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      Caption         =   "PP���H"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   7920
      TabIndex        =   213
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      Caption         =   "�O��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   6960
      TabIndex        =   212
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   6000
      TabIndex        =   211
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   5040
      TabIndex        =   209
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      Caption         =   "(����)"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   5040
      TabIndex        =   210
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      Caption         =   "�H��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   4080
      TabIndex        =   208
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      Caption         =   "(����)"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3120
      TabIndex        =   207
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      Caption         =   "�H��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   206
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      Caption         =   "�H��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   205
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      Caption         =   "��ٰ�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   204
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      Caption         =   "ۯĐ�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   203
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "�i���ú�ذ"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   63
      Left            =   240
      TabIndex        =   201
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "�~/��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   61
      Left            =   1680
      TabIndex        =   200
      Top             =   9480
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "�H����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   60
      Left            =   480
      TabIndex        =   199
      Top             =   9480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "�~/��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   59
      Left            =   1680
      TabIndex        =   198
      Top             =   9150
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "��ڰ�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   58
      Left            =   480
      TabIndex        =   197
      Top             =   9150
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "�b"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   57
      Left            =   1680
      TabIndex        =   196
      Top             =   7170
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "�t���H��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   56
      Left            =   480
      TabIndex        =   195
      Top             =   7170
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "��/��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   54
      Left            =   480
      TabIndex        =   193
      Top             =   8820
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "��Ǝ��Ԍv"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   51
      Left            =   480
      TabIndex        =   191
      Top             =   8490
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "����ƍH��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   49
      Left            =   480
      TabIndex        =   189
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "�i�]�T���j"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   47
      Left            =   480
      TabIndex        =   187
      Top             =   7830
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "����ƍH��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   45
      Left            =   480
      TabIndex        =   185
      Top             =   7500
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "��/��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   55
      Left            =   1680
      TabIndex        =   194
      Top             =   8820
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "�b"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   52
      Left            =   1680
      TabIndex        =   192
      Top             =   8490
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "�b"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   50
      Left            =   1680
      TabIndex        =   190
      Top             =   8160
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   48
      Left            =   1680
      TabIndex        =   188
      Top             =   7830
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "�b"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   46
      Left            =   1680
      TabIndex        =   186
      Top             =   7500
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "�b"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   36
      Left            =   1680
      TabIndex        =   184
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  '�ׯ�
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   107
      Left            =   0
      TabIndex        =   183
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�ؑ֋敪"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   105
      Left            =   13350
      TabIndex        =   182
      Top             =   9990
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "(1:�V�K 2:���s)"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   103
      Left            =   11490
      TabIndex        =   180
      Top             =   9990
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���ϋ敪"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   102
      Left            =   10290
      TabIndex        =   179
      Top             =   9990
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�d�l����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   101
      Left            =   10320
      TabIndex        =   178
      Top             =   9600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "���Ϗ����l"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   100
      Left            =   10710
      TabIndex        =   177
      Top             =   8370
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "��ƍH���v(�b)"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   76
      Left            =   6750
      TabIndex        =   175
      Top             =   9840
      Width           =   1425
   End
   Begin VB.Label Label1 
      Appearance      =   0  '�ׯ�
      BorderStyle     =   1  '����
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   75
      Left            =   5520
      TabIndex        =   174
      Top             =   9480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  '�ׯ�
      BorderStyle     =   1  '����
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   74
      Left            =   5520
      TabIndex        =   173
      Top             =   9150
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  '�ׯ�
      BorderStyle     =   1  '����
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   73
      Left            =   5520
      TabIndex        =   172
      Top             =   8820
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  '�ׯ�
      BorderStyle     =   1  '����
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   53
      Left            =   5520
      TabIndex        =   171
      Top             =   8490
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  '�ׯ�
      BorderStyle     =   1  '����
      Caption         =   "���x���\��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   44
      Left            =   5520
      TabIndex        =   170
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  '�ׯ�
      BorderStyle     =   1  '����
      Caption         =   "�����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   43
      Left            =   5520
      TabIndex        =   169
      Top             =   7170
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  '�ׯ�
      BorderStyle     =   1  '����
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   42
      Left            =   5520
      TabIndex        =   168
      Top             =   7500
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  '�ׯ�
      BorderStyle     =   1  '����
      Caption         =   "���H���"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   41
      Left            =   5520
      TabIndex        =   167
      Top             =   7830
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  '�ׯ�
      BorderStyle     =   1  '����
      Caption         =   "�W��������"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   37
      Left            =   5520
      TabIndex        =   166
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BorderStyle     =   1  '����
      Caption         =   "��ƍH��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   82
      Left            =   5520
      TabIndex        =   165
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BorderStyle     =   1  '����
      Caption         =   "�P��/�b"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   40
      Left            =   6720
      TabIndex        =   164
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BorderStyle     =   1  '����
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   39
      Left            =   7560
      TabIndex        =   163
      Top             =   6600
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BorderStyle     =   1  '����
      Caption         =   "�H��/�b"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   38
      Left            =   8160
      TabIndex        =   162
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "�w�}�[���l"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   81
      Left            =   10680
      TabIndex        =   161
      Top             =   6570
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BorderStyle     =   1  '����
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   72
      Left            =   3360
      TabIndex        =   160
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BorderStyle     =   1  '����
      Caption         =   "�W��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   71
      Left            =   2520
      TabIndex        =   159
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "�o�א�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   20
      Left            =   0
      TabIndex        =   157
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "�R"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   35
      Left            =   14040
      TabIndex        =   156
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "�Q"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   34
      Left            =   12960
      TabIndex        =   155
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "�P"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   33
      Left            =   11880
      TabIndex        =   154
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "�P�Q"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   32
      Left            =   10800
      TabIndex        =   153
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "�P�P"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   31
      Left            =   9720
      TabIndex        =   152
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "�P�O"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   30
      Left            =   8640
      TabIndex        =   151
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "�X"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   29
      Left            =   7560
      TabIndex        =   150
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "�W"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   28
      Left            =   6480
      TabIndex        =   149
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "�V"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   27
      Left            =   5400
      TabIndex        =   148
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "�U"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   26
      Left            =   4320
      TabIndex        =   147
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "�T"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   25
      Left            =   3240
      TabIndex        =   146
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "�S"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   24
      Left            =   2160
      TabIndex        =   145
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   23
      Left            =   1080
      TabIndex        =   144
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "���N�x"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   22
      Left            =   0
      TabIndex        =   143
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "�O�N�x"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   21
      Left            =   0
      TabIndex        =   142
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Appearance      =   0  '�ׯ�
      Caption         =   "�V"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   17
      Left            =   0
      TabIndex        =   141
      Top             =   3000
      Width           =   495
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   15600
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label1 
      Caption         =   "�S����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   720
      TabIndex        =   139
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   10800
      TabIndex        =   138
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   10440
      TabIndex        =   137
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   9840
      TabIndex        =   136
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "���i�i��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   135
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "�d����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   134
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�P���ؑ֓�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   104
      Left            =   9810
      TabIndex        =   181
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "�O��H��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   62
      Left            =   480
      TabIndex        =   158
      Top             =   6840
      Width           =   1215
   End
End
Attribute VB_Name = "SEI00181"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'------------------------------------   '�e�L�X�g��`

Private Const ptxTanto_Code% = 0            '�S���҃R�[�h
Private Const ptxTanto_Name% = 1            '�S���Җ���
Private Const ptxHin_Gai% = 2               '�i��
Private Const ptxHin_Name% = 3              '�i��

Private Const ptxST_SOKO% = 4               '�W���I�ԁ@ �q��
Private Const ptxST_RETU% = 5               '�W���I��   ��
Private Const ptxST_REN% = 6                '�W���I�ԁ@ �A
Private Const ptxST_DAN% = 7                '�W���I�ԁ@ �i

Private Const ptxCATEGORY_CODE% = 72        '�i���ú�ذ����



Private Const ptxBEF_SEI_LOT% = 8           '�ύX�O�@   ���b�g��
Private Const ptxBEF_SEI_RATE% = 9          '           �����[�g
Private Const ptxBEF_S_KOUSU% = 10          '           �����[�g
Private Const ptxBEF_S_KOUSU_GENKA% = 11    '           (����)���i���H��
Private Const ptxBEF_S_KOUSU_BAIKA% = 12    '           (����)���i���H��
Private Const ptxBEF_S_SHIZAI_GENKA% = 13   '           (����)����
Private Const ptxBEF_S_SHIZAI_BAIKA% = 14   '           (����)����

Private Const ptxBEF_S_GAISO_TANKA% = 165   '           �O���P��
Private Const ptxBEF_S_PPSC_KAKO_KOSU% = 161 '          PPSC���H�P��
Private Const ptxBEF_S_BU_KAKO_KOSU% = 162  '           BU���H�P��




Private Const ptxBEF_S_KOUSU_SET_DATE% = 15 '          �ݒ��
Private Const ptxBEF_SEI_TANKA_TANTO% = 16  '          �S����
Private Const ptxBEF_SE_TANKA_MEMO% = 17    '          ����

Private Const ptxAFT_SEI_LOT% = 18          '�ύX��@   ���b�g��
Private Const ptxAFT_SEI_RATE% = 19         '           �����[�g
Private Const ptxAFT_S_KOUSU% = 20          '           �H��
Private Const ptxAFT_S_KOUSU_GENKA% = 21    '           (����)���i���H��
Private Const ptxAFT_S_KOUSU_BAIKA% = 22    '           (����)���i���H��
Private Const ptxAFT_S_SHIZAI_GENKA% = 23   '           (����)����
Private Const ptxAFT_S_SHIZAI_BAIKA% = 24   '           (����)����




Private Const ptxAFT_S_GAISO_TANKA% = 166   '           �O���P��
Private Const ptxAFT_S_PPSC_KAKO_KOSU% = 163 '          PPSC���H�P��
Private Const ptxAFT_S_BU_KAKO_KOSU% = 164  '           BU���H�P��


Private Const ptxAFT_S_KOUSU_SET_DATE% = 25 '          �ݒ��
Private Const ptxAFT_SEI_TANKA_TANTO% = 26  '          �S����
Private Const ptxAFT_SE_TANKA_MEMO% = 27    '          ����


Private Const ptxZEN_AVE% = 28              '�����Ϗo�א�   �O�N�x�@����
Private Const ptxZEN_SYUKAQTY04% = 29       '�����Ϗo�א�   �O�N�x�@4��
Private Const ptxZEN_SYUKAQTY05% = 30       '�@                     5��
Private Const ptxZEN_SYUKAQTY06% = 31       '�@                     6��
Private Const ptxZEN_SYUKAQTY07% = 32       '�@                     7��
Private Const ptxZEN_SYUKAQTY08% = 33       '�@                     8��
Private Const ptxZEN_SYUKAQTY09% = 34       '�@                     9��
Private Const ptxZEN_SYUKAQTY10% = 35       '�@                     10��
Private Const ptxZEN_SYUKAQTY11% = 36       '�@                     11��
Private Const ptxZEN_SYUKAQTY12% = 37       '�@                     12��
Private Const ptxZEN_SYUKAQTY01% = 38       '�@                     1��
Private Const ptxZEN_SYUKAQTY02% = 39       '�@                     2��
Private Const ptxZEN_SYUKAQTY03% = 40       '�@                     3��

Private Const ptxTOU_AVE% = 41              '�����Ϗo�א�   ���N�x�@����
Private Const ptxTOU_SYUKAQTY04% = 42       '�����Ϗo�א�   ���N�x�@4��
Private Const ptxTOU_SYUKAQTY05% = 43       '�@                     5��
Private Const ptxTOU_SYUKAQTY06% = 44       '�@                     6��
Private Const ptxTOU_SYUKAQTY07% = 45       '�@                     7��
Private Const ptxTOU_SYUKAQTY08% = 46       '�@                     8��
Private Const ptxTOU_SYUKAQTY09% = 47       '�@                     9��
Private Const ptxTOU_SYUKAQTY10% = 48       '�@                     10��
Private Const ptxTOU_SYUKAQTY11% = 49       '�@                     11��
Private Const ptxTOU_SYUKAQTY12% = 50       '�@                     12��
Private Const ptxTOU_SYUKAQTY01% = 51       '�@                     1��
Private Const ptxTOU_SYUKAQTY02% = 52       '�@                     2��
Private Const ptxTOU_SYUKAQTY03% = 53       '�@                     3��

'-------------------------------------------'   �O��H��    2011.12.12
Private Const ptxCATE_ST_KOUTEI% = 54       ' �O��H���i�b�j�W��
Private Const ptxCATE_AD_KOUTEI% = 55       ' �O��H���i�b�j����

Private Const ptxCATE_ST_FUKA% = 56         ' �t���H���i�b�j�W��
Private Const ptxCATE_AD_FUKA% = 57         ' �t���H���i�b�j����

Private Const ptxCATE_ST_JITU1% = 58        ' ����ƍH���i�b�j�W��
Private Const ptxCATE_AD_JITU1% = 59        ' ����ƍH���i�b�j����

Private Const ptxCATE_ST_YOYU_RITU% = 60    ' �]�T���i���j�W��
Private Const ptxCATE_AD_YOYU_RITU% = 61    ' �]�T���i���j����

Private Const ptxCATE_ST_JITU2% = 62        ' ����ƍH���i�b�j�W��
Private Const ptxCATE_AD_JITU2% = 63        ' ����ƍH���i�b�j����

Private Const ptxCATE_ST_TOTAL% = 64        ' ��Ǝ��Ԍv�i�b�j�W��
Private Const ptxCATE_AD_TOTAL% = 65        ' ��Ǝ��Ԍv�i�b�j����

Private Const ptxCATE_ST_FUN% = 66          ' ��/�i��/�j�W��
Private Const ptxCATE_AD_FUN% = 67          ' ��/�i��/�j����

Private Const ptxCATE_ST_FUN_RATE% = 68     ' ��ڰāi�~/���j�W��
Private Const ptxCATE_AD_FUN_RATE% = 69     ' ��ڰāi�~/���j����

Private Const ptxCATE_ST_KOURYO% = 70       ' �H�����i�~/�j�W��
Private Const ptxCATE_AD_KOURYO% = 71       ' �H�����i�~/�j����

'-------------------------------------------'   �O��s��    2011.12.12

Private Const ptxMAIN_KOUTEI_TANI01% = 86   '��ƍH��01 �P��
Private Const ptxMAIN_KOUTEI_QTY01% = 87    '           ����
Private Const ptxMAIN_KOUTEI_KOUSU01% = 88  '           �H��
Private Const ptxMAIN_KOUTEI_TANI02% = 89   '��ƍH��02 �P��
Private Const ptxMAIN_KOUTEI_QTY02% = 90    '           ����
Private Const ptxMAIN_KOUTEI_KOUSU02% = 91  '           �H��
Private Const ptxMAIN_KOUTEI_TANI03% = 92   '��ƍH��03 �P��
Private Const ptxMAIN_KOUTEI_QTY03% = 93    '           ����
Private Const ptxMAIN_KOUTEI_KOUSU03% = 94  '           �H��
Private Const ptxMAIN_KOUTEI_TANI04% = 95   '��ƍH��04 �P��
Private Const ptxMAIN_KOUTEI_QTY04% = 96    '           ����
Private Const ptxMAIN_KOUTEI_KOUSU04% = 97  '           �H��
Private Const ptxMAIN_KOUTEI_TANI05% = 98   '��ƍH��05 �P��
Private Const ptxMAIN_KOUTEI_QTY05% = 99    '           ����
Private Const ptxMAIN_KOUTEI_KOUSU05% = 100 '           �H��
Private Const ptxMAIN_KOUTEI_TANI06% = 101  '��ƍH��06 �P��
Private Const ptxMAIN_KOUTEI_QTY06% = 102   '           ����
Private Const ptxMAIN_KOUTEI_KOUSU06% = 103 '           �H��
Private Const ptxMAIN_KOUTEI_TANI07% = 104  '��ƍH��07 �P��
Private Const ptxMAIN_KOUTEI_QTY07% = 105   '           ����
Private Const ptxMAIN_KOUTEI_KOUSU07% = 106 '           �H��
Private Const ptxMAIN_KOUTEI_TANI08% = 107  '��ƍH��08 �P��
Private Const ptxMAIN_KOUTEI_QTY08% = 108   '           ����
Private Const ptxMAIN_KOUTEI_KOUSU08% = 109 '           �H��
Private Const ptxMAIN_KOUTEI_TANI09% = 110  '��ƍH��09 �P��
Private Const ptxMAIN_KOUTEI_QTY09% = 111   '           ����
Private Const ptxMAIN_KOUTEI_KOUSU09% = 112 '           �H��

Private Const ptxMAIN_KOUTEI_KEI1% = 113    '��ƍH���v �v


Private Const ptxS_CLASS_CODE% = 155        '���i���׽
Private Const ptxF_CLASS_CODE% = 156        '�t���׽
Private Const ptxN_CLASS_CODE% = 157        '���E�׽

Private Const ptxIO_TANKA_No% = 158         '�I�敪
Private Const ptxSE_Name% = 159             '�I�敪����





Private Const ptxSHIYOU_NO% = 167           '�d�l����       2009.06.02
Private Const ptxMITSUMORI_KBN% = 168       '���ς�敪     2009.06.02
Private Const ptxKIRIKAE_KBN% = 170         '�ؑ֋敪       2009.06.02
    







'------2009.07.24
Private Const ptxOLD_S_KOUSU_BAIKA% = 171       ' ��  (����)���i���H��
Private Const ptxOLD_S_SHIZAI_BAIKA% = 172      ' ��  (����)����

Private Const ptxOLD_S_GAISO_TANKA% = 173       ' ��  �O���P��
Private Const ptxOLD_S_PPSC_KAKO_KOSU% = 174    ' ��  PPSC���H�P��
Private Const ptxOLD_S_BU_KAKO_KOSU% = 175      ' ��  BU���H�P��
Private Const ptxTANKA_KIRIKAE_DT% = 176        ' ��  �P���ؑ֓��t
'------2009.07.24




'------------------------------------   '�R���{��`
Private Const pcmbSHIMUKE% = 0          '�d������
Private Const pcmbCATEGORY_Name% = 1    '�i���ú�ذ


'------------------------------------   '���b�`�e�L�X�g�{�b�N�X��`
Private Const prchBIKOU% = 0            '���l
Private Const prchM_BIKOU% = 1          '���Ϗ����l



'------------------------------------   '�\���i
Private Const pGrdKOUSEI% = 0

Dim KOUSEI      As New XArrayDB

Private Const Min_Row% = 1              '�ŏ��s��

Dim Max_Row   As Integer                '�O���b�h�ő�\������

Private Const Min_Col% = 0              '�ŏ���
Private Const Max_Col% = 14             '�ő��

Private Const ColKO_JGYOBU% = 0         '���ƕ�
Private Const ColKO_NAIGAI% = 1         '�����O


Private Const ColKO_SYUBETSU% = 2       '���
Private Const ColKO_HIN_GAI% = 3        '�i��
Private Const ColKO_HIN_NAME% = 4       '�i��
Private Const ColKO_QTY% = 5            '����
Private Const ColG_ST_SHITAN% = 6       '�d����
Private Const ColG_ST_URITAN% = 7       '���し

Private Const ColG_SPTAN% = 8           '���ʒP����

Private Const ColG_ST_SHIKIN% = 9       '�d�����z
Private Const ColG_ST_URIKIN% = 10      '������z
Private Const ColS_KOUSU% = 11          '��Ǝ���
Private Const ColSEI_SYU_KON% = 12      '�W������
Private Const ColKO_BIKOU% = 13         '���l
                                        
                                        '���� ���z�o�͗p
Private Const ColG_ST_URIKIN_KUSATU% = 14

'-----------------------------------    �h���b�v�_�E��
Dim SYUBETSU        As New XArrayDB


'-----------------------------------

Dim KOSOU_KBN       As String * 2       '���敪
Dim GAISO_KBN       As String * 2       '�O���敪


Dim INV_IO_TANKA_No As String * 2       '�W���I���o�^���̏o�ɋ敪
Dim HIN_INV         As Boolean          '���o�^�i�Ԃ̓o�^��


Dim KUSATU_F        As Boolean          '�ΏۃZ���^�[�@���� OR ���ÈȊO


Dim SHIZAI_T        As Variant          '���ޑΏ�
Dim DOUKON_T        As Variant          '�����Ώ�
Dim KAKOU_T         As Variant          '���H�Ώ�

Dim BU_T            As Variant          'BU�t���Ώ�
Dim PPSC_T          As Variant          'PPSC�t���Ώ�

Private Const KUSATU_ETC$ = "���̑�"


Dim svHin_Gai       As String           '�i��
Dim svSHIMUKE_CODE  As String           '�d������
Dim svCATEGORY_CODE As String           '�i���ú�ذ����


Dim FUTAI_KBN       As String * 2       '�t�э�� 2009.09.05


Dim ITEM_CATEGORY_SUMI  As Variant      '���i���ς�    �i�ڶú�ذ(���ޕi�����p) 2013.01.16
Dim CHK_SHIZAI_T        As Variant      '�����Ώێ���                           2013.01.16

'-----------------------------------    �d�w�b�d�k �������Z��

Dim EX_NAME1        As String           '�����P
Dim EX_NAME2        As String           '�����Q

Dim EX_SYAMEI       As String           '���Ё@����
Dim EX_ADDR1        As String           '���Ё@�Z���P
Dim EX_ADDR2        As String           '���Ё@�Z���Q


Dim EX_CENTER_NAME  As String           '�Z���^�[   ����
Dim EX_CENTER_ADDR1 As String           '�Z���^�[   �Z���P
Dim EX_CENTER_ADDR2 As String           '�Z���^�[   �Z���Q

Dim EX_BIKOU1       As String           '���l�P
Dim EX_BIKOU2       As String           '���l�Q




'2009.06.02
Dim EX_SHIZAI_T     As Variant          '���ޑΏ�
Dim EX_SHIZAI_F     As Boolean          '���ޑΏ�

Dim EX_DOUKON_T     As Variant          '�����Ώ�
Dim EX_DOUKON_F     As Boolean          '�����Ώ�

Dim EX_FUKA_T       As Variant          '�t�����
Dim EX_FUKA_F       As Boolean          '�t�����
'2009.06.02


Dim SP_KOUSU_T      As Variant          '���ʒP��(��ƍH���@�b/��)
Dim SP_KOURYO_T     As Variant          '���ʒP��(�H��@)
Dim SP_HAKO_T       As Variant          '���ʒP��(����@)




Dim EX_BCR_CODE     As String           '�ް�������ٺ���


Dim EXCEL_TEMPLATE  As String           'EXCEL����ڰ�


'--------------------------------------- EXCEL�p�萔
Private Const xlCalculationManual% = -4135
Private Const xlLeft% = -4131
Private Const xlCenter% = -4108
Private Const xlBottom% = -4107
Private Const xlNone% = -4142
Private Const xlContinuous% = 1
Private Const xlThin% = 2
Private Const xlAutomatic% = -4105
Private Const xlRight% = -4152
Private Const xlDiagonalDown% = 5
Private Const xlDiagonalUp% = 6
Private Const xlEdgeLeft% = 7
Private Const xlEdgeTop% = 8
Private Const xlEdgeBottom% = 9
Private Const xlEdgeRight% = 10
Private Const xlInsideVertical% = 11
Private Const xlInsideHorizontal% = 12
Private Const xlThick% = 4
Private Const xlCalculationAutomatic% = -4105
Private Const xlPortrait% = 1
Private Const xlDot% = -4118

'--------------------------------------- EXCEL�p�萔
Dim Insert_Pic       As String           '���

Dim SYONIN_Pic       As String           '���F��


Dim MAIN_HIN_GAI    As String * 20

Dim Save_Dir        As String

Dim SEI0018_LOG     As String


Private IN_cnt  As Integer
Private OK_cnt  As Integer
Private NG_cnt  As Integer

Private KIN_NG_CNT  As Integer





'--------   ���O�ް��ݒ�    2018.05.16
Private DATA_KBN_TBL()      As String * 1
Private SYUBETSU_TBL()      As String * 2
'--------   ���O�ް��ݒ�    2018.05.16







'Private Const LAST_UPDATE_DAY$ = "[SEI0018] 2018.06.11 15:00"
Private Const LAST_UPDATE_DAY$ = "[SEI0018] 2018.07.18 15:30"
Private Sub Combo1_Change(Index As Integer)
Dim i   As Integer
    
    
    Select Case Index
    
        Case pcmbSHIMUKE
        
        
            If Trim(svHin_Gai) = (Text1(ptxHin_Gai).Text) Then
            Else
                For i = 2 To 5
                    Command1(i).Enabled = False
                Next i
            End If
    
    
    
                        '�i���ú�ذ�̃Z�b�g
'            If ITEM_CATEGORY_Set_Proc() Then
'                Unload Me
'            End If
    
        Case pcmbCATEGORY_Name
    
            If Trim(Right(Combo1(Index).Text, 8)) = Trim(Text1(ptxCATEGORY_CODE).Text) Then
            Else
                Text1(ptxCATEGORY_CODE).Text = Trim(Right(Combo1(Index).Text, 8))
            End If
    End Select

End Sub

Private Sub Combo1_GotFocus(Index As Integer)


    Select Case Index
        Case pcmbSHIMUKE
            svSHIMUKE_CODE = Right(Combo1(pcmbSHIMUKE).Text, 2)
    
        Case pcmbCATEGORY_Name
            svCATEGORY_CODE = Text1(ptxCATEGORY_CODE).Text
    
    End Select

End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If


    Select Case Index
        Case pcmbSHIMUKE
            svSHIMUKE_CODE = Right(Combo1(pcmbSHIMUKE).Text, 2)
    
        Case pcmbCATEGORY_Name
            If Trim(Right(Combo1(Index).Text, 8)) = Trim(Text1(ptxCATEGORY_CODE).Text) Then
            Else
                Text1(ptxCATEGORY_CODE).Text = Trim(Right(Combo1(Index).Text, 8))
            
                If CATEGORY_Disp_Proc() Then
                    Unload Me
                End If
            
            End If
    End Select



End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim i   As Integer
    
    Select Case Index
        Case pcmbSHIMUKE
        
            If Trim(svSHIMUKE_CODE) = Right(Combo1(pcmbSHIMUKE).Text, 2) Then
            Else
                For i = 2 To 5
                    Command1(i).Enabled = False
                Next i
            End If
                        '�i���ú�ذ�̃Z�b�g
            If ITEM_CATEGORY_Set_Proc() Then
                Unload Me
            End If
        
        
            '�i���J�e�S���B
'2011.12.26            Text1(ptxCATEGORY_CODE).Text = Trim(StrConv(ITEMREC.CATEGORY_CODE, vbUnicode))
            For i = 0 To Combo1(pcmbCATEGORY_Name).ListCount - 1
                If Trim(Text1(ptxCATEGORY_CODE).Text) = Trim(Right(Combo1(pcmbCATEGORY_Name).List(i), 8)) Then
                    Text1(ptxCATEGORY_CODE).Text = Trim(Right(Combo1(pcmbCATEGORY_Name).List(i), 8))
                    Combo1(pcmbCATEGORY_Name).ListIndex = i
                    Exit For
                End If
            Next i
            If i > Combo1(pcmbCATEGORY_Name).ListCount - 1 Then
                Combo1(pcmbCATEGORY_Name).ListIndex = 0
            End If
        
        
        
        
        Case pcmbCATEGORY_Name
            If Trim(Right(Combo1(Index).Text, 8)) = Trim(Text1(ptxCATEGORY_CODE).Text) Then
            Else
                Text1(ptxCATEGORY_CODE).Text = Trim(Right(Combo1(Index).Text, 8))
            
                If CATEGORY_Disp_Proc() Then
                    Unload Me
                End If
            
            End If
    End Select
End Sub

Private Sub Command1_Click(Index As Integer)


Dim ans     As Integer
Dim i       As Integer

Dim MESG    As String
Dim Errflg  As Integer


    Select Case Index
    
        Case 0      '�I��
            Unload Me
    
        Case 1      '�����i�\���j
        
        
            If Detail_Disp_Proc(Errflg) Then
                Unload Me
            End If
        
            Text1(ptxCATEGORY_CODE).SetFocus
        
        
        Case 2      '�ۑ�
            
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.02.16
            If Trim(Text1(ptxCATEGORY_CODE).Text) = "" Then
                MsgBox "�i���J�e�S���[�́A�K�{���͂ł��B�ē��͂��ĉ������ "
                Text1(ptxCATEGORY_CODE).SetFocus
                Exit Sub
            End If
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.02.16
            
            For i = ptxTanto_Code To ptxAFT_S_BU_KAKO_KOSU
            
                If Error_Check_Proc(i) Then
                    Exit Sub
                End If
            
            
            Next i
            '2009.06.02
            For i = ptxSHIYOU_NO To ptxKIRIKAE_KBN
            
                If Error_Check_Proc(i) Then
                    Exit Sub
                End If
            
            
            Next i
            If Grid_Error_Check_Proc() Then
                Exit Sub
            End If
        
            MESG = "���i���\���f�[�^��ۑ����܂��B" & vbCrLf
            MESG = MESG & "�@�@��ʁ^�i�ԁ^����" & vbCrLf
            MESG = MESG & "�@�@�w�}�[���l" & vbCrLf
            MESG = MESG & "��낵���ł����H" & vbCrLf
        
        
            ans = MsgBox(MESG, vbYesNo + vbDefaultButton2 + vbExclamation, "���i���\���̕ۑ��m�F")
            If ans = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
            
                If Detail_Disp_Proc(Errflg) Then
                    Unload Me
                End If
            
            End If
        
            Command1(4).Enabled = True          '2013.01.17
                    
            Text1(ptxAFT_SEI_LOT).SetFocus
        
        Case 3      '�P���v�Z
        
            For i = ptxTanto_Code To ptxAFT_S_BU_KAKO_KOSU
            
                If Error_Check_Proc(i) Then
                    Exit Sub
                End If
            
            Next i
        
        
            If TANKA_KEISAN_Proc() Then
                Unload Me
            End If
        
            Command1(4).Enabled = True          '2013.01.17
        
        Case 4      '���Ϗ����s
            
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.02.16
            If Trim(Text1(ptxCATEGORY_CODE).Text) = "" Then
                MsgBox "�i���J�e�S���[�́A�K�{���͂ł��B�ē��͂��ĉ������ "
                Text1(ptxCATEGORY_CODE).SetFocus
                Exit Sub
            End If
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.02.16
            
            
            
            If Estimate_Proc() Then
                Unload Me
            End If
        
        Case 5      '�P���o�^
            
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.02.16
            If Trim(Text1(ptxCATEGORY_CODE).Text) = "" Then
                MsgBox "�i���J�e�S���[�́A�K�{���͂ł��B�ē��͂��ĉ������ "
                Text1(ptxCATEGORY_CODE).SetFocus
                Exit Sub
            End If
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.02.16
            
            
            
            For i = ptxTanto_Code To ptxAFT_S_BU_KAKO_KOSU
            
                If Error_Check_Proc(i) Then
                    Exit Sub
                End If
            
            
            Next i
            
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2011.12.21
            If TANKA_KEISAN_Proc() Then
                Unload Me
            End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2011.12.21
            
            
            MESG = "�P����o�^���܂��B��낵���ł����H" & vbCrLf
            MESG = MESG & "�@���b�g���F" & Text1(ptxAFT_SEI_LOT).Text & vbCrLf
            MESG = MESG & "�@�����[�g�F" & Text1(ptxAFT_SEI_RATE).Text & vbCrLf
            MESG = MESG & "�@�H���F" & Text1(ptxAFT_S_KOUSU).Text & vbCrLf
            MESG = MESG & "�@�i�����j�H���F" & Text1(ptxAFT_S_KOUSU_GENKA).Text & vbCrLf
            MESG = MESG & "�@ (����) �H���F" & Text1(ptxAFT_S_KOUSU_BAIKA).Text & vbCrLf
            MESG = MESG & "�@�i�����j����F" & Text1(ptxAFT_S_SHIZAI_GENKA).Text & vbCrLf
            MESG = MESG & "�@ (����) ����F" & Text1(ptxAFT_S_SHIZAI_BAIKA).Text & vbCrLf
            MESG = MESG & "�@ �ݒ���F" & Text1(ptxAFT_S_KOUSU_SET_DATE).Text & vbCrLf
            MESG = MESG & "�@ �S���ҁF" & Text1(ptxAFT_SEI_TANKA_TANTO).Text & vbCrLf
            MESG = MESG & "�@ �����F" & Text1(ptxAFT_SE_TANKA_MEMO).Text & vbCrLf

            
            
            
            ans = MsgBox(MESG, vbYesNo + vbDefaultButton1 + vbExclamation, "�m�F����")
            If ans = vbYes Then
                If Tanka_Update_Proc() Then
                    Unload Me
                End If
            
                If Detail_Disp_Proc(Errflg) Then
                    Unload Me
                End If
            
            
            End If
                    
            Command1(4).Enabled = True          '2013.01.17
            
            Text1(ptxAFT_SEI_LOT).SetFocus
    
    
        Case 6  '���@�P���v�Z   2013.01.16
            
            For i = ptxTanto_Code To ptxAFT_S_BU_KAKO_KOSU
            
                If Error_Check_Proc(i) Then
                    Exit Sub
                End If
            
            Next i
        
        
            If KARI_TANKA_KEISAN_Proc() Then
                Unload Me
            End If
    
    
            Command1(4).Enabled = False
    End Select






End Sub

Private Sub Command2_Click(Index As Integer)
Dim i               As Integer

Dim wkLine          As Variant
Dim wkItem          As Variant

Dim ans             As Integer
Dim sts             As Integer

Dim S_DATETIME      As String


    Select Case Index
        Case 0
            Text2.Text = ""         '2018.03.12
            Text3.Text = ""         '2018.03.12
        Case 1
        
        
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, txtTANTO_CODE.Text)
        
            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                    txtTanto_Name.Text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    txtTanto_Name.Text = ""
            
                    MsgBox "���͂������ڂ̓G���[�ł��B(�S����)"
                    txtTANTO_CODE.SetFocus
                    Exit Sub
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
                    Unload Me
                    Exit Sub
            
            End Select
        
                    
            '>>>>>  �����敪    2018.05.25
            If Trim(txtBUHIN) = "" Or txtBUHIN = "0" Or txtBUHIN = "1" Or txtBUHIN = "2" Or txtBUHIN = "3" Then
            Else
        
                MsgBox "���͂������ڂ̓G���[�ł��B(�����敪)"
                txtBUHIN.SetFocus
                Exit Sub
            End If
                    
        
        
        
        
        
        
        
            Beep
            ans = MsgBox("[���Ϗ��ꊇ���s]���s���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbNo Then
                Exit Sub
            End If
Call LOG_OUT(SEI0018_LOG, "���Ϗ��ꊇ�쐬�@�����J�n[" & Now & "]")
        
            S_DATETIME = Now
        
            For i = 0 To 2
                Command2(i).Enabled = False
            Next i
            
            Text2.Locked = True
            Text3.Locked = True
            
            
            SEI00181.MousePointer = vbHourglass
            DoEvents
        
        
            List2.Clear
            
            IN_cnt = 0
            OK_cnt = 0
            NG_cnt = 0
            
            KIN_NG_CNT = 0
            
            txtIN_CNT.Text = Format(OK_cnt, "#,##0")
            txtOK_CNT.Text = Format(OK_cnt, "#,##0")
            txtNG_CNT.Text = Format(NG_cnt, "#,##0")
                                
            txtKIN_NG_CNT.Text = Format(KIN_NG_CNT, "#,##0")
                                
            wkLine = Split(Text2.Text, vbCrLf, -1)
    
    
            Combo1(pcmbSHIMUKE).ListIndex = cmbSHIMUKE.ListIndex           '2018.03.07
    
    
            For i = 0 To UBound(wkLine)
                
                
                wkItem = Split(wkLine(i), vbTab, -1)
                
                
                If UBound(wkItem) < 0 Then
                Else
                    IN_cnt = IN_cnt + 1
                    txtIN_CNT.Text = Format(IN_cnt, "#,##0")
                
                    MAIN_HIN_GAI = wkItem(0)
                
                    If Main_Update_Proc() Then
                        Unload Me
                    End If
                
                
                
                
                    DoEvents
                
                End If
    
            Next i
                    
                    
'>>>>>>>>>>>>>>>>>  �e�i�ԕ�    2018.03.12
            List3.Clear
            
            
            txtIN_CNT.Text = Format(OK_cnt, "#,##0")
            txtOK_CNT.Text = Format(OK_cnt, "#,##0")
            txtNG_CNT.Text = Format(NG_cnt, "#,##0")
                                
            txtKIN_NG_CNT.Text = Format(KIN_NG_CNT, "#,##0")
                                
            wkLine = Split(Text3.Text, vbCrLf, -1)
    
    
            Combo1(pcmbSHIMUKE).ListIndex = cmbSHIMUKE.ListIndex           '2018.03.07
    
    
            For i = 0 To UBound(wkLine)
                
                
                wkItem = Split(wkLine(i), vbTab, -1)
                
                
                If UBound(wkItem) < 0 Then
                Else
                    IN_cnt = IN_cnt + 1
                    txtIN_CNT.Text = Format(IN_cnt, "#,##0")
                
                    
                    
                    MAIN_HIN_GAI = wkItem(0)
                
                
                
                    If Main_Update_OYA_Proc() Then
                        Unload Me
                    End If
                
                
                
                
                    DoEvents
                
                End If
    
            Next i
'>>>>>>>>>>>>>>>>>  �e�i�ԕ�    2018.03.12
                    
                    
                    
            DoEvents
        
Call LOG_OUT(SEI0018_LOG, "���Ϗ��ꊇ�쐬�@����I��[" & Now & "]")
            MsgBox "����I�����܂����B[" & S_DATETIME & "��" & Now & "]"
        
            For i = 0 To 2
                Command2(i).Enabled = True
            Next i
        
            Text2.Locked = False
            Text3.Locked = False
            
        
        
        
           SEI00181.MousePointer = vbDefault
           DoEvents
        
        Case 2
    
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, txtTANTO_CODE.Text)
        
            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                    txtTanto_Name.Text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    txtTanto_Name.Text = ""
            
                    MsgBox "���͂������ڂ̓G���[�ł��B(�S����)"
                    txtTANTO_CODE.SetFocus
                    Exit Sub
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
                    Unload Me
                    Exit Sub
            
            End Select
    
    
    
    
            '>>>>>  �����敪    2018.05.25
            If Trim(txtBUHIN) = "" Or txtBUHIN = "0" Or txtBUHIN = "1" Or txtBUHIN = "2" Or txtBUHIN = "3" Then
            Else
        
                MsgBox "���͂������ڂ̓G���[�ł��B(�����敪)"
                txtBUHIN.SetFocus
                Exit Sub
            End If
    
    
    
            List2.Clear
    
            IN_cnt = 0
            
            OK_cnt = 0
            NG_cnt = 0
            
            KIN_NG_CNT = 0
            
            txtIN_CNT.Text = Format(OK_cnt, "#,##0")
            txtOK_CNT.Text = Format(OK_cnt, "#,##0")
            txtNG_CNT.Text = Format(NG_cnt, "#,##0")
    
    
            txtKIN_NG_CNT.Text = Format(KIN_NG_CNT, "#,##0")
    
    
    
            txtOUT_CNT = ""
            IN_cnt = 0
    
            For i = 0 To 1
                Command2(i).Enabled = False
            Next i
            
            Text2.Locked = True
            Text3.Locked = True
            
            
            
            SEI00181.MousePointer = vbHourglass
            DoEvents
    
    
    
            wkLine = Split(Text2.Text, vbCrLf, -1)
    
            Combo1(pcmbSHIMUKE).ListIndex = cmbSHIMUKE.ListIndex           '2018.03.07
            
            
            For i = 0 To UBound(wkLine)
                
                
                wkItem = Split(wkLine(i), vbTab, -1)
                
                
                If UBound(wkItem) < 0 Then
                Else
                
                    MAIN_HIN_GAI = wkItem(0)
                
                    If COUNT_Proc() Then
                        Unload Me
                    End If
                
                
                
                
                    DoEvents
                
                End If
    
            Next i
'>>>>>>>>>>>>>>>>>  �e�i�ԕ�    2018.03.12

            List3.Clear
    
            
            txtIN_CNT.Text = Format(OK_cnt, "#,##0")
            txtOK_CNT.Text = Format(OK_cnt, "#,##0")
            txtNG_CNT.Text = Format(NG_cnt, "#,##0")
    
    
            txtKIN_NG_CNT.Text = Format(KIN_NG_CNT, "#,##0")
    
    
    
    
    
            For i = 0 To 1
                Command2(i).Enabled = False
            Next i
            SEI00181.MousePointer = vbHourglass
            DoEvents
    
    
    
            wkLine = Split(Text3.Text, vbCrLf, -1)
    
            Combo1(pcmbSHIMUKE).ListIndex = cmbSHIMUKE.ListIndex           '2018.03.07
            
            
            For i = 0 To UBound(wkLine)
                
                
                wkItem = Split(wkLine(i), vbTab, -1)
                
                
                If UBound(wkItem) < 0 Then
                Else
                
                    MAIN_HIN_GAI = wkItem(0)
                    
                    List3.AddItem MAIN_HIN_GAI


                    IN_cnt = IN_cnt + 1
                    txtOUT_CNT.Text = Format(IN_cnt, "#,##0")
                
                
                
                
                    DoEvents
                
                End If
    
            Next i



'>>>>>>>>>>>>>>>>>  �e�i�ԕ�    2018.03.12
    
    
    
        
            For i = 0 To 1
                Command2(i).Enabled = True
            Next i
        
            Text2.Locked = False
            Text3.Locked = False
            
        
        
            SEI00181.MousePointer = vbDefault
            DoEvents
        
        
        
        Case 3
            Unload Me
    
    
    
    
    
    
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
Dim c       As String * 128
Dim sts     As Integer


Dim wkVAL   As Variant  '2018.05.16
Dim i       As Integer  '2018.05.16


'    If App.PrevInstance Then
'        Beep
'        MsgBox "����v���O�������s���ł��B"
'        End
'    End If


    
    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[�����V�X�e��]���i���P�����ύ쐬���� �i���J�e�S���[�Ή�", Me.hwnd, 0)
    '�Ō�̗v�f��-1�ɂ����
    '�e�E�B���h�E�̑S�̂̕��̎c��̕���
    '�����I�Ɋ��蓖�Ă�
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)



    Show
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
                                '�Z���^�[�̎���
    If GetIni(App.EXEName, "KUSATU", App.EXEName, c) Then
        KUSATU_F = False
    Else
        If Trim(c) = "1" Then
            KUSATU_F = True
        Else
            KUSATU_F = False
        End If
    End If
                                '�����ދ敪�̊l��
    If GetIni(App.EXEName, "KOSOU", App.EXEName, c) Then
        Beep
        MsgBox "�����ދ敪�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    Else
        KOSOU_KBN = Trim(c)
    End If
                                '�O�����ދ敪�̊l��
    If GetIni(App.EXEName, "GAISO", App.EXEName, c) Then
        Beep
        MsgBox "�O�����ދ敪�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    Else
        GAISO_KBN = Trim(c)
    End If
                                '���o�^���̏o�ɋ敪�̊l��
    If GetIni(App.EXEName, "INV_IO_TANKA_No", App.EXEName, c) Then
        INV_IO_TANKA_No = ""
    Else
        INV_IO_TANKA_No = Trim(c)
    End If
                                '���o�^�i�Ԃ̓o�^�ۂ̊l��
    If GetIni(App.EXEName, "HIN_INV", App.EXEName, c) Then
        HIN_INV = False
    Else
        If Trim(c) = "0" Then
            HIN_INV = False
        Else
            HIN_INV = True
        End If
    End If
                                '���ޑΏێ��
    If GetIni(App.EXEName, "SHIZAI", App.EXEName, c) Then
        Beep
        MsgBox "���ޑΏۂ̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    Else
        SHIZAI_T = Split(Trim(c), ",", -1)
    End If
                                '�����Ώێ��
    If GetIni(App.EXEName, "DOUKON", App.EXEName, c) Then
        Beep
        MsgBox "�����Ώۂ̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    Else
        DOUKON_T = Split(Trim(c), ",", -1)
    End If
                                '���H�Ώێ��
   If GetIni(App.EXEName, "KAKOU", App.EXEName, c) Then
        Beep
        MsgBox "���H�Ώۂ̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    Else
        KAKOU_T = Split(Trim(c), ",", -1)
    End If
                                'PPSC�Ώێ��
    If GetIni(App.EXEName, "PPSC", App.EXEName, c) Then
        Beep
        MsgBox "PPSC�Ώۂ̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    Else
        PPSC_T = Split(Trim(c), ",", -1)
    End If
                                'BU�Ώێ��
    If GetIni(App.EXEName, "BU", App.EXEName, c) Then
        Beep
        MsgBox "BU�Ώۂ̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    Else
        BU_T = Split(Trim(c), ",", -1)
    End If
                                '�t�э�Ƃ̊l�� 2009.09.05
    If GetIni(App.EXEName, "FUTAI", App.EXEName, c) Then
        FUTAI_KBN = ""
    Else
        FUTAI_KBN = Trim(c)
    End If
                                '���ʒP��(��ƍH���@�b/��)
    If GetIni("SpecialPrice", "SP_KOUSU", App.EXEName, c) Then
        Beep
        MsgBox "���ʒP��(��ƍH���@�b/��)�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    Else
        SP_KOUSU_T = Split(Trim(c), ",", -1)
    End If
                                '���ʒP��(�H��@)
    If GetIni("SpecialPrice", "SP_KOURYO", App.EXEName, c) Then
        Beep
        MsgBox "���ʒP��(�H��@)�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    Else
        SP_KOURYO_T = Split(Trim(c), ",", -1)
    End If
                                '���ʒP��(����@)
    If GetIni("SpecialPrice", "SP_HAKO", App.EXEName, c) Then
        Beep
        MsgBox "���ʒP��(����@)�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    Else
        SP_HAKO_T = Split(Trim(c), ",", -1)
    End If


                                '���������� �i�ڶú�ذ  2013.01.16
    If GetIni(App.EXEName, "ITEM_CATEGORY_SUMI", App.EXEName, c) Then
        c = "********"
        ITEM_CATEGORY_SUMI = Split(Trim(c), ",", -1)
    Else
        ITEM_CATEGORY_SUMI = Split(Trim(c), ",", -1)
    End If
                                '�������� ��ʺ���      2013.01.16
    If GetIni(App.EXEName, "CHK_SHIZAI", App.EXEName, c) Then
        c = "**"
        CHK_SHIZAI_T = Split(Trim(c), ",", -1)
    Else
        CHK_SHIZAI_T = Split(Trim(c), ",", -1)
    End If








'------------------------------------------------------ EXCEL�p����
                                '���Ϗ� �����P
    If GetIni("Estimate", "NAME1", App.EXEName, c) Then
        EX_NAME1 = ""
    Else
        EX_NAME1 = Trim(c)
    End If
                                '���Ϗ� �����Q
    If GetIni("Estimate", "NAME2", App.EXEName, c) Then
        EX_NAME2 = ""
    Else
        EX_NAME2 = Trim(c)
    End If
                                '���Ϗ� ���Ё@����
    If GetIni("Estimate", "SYAMEI", App.EXEName, c) Then
        EX_SYAMEI = ""
    Else
        EX_SYAMEI = Trim(c)
    End If
                                '���Ϗ� ���Ё@�Z���P
    If GetIni("Estimate", "ADDR1", App.EXEName, c) Then
        EX_ADDR1 = ""
    Else
        EX_ADDR1 = Trim(c)
    End If
                                '���Ϗ� ���Ё@�Z���Q
    If GetIni("Estimate", "ADDR2", App.EXEName, c) Then
        EX_ADDR2 = ""
    Else
        EX_ADDR2 = Trim(c)
    End If
                                '���Ϗ� �Z���^�[   ����
    If GetIni("Estimate", "CENTER_NAME", App.EXEName, c) Then
        EX_CENTER_NAME = ""
    Else
        EX_CENTER_NAME = Trim(c)
    End If
                                '���Ϗ� �Z���^�[   �Z���P
    If GetIni("Estimate", "CENTER_ADDR1", App.EXEName, c) Then
        EX_CENTER_ADDR1 = ""
    Else
        EX_CENTER_ADDR1 = Trim(c)
    End If
                                '���Ϗ� �Z���^�[   �Z���Q
    If GetIni("Estimate", "CENTER_ADDR2", App.EXEName, c) Then
        EX_CENTER_ADDR2 = ""
    Else
        EX_CENTER_ADDR2 = Trim(c)
    End If
                                '���Ϗ� ���l�P
    If GetIni("Estimate", "BIKOU1", App.EXEName, c) Then
        EX_BIKOU1 = ""
    Else
        EX_BIKOU1 = Trim(c)
    End If
                                '���Ϗ� ���l�Q
    If GetIni("Estimate", "BIKOU2", App.EXEName, c) Then
        EX_BIKOU2 = ""
    Else
        EX_BIKOU2 = Trim(c)
    End If
                                '���ޑΏێ��
    If GetIni("Estimate", "EX_SHIZAI", App.EXEName, c) Then
        EX_SHIZAI_F = False
    Else
        EX_SHIZAI_F = True
        EX_SHIZAI_T = Split(Trim(c), ",", -1)
    End If
                                '�����Ώێ��
    If GetIni("Estimate", "EX_DOUKON", App.EXEName, c) Then
        EX_DOUKON_F = False
    Else
        EX_DOUKON_F = True
        EX_DOUKON_T = Split(Trim(c), ",", -1)
    End If

                                '�t����ƑΏێ��
    If GetIni("Estimate", "EX_FUKA", App.EXEName, c) Then
        EX_FUKA_F = False
    Else
        EX_FUKA_F = True
        EX_FUKA_T = Split(Trim(c), ",", -1)
    End If

                                '�ް�������ٺ���
    If GetIni("Estimate", "EX_BCR_CODE", App.EXEName, c) Then
        EX_BCR_CODE = ""
    Else
        EX_BCR_CODE = Trim(c)
    End If
    If GetIni("Estimate", "EXCEL_TEMPLATE", App.EXEName, c) Then
        EXCEL_TEMPLATE = ""
    Else
        EXCEL_TEMPLATE = Trim(c)
    End If
    If GetIni("Estimate", "INSERT_PIC", App.EXEName, c) Then
        Insert_Pic = ""
    Else
        Insert_Pic = Trim(c)
    End If
    If GetIni("Estimate", "SYONIN_PIC", App.EXEName, c) Then
        SYONIN_Pic = ""
    Else
        SYONIN_Pic = Trim(c)
    End If
'------------------------------------------------------ EXCEL�p����

                                '�q�Ƀ}�X�^�n�o�d�m
    If SOKO_Open(BtOpenRead) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i���J�e�S���}�X�^�n�o�d�m
    If ITEM_CATEGORY_Open(BtOpenRead) Then
        Unload Me
    End If
                                '�\���}�X�^�n�o�d�m
    If P_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '�\���}�X�^�n�o�d�m
    If wP_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                '�R�[�h�}�X�^�n�o�d�m
    If P_CODE_Open(BtOpenRead) Then
        Unload Me
    End If
                                '�S���҃}�X�^�n�o�d�m
    If TANTO_Open(BtOpenRead) Then
        Unload Me
    End If
                                '�����Ϗo�א�(���ʏW�v)�n�o�d�m
    If MONTHLYQTY_Open(BtOpenRead) Then
        Unload Me
    End If
                                
                                '���o�ɒP���}�X�^�n�o�d�m
    If SE_LOC_TANKA_M_Open(BtOpenRead) Then
        Unload Me
    End If
                                '�i�ڒP���ύX�����n�o�d�m
    If ITEM_HST_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                '�Ǘ��}�X�^�n�o�d�m
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If

    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)
    sts = BTRV(BtOpGetEqual, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    Select Case sts
        Case BtNoErr
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�Ǘ��}�X�^(KEY=01)")
        Unload Me
    End Select

    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_DEF_No)
    sts = BTRV(BtOpGetEqual, P_KANRI_POS, P_KANRIREC02, Len(P_KANRIREC02), K0_P_KANRI, Len(K0_P_KANRI), 0)
    Select Case sts
        Case BtNoErr
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�Ǘ��}�X�^(KEY=02)")
        Unload Me
    End Select
    
    
    
    If GetIni("Estimate", "Save_Dir", App.EXEName, c) Then
        Save_Dir = ""
    Else
        Save_Dir = Trim(c)
    End If



    If GetIni(App.EXEName, "SEI0018_LOG", App.EXEName, c) Then
        SEI0018_LOG = ""
    Else
        SEI0018_LOG = Trim(c)
    End If
    
    
    '�Ώ��ް��敪�捞��     2018.05.16
    Erase DATA_KBN_TBL
    Erase SYUBETSU_TBL
    
    If GetIni("Lump_SEIKYU", "SEL_DATA_KBN", "SEI_SYS", c) Then
        c = "*"
    End If
    wkVAL = Split(Trim(c), ",", -1)
    For i = 0 To UBound(wkVAL)
    
        ReDim Preserve DATA_KBN_TBL(0 To i)
        DATA_KBN_TBL(i) = wkVAL(i)
    
    Next i
    
    
    If GetIni("Lump_SEIKYU", "SEL_SYUBETSU", "SEI_SYS", c) Then
        c = "*"
    End If
    wkVAL = Split(Trim(c), ",", -1)
    For i = 0 To UBound(wkVAL)
    
        ReDim Preserve SYUBETSU_TBL(0 To i)
        SYUBETSU_TBL(i) = wkVAL(i)
    
    Next i
    
    
    
    
    
    
    '����Ͻ���`
    Call P_CODE_TBL_Proc
    
    '�d������̃Z�b�g
    If Code_Set_Proc(pcmbSHIMUKE, P_KBN04_CD, 0) Then
        Unload Me
    End If
    Combo1(pcmbSHIMUKE).ListIndex = 0


    '�i���ú�ذ�̃Z�b�g
    If ITEM_CATEGORY_Set_Proc() Then
        Unload Me
    End If


    '�d������̃Z�b�g
    If Code_Set_Proc(pcmbSHIMUKE, P_KBN04_CD, 0, 1) Then
        Unload Me
    End If
    Combo1(pcmbSHIMUKE).ListIndex = 0


    '��ʃZ�b�g
    If SYUBETSU_Set_Proc() Then
        Unload Me
    End If







    SEI00181.Caption = SEI00181.Caption & " " & LAST_UPDATE_DAY

    Call Init_Proc


    
    cmbSHIMUKE.ListIndex = 0
    
    
    
    
    
    txtTANTO_CODE.SetFocus


End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
Dim yn  As Integer
                                            
                                            
    yn = MsgBox("�I�����܂����H", vbYesNo, "�m�F����")
    If yn = vbNo Then
        Cancel = True
        Exit Sub
    End If
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '�\���}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�\���}�X�^")
        End If
    End If
                                            '�R�[�h�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�\���}�X�^")
        End If
    End If
                                            '�S���҃}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�S���҃}�X�^")
        End If
    End If
                                            '�Ǘ��}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�Ǘ��}�X�^")
        End If
    End If
    
    
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    End
End Sub

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------
Dim i   As Integer


    SEI00181.MousePointer = vbHourglass

    Call Ctrl_Lock(SEI00181)



End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------
Dim i   As Integer

    Call Ctrl_UnLock(SEI00181)


    SEI00181.MousePointer = vbDefault

End Sub


Private Sub SHORI_Click(Index As Integer)
    Select Case Index
        Case 0 To 5
            Command1(Index).Value = True

'        Case 6      '��ʈ��                                          2017.03.29
'                                                                       2017.03.29
'            Call Form_HCopy(Picture1, vbPRPSA4, vbPRORLandscape)       2017.03.29


    End Select
                    
    
    


End Sub






Private Function Init_Proc() As Integer
'----------------------------------------------------------------------------
'                   ��ʏ�����
'----------------------------------------------------------------------------
Dim i           As Integer

Dim Row         As Integer
Dim KOTEI_NO    As Integer

Dim c           As String * 128
                                
Dim wkKOTEI As Variant
                                
                                
                                
                                
                                
    Init_Proc = True
                                
                                
    If SYUBETSU_Set_Proc() Then
        Exit Function
    End If
                                
                                
                                
                                '��ƍH������荞��
'    Set SAGYO = Nothing
    
    
    
    
    
'    Text1(ptxDEF_LOT).Text = Format(StrConv(P_KANRIREC.KOUTEI_LOT, vbUnicode), "#0")
    
    
    
    
    Row = 0
    KOTEI_NO = 0
    For i = 1 To 10
        
        If GetIni("KOUTEI", "BEF" & Format(i, "00"), "SEI_SYS", c) Then
        Else
            If Trim(c) = "*" Then
            Else
                Row = Row + 1
'                SAGYO.ReDim S_Min_Row, Row, S_Min_Col, S_Max_Col
                
                KOTEI_NO = KOTEI_NO + 10
'                SAGYO(Row, ColS_No) = Format(KOTEI_NO, "00")
                
                
                
                wkKOTEI = Split(Trim(c), ",", -1)
                
'                SAGYO(Row, ColS_NAME) = Trim(wkKOTEI(0))
'                SAGYO(Row, ColS_KOUSU) = Trim(wkKOTEI(1))
                
                If IsNumeric(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)) Then
'                    SAGYO(Row, ColS_TANKA) = Format(SAGYO(Row, ColS_KOUSU) * CDbl(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)), "#,##0.00")
                Else
'                    SAGYO(Row, ColS_TANKA) = 0
                End If
                
                If IsNumeric(StrConv(P_KANRIREC.KOUTEI_LOT, vbUnicode)) Then
'                    SAGYO(Row, ColS_KIN) = Format(SAGYO(Row, ColS_TANKA) * CLng(StrConv(P_KANRIREC.KOUTEI_LOT, vbUnicode)), "#,##0")
                Else
'                    SAGYO(Row, ColS_KIN) = 0
                End If
            End If
        End If
    
    Next i
    
    For i = 1 To 10
        
        If GetIni("KOUTEI", "MAIN" & Format(i, "00"), "SEI_SYS", c) Then
        Else
            If Trim(c) = "*" Then
            Else
                Row = Row + 1
'                SAGYO.ReDim S_Min_Row, Row, S_Min_Col, S_Max_Col
                
                KOTEI_NO = KOTEI_NO + 10
'                SAGYO(Row, ColS_No) = Format(KOTEI_NO, "00")
                
                wkKOTEI = Split(Trim(c), ",", -1)
                
'                SAGYO(Row, ColS_NAME) = Trim(wkKOTEI(0))
'                SAGYO(Row, ColS_KOUSU) = Trim(wkKOTEI(1))
                
                If IsNumeric(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)) Then
'                    SAGYO(Row, ColS_TANKA) = Format(SAGYO(Row, ColS_KOUSU) * CDbl(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)), "#,##0.00")
                Else
'                    SAGYO(Row, ColS_TANKA) = 0
                End If
                If IsNumeric(StrConv(P_KANRIREC.KOUTEI_LOT, vbUnicode)) Then
'                    SAGYO(Row, ColS_KIN) = Format(SAGYO(Row, ColS_TANKA) * CLng(StrConv(P_KANRIREC.KOUTEI_LOT, vbUnicode)), "#,##0")
                Else
'                    SAGYO(Row, ColS_KIN) = 0
                End If
            End If
        End If
    
    Next i
                                
    For i = 1 To 10
        
        If GetIni("KOUTEI", "AFT" & Format(i, "00"), "SEI_SYS", c) Then
        Else
            If Trim(c) = "*" Then
            Else
                Row = Row + 1
'                SAGYO.ReDim S_Min_Row, Row, S_Min_Col, S_Max_Col
                
                KOTEI_NO = KOTEI_NO + 10
'                SAGYO(Row, ColS_No) = Format(KOTEI_NO, "00")
                
                wkKOTEI = Split(Trim(c), ",", -1)
                
'                SAGYO(Row, ColS_NAME) = Trim(wkKOTEI(0))
'                SAGYO(Row, ColS_KOUSU) = Trim(wkKOTEI(1))
                
                If IsNumeric(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)) Then
'                    SAGYO(Row, ColS_TANKA) = Format(SAGYO(Row, ColS_KOUSU) * CDbl(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)), "#,##0.00")
                Else
'                    SAGYO(Row, ColS_TANKA) = 0
                End If
                
                If IsNumeric(StrConv(P_KANRIREC.KOUTEI_LOT, vbUnicode)) Then
'                    SAGYO(Row, ColS_KIN) = Format(SAGYO(Row, ColS_TANKA) * CLng(StrConv(P_KANRIREC.KOUTEI_LOT, vbUnicode)), "#,##0")
                Else
'                    SAGYO(Row, ColS_KIN) = 0
                End If
            End If
        End If
    
    Next i
                                
                                
                                
'    Set TDBGrid1(pGrdSAGYO).Array = SAGYO
    
    
'    TDBGrid1(pGrdSAGYO).Bookmark = Null
    
'    TDBGrid1(pGrdSAGYO).ReBind
'    TDBGrid1(pGrdSAGYO).Update
'    TDBGrid1(pGrdSAGYO).ScrollBars = dbgAutomatic

    Init_Proc = True


End Function
Private Function ITEM_CATEGORY_Set_Proc() As Integer
'----------------------------------------------------------------------------
'                   �i���J�e�S���B�[�}�X�^���h���b�v�_�E�����X�g�ɃZ�b�g����B
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer



Dim i           As Integer
    
    ITEM_CATEGORY_Set_Proc = True
    
    Call UniCode_Conv(K0_ITEM_CATEGORY.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_ITEM_CATEGORY.CATEGORY_CODE, "")


    Combo1(pcmbCATEGORY_Name).Clear


    Combo1(pcmbCATEGORY_Name).AddItem "�Ȃ�" & Space(76) & Space(8)


    com = BtOpGetGreater

    Do
        DoEvents
    
        sts = BTRV(com, ITEM_CATEGORY_POS, ITEM_CATEGORYREC, Len(ITEM_CATEGORYREC), K0_ITEM_CATEGORY, Len(K0_ITEM_CATEGORY), 0)
            
        Select Case sts
            Case BtNoErr
            
                                
                If StrConv(ITEM_CATEGORYREC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Then

                    Exit Do

                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�i���J�e�S���}�X�^")
                Exit Function
        
        End Select

        
        Combo1(pcmbCATEGORY_Name).AddItem StrConv(ITEM_CATEGORYREC.CATEGORY_NAME, vbUnicode) & StrConv(ITEM_CATEGORYREC.CATEGORY_CODE, vbUnicode)
        
        com = BtOpGetNext
    
    Loop



    If Combo1(pcmbCATEGORY_Name).ListCount > 1 Then
        Combo1(pcmbCATEGORY_Name).ListIndex = 0
    End If

    ITEM_CATEGORY_Set_Proc = False
    



End Function
Private Function SYUBETSU_Set_Proc() As Integer
'----------------------------------------------------------------------------
'                   �R�[�h�}�X�^���h���b�v�_�E�����X�g�ɃZ�b�g����B
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer



Dim i           As Integer
    
    SYUBETSU_Set_Proc = True
    
    Set SYUBETSU = Nothing
    
    
    
    
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN06_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, "")

    com = BtOpGetGreater

    i = 0
    Do
        DoEvents
    
        sts = BTRV(com, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            
        Select Case sts
            Case BtNoErr
            
                                
                If StrConv(P_CODEREC.DATA_KBN, vbUnicode) <> P_KBN06_CD Then
                    
                    Exit Do
                
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�R�[�h�}�X�^")
                Exit Function
        
        End Select

        
        i = i + 1
        SYUBETSU.ReDim 1, i, 0, 0
        
        
        SYUBETSU(i, 0) = Trim(StrConv(P_CODEREC.C_RNAME, vbUnicode)) & "            " & _
                                Left(StrConv(P_CODEREC.C_Code, vbUnicode), 2)
        
        
        com = BtOpGetNext
    
    Loop

    Set TDBDropDown1.Array = SYUBETSU
    TDBDropDown1.ReBind

    SYUBETSU_Set_Proc = False
    



End Function



Private Sub TDBGrid1_AfterColUpdate(Index As Integer, ByVal ColIndex As Integer)

Dim sts         As Integer
Dim Bookmark    As Variant
    
    
Dim i           As Integer
    
    
Dim wkDouble    As Double
    
    
    Set TDBGrid1(pGrdKOUSEI).Array = KOUSEI
    TDBGrid1(pGrdKOUSEI).Update
    
    
    
    If TDBGrid1(pGrdKOUSEI).Bookmark = Null Then
        Exit Sub
    End If
    
    If TDBGrid1(pGrdKOUSEI).Bookmark <= 0 Then
        Exit Sub
    End If
    
                    
    Select Case ColIndex
        
        Case ColKO_HIN_GAI
        
            If Trim(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_HIN_GAI)) = "" Then
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_JGYOBU) = ""
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_NAIGAI) = ""
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_HIN_NAME) = ""
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY) = ""
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN) = ""
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = ""
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHIKIN) = ""
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = ""
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColS_KOUSU) = ""
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColSEI_SYU_KON) = ""
            
            
            
            Else
                '�i��
                If Trim(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_JGYOBU)) = "" And _
                    Trim(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_NAIGAI)) = "" Then
                    Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
                Else
                    Call UniCode_Conv(K0_ITEM.JGYOBU, KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_JGYOBU))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_NAIGAI))
                End If
                
                '2013.01.17
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_HIN_GAI) = StrConv(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_HIN_GAI), vbUpperCase)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_HIN_GAI))
            
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        '���ޕi�œǂݑւ�
                        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_HIN_GAI))
                        
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                                
                                If HIN_INV Then
                                    '���o�^�i�ԁ@�@���ނƂ��Ă���
                                    Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)
                                    Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
                                    Call UniCode_Conv(ITEMREC.HIN_NAME, "���o�^�i��")
                                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                                Else
                                    MsgBox "[" & Format(TDBGrid1(pGrdKOUSEI).Bookmark, "0") & "]�s�� ���͂������ڂ̓G���[�ł��B(�i��)"
                                    Exit Sub
                                End If
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                                Unload Me
                        
                        End Select
                
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                        Unload Me
                
                End Select
            
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_JGYOBU) = StrConv(ITEMREC.JGYOBU, vbUnicode)
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_NAIGAI) = StrConv(ITEMREC.NAIGAI, vbUnicode)
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_HIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
            
            
                '����
                If Trim(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) = "" Then
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY) = "1.00"
                End If
                If IsNumeric(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) Then
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY) = Format(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)), "#0.00")
                Else
                    MsgBox "[" & Format(TDBGrid1(pGrdKOUSEI).Bookmark, "0") & "]�s�� ���͂������ڂ̓G���[�ł��B(����)"
                    Exit Sub
                End If
            
            
                '�d���� >>>>>>>>>>>>>>  ��ɍŐV�̕i�ړ��e����荞��    2013.04.01
                'If Trim(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN)) = "" Then
                    If IsNumeric(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) Then
                        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN) = Format(CDbl(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)), "#0.00")
                    Else
                        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN) = "0.00"
                    End If
                'Else
                '    If IsNumeric(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN)) Then
                '        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN) = Format(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN)), "#0.00")
                '    Else
                '        MsgBox "[" & Format(TDBGrid1(pGrdKOUSEI).Bookmark, "0") & "]�s�� ���͂������ڂ̓G���[�ł��B(�d����)"
                '        Exit Sub
                '
                '    End If
                'End If
                '�d���� >>>>>>>>>>>>>>  ��ɍŐV�̕i�ړ��e����荞��    2013.04.01
                
                '�d�����z�v
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHIKIN) = 0
            
                For i = 0 To UBound(SHIZAI_T)
                    If Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU), 2) = SHIZAI_T(i) Then
                        
                        
        '                If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
                        If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
                            
                            
                            If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
                                
                                If CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) = 0 Then      '2010.02.22
                                
                                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHIKIN) = "0.00"
                                Else
                                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN)) / CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY))), 2), "#,##0.00")
                                End If
                            Else
                                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) * CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN))), 2), "#,##0.00")
                            End If
                            
                        End If
                        Exit For
                    End If
                
                Next i
                If CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHIKIN)) = 0 Then
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHIKIN) = ""
                End If
            
                '�̔���
                
                Select Case StrConv(ITEMREC.SEI_KBN, vbUnicode)
                
                
                    Case "1"
                        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = "�ʔ�"
                    Case "2"
                        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = "�x��"
                    Case Else
                        ' >>>>>>>>>>>>>>  ��ɍŐV�̕i�ړ��e����荞��    2013.04.01
                        'If Trim(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN)) = "" Then
                            If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), "#0.00")
                            Else
                                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = "0.00"
                            End If
                        'Else
                        '    If IsNumeric(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN)) Then
                        '        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = Format(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN)), "#0.00")
                        '    Else
                        '        MsgBox "[" & Format(TDBGrid1(pGrdKOUSEI).Bookmark, "0") & "]�s�� ���͂������ڂ̓G���[�ł��B(�̔���)"
                        '        Exit Sub
                        '    End If
                        'End If
                        ' >>>>>>>>>>>>>>  ��ɍŐV�̕i�ړ��e����荞��    2013.04.01
                
                End Select
                    
                '������z�v
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = 0
            
                For i = 0 To UBound(SHIZAI_T)
                
                    If Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU), 2) = SHIZAI_T(i) Then
        '                If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
                        If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
                            If Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU), 2) = GAISO_KBN Then
                                
                                If CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) = 0 Then
                                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = "0.00"
                                Else
                                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN)) / CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY))), 2), "#,##0.00")
                                End If
                            Else
                                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) * CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN))), 2), "#,##0.00")
                            End If
                        End If
                    Else
                    
                        If KUSATU_F Then
        '                    If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
                            If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
                    
                                If Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU), 2) = GAISO_KBN Then
                                
                                    If CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) = 0 Then      '2010.02.22
                                    
                                        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN_KUSATU) = "0.00"
                                    Else
                                        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN_KUSATU) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN)) / CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY))), 2), "#,##0.00")
                                    End If
                                Else
                                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN_KUSATU) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) * CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN))), 2), "#,##0.00")
                                End If
                            End If
                        
                        End If
                        
                    End If
                
                Next i
                
                If CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN)) = 0 Then
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = ""
                End If
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> ���ʒP���ł̏���
'                If Trim(Text1(ptxCATEGORY_CODE).Text) = "" Then
'                Else
'                    '��ƍH���@�b/��
'                    For i = 0 To UBound(SP_KOUSU_T)
'                        If SP_KOUSU_T(i) = Trim(Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU), 2)) Then
'                            Call UniCode_Conv(K0_ITEM_CATEGORY.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
'                            Call UniCode_Conv(K0_ITEM_CATEGORY.CATEGORY_CODE, Text1(ptxCATEGORY_CODE).Text)
'                            sts = BTRV(BtOpGetEqual, ITEM_CATEGORY_POS, ITEM_CATEGORYREC, Len(ITEM_CATEGORYREC), K0_ITEM_CATEGORY, Len(K0_ITEM_CATEGORY), 0)
'                            Select Case sts
'                                Case BtNoErr
'                                    If IsNumeric(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_QTY, vbUnicode)) Then
'                                        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_SPTAN) = Val(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_QTY, vbUnicode))
'                                    Else
'                                        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_SPTAN) = "0"
'                                    End If
'                                Case BtErrKeyNotFound
'                                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_SPTAN) = "0"
'                                Case Else
'                                    Call File_Error(sts, BtOpGetEqual, "�i���J�e�S���}�X�^")
'                                    Unload Me
'                            End Select
'                            KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = Val(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_SPTAN))
'                        End If
'                    Next i
'                    '�H����
'                    For i = 0 To UBound(SP_KOURYO_T)
'                        If SP_KOURYO_T(i) = Trim(Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU), 2)) Then
'                            Call UniCode_Conv(K0_ITEM_CATEGORY.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
'                            Call UniCode_Conv(K0_ITEM_CATEGORY.CATEGORY_CODE, Text1(ptxCATEGORY_CODE).Text)
'                            sts = BTRV(BtOpGetEqual, ITEM_CATEGORY_POS, ITEM_CATEGORYREC, Len(ITEM_CATEGORYREC), K0_ITEM_CATEGORY, Len(K0_ITEM_CATEGORY), 0)
'                            Select Case sts
'                                Case BtNoErr
'                                    If IsNumeric(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_KOURYO, vbUnicode)) Then
'                                        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_SPTAN) = Val(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_KOURYO, vbUnicode))
'                                    Else
'                                        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_SPTAN) = "0"
'                                    End If
'                                Case BtErrKeyNotFound
'                                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_SPTAN) = "0"
'                                Case Else
'                                    Call File_Error(sts, BtOpGetEqual, "�i���J�e�S���}�X�^")
'                                    Unload Me
'                            End Select
'                            KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = Val(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_SPTAN))
'                        End If
'                    Next i
'                    '���し
'                    For i = 0 To UBound(SP_HAKO_T)
'                        If SP_HAKO_T(i) = Trim(Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU), 2)) Then
'                            Call UniCode_Conv(K0_ITEM_CATEGORY.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
'                            Call UniCode_Conv(K0_ITEM_CATEGORY.CATEGORY_CODE, Text1(ptxCATEGORY_CODE).Text)
'                            sts = BTRV(BtOpGetEqual, ITEM_CATEGORY_POS, ITEM_CATEGORYREC, Len(ITEM_CATEGORYREC), K0_ITEM_CATEGORY, Len(K0_ITEM_CATEGORY), 0)
'                            Select Case sts
'                                Case BtNoErr
'                                    If IsNumeric(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_HAKO, vbUnicode)) Then
'                                        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_SPTAN) = Val(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_HAKO, vbUnicode))
'                                    Else
'                                        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_SPTAN) = "0"
'                                    End If
'                                Case BtErrKeyNotFound
'                                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_SPTAN) = "0"
'                                Case Else
'                                    Call File_Error(sts, BtOpGetEqual, "�i���J�e�S���}�X�^")
'                                    Unload Me
'                            End Select
'                            KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = Val(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_SPTAN))
'                        End If
'                    Next i
'                End If
        
                If Not IsEmpty(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_SPTAN)) Then          '2013.04.01
                    If IsNumeric(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_SPTAN)) Then
                        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = Val(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_SPTAN))
                    End If
                End If
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                
                
                
                
                
                
                
                
                '��Ǝ���
                If Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU), 2) = GAISO_KBN Then
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColS_KOUSU) = ""
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColSEI_SYU_KON) = ""
                Else
                
                    ' >>>>>>>>>>>>>>  ��ɍŐV�̕i�ړ��e����荞��    2013.04.01
                    'If Trim(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColS_KOUSU)) = "" Then
                        If IsNumeric(StrConv(ITEMREC.S_KOUSU, vbUnicode)) Then
                            KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColS_KOUSU) = Format(CDbl(StrConv(ITEMREC.S_KOUSU, vbUnicode)), "#")
                        Else
                            KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColS_KOUSU) = ""
                        End If
                    'Else
                    '    If IsNumeric(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColS_KOUSU)) Then
                    '        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColS_KOUSU) = Format(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColS_KOUSU)), "#")
                    '    Else
                    '        MsgBox "[" & Format(TDBGrid1(pGrdKOUSEI).Bookmark, "0") & "]�s�� ���͂������ڂ̓G���[�ł��B(��Ǝ���)"
                    '    End If
                    'End If
                    ' >>>>>>>>>>>>>>  ��ɍŐV�̕i�ړ��e����荞��    2013.04.01
                    
                    '�W�������
                    
                    ' >>>>>>>>>>>>>>  ��ɍŐV�̕i�ړ��e����荞��    2013.04.01
                    'If Trim(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColSEI_SYU_KON)) = "" Then
                        If IsNumeric(StrConv(ITEMREC.SEI_SYU_KON, vbUnicode)) Then
                            KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColSEI_SYU_KON) = Format(CDbl(StrConv(ITEMREC.SEI_SYU_KON, vbUnicode)), "#")
                        Else
                            KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColSEI_SYU_KON) = ""
                        End If
                    'Else
                    '    If IsNumeric(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColSEI_SYU_KON)) Then
                    '        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColSEI_SYU_KON) = Format(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColSEI_SYU_KON)), "#")
                    '    Else
                    '        MsgBox "[" & Format(TDBGrid1(pGrdKOUSEI).Bookmark, "0") & "]�s�� ���͂������ڂ̓G���[�ł��B(�W�������)"
                    '    End If
                    'End If
                    ' >>>>>>>>>>>>>>  ��ɍŐV�̕i�ړ��e����荞��    2013.04.01
            
                End If
            End If
                
            Set TDBGrid1(pGrdKOUSEI).Array = KOUSEI
                
            
            TDBGrid1(pGrdKOUSEI).Refresh
            TDBGrid1(pGrdKOUSEI).Update
        '    TDBGrid1.ScrollBars = dbgAutomatic
            
            TDBGrid1(pGrdKOUSEI).SetFocus
        
        
        
        
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   '���� 2017.01.14
        Case ColKO_QTY
            
            
            
            If Trim(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) = "" Then
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY) = "1.00"
            End If
            If IsNumeric(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) Then
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY) = Format(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)), "#0.00")
            Else
                MsgBox "[" & Format(TDBGrid1(pGrdKOUSEI).Bookmark, "0") & "]�s�� ���͂������ڂ̓G���[�ł��B(����)"
                Exit Sub
            End If
            
            
            
            '�i��
            If Trim(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_JGYOBU)) = "" And _
                Trim(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_NAIGAI)) = "" Then
                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
            Else
                Call UniCode_Conv(K0_ITEM.JGYOBU, KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_JGYOBU))
                Call UniCode_Conv(K0_ITEM.NAIGAI, KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_NAIGAI))
            End If
            
            '2013.01.17
            KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_HIN_GAI) = StrConv(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_HIN_GAI), vbUpperCase)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_HIN_GAI))
        
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    '���ޕi�œǂݑւ�
                    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_HIN_GAI))
                    
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                            
                            If HIN_INV Then
                                '���o�^�i�ԁ@�@���ނƂ��Ă���
                                Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)
                                Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
                                Call UniCode_Conv(ITEMREC.HIN_NAME, "���o�^�i��")
                                Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                            Else
                                MsgBox "[" & Format(TDBGrid1(pGrdKOUSEI).Bookmark, "0") & "]�s�� ���͂������ڂ̓G���[�ł��B(�i��)"
                                Exit Sub
                            End If
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                            Unload Me
                    
                    End Select
            
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                    Unload Me
            
            End Select



            '�d�����z�v
            KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHIKIN) = 0
        
            For i = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU), 2) = SHIZAI_T(i) Then
                    
                    
    '                If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
                    If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
                        
                        
                        If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
                            
                            If CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) = 0 Then      '2010.02.22
                            
                                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHIKIN) = "0.00"
                            Else
                                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN)) / CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY))), 2), "#,##0.00")
                            End If
                        Else
                            KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) * CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN))), 2), "#,##0.00")
                        End If
                        
                    End If
                    Exit For
                End If
            
            Next i
            If CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHIKIN)) = 0 Then
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHIKIN) = ""
            End If

            '�̔���
            Select Case StrConv(ITEMREC.SEI_KBN, vbUnicode)
            
            
                Case "1"
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = "�ʔ�"
                Case "2"
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = "�x��"
                Case Else
                    ' >>>>>>>>>>>>>>  ��ɍŐV�̕i�ړ��e����荞��    2013.04.01
                    If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), "#0.00")
                    Else
                        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = "0.00"
                    End If
                    ' >>>>>>>>>>>>>>  ��ɍŐV�̕i�ړ��e����荞��    2013.04.01
            End Select

            '������z�v
            KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = 0
        
            For i = 0 To UBound(SHIZAI_T)
            
                If Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU), 2) = SHIZAI_T(i) Then
    '                If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
                    If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
                        If Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU), 2) = GAISO_KBN Then
                            
                            If CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) = 0 Then
                                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = "0.00"
                            Else
                                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN)) / CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY))), 2), "#,##0.00")
                            End If
                        Else
                            KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) * CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN))), 2), "#,##0.00")
                        End If
                    End If
                Else
                
                    If KUSATU_F Then
    '                    If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
                        If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
                
                            If Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU), 2) = GAISO_KBN Then
                            
                                If CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) = 0 Then      '2010.02.22
                                
                                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN_KUSATU) = "0.00"
                                Else
                                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN_KUSATU) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN)) / CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY))), 2), "#,##0.00")
                                End If
                            Else
                                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN_KUSATU) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) * CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN))), 2), "#,##0.00")
                            End If
                        End If
                    
                    End If
                    
                End If
            
            Next i
            
            If CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN)) = 0 Then
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = ""
            End If


            If Not IsEmpty(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_SPTAN)) Then          '2013.04.01
                If IsNumeric(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_SPTAN)) Then
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = Val(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_SPTAN))
                End If
            End If



            '��Ǝ���
            If Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU), 2) = GAISO_KBN Then
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColS_KOUSU) = ""
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColSEI_SYU_KON) = ""
            Else
            
            ' >>>>>>>>>>>>>>  ��ɍŐV�̕i�ړ��e����荞��    2013.04.01
                If IsNumeric(StrConv(ITEMREC.S_KOUSU, vbUnicode)) Then
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColS_KOUSU) = Format(CDbl(StrConv(ITEMREC.S_KOUSU, vbUnicode)), "#")
                Else
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColS_KOUSU) = ""
                End If
            
            '�W�������
            
            ' >>>>>>>>>>>>>>  ��ɍŐV�̕i�ړ��e����荞��    2013.04.01
                If IsNumeric(StrConv(ITEMREC.SEI_SYU_KON, vbUnicode)) Then
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColSEI_SYU_KON) = Format(CDbl(StrConv(ITEMREC.SEI_SYU_KON, vbUnicode)), "#")
                Else
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColSEI_SYU_KON) = ""
                End If
            ' >>>>>>>>>>>>>>  ��ɍŐV�̕i�ړ��e����荞��    2013.04.01
        
            End If


            Set TDBGrid1(pGrdKOUSEI).Array = KOUSEI





            TDBGrid1(pGrdKOUSEI).Refresh
            TDBGrid1(pGrdKOUSEI).Update
            
            TDBGrid1(pGrdKOUSEI).SetFocus

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   '���� 2017.01.14
        
        Case ColG_SPTAN



'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2017.01.14
            '�i��
            If Trim(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_JGYOBU)) = "" And _
                Trim(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_NAIGAI)) = "" Then
                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
            Else
                Call UniCode_Conv(K0_ITEM.JGYOBU, KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_JGYOBU))
                Call UniCode_Conv(K0_ITEM.NAIGAI, KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_NAIGAI))
            End If
            
            '2013.01.17
            KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_HIN_GAI) = StrConv(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_HIN_GAI), vbUpperCase)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_HIN_GAI))
        
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    '���ޕi�œǂݑւ�
                    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_HIN_GAI))
                    
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                            
                            If HIN_INV Then
                                '���o�^�i�ԁ@�@���ނƂ��Ă���
                                Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)
                                Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
                                Call UniCode_Conv(ITEMREC.HIN_NAME, "���o�^�i��")
                                Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                            Else
                                MsgBox "[" & Format(TDBGrid1(pGrdKOUSEI).Bookmark, "0") & "]�s�� ���͂������ڂ̓G���[�ł��B(�i��)"
                                Exit Sub
                            End If
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                            Unload Me
                    
                    End Select
            
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                    Unload Me
            
            End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2017.01.14

            If KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_SPTAN) = "" Then
                
                Select Case StrConv(ITEMREC.SEI_KBN, vbUnicode)
                
                
                    Case "1"
                        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = "�ʔ�"
                    Case "2"
                        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = "�x��"
                    Case Else
                        If Trim(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN)) = "" Then
                            If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), "#0.00")
                            Else
                                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = "0.00"
                            End If
                        Else
                            If IsNumeric(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN)) Then
                                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = Format(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN)), "#0.00")
                            Else
                                MsgBox "[" & Format(TDBGrid1(pGrdKOUSEI).Bookmark, "0") & "]�s�� ���͂������ڂ̓G���[�ł��B(�̔���)"
                                Exit Sub
                            End If
                        End If
                
                End Select
                    
                '������z�v
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = 0
            
                For i = 0 To UBound(SHIZAI_T)
                
                    If Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU), 2) = SHIZAI_T(i) Then
        '                If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
                        If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
                            If Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU), 2) = GAISO_KBN Then
                                
                                If CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) = 0 Then
                                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = "0.00"
                                Else
                                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN)) / CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY))), 2), "#,##0.00")
                                End If
                            Else
                                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) * CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN))), 2), "#,##0.00")
                            End If
                        End If
                    Else
                    
                        If KUSATU_F Then
        '                    If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
                            If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
                    
                                If Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU), 2) = GAISO_KBN Then
                                
                                    If CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) = 0 Then      '2010.02.22
                                    
                                        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN_KUSATU) = "0.00"
                                    Else
                                        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN_KUSATU) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN)) / CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY))), 2), "#,##0.00")
                                    End If
                                Else
                                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN_KUSATU) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) * CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN))), 2), "#,##0.00")
                                End If
                            End If
                        
                        End If
                        
                    End If
                
                Next i
                
                If CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN)) = 0 Then
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = ""
                End If

            Else
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = Val(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_SPTAN))
            
            
                '������z�v
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = 0  '2013.04.01
Debug.Print StrConv(ITEMREC.HIN_GAI, vbUnicode)
                For i = 0 To UBound(SHIZAI_T)
                
                    If Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU), 2) = SHIZAI_T(i) Then
        '                If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
                        If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
                            If Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU), 2) = GAISO_KBN Then
                                
                                If CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) = 0 Then
                                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = "0.00"
                                Else
                                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN)) / CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY))), 2), "#,##0.00")
                                End If
                            Else
                                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) * CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN))), 2), "#,##0.00")
                            End If
                        End If
                    Else
                    
                        If KUSATU_F Then
        '                    If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
                            If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
                    
                                If Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU), 2) = GAISO_KBN Then
                                
                                    If CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) = 0 Then      '2010.02.22
                                    
                                        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN_KUSATU) = "0.00"
                                    Else
                                        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN_KUSATU) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN)) / CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY))), 2), "#,##0.00")
                                    End If
                                Else
                                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN_KUSATU) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) * CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN))), 2), "#,##0.00")
                                End If
                            End If
                        
                        End If
                        
                    End If
                
                Next i
                
                If CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN)) = 0 Then
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = ""
                End If
            
            
            End If
                
            Set TDBGrid1(pGrdKOUSEI).Array = KOUSEI
            
        
            TDBGrid1(pGrdKOUSEI).Refresh
            TDBGrid1(pGrdKOUSEI).Update
        
            TDBGrid1(pGrdKOUSEI).SetFocus


    End Select
End Sub


Private Sub TDBGrid1_BeforeInsert(Index As Integer, Cancel As Integer)
    
    KOUSEI.ReDim Min_Row, KOUSEI.Count(1), Min_Col, Max_Col

End Sub

Private Sub Text1_Change(Index As Integer)
Dim i   As Integer
    
    
    Select Case Index
        Case ptxHin_Gai
            If Trim(svHin_Gai) = (Text1(ptxHin_Gai).Text) Then
            Else
                For i = 2 To 5
                    Command1(i).Enabled = False
                Next i
            
            
'2018.04.02                Text1(ptxMAIN_KOUTEI_QTY01).Text = ""
            
            
            End If
    
    
    
    
    End Select



End Sub

Private Sub Text1_GotFocus(Index As Integer)
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If


    
    Select Case Index
        Case ptxHin_Gai
            svHin_Gai = Text1(ptxHin_Gai).Text
        Case ptxCATEGORY_CODE
            svCATEGORY_CODE = Text1(ptxCATEGORY_CODE).Text
    End Select



End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If

    If Error_Check_Proc(Index) Then   '�G���[�`�F�b�N
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '�ړ�
End Sub
Private Function Error_Check_Proc(Mode As Integer) As Integer
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
        
Dim INV_F       As Boolean
    
    
    Error_Check_Proc = True
    
    Select Case Mode
    
    
        Case ptxTanto_Code     '�S����
        
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxTanto_Code).Text)

            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                    Text1(ptxTanto_Name).Text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    Text1(ptxTanto_Name).Text = ""
            
Call LOG_OUT(SEI0018_LOG, "�S���҃G���[= " & Text1(ptxTanto_Code).Text)
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
                    Exit Function
                
            
            
            End Select
        Case ptxHin_Gai         '�i��
    
            
            Text1(ptxHin_Gai).Text = Trim(StrConv(Text1(ptxHin_Gai).Text, vbUpperCase))
            
            Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
            Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHin_Gai).Text)


            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                
                    Text1(ptxHin_Name).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                
                    If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                        Text1(ptxST_SOKO).Text = ""
                        Text1(ptxST_RETU).Text = ""
                        Text1(ptxST_REN).Text = ""
                        Text1(ptxST_DAN).Text = ""
                    Else
                        Text1(ptxST_SOKO).Text = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                        Text1(ptxST_RETU).Text = StrConv(ITEMREC.ST_RETU, vbUnicode)
                        Text1(ptxST_REN).Text = StrConv(ITEMREC.ST_REN, vbUnicode)
                        Text1(ptxST_DAN).Text = StrConv(ITEMREC.ST_DAN, vbUnicode)
                    End If
                
                
                Case BtErrKeyNotFound

                    Text1(ptxHin_Name).Text = ""

Call LOG_OUT(SEI0018_LOG, "�i�Ԗ��o�^�G���[= " & Text1(ptxHin_Gai).Text)
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                    Exit Function

            End Select
        
        
        
'>>>>>>>>>>>    2018.05.25  �����敪�̃`�F�b�N
            If Trim(txtBUHIN) <> "" Then
                If txtBUHIN <> StrConv(ITEMREC.NAI_BUHIN, vbUnicode) And txtBUHIN <> StrConv(ITEMREC.GAI_BUHIN, vbUnicode) Then
                
Call LOG_OUT(SEI0018_LOG, "�����敪�G���[ HIN_GAI= " & Text1(ptxHin_Gai).Text & "�������敪=" & StrConv(ITEMREC.NAI_BUHIN, vbUnicode) & "�O�����敪=" & StrConv(ITEMREC.GAI_BUHIN, vbUnicode))
                    Exit Function
                End If
            End If






'>>>>>>>>>>>    2018.05.25
        
        
        
'>>>>>>>>>>>    2018.06.11  �i�ڃJ�e�S���[�̃`�F�b�N
            If Trim(StrConv(ITEMREC.CATEGORY_CODE, vbUnicode)) = "" Then
                
Call LOG_OUT(SEI0018_LOG, "�i�ڃJ�e�S���[�G���[ HIN_GAI= " & Text1(ptxHin_Gai).Text & "�i�ڃJ�e�S���[<��>")
                Exit Function
            End If
'>>>>>>>>>>>    2018.06.11
        
        
        
        
        
        
        
            INV_F = False
            Call UniCode_Conv(K0_SOKO.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
            Select Case sts
                Case BtNoErr
        
                    Call UniCode_Conv(K0_SE_LOC_TANKA_M.SE_IO_TANKA_No, StrConv(SOKOREC.IO_TANKA_No, vbUnicode))
                    sts = BTRV(BtOpGetEqual, SE_LOC_TANKA_M_POS, SE_LOC_TANKA_M_REC, Len(SE_LOC_TANKA_M_REC), K0_SE_LOC_TANKA_M, Len(K0_SE_LOC_TANKA_M), 0)
                    Select Case sts
                        Case BtNoErr
                        
                            Text1(ptxIO_TANKA_No).Text = StrConv(SE_LOC_TANKA_M_REC.SE_IO_TANKA_No, vbUnicode)
                            Text1(ptxSE_Name).Text = StrConv(SE_LOC_TANKA_M_REC.SE_Name, vbUnicode)
                        
                        
                        Case BtErrKeyNotFound
                
                            INV_F = True
                
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "���o�ɒP���ݒ�}�X�^")
                        Exit Function
                    End Select
        
                Case BtErrKeyNotFound
        
                    INV_F = True
        
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^")
                    Exit Function
    
            End Select
    
    
            If INV_F Then
                
                Call UniCode_Conv(K0_SE_LOC_TANKA_M.SE_IO_TANKA_No, INV_IO_TANKA_No)
                sts = BTRV(BtOpGetEqual, SE_LOC_TANKA_M_POS, SE_LOC_TANKA_M_REC, Len(SE_LOC_TANKA_M_REC), K0_SE_LOC_TANKA_M, Len(K0_SE_LOC_TANKA_M), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                    
                            Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_Name, "")
                    
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "���o�ɒP���ݒ�}�X�^")
                        Exit Function
                End Select
            
            
                Text1(ptxIO_TANKA_No).Text = INV_IO_TANKA_No
                Text1(ptxSE_Name).Text = ""
            
            End If
        
        
        
        
        Case ptxCATEGORY_CODE               ' �i���ú�ذ����
        
            For i = 0 To Combo1(pcmbCATEGORY_Name).ListCount - 1
                If Trim(Text1(Mode).Text) = Trim(Right(Combo1(pcmbCATEGORY_Name).List(i), 8)) Then
                    Combo1(pcmbCATEGORY_Name).ListIndex = i
                    Exit For
                End If
            Next i
            If i > Combo1(pcmbCATEGORY_Name).ListCount - 1 Then
'                MsgBox "���͂������ڂ̓G���[�ł��B(�i���J�e�S���[�@���o�^)"
                
Call LOG_OUT(SEI0018_LOG, "�i���J�e�S���[���o�^�G���[= " & Text1(ptxHin_Gai).Text)
                
                
                Text1(Mode).SetFocus
                Exit Function
            End If


            If svCATEGORY_CODE = Trim(Text1(Mode).Text) Then
            
                If IsNumeric(Text1(ptxCATE_AD_KOURYO).Text) Then
                    Text1(ptxAFT_S_KOUSU_BAIKA).Text = Format(CDbl(Text1(ptxCATE_AD_KOURYO).Text), "#0.00")
                Else
                    Text1(ptxAFT_S_KOUSU_BAIKA).Text = ""
                End If
            
            
            Else
                If CATEGORY_Disp_Proc() Then
                    Exit Function
                End If
            End If
            
        Case ptxOLD_S_BU_KAKO_KOSU          ' ��  BU���H�P��
        
        
        
        
        
        
        
        
        
        
        
        Case ptxOLD_S_KOUSU_BAIKA           '��(����)���i���H��
        
        
            If Text1(ptxOLD_S_KOUSU_BAIKA).Text = "" Then
                Text1(ptxOLD_S_KOUSU_BAIKA).Text = "0.00"
            End If
            
            If Not IsNumeric(Text1(ptxOLD_S_KOUSU_BAIKA).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B(�H������)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxOLD_S_KOUSU_BAIKA).Text = Format(CDbl(Text1(ptxOLD_S_KOUSU_BAIKA).Text), "#0.00")
            End If
        
        
        
        
        
        
        
        
        
        
        
        
        
        Case ptxOLD_S_SHIZAI_BAIKA          '��(����)����

            If Text1(ptxOLD_S_SHIZAI_BAIKA).Text = "" Then
                Text1(ptxOLD_S_SHIZAI_BAIKA).Text = "0.00"
            End If
            
            If Not IsNumeric(Text1(ptxOLD_S_SHIZAI_BAIKA).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B(���ޔ���)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxOLD_S_SHIZAI_BAIKA).Text = Format(CDbl(Text1(ptxOLD_S_SHIZAI_BAIKA).Text), "#0.00")
            End If


        Case ptxOLD_S_GAISO_TANKA           '���O���P��
        
        
            If Text1(ptxOLD_S_GAISO_TANKA).Text = "" Then
                Text1(ptxOLD_S_GAISO_TANKA).Text = "0.00"
            End If
            
            If Not IsNumeric(Text1(ptxOLD_S_GAISO_TANKA).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B(�O���P��)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxOLD_S_GAISO_TANKA).Text = Format(CDbl(Text1(ptxOLD_S_GAISO_TANKA).Text), "#0.00")
            End If
        
        
        
        
        
        Case ptxOLD_S_PPSC_KAKO_KOSU        '��PPSC���H�P��
            
            If Text1(ptxOLD_S_PPSC_KAKO_KOSU).Text = "" Then
                Text1(ptxOLD_S_PPSC_KAKO_KOSU).Text = "0.00"
            End If
            
            If Not IsNumeric(Text1(ptxOLD_S_PPSC_KAKO_KOSU).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B(PPSC���H�P��)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxOLD_S_PPSC_KAKO_KOSU).Text = Format(CDbl(Text1(ptxOLD_S_PPSC_KAKO_KOSU).Text), "#0.00")
            End If
        
        Case ptxOLD_S_BU_KAKO_KOSU          '��BU���H�P��
    
            If Text1(ptxOLD_S_BU_KAKO_KOSU).Text = "" Then
                Text1(ptxOLD_S_BU_KAKO_KOSU).Text = "0.00"
            End If
            
            If Not IsNumeric(Text1(ptxOLD_S_BU_KAKO_KOSU).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B(PPSC���H�P��)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxOLD_S_BU_KAKO_KOSU).Text = Format(CDbl(Text1(ptxOLD_S_BU_KAKO_KOSU).Text), "#0.00")
            End If
        
        
        Case ptxBEF_SEI_LOT                 '�ύX�O�@   ���b�g��
        
            If Text1(ptxBEF_SEI_LOT).Text = "" Then
'Call LOG_OUT(SEI0018_LOG, Text1(ptxHin_Gai).Text & " ���b�g�G���[= " & Text1(ptxBEF_SEI_LOT).Text)
'Exit Function
            
            Else
            
                If Not IsNumeric(Text1(ptxBEF_SEI_LOT).Text) Then
                    
'Call LOG_OUT(SEI0018_LOG, Text1(ptxHin_Gai).Text & " ���b�g�G���[= " & Text1(ptxBEF_SEI_LOT).Text)
'Exit Function
                    
                    
                    MsgBox "���͂������ڂ̓G���[�ł��B(���b�g��)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxBEF_SEI_LOT).Text = Format(CLng(Text1(ptxBEF_SEI_LOT).Text), "#0")
                End If
            
            End If
        
        Case ptxBEF_SEI_RATE                '           �����[�g
        
            If Text1(ptxBEF_SEI_RATE).Text = "" Then
            
            Else
            
            
                If Not IsNumeric(Text1(ptxBEF_SEI_RATE).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(�����[�g)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxBEF_SEI_RATE).Text = Format(CLng(Text1(ptxBEF_SEI_RATE).Text), "#0.00")
                End If
            End If
        
        
        Case ptxBEF_S_KOUSU                 '           �����[�g
        
        
            If Text1(ptxBEF_S_KOUSU).Text = "" Then
            
            Else
                If Not IsNumeric(Text1(ptxBEF_S_KOUSU).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(�H��)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxBEF_S_KOUSU).Text = Format(CDbl(Text1(ptxBEF_S_KOUSU).Text), "#0.00")
                End If
            End If
        
        Case ptxBEF_S_KOUSU_GENKA           '           (����)���i���H��
        
            If Text1(ptxBEF_S_KOUSU_GENKA).Text = "" Then
            
            Else
            
            
                If Not IsNumeric(Text1(ptxBEF_S_KOUSU_GENKA).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(�H������)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxBEF_S_KOUSU_GENKA).Text = Format(CDbl(Text1(ptxBEF_S_KOUSU_GENKA).Text), "#0.00")
                End If
            End If
        
        
        Case ptxBEF_S_KOUSU_BAIKA           '           (����)���i���H��
        
        
            If Text1(ptxBEF_S_KOUSU_BAIKA).Text = "" Then
            
            Else
            
                If Not IsNumeric(Text1(ptxBEF_S_KOUSU_BAIKA).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(�H������)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxBEF_S_KOUSU_BAIKA).Text = Format(CDbl(Text1(ptxBEF_S_KOUSU_BAIKA).Text), "#0.00")
                End If
            End If
        
        Case ptxBEF_S_SHIZAI_GENKA          '           (����)����
        
        
            If Text1(ptxBEF_S_SHIZAI_GENKA).Text = "" Then
            
            Else
            
            
                If Not IsNumeric(Text1(ptxBEF_S_SHIZAI_GENKA).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(���ތ���)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxBEF_S_SHIZAI_GENKA).Text = Format(CDbl(Text1(ptxBEF_S_SHIZAI_GENKA).Text), "#0.00")
                End If
            End If
        
        
        
        
        Case ptxBEF_S_SHIZAI_BAIKA          '           (����)����

            If Text1(ptxBEF_S_SHIZAI_BAIKA).Text = "" Then
            
            Else
            
                If Not IsNumeric(Text1(ptxBEF_S_SHIZAI_BAIKA).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(���ޔ���)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxBEF_S_SHIZAI_BAIKA).Text = Format(CDbl(Text1(ptxBEF_S_SHIZAI_BAIKA).Text), "#0.00")
                End If
            End If

        Case ptxBEF_S_GAISO_TANKA           '           �O���P��
        
        
            If Text1(ptxBEF_S_GAISO_TANKA).Text = "" Then
            
            Else
            
                If Not IsNumeric(Text1(ptxBEF_S_GAISO_TANKA).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(�O���P��)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxBEF_S_GAISO_TANKA).Text = Format(CDbl(Text1(ptxBEF_S_GAISO_TANKA).Text), "#0.00")
                End If
            End If
        
        
        
        
        Case ptxBEF_S_PPSC_KAKO_KOSU        '           PPSC���H�P��
            
            If Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text = "" Then
            
            Else
            
            
                If Not IsNumeric(Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(PPSC���H�P��)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text = Format(CDbl(Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text), "#0.00")
                End If
        
            End If
        
        
        Case ptxBEF_S_BU_KAKO_KOSU          '           BU���H�P��
    
            If Text1(ptxBEF_S_BU_KAKO_KOSU).Text = "" Then
            
            Else
            
            
                If Not IsNumeric(Text1(ptxBEF_S_BU_KAKO_KOSU).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(PPSC���H�P��)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxBEF_S_BU_KAKO_KOSU).Text = Format(CDbl(Text1(ptxBEF_S_BU_KAKO_KOSU).Text), "#0.00")
                End If
            End If
        
        
        
        Case ptxBEF_S_KOUSU_SET_DATE        '           �ݒ��
        
        
        
            If Text1(ptxBEF_S_KOUSU_SET_DATE).Text = "" Then
            
            Else
            
            
            
                If Len(Text1(ptxBEF_S_KOUSU_SET_DATE).Text) < 8 Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(�ݒ��)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
            
                    If Not IsDate(Mid(Text1(ptxBEF_S_KOUSU_SET_DATE).Text, 1, 4) & "/" & _
                                    Mid(Text1(ptxBEF_S_KOUSU_SET_DATE).Text, 5, 2) & "/" & _
                                    Mid(Text1(ptxBEF_S_KOUSU_SET_DATE).Text, 7, 2)) Then
                        MsgBox "���͂������ڂ̓G���[�ł��B(�ݒ��)"
                        Text1(Mode).SetFocus
                        Exit Function
                    End If
                End If
            End If
        
        Case ptxBEF_SEI_TANKA_TANTO         '          �S����
            If Text1(ptxBEF_SEI_TANKA_TANTO).Text = "" Then
            Else
                
                Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxBEF_SEI_TANKA_TANTO).Text)
    
                sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                
'>>>>>>>>>>>>>>>>>>>    2012.01.07
'                        MsgBox "���͂������ڂ̓G���[�ł��B(�S����)"
'                        Text1(Mode).SetFocus
'                        Exit Function
'>>>>>>>>>>>>>>>>>>>    2012.01.07
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
                        Exit Function
                End Select
            End If
        
        Case ptxBEF_SE_TANKA_MEMO           '          ����
        
        
        
        
        Case ptxAFT_SEI_LOT         '���b�g��
            
            If Text1(ptxAFT_SEI_LOT).Text = "" Then
                Text1(ptxAFT_SEI_LOT).Text = "1"
            End If
            
            If Not IsNumeric(Text1(ptxAFT_SEI_LOT).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B(���b�g��)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxAFT_SEI_LOT).Text = Format(CLng(Text1(ptxAFT_SEI_LOT).Text), "#0")
            End If
        
        Case ptxAFT_SEI_RATE        '�����[�g
            
            If Text1(ptxAFT_SEI_RATE).Text = "" Then
                Text1(ptxAFT_SEI_RATE).Text = "0"
            End If
            
            If Not IsNumeric(Text1(ptxAFT_SEI_RATE).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B(�����[�g)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxAFT_SEI_RATE).Text = Format(CLng(Text1(ptxAFT_SEI_RATE).Text), "#0.00")
            End If
    
        Case ptxAFT_S_KOUSU         '�H��
            
            If Text1(ptxAFT_S_KOUSU).Text = "" Then
                Text1(ptxAFT_S_KOUSU).Text = "0.00"
            End If
            
            If Not IsNumeric(Text1(ptxAFT_SEI_RATE).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B(�H��)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxAFT_S_KOUSU).Text = Format(CDbl(Text1(ptxAFT_S_KOUSU).Text), "#0.00")
            End If
    
    
        Case ptxAFT_S_KOUSU_GENKA   '�H������
            
            If Text1(ptxAFT_S_KOUSU_GENKA).Text = "" Then
                Text1(ptxAFT_S_KOUSU_GENKA).Text = "0.00"
            End If
            
            If Not IsNumeric(Text1(ptxAFT_S_KOUSU_GENKA).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B(�H������)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxAFT_S_KOUSU_GENKA).Text = Format(CDbl(Text1(ptxAFT_S_KOUSU_GENKA).Text), "#0.00")
            End If
        
        Case ptxAFT_S_KOUSU_BAIKA   '�H������
            
            If Text1(ptxAFT_S_KOUSU_BAIKA).Text = "" Then
                Text1(ptxAFT_S_KOUSU_BAIKA).Text = "0.00"
            End If
            
            If Not IsNumeric(Text1(ptxAFT_S_KOUSU_BAIKA).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B(�H������)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxAFT_S_KOUSU_BAIKA).Text = Format(CDbl(Text1(ptxAFT_S_KOUSU_BAIKA).Text), "#0.00")
            End If
    
    
    
    
        Case ptxAFT_S_SHIZAI_GENKA   '���ތ���
            
            If Text1(ptxAFT_S_SHIZAI_GENKA).Text = "" Then
                Text1(ptxAFT_S_SHIZAI_GENKA).Text = "0.00"
            End If
            
            If Not IsNumeric(Text1(ptxAFT_S_SHIZAI_GENKA).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B(���ތ���)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxAFT_S_SHIZAI_GENKA).Text = Format(CDbl(Text1(ptxAFT_S_SHIZAI_GENKA).Text), "#0.00")
            End If
    
    
        Case ptxAFT_S_SHIZAI_BAIKA  '���ޔ���
            
            If Text1(ptxAFT_S_SHIZAI_BAIKA).Text = "" Then
                Text1(ptxAFT_S_SHIZAI_BAIKA).Text = "0.00"
            End If
            
            If Not IsNumeric(Text1(ptxAFT_S_SHIZAI_BAIKA).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B(���ޔ���)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxAFT_S_SHIZAI_BAIKA).Text = Format(CDbl(Text1(ptxAFT_S_SHIZAI_BAIKA).Text), "#0.00")
            End If
    
        Case ptxAFT_S_GAISO_TANKA       '�O���P��
    
            If Text1(ptxAFT_S_GAISO_TANKA).Text = "" Then
                Text1(ptxAFT_S_GAISO_TANKA).Text = "0.00"
            End If
            
            If Not IsNumeric(Text1(ptxAFT_S_GAISO_TANKA).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B(�O���P��)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxAFT_S_GAISO_TANKA).Text = Format(CDbl(Text1(ptxAFT_S_GAISO_TANKA).Text), "#0.00")
            End If
    
    
    
        Case ptxAFT_S_PPSC_KAKO_KOSU    'PPSC���H�P��
        
        
            If Text1(ptxAFT_S_PPSC_KAKO_KOSU).Text = "" Then
                Text1(ptxAFT_S_PPSC_KAKO_KOSU).Text = "0.00"
            End If
            
            If Not IsNumeric(Text1(ptxAFT_S_PPSC_KAKO_KOSU).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B(PPSC���H�P��)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxAFT_S_PPSC_KAKO_KOSU).Text = Format(CDbl(Text1(ptxAFT_S_PPSC_KAKO_KOSU).Text), "#0.00")
            End If
        
        
        
        
        Case ptxAFT_S_BU_KAKO_KOSU      'BU���H�P��
    
            If Text1(ptxAFT_S_BU_KAKO_KOSU).Text = "" Then
                Text1(ptxAFT_S_BU_KAKO_KOSU).Text = "0.00"
            End If
            
            If Not IsNumeric(Text1(ptxAFT_S_BU_KAKO_KOSU).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B(PPSC���H�P��)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxAFT_S_BU_KAKO_KOSU).Text = Format(CDbl(Text1(ptxAFT_S_BU_KAKO_KOSU).Text), "#0.00")
            End If
    
    
    
        Case ptxAFT_SEI_TANKA_TANTO     '�S����
        
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxAFT_SEI_TANKA_TANTO).Text)

            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
            
                    MsgBox "���͂������ڂ̓G���[�ł��B(�S����)"
                    Text1(Mode).SetFocus
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
                    Exit Function
                
            
            
            End Select
    
        Case ptxAFT_SE_TANKA_MEMO       '����
        
        Case ptxCATE_ST_KOUTEI          ' �O��H���i�b�j�W��
        
        Case ptxCATE_AD_KOUTEI          ' �O��H���i�b�j����
        
        
            If Trim(Text1(Mode).Text) = "" Then
                Text1(Mode).Text = "0"
            End If
        
            If Not IsNumeric(Text1(Mode).Text) Then
            
                MsgBox "���͂������ڂ̓G���[�ł��B(�O��H��)"
                Text1(Mode).SetFocus
                Exit Function
            
            End If
        
        
            '�Čv�Z
            Call CATEGORY_KEISAN_PROC
        
        
        
        
        Case ptxCATE_ST_FUKA            ' �t���H���i�b�j�W��
        
        
            If Trim(Text1(Mode).Text) = "" Then
                Text1(Mode).Text = "0"
            End If
        
            If Not IsNumeric(Text1(Mode).Text) Then
            
                MsgBox "���͂������ڂ̓G���[�ł��B(�t���H��)"
                Text1(Mode).SetFocus
                Exit Function
            
            Else
                Text1(Mode).Text = Val(Text1(Mode).Text)
            End If
        
            Text1(ptxCATE_AD_FUKA).Text = Text1(ptxCATE_ST_FUKA).Text
        
        
            '�Čv�Z
            Call CATEGORY_KEISAN_PROC
        
        Case ptxCATE_AD_FUKA            ' �t���H���i�b�j����

        Case ptxCATE_ST_JITU1           ' ����ƍH���i�b�j�W��
        
        Case ptxCATE_AD_JITU1           ' ����ƍH���i�b�j����

        Case ptxCATE_ST_YOYU_RITU       ' �]�T���i���j�W��
        
        Case ptxCATE_AD_YOYU_RITU       ' �]�T���i���j����

            If Trim(Text1(Mode).Text) = "" Then
                Text1(Mode).Text = "0"
            End If
        
            If Not IsNumeric(Text1(Mode).Text) Then
            
                MsgBox "���͂������ڂ̓G���[�ł��B(�]�T��)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(Mode).Text = Val(Text1(Mode).Text)
            
            End If

            '�Čv�Z
            Call CATEGORY_KEISAN_PROC

        Case ptxCATE_ST_JITU2           ' ����ƍH���i�b�j�W��
        
        Case ptxCATE_AD_JITU2           ' ����ƍH���i�b�j����

        Case ptxCATE_ST_TOTAL           ' ��Ǝ��Ԍv�i�b�j�W��
        
        Case ptxCATE_AD_TOTAL           ' ��Ǝ��Ԍv�i�b�j����

        Case ptxCATE_ST_FUN             ' ��/�i��/�j�W��
        
        Case ptxCATE_AD_FUN             ' ��/�i��/�j����

        Case ptxCATE_ST_FUN_RATE        ' ��ڰāi�~/���j�W��
        
        Case ptxCATE_AD_FUN_RATE        ' ��ڰāi�~/���j����

            If Trim(Text1(Mode).Text) = "" Then
                Text1(Mode).Text = "0.00"
            End If
        
            If Not IsNumeric(Text1(Mode).Text) Then
            
                MsgBox "���͂������ڂ̓G���[�ł��B(��ڰ�)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(Mode).Text = Val(Text1(Mode).Text)
            
            End If

            '�Čv�Z
            Call CATEGORY_KEISAN_PROC

        Case ptxCATE_ST_KOURYO          ' �H�����i�~/�j�W��
        
        Case ptxCATE_AD_KOURYO          ' �H�����i�~/�j����
        
            If Trim(Text1(Mode).Text) = "" Then
                Text1(Mode).Text = "0.00"
            End If
        
            If Not IsNumeric(Text1(Mode).Text) Then
            
                MsgBox "���͂������ڂ̓G���[�ł��B(�H����)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(Mode).Text = Val(Text1(Mode).Text)
            
            End If
        
            '�Čv�Z
            Call CATEGORY_KEISAN_PROC
        
        Case ptxMAIN_KOUTEI_QTY01       '���x���\��t������
            
            If Not IsNumeric(Text1(ptxMAIN_KOUTEI_QTY01).Text) Then
                Text1(ptxMAIN_KOUTEI_QTY01).Text = "1"
            Else
                Text1(ptxMAIN_KOUTEI_QTY01).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_QTY01).Text), "#0")
            End If
    
            If IsNumeric(Text1(ptxMAIN_KOUTEI_TANI01)) Then
                Text1(ptxMAIN_KOUTEI_KOUSU01).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI01).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY01).Text), "#0")
            End If
    
        Case ptxSHIYOU_NO               '�d�l����       2009.06.02
        Case ptxMITSUMORI_KBN           '���ϋ敪       2009.06.02
        
            If Text1(ptxMITSUMORI_KBN).Text = "1" Or Text1(ptxMITSUMORI_KBN).Text = "2" Then
            Else
                MsgBox "���͂������ڂ̓G���[�ł��B(���ϋ敪)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
        Case ptxTANKA_KIRIKAE_DT        '�P���ؑ֓��t   2009.06.02
            
            If Trim(Text1(ptxTANKA_KIRIKAE_DT).Text) = "" Then
            Else
                If Len(Trim(Text1(ptxTANKA_KIRIKAE_DT).Text)) <> 8 Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(�P���ؑ֓��t)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    If IsDate(Mid(Text1(ptxTANKA_KIRIKAE_DT).Text, 1, 4) & "/" & Mid(Text1(ptxTANKA_KIRIKAE_DT).Text, 5, 2) & "/" & Mid(Text1(ptxTANKA_KIRIKAE_DT).Text, 7, 2)) Then
                    Else
                        MsgBox "���͂������ڂ̓G���[�ł��B(�P���ؑ֓��t)"
                        Text1(Mode).SetFocus
                        Exit Function
                    End If
                End If
            End If
                
        
        Case ptxKIRIKAE_KBN             '�ؑ֋敪       2009.06.02
    
    
    End Select
        
    Error_Check_Proc = False
    

End Function


Private Function Code_Set_Proc(Index As Integer, KBN As String, Mode As Integer, Optional flg As Integer = 0) As Integer
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
    If flg = 1 Then
        cmbSHIMUKE.Clear
    End If
    
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
            wkOption = Trim(StrConv(P_CODEREC.OPTION1, vbUnicode))
        End If
        If P_KBN_TBL(i).KBN_OP2 Then
            wkOption = wkOption & Trim(StrConv(P_CODEREC.OPTION2, vbUnicode))
        End If
        
        
        
        Combo1(Index).AddItem StrConv(P_CODEREC.C_RNAME, vbUnicode) & "            " & _
                                Left(StrConv(P_CODEREC.C_Code, vbUnicode), Key_Len) & wkOption
        
        
        
        If flg = 1 Then
            cmbSHIMUKE.AddItem StrConv(P_CODEREC.C_RNAME, vbUnicode) & "            " & _
                                    Left(StrConv(P_CODEREC.C_Code, vbUnicode), Key_Len) & wkOption
        End If
        
        
        com = BtOpGetNext
    
    Loop

    Code_Set_Proc = False
    



End Function


Private Function P_COMPO_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   �\���}�X�^�̓ǂݍ��݁��\��
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
    
Dim Row         As Long
    
Dim FAST_FLG    As Boolean
    
    P_COMPO_Disp_Proc = True
    Call Input_Lock             '2008.01.15
    
        
    
    
            

    

    Set KOUSEI = Nothing

    
    
    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHin_Gai).Text)
       
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_HEAD)
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")
       
    sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        
    Select Case sts
        Case BtNoErr
        
            FAST_FLG = True
        
            '���l
            RichTextBox1(prchBIKOU).Text = RTrim(StrConv(P_COMPO_O_REC.BIKOU, vbUnicode))
        
            '���i���׽
            Text1(ptxS_CLASS_CODE).Text = Trim(StrConv(P_COMPO_O_REC.CLASS_CODE, vbUnicode))
            '�t���׽
            Text1(ptxF_CLASS_CODE).Text = Trim(StrConv(P_COMPO_O_REC.F_CLASS_CODE, vbUnicode))
            '���E�׽
            Text1(ptxN_CLASS_CODE).Text = Trim(StrConv(P_COMPO_O_REC.N_CLASS_CODE, vbUnicode))

        
        Case BtErrKeyNotFound
            
            FAST_FLG = False
            
            '���l
            RichTextBox1(prchBIKOU).Text = ""
        
            '���i���׽
            Text1(ptxS_CLASS_CODE).Text = ""
            '�t���׽
            Text1(ptxF_CLASS_CODE).Text = ""
            '���E�׽
            Text1(ptxN_CLASS_CODE).Text = ""
        
        
        Case Else
            
            Set KOUSEI = Nothing
            
            
            Call Input_UnLock           '2008.01.15
            P_COMPO_Disp_Proc = sts
            Exit Function
    End Select

    '--------------------------------   �u�q�v���
    
    Set KOUSEI = Nothing
    
    
    
    If FAST_FLG Then
    
        Row = Min_Row - 1
        
        Do
            DoEvents
            
            sts = BTRV(BtOpGetNext, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                
                                
                    If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Or _
                        StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                        StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Or _
                        Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxHin_Gai).Text) Then
                    
                        Exit Do
                
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call Input_UnLock             '2008.01.15
                    Call File_Error(sts, BtOpGetNext, "�\���}�X�^")
                    Exit Function
            End Select
            
            
            
            If StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) = P_KOSOU Then
                Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, KOSOU_KBN)
            End If
            If StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) = P_GAISOU Then
                Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, GAISO_KBN)
            End If
            
            Row = Row + 1
                        
            If Grid_Set_Proc(Row) Then
                Exit Function
            End If
            
            
            
        Loop
    End If

    Set TDBGrid1(pGrdKOUSEI).Array = KOUSEI
    
    
    TDBGrid1(pGrdKOUSEI).Bookmark = Null
    
    TDBGrid1(pGrdKOUSEI).ReBind
    TDBGrid1(pGrdKOUSEI).Update
    TDBGrid1(pGrdKOUSEI).ScrollBars = dbgAutomatic
    
    If KOUSEI.Count(1) > 0 Then
        TDBGrid1(pGrdKOUSEI).MoveFirst
    End If















    Call Input_UnLock

    
    
    P_COMPO_Disp_Proc = False

End Function
Private Function Grid_Set_Proc(Row As Long) As Integer
'----------------------------------------------------------------------------
'                   �\���}�X�^==>Grid�e�[�u��
'----------------------------------------------------------------------------

Dim sts As Integer
Dim i   As Integer
Dim j   As Integer
    
    Grid_Set_Proc = True

    

    KOUSEI.ReDim Min_Row, Row, Min_Col, Max_Col
    
    
    '���ƕ�
    KOUSEI(Row, ColKO_JGYOBU) = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode)
    '�����O
    KOUSEI(Row, ColKO_NAIGAI) = StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode)
    
    '���
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN06_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        
    Select Case sts
        Case BtNoErr
            KOUSEI(Row, ColKO_SYUBETSU) = Trim(StrConv(P_CODEREC.C_RNAME, vbUnicode)) & "            " & _
                                Left(StrConv(P_CODEREC.C_Code, vbUnicode), 2)
        
        
        
        Case BtErrKeyNotFound
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�R�[�h�}�X�^")
            Exit Function
    
    End Select
    '�i��
    KOUSEI(Row, ColKO_HIN_GAI) = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
    '�i��
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        
    Select Case sts
        Case BtNoErr
            KOUSEI(Row, ColKO_HIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
        Case BtErrKeyNotFound
            KOUSEI(Row, ColKO_HIN_NAME) = "���o�^�i��"
            
            Call UniCode_Conv(ITEMREC.S_KOUSU, "00000000")
        
            Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "00000000.00")
            Call UniCode_Conv(ITEMREC.G_ST_URITAN, "00000000.00")
        
        
            Call UniCode_Conv(ITEMREC.SEI_SYU_KON, "000.00")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Function
    
    End Select
    '����
    If IsNumeric(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) Then
        KOUSEI(Row, ColKO_QTY) = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
    Else
        KOUSEI(Row, ColKO_QTY) = "1.00"
    End If
    
    '�d���P��
    If IsNumeric(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) Then
        KOUSEI(Row, ColG_ST_SHITAN) = Format(CDbl(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)), "#0.00")
    Else
        KOUSEI(Row, ColG_ST_SHITAN) = "0.00"
    End If
    
    Select Case StrConv(ITEMREC.SEI_KBN, vbUnicode)
    
        Case "1"
            KOUSEI(Row, ColG_ST_URITAN) = "�ʔ�"
        Case "2"
            KOUSEI(Row, ColG_ST_URITAN) = "�x��"
        Case Else
            
Debug.Print StrConv(ITEMREC.G_SPTAN, vbUnicode)
            
            If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                KOUSEI(Row, ColG_ST_URITAN) = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), "#0.00")
            Else
                KOUSEI(Row, ColG_ST_URITAN) = "0.00"
            End If
    End Select
    '�d�����z�v
    KOUSEI(Row, ColG_ST_SHIKIN) = 0

    For i = 0 To UBound(SHIZAI_T)
        If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = SHIZAI_T(i) Then
            
            
            If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
                
                If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
                    
                    If CDbl(KOUSEI(Row, ColKO_QTY)) = 0 Then
                        KOUSEI(Row, ColG_ST_SHIKIN) = "0.00"
                    Else
                        KOUSEI(Row, ColG_ST_SHIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(Row, ColG_ST_SHITAN)) / CDbl(KOUSEI(Row, ColKO_QTY))), 2), "#,##0.00")
                    End If
                Else
                    KOUSEI(Row, ColG_ST_SHIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(Row, ColKO_QTY)) * CDbl(KOUSEI(Row, ColG_ST_SHITAN))), 2), "#,##0.00")
                End If
            End If
            Exit For
        End If
    
    Next i
    If CDbl(KOUSEI(Row, ColG_ST_SHIKIN)) = 0 Then
        KOUSEI(Row, ColG_ST_SHIKIN) = ""
    End If
    
    '������z�v
    KOUSEI(Row, ColG_ST_URIKIN) = 0
    KOUSEI(Row, ColG_ST_URIKIN_KUSATU) = 0

    For i = 0 To UBound(SHIZAI_T)
    
        If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
    
    
            If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = SHIZAI_T(i) Then
                If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
                    
                    If CCur(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) = 0 Then  '2010.02.22
                        KOUSEI(Row, ColG_ST_URIKIN) = "0.00"
                    Else
                        KOUSEI(Row, ColG_ST_URIKIN) = CDbl(ToHalfAdjust(1 / CCur(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), 3))
                    End If
                    KOUSEI(Row, ColG_ST_URIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(Row, ColG_ST_URITAN)) * CDbl(KOUSEI(Row, ColG_ST_URIKIN))), 2), "#,##0.00")
                
                
                
                
                
                Else
                    KOUSEI(Row, ColG_ST_URIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(Row, ColKO_QTY)) * CDbl(KOUSEI(Row, ColG_ST_URITAN))), 2), "#,##0.00")
                End If
    
            
            Else
            
                If KUSATU_F Then
            
                    If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
                    
                        If CCur(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) = 0 Then  '2010.02.22
                            KOUSEI(Row, ColG_ST_URIKIN_KUSATU) = 0
                        Else
                            KOUSEI(Row, ColG_ST_URIKIN_KUSATU) = CDbl(ToHalfAdjust(1 / CCur(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), 3))
                        End If
                        KOUSEI(Row, ColG_ST_URIKIN_KUSATU) = Format(ToRoundUp(CCur(CDbl(KOUSEI(Row, ColG_ST_URITAN)) * CDbl(KOUSEI(Row, ColG_ST_URIKIN_KUSATU))), 2), "#,##0.00")
                    
                    
                    
                    
                    
                    Else
                        KOUSEI(Row, ColG_ST_URIKIN_KUSATU) = Format(ToRoundUp(CCur(CDbl(KOUSEI(Row, ColKO_QTY)) * CDbl(KOUSEI(Row, ColG_ST_URITAN))), 2), "#,##0.00")
                    End If
                
                
                End If
            
            
            
            End If
        End If
    Next i
    
    If CDbl(KOUSEI(Row, ColG_ST_URIKIN)) = 0 Then
        KOUSEI(Row, ColG_ST_URIKIN) = ""
    End If
    
    If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
        KOUSEI(Row, ColS_KOUSU) = ""
        KOUSEI(Row, ColSEI_SYU_KON) = ""
    Else
        '��Ǝ���
        If IsNumeric(StrConv(ITEMREC.S_KOUSU, vbUnicode)) Then
            KOUSEI(Row, ColS_KOUSU) = Format(CDbl(StrConv(ITEMREC.S_KOUSU, vbUnicode)), "#")
        Else
            KOUSEI(Row, ColS_KOUSU) = ""
        End If
        '�W������
        If IsNumeric(StrConv(ITEMREC.SEI_SYU_KON, vbUnicode)) Then
            KOUSEI(Row, ColSEI_SYU_KON) = Format(CInt(StrConv(ITEMREC.SEI_SYU_KON, vbUnicode)), "#")
        Else
            KOUSEI(Row, ColSEI_SYU_KON) = ""
        End If
    End If
    '���l
    KOUSEI(Row, ColKO_BIKOU) = StrConv(P_COMPO_K_REC.KO_BIKOU, vbUnicode)
    
    Grid_Set_Proc = False
End Function

' ------------------------------------------------------------------------
'       �w�肵�����x�̐��l�ɐ؂�グ���܂��B
'
' @Param    dValue      �ۂߑΏۂ̔{���x���������_���B
' @Param    iDigits     �߂�l�̗L�������̐��x�B
' @Return               iDigits �ɓ��������x�̐��l�ɐ؂�グ��ꂽ���l�B
' ------------------------------------------------------------------------
Private Function ToRoundUp(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
    Dim dCoef As Double

    
        


    dCoef = (10 ^ iDigits)



    
    
    
    Select Case CDbl(dValue * dCoef) - Fix(dValue * dCoef)
        Case Is > 0
            ToRoundUp = (Int(dValue * dCoef) + 1) / dCoef
        Case Is < 0
            ToRoundUp = (Fix(dValue * dCoef) - 1) / dCoef
        Case Else
            ToRoundUp = dValue
    End Select


'    Select Case CDbl(dValue * dCoef) - Fix(dValue * dCoef)
'        Case Is > 0
'            ToRoundUp = (Int(dValue * dCoef + 0.9)) / dCoef
'        Case Is < 0
'            ToRoundUp = (Fix(dValue * dCoef - 0.9)) / dCoef
'        Case Else
'            ToRoundUp = dValue
'    End Select



End Function

' ------------------------------------------------------------------------
'       �w�肵�����x�̐��l�ɐ؂�̂Ă��܂��B
'
' @Param    dValue      �ۂߑΏۂ̔{���x���������_���B
' @Param    iDigits     �߂�l�̗L�������̐��x�B
' @Return               iDigits �ɓ��������x�̐��l�ɐ؂�̂Ă�ꂽ���l�B
' ------------------------------------------------------------------------
Public Function ToRoundDown(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
    Dim dCoef As Double

    dCoef = (10 ^ iDigits)

    Select Case CDbl(dValue * dCoef) - Fix(dValue * dCoef)
        Case Is > 0
            ToRoundDown = Int(dValue * dCoef) / dCoef
        Case Is < 0
            ToRoundDown = Fix(dValue * dCoef) / dCoef
        Case Else
            ToRoundDown = dValue
    End Select
End Function





' ------------------------------------------------------------------------
'       �w�肵�����x�̐��l�Ɏl�̌ܓ����܂��B
'
' @Param    dValue      �ۂߑΏۂ̔{���x���������_���B
' @Param    iDigits     �߂�l�̗L�������̐��x�B
' @Return               iDigits �ɓ��������x�̐��l�Ɏl�̌ܓ����ꂽ���l�B
' ------------------------------------------------------------------------
Public Function ToHalfAdjust(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
    Dim dCoef As Double

    dCoef = (10 ^ iDigits)

    If dValue > 0 Then
        ToHalfAdjust = Int(CDbl(dValue * dCoef + 0.5)) / dCoef
    Else
        ToHalfAdjust = Fix(CDbl(dValue * dCoef - 0.5)) / dCoef
    End If
End Function






Private Function Estimate_Proc() As Integer
'----------------------------------------------------------------------------
'                   �d�w�b�d�k�i�䌩�Ϗ��j�o��
'       2009.06.02
'----------------------------------------------------------------------------
Dim excelApplication    As Object
Dim excelWorkBook       As Object
Dim excelSheet          As Object

Dim i                   As Integer
Dim j                   As Integer

Dim com                 As Integer
Dim sts                 As Integer

Dim wkBikou             As Variant

Dim Row                 As Integer
Dim SHIZAI_TOTAL_ROW    As Integer
Dim FUKA_TOTAL_ROW      As Integer
Dim TOTAL_ROW           As Integer
    
    
Dim wkint               As Integer
Dim BEF_KOTEI           As Double
Dim AFT_KOTEI           As Double
Dim MAIN_KOTEI          As Double
    
    
Dim stTime              As String
    
    
Dim wkNum1              As Currency
Dim wkNum2              As Currency
    
    
    
    
'2011.01.11
Dim S_Start             As String
Dim CREATE_EXCEL        As String
Dim HEAD_EXCEL          As String

Dim BODY1_EXCEL         As String
Dim BODY2_EXCEL         As String
Dim BODY3_EXCEL         As String

Dim INS1_EXCEL          As String
Dim INS2_EXCEL          As String
Dim INS3_EXCEL          As String


Dim TOTAL_EXCEL         As String

Dim FOOT_EXCEL          As String
Dim DSP_EXCEL           As String
Dim S_END               As String

Dim S_TITLE             As String
'2011.01.11
    
    
    
Dim SP_TANKA_F          As Boolean          '2011.12.22
    
    
Dim ED_HIN_GAI          As String * 20
Dim ED_I                As Integer
    
    
    
    Estimate_Proc = True
    
    
    Call Input_Lock
    
    
S_TITLE = "�����v�ZOFF"
    
S_Start = Right(Format(Now, "hh:mm:ss"), 5)
    
    Set excelApplication = CreateObject("Excel.Application")
    

    If Trim(EXCEL_TEMPLATE) = "" Then
        Set excelWorkBook = excelApplication.Workbooks.Add
    
    Else
                                                        '����ڰ��ޯ����J��
        Set excelWorkBook = excelApplication.Workbooks.Open(EXCEL_TEMPLATE)
    End If

    Set excelSheet = excelWorkBook.Worksheets(1)
    
    
    
    
    
'excelApplication.Visible = True
    
excelApplication.Calculation = xlCalculationManual
excelApplication.ScreenUpdating = False

    
    
    
CREATE_EXCEL = Right(Format(Now, "hh:mm:ss"), 5)
DoEvents
    
    
    excelSheet.Application.ActiveWindow.DisplayGridlines = False
    
'---    �w�b�_�[�o��
HEAD_EXCEL = Right(Format(Now, "hh:mm:ss"), 5)
DoEvents
    
    Call Estimate_Head_Proc(excelApplication, excelWorkBook, excelSheet)
    
    
    
'---    11�s��
    excelSheet.Application.Rows(11).RowHeight = 13.5
    
    
'---    12�s��
    Call Estimate_Line11_13_Proc(excelApplication, excelWorkBook, excelSheet)   '2011.01.11
    

'---    ���ޕ��o��
BODY1_EXCEL = Right(Format(Now, "hh:mm:ss"), 5)
DoEvents

    If Estimate_SHIZAI_Proc(excelApplication, excelWorkBook, excelSheet, Row) Then
        Call Input_UnLock
        Exit Function
    End If
    SHIZAI_TOTAL_ROW = Row

'---    �������o��
BODY2_EXCEL = Right(Format(Now, "hh:mm:ss"), 5)
DoEvents

    If Estimate_DOUKON_Proc(excelApplication, excelWorkBook, excelSheet, Row) Then
        Call Input_UnLock
        Exit Function
    End If

'---    �t�����o��

BODY3_EXCEL = Right(Format(Now, "hh:mm:ss"), 5)
DoEvents
    If Estimate_FUKA_Proc(excelApplication, excelWorkBook, excelSheet, Row) Then
        Call Input_UnLock
        Exit Function
    End If

    FUKA_TOTAL_ROW = Row

    
'---    42�s��
    Row = Row + 2
    excelSheet.Application.Cells(Row, 2).Font.Size = 10
    
    excelSheet.Application.Cells(Row, 2).HorizontalAlignment = xlLeft
    excelSheet.Application.Cells(Row, 2).Value = "�y��Ɣ�z"
    
    
    
'---    43�s��
    Row = Row + 1
    excelSheet.Application.Rows(Row).RowHeight = 20.25
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 3)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 3)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 3)).MergeCells = True
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row, 5)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row, 5)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row, 5)).MergeCells = True
    

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 6), excelSheet.Application.Cells(Row, 7)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 6), excelSheet.Application.Cells(Row, 7)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 6), excelSheet.Application.Cells(Row, 7)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 6), excelSheet.Application.Cells(Row, 7)).ShrinkToFit = True

    excelSheet.Application.Cells(Row, 8).HorizontalAlignment = xlCenter
    excelSheet.Application.Cells(Row, 8).VerticalAlignment = xlCenter

    excelSheet.Application.Cells(Row, 9).HorizontalAlignment = xlCenter
    excelSheet.Application.Cells(Row, 9).VerticalAlignment = xlCenter



    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).MergeCells = True

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 9)).Font.Size = 10
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).Font.Size = 12

    excelSheet.Application.Cells(Row, 2).Value = "�O��H��(�b)"
    excelSheet.Application.Cells(Row, 4).Value = "����ƍH��(�b)"
    
    excelSheet.Application.Cells(Row, 6).Value = "��Ǝ��Ԍv(�b/��)"
    excelSheet.Application.Cells(Row, 8).Value = "��/��"
    excelSheet.Application.Cells(Row, 9).Value = "�����[�g"
    excelSheet.Application.Cells(Row, 10).Value = "�B�H���P��"








'2010.05.13
INS1_EXCEL = Right(Format(Now, "hh:mm:ss"), 5)
DoEvents
    
    excelSheet.Application.Cells(Row, 14).Font.Size = 12
    excelSheet.Application.Cells(Row, 14).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Cells(Row, 14).VerticalAlignment = xlBottom
    excelSheet.Application.Cells(Row, 14).Value = "�P��"

    excelSheet.Application.Cells(Row, 15).Font.Size = 12
    excelSheet.Application.Cells(Row, 15).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Cells(Row, 15).VerticalAlignment = xlBottom
    excelSheet.Application.Cells(Row, 15).Value = "�`�F�b�N"

    excelSheet.Application.Cells(Row, 17).Font.Size = 12
'2011.11.21    excelSheet.Application.Cells(Row, 17).VerticalAlignment = xlBottom
    excelSheet.Application.Cells(Row, 17).Value = "�r�X�E����E�ۏ؏��`�F�b�N"

'2010.05.13





'---    44�s��
    Row = Row + 1
    excelSheet.Application.Rows(Row).RowHeight = 20.25
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 3)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 3)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 3)).MergeCells = True
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row, 5)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row, 5)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row, 5)).MergeCells = True
    

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 6), excelSheet.Application.Cells(Row, 7)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 6), excelSheet.Application.Cells(Row, 7)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 6), excelSheet.Application.Cells(Row, 7)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 6), excelSheet.Application.Cells(Row, 7)).ShrinkToFit = True

    excelSheet.Application.Cells(Row, 8).HorizontalAlignment = xlCenter
    excelSheet.Application.Cells(Row, 8).VerticalAlignment = xlCenter
    excelSheet.Application.Cells(Row, 8).NumberFormatLocal = "#,##0.00_ "

    excelSheet.Application.Cells(Row, 9).HorizontalAlignment = xlCenter
    excelSheet.Application.Cells(Row, 9).VerticalAlignment = xlCenter



    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).MergeCells = True

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 9)).Font.Size = 10
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).Font.Size = 12




    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 10)).Font.Size = 12
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 11), excelSheet.Application.Cells(Row, 12)).Font.Size = 14
    
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 12)).NumberFormatLocal = "#,##0_ "

    
'2009.07.01
        
    Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHin_Gai).Text)

    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Function
    
    End Select
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �i���ú��   2011.12.12
    
    SP_TANKA_F = False
    Call UniCode_Conv(K0_ITEM_CATEGORY.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_ITEM_CATEGORY.CATEGORY_CODE, Text1(ptxCATEGORY_CODE).Text)
    sts = BTRV(BtOpGetEqual, ITEM_CATEGORY_POS, ITEM_CATEGORYREC, Len(ITEM_CATEGORYREC), K0_ITEM_CATEGORY, Len(K0_ITEM_CATEGORY), 0)
    Select Case sts
        Case BtNoErr
            If IsNumeric(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_KOURYO, vbUnicode)) Then
              SP_TANKA_F = True
            End If
        Case BtErrKeyNotFound
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i���J�e�S���[�}�X�^")
            Exit Function
    
    End Select
        
    If SP_TANKA_F Then
    
    
        '�O��H���i�b�j
        excelSheet.Application.Cells(Row, 2).Value = ""
        '����ƍH�� (�b)
        excelSheet.Application.Cells(Row, 4).Value = ""
        
        '�H���P��
        excelSheet.Application.Cells(Row, 10).FormulaR1C1 = Val(Text1(ptxCATE_AD_KOURYO).Text)
        '��ڰ�
        excelSheet.Application.Cells(Row, 9).Value = Val(Text1(ptxCATE_AD_FUN_RATE).Text)
        '��/��
        If Val(Text1(ptxCATE_AD_FUN_RATE).Text) = 0 Then
            excelSheet.Application.Cells(Row, 8).Value = 0
        Else
            excelSheet.Application.Cells(Row, 8).FormulaR1C1 = "=round(RC[+2]/RC[+1],2)"
        End If
        '��Ǝ��Ԍv�i�b�j
        excelSheet.Application.Cells(Row, 6).FormulaR1C1 = "=round(RC[+2]*60,2)"
    
    
        '�H���P��
        excelSheet.Application.Cells(Row, 10).FormulaR1C1 = "=round(RC[-2]*RC[-1],2)"
    
    Else
        '�O��H���i�b�j
        excelSheet.Application.Cells(Row, 2).Value = Val(Text1(ptxCATE_AD_KOUTEI).Text)
        '����ƍH�� (�b)
        excelSheet.Application.Cells(Row, 4).Value = Val(Text1(ptxCATE_AD_JITU2).Text) + Val(Text1(ptxCATE_AD_FUKA).Text)
        '��Ǝ��Ԍv�i�b�j
        excelSheet.Application.Cells(Row, 6).FormulaR1C1 = "=sum(RC[-5]:RC[-1]"
        '��/��
        excelSheet.Application.Cells(Row, 8).Value = Val(Text1(ptxCATE_AD_FUN).Text)
        '��ڰ�
        excelSheet.Application.Cells(Row, 9).Value = Val(Text1(ptxCATE_AD_FUN_RATE).Text)
        '�H���P��
        excelSheet.Application.Cells(Row, 10).FormulaR1C1 = "=round(RC[-2]*RC[-1],2)"
    End If
    
    If IsNumeric(excelSheet.Application.Cells(Row, 10).Value) Then
        wkNum1 = CCur(excelSheet.Application.Cells(Row, 10).Value)
    Else
        wkNum1 = 0
    End If
    If IsNumeric(Text1(ptxCATE_AD_KOURYO).Text) Then
        wkNum2 = CCur(Text1(ptxCATE_AD_KOURYO).Text)
    Else
        wkNum2 = 0
    End If
    
    If wkNum1 <> wkNum2 Then
        MsgBox "�B�H���P�����v�Z�l(��/�~�����[�g)�ƈقȂ�܂��B"
        excelSheet.Application.Cells(Row, 13).Value = "�B�H���P�����v�Z�l(��/�~�����[�g)�ƈقȂ�܂��B"
    End If
    
'>>>>>>>>>>>>   2018.07.18
'    If IsNumeric(Text1(ptxCATE_AD_KOURYO).Text) Then
'        excelSheet.Application.Cells(Row, 10).Value = CDbl(Text1(ptxCATE_AD_KOURYO).Text)
'        excelSheet.Application.Cells(Row, 10).NumberFormatLocal = "#,##0.00_ "
'    Else
'        excelSheet.Application.Cells(Row, 10).Value = ""
'
'    End If

    If IsNumeric(Text1(ptxBEF_S_KOUSU_BAIKA).Text) Then
        excelSheet.Application.Cells(Row, 10).Value = CDbl(Text1(ptxBEF_S_KOUSU_BAIKA).Text)
        excelSheet.Application.Cells(Row, 10).NumberFormatLocal = "#,##0.00_ "
    Else
        excelSheet.Application.Cells(Row, 10).Value = ""

    End If
'>>>>>>>>>>>>   2018.07.18
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �i���ú��   2011.12.12



'2010.05.13
INS2_EXCEL = Right(Format(Now, "hh:mm:ss"), 5)
DoEvents
    
    
    excelSheet.Application.Cells(Row, 14).HorizontalAlignment = xlRight
'2011.11.21    excelSheet.Application.Cells(Row, 14).VerticalAlignment = xlBottom
    excelSheet.Application.Cells(Row, 14).NumberFormatLocal = "#,##0.00_ "
    excelSheet.Application.Cells(Row, 14).FormulaR1C1 = "=round(round((RC[-12]+RC[-10])/60,2)*RC[-5],2)"


    excelSheet.Application.Cells(Row, 15).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Cells(Row, 15).VerticalAlignment = xlBottom
    excelSheet.Application.Cells(Row, 15).FormulaR1C1 = "=IF(RC[-5]=RC[-1],""��"",""�~"")"


    excelSheet.Application.Cells(Row, 17).HorizontalAlignment = xlRight
'2011.11.21    excelSheet.Application.Cells(Row, 17).VerticalAlignment = xlBottom
    excelSheet.Application.Cells(Row, 17).NumberFormatLocal = "#,##0.00_ "
    excelSheet.Application.Cells(Row, 17).FormulaR1C1 = "=round(RC[-11]/60,2)"

    excelSheet.Application.Cells(Row, 18).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Cells(Row, 18).VerticalAlignment = xlBottom
    excelSheet.Application.Cells(Row, 18).FormulaR1C1 = "=IF(RC[-10]=RC[-1],""��"",""�~"")"

'2010.05.13













    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 2), excelSheet.Application.Cells(Row, 9)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 2), excelSheet.Application.Cells(Row, 9)).Borders(xlDiagonalUp).LineStyle = xlNone

    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 2), excelSheet.Application.Cells(Row, 9)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 2), excelSheet.Application.Cells(Row, 9)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 2), excelSheet.Application.Cells(Row, 9)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 2), excelSheet.Application.Cells(Row, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 2), excelSheet.Application.Cells(Row, 9)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 2), excelSheet.Application.Cells(Row, 9)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 2), excelSheet.Application.Cells(Row, 9)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 2), excelSheet.Application.Cells(Row, 9)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 2), excelSheet.Application.Cells(Row, 9)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 2), excelSheet.Application.Cells(Row, 9)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 2), excelSheet.Application.Cells(Row, 9)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 2), excelSheet.Application.Cells(Row, 9)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 2), excelSheet.Application.Cells(Row, 9)).Borders(xlInsideVertical).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 2), excelSheet.Application.Cells(Row, 9)).Borders(xlInsideVertical).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 2), excelSheet.Application.Cells(Row, 9)).Borders(xlInsideVertical).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 2), excelSheet.Application.Cells(Row, 9)).Borders(xlInsideHorizontal).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 2), excelSheet.Application.Cells(Row, 9)).Borders(xlInsideHorizontal).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 2), excelSheet.Application.Cells(Row, 9)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic


    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlDiagonalUp).LineStyle = xlNone

    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeLeft).Weight = xlThick
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeTop).Weight = xlThick
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeBottom).Weight = xlThick
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeRight).Weight = xlThick
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideVertical).LineStyle = xlNone
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideHorizontal).Weight = xlThick
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic








'2010.05.13
INS3_EXCEL = Right(Format(Now, "hh:mm:ss"), 5)
DoEvents
    
    
    excelSheet.Application.Cells(Row + 1, 14).Font.Size = 12
    excelSheet.Application.Cells(Row + 1, 14).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Cells(Row + 1, 14).VerticalAlignment = xlBottom
    excelSheet.Application.Cells(Row + 1, 14).Value = "�P��"

    excelSheet.Application.Cells(Row + 1, 15).Font.Size = 12
    excelSheet.Application.Cells(Row + 1, 15).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Cells(Row + 1, 15).VerticalAlignment = xlBottom
    excelSheet.Application.Cells(Row + 1, 15).Value = "�`�F�b�N"
'2010.05.13



'---    46�s��
TOTAL_EXCEL = Right(Format(Now, "hh:mm:ss"), 5)
DoEvents
    
    
    Row = Row + 2
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row, 10)).HorizontalAlignment = xlCenter
    
    excelSheet.Application.Cells(Row, 9).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(Row, 9).VerticalAlignment = xlCenter
    excelSheet.Application.Cells(Row, 9).Font.Size = 14
    excelSheet.Application.Cells(Row, 9).Value = "���i����p�@�{�A�{�B"

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).Font.Size = 14
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).Font.FontStyle = "����"
        
    If SHIZAI_TOTAL_ROW = 15 Then
        excelSheet.Application.Cells(Row, 10).FormulaR1C1 = "=R[-2]C+R[" & FUKA_TOTAL_ROW - Row & "]C"
    Else
        excelSheet.Application.Cells(Row, 10).FormulaR1C1 = "=R[-2]C+R[" & SHIZAI_TOTAL_ROW - Row & "]C+R[" & FUKA_TOTAL_ROW - Row & "]C"
    End If
    excelSheet.Application.Cells(Row, 10).NumberFormatLocal = "#,##0.00_ "
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlDiagonalUp).LineStyle = xlNone

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeLeft).Weight = xlThick
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeTop).Weight = xlThick
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeBottom).Weight = xlThick
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeRight).Weight = xlThick
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideVertical).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideHorizontal).LineStyle = xlNone

    excelSheet.Application.Cells(Row, 10).Borders(xlLeft).LineStyle = xlContinuous
    excelSheet.Application.Cells(Row, 10).Borders(xlLeft).Weight = xlThick
    excelSheet.Application.Cells(Row, 10).Borders(xlLeft).ColorIndex = xlAutomatic



'2010.05.13
    excelSheet.Application.Cells(Row, 14).HorizontalAlignment = xlRight
'2011.11.21    excelSheet.Application.Cells(Row, 14).VerticalAlignment = xlBottom
    excelSheet.Application.Cells(Row, 14).NumberFormatLocal = "#,##0.00_ "
    If SHIZAI_TOTAL_ROW = 15 Then
        excelSheet.Application.Cells(Row, 14).FormulaR1C1 = "=R[-2]C+R[" & FUKA_TOTAL_ROW - Row & "]C"
    Else
        excelSheet.Application.Cells(Row, 14).FormulaR1C1 = "=R[-2]C+R[" & SHIZAI_TOTAL_ROW - Row & "]C+R[" & FUKA_TOTAL_ROW - Row & "]C"
    End If


    excelSheet.Application.Cells(Row, 15).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Cells(Row, 15).VerticalAlignment = xlBottom
    excelSheet.Application.Cells(Row, 15).FormulaR1C1 = "=IF(RC[-5]=RC[-1],""��"",""�~"")"


    excelSheet.Application.Cells(Row + 1, 17).Font.Size = 11
'2011.11.21    excelSheet.Application.Cells(Row + 1, 17).VerticalAlignment = xlBottom
    excelSheet.Application.Cells(Row + 1, 17).Value = "���b�g��"

    excelSheet.Application.Cells(Row + 2, 17).Font.Size = 11
'2011.11.21    excelSheet.Application.Cells(Row + 2, 17).VerticalAlignment = xlBottom
    excelSheet.Application.Cells(Row + 2, 17).Value = Text1(ptxBEF_SEI_LOT).Text

'2010.05.13



'---    48�s��
    Row = Row + 2
    excelSheet.Application.Cells(Row, 2).Font.Size = 10
    
    excelSheet.Application.Cells(Row, 2).HorizontalAlignment = xlLeft
    excelSheet.Application.Cells(Row, 2).Value = "�y���l�z"


'---    49�`51�s��
    
    
    Row = Row + 1
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 2, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 2, 11)).Borders(xlDiagonalUp).LineStyle = xlNone

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 2, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 2, 11)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 2, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 2, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 2, 11)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 2, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 2, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 2, 11)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 2, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 2, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 2, 11)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 2, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 2, 11)).Borders(xlInsideVertical).LineStyle = xlNone
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 2, 11)).Borders(xlInsideHorizontal).LineStyle = xlNone


    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).HorizontalAlignment = xlLeft
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).MergeCells = True
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 2), excelSheet.Application.Cells(Row + 1, 11)).HorizontalAlignment = xlLeft
    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 2), excelSheet.Application.Cells(Row + 1, 11)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 2), excelSheet.Application.Cells(Row + 1, 11)).MergeCells = True

    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 2, 2), excelSheet.Application.Cells(Row + 2, 11)).HorizontalAlignment = xlLeft
    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 2, 2), excelSheet.Application.Cells(Row + 2, 11)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 2, 2), excelSheet.Application.Cells(Row + 2, 11)).MergeCells = True



    If Trim(RichTextBox1(prchM_BIKOU).Text) = "" Then
    Else
        wkBikou = Split(RTrim(RichTextBox1(prchM_BIKOU).Text), vbCrLf, -1)
        For i = Row To UBound(wkBikou) + Row
            excelSheet.Application.Cells(i, 2).Value = wkBikou(i - Row)
        Next i
    End If
    
    
    
    



'---    53�`56�s��
FOOT_EXCEL = Right(Format(Now, "hh:mm:ss"), 5)
DoEvents
    
    
    Row = Row + 5
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 3)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 3)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 3)).MergeCells = True
    
    Select Case Trim(Text1(ptxMITSUMORI_KBN).Text)
        Case "1"
            excelSheet.Application.Cells(Row, 2).Value = "�V�K�d�l"
        Case "2"
            excelSheet.Application.Cells(Row, 2).Value = "���s�d�l"
    End Select

    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 2, 2), excelSheet.Application.Cells(Row + 3, 3)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 2, 2), excelSheet.Application.Cells(Row + 3, 3)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 2, 2), excelSheet.Application.Cells(Row + 3, 3)).MergeCells = True

    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 2, 2), excelSheet.Application.Cells(Row + 3, 3)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 2, 2), excelSheet.Application.Cells(Row + 3, 3)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 2, 2), excelSheet.Application.Cells(Row + 3, 3)).MergeCells = True

   

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row + 3, 4)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row + 3, 4)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row + 3, 4)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row + 3, 4)).WrapText = True
    
    excelSheet.Application.Cells(Row, 4).Value = "�d�l����" & Left(Combo1(pcmbSHIMUKE).Text, Len(Combo1(pcmbSHIMUKE).Text) - 4)
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row + 3, 5)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row + 3, 5)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row + 3, 5)).MergeCells = True
    
    
    excelSheet.Application.Cells(Row, 5).Value = Trim(Text1(ptxSHIYOU_NO).Text)







    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 3, 3)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 3, 3)).Borders(xlDiagonalUp).LineStyle = xlNone

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 3, 3)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 3, 3)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 3, 3)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 3, 3)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 3, 3)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 3, 3)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 3, 3)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 3, 3)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 3, 3)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 3, 3)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 3, 3)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 3, 3)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 3, 3)).Borders(xlInsideVertical).LineStyle = xlNone

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 3, 3)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 3, 3)).Borders(xlInsideHorizontal).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 3, 3)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic


    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row + 3, 4)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row + 3, 4)).Borders(xlDiagonalUp).LineStyle = xlNone

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row + 3, 4)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row + 3, 4)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row + 3, 4)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row + 3, 4)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row + 3, 4)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row + 3, 4)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row + 3, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row + 3, 4)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row + 3, 4)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row + 3, 4)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row + 3, 4)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row + 3, 4)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row + 3, 4)).Borders(xlInsideVertical).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row + 3, 4)).Borders(xlInsideHorizontal).LineStyle = xlNone

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row + 3, 5)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row + 3, 5)).Borders(xlDiagonalUp).LineStyle = xlNone

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row + 3, 5)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row + 3, 5)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row + 3, 5)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row + 3, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row + 3, 5)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row + 3, 5)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row + 3, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row + 3, 5)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row + 3, 5)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row + 3, 5)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row + 3, 5)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row + 3, 5)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row + 3, 5)).Borders(xlInsideVertical).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row + 3, 5)).Borders(xlInsideHorizontal).LineStyle = xlNone








'''2011.01.21
    If Trim(Insert_Pic) = "" Then
        excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 9), excelSheet.Application.Cells(Row + 3, 9)).HorizontalAlignment = xlCenter
        excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 9), excelSheet.Application.Cells(Row + 3, 9)).VerticalAlignment = xlCenter
        excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 9), excelSheet.Application.Cells(Row + 3, 9)).MergeCells = True
    
        excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 10), excelSheet.Application.Cells(Row + 3, 10)).HorizontalAlignment = xlCenter
        excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 10), excelSheet.Application.Cells(Row + 3, 10)).VerticalAlignment = xlCenter
        excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 10), excelSheet.Application.Cells(Row + 3, 10)).MergeCells = True
    
        excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 11), excelSheet.Application.Cells(Row + 3, 11)).HorizontalAlignment = xlCenter
        excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 11), excelSheet.Application.Cells(Row + 3, 11)).VerticalAlignment = xlCenter
        excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 11), excelSheet.Application.Cells(Row + 3, 11)).MergeCells = True
    
    
    
    
        excelSheet.Application.Cells(Row, 9).HorizontalAlignment = xlCenter
        excelSheet.Application.Cells(Row, 9).VerticalAlignment = xlCenter
        excelSheet.Application.Cells(Row, 9).Font.Size = 10
        excelSheet.Application.Cells(Row, 9).Value = "���F��"
    
        excelSheet.Application.Cells(Row, 10).HorizontalAlignment = xlCenter
        excelSheet.Application.Cells(Row, 10).VerticalAlignment = xlCenter
        excelSheet.Application.Cells(Row, 10).Font.Size = 10
        excelSheet.Application.Cells(Row, 10).Value = "����"
    
        excelSheet.Application.Cells(Row, 11).HorizontalAlignment = xlCenter
        excelSheet.Application.Cells(Row, 11).VerticalAlignment = xlCenter
        excelSheet.Application.Cells(Row, 11).Font.Size = 10
        excelSheet.Application.Cells(Row, 11).Value = "�S����"
    
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 3, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 3, 11)).Borders(xlDiagonalUp).LineStyle = xlNone
    
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 3, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 3, 11)).Borders(xlEdgeLeft).Weight = xlThin
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 3, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 3, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 3, 11)).Borders(xlEdgeTop).Weight = xlThin
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 3, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
    
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 3, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 3, 11)).Borders(xlEdgeBottom).Weight = xlThin
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 3, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 3, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 3, 11)).Borders(xlEdgeRight).Weight = xlThin
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 3, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
    
    
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 3, 11)).Borders(xlInsideVertical).LineStyle = xlContinuous
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 3, 11)).Borders(xlInsideVertical).Weight = xlThin
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 3, 11)).Borders(xlInsideVertical).ColorIndex = xlAutomatic
    
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 3, 11)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 3, 11)).Borders(xlInsideHorizontal).Weight = xlThin
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 3, 11)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
    End If
'''2011.01.21
    If Trim(Insert_Pic) <> "" Then
        
        
        


        
        
        
        
'        excelSheet.Application.Pictures.Insert (Insert_Pic)


'        excelSheet.Pictures.Insert(Insert_Pic).Top = excelSheet.Application.Cells(row, 7).Top
'        excelSheet.Pictures.Insert(Insert_Pic).Left = excelSheet.Application.Cells(row, 7).Left
        
        
'----------------   2013.07.02 Pictures.Insert-->Shapes.AddPicture
'
'         With excelSheet.Pictures.Insert(Insert_Pic)
'            .Top = excelSheet.Application.Cells(Row - 1, 7).Top
'            .Left = excelSheet.Application.Cells(Row - 1, 7).Left
'''            .Height = 3.15 / 0.0378
'            .Width = (excelSheet.Application.Cells(Row - 1, 7).Width + _
'                        excelSheet.Application.Cells(Row - 1, 8).Width + _
'                        excelSheet.Application.Cells(Row - 1, 9).Width + _
'                        excelSheet.Application.Cells(Row - 1, 10).Width + _
'                        excelSheet.Application.Cells(Row - 1, 11).Width)
''            .Width = (excelSheet.Application.Cells(row - 1, 11).Top + excelSheet.Application.Cells(row - 1, 11).Width)
'
'
'
'
''            .Height = 2.93 / 0.0378
'
'
''            .Width = 8.62 / 0.0378
'
'
'
'        End With


        excelSheet.Shapes.AddPicture(Insert_Pic, _
                                            False, _
                                            True, _
                                            excelSheet.Application.Cells(Row - 1, 7).Left, _
                                            excelSheet.Application.Cells(Row - 1, 7).Top, _
                                            (excelSheet.Application.Cells(Row - 1, 7).Width + _
                                            excelSheet.Application.Cells(Row - 1, 8).Width + _
                                            excelSheet.Application.Cells(Row - 1, 9).Width + _
                                            excelSheet.Application.Cells(Row - 1, 10).Width + _
                                            excelSheet.Application.Cells(Row - 1, 11).Width), _
                                            100).Apply
'----------------   2013.07.02 Pictures.Insert-->Shapes.AddPicture





'        With excelSheet.Shapes(8)
'            .LockAspectRatio = True     '---(1)�}�`�̏c���̔䗦���Œ�
'        End With


    End If



'---    ��O�g
    Row = Row + 4

    excelSheet.Application.Rows(Row).RowHeight = 45     '2011.01.24
    
    excelSheet.Application.Range(excelSheet.Application.Cells(11, 1), excelSheet.Application.Cells(11, 12)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(11, 1), excelSheet.Application.Cells(11, 12)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(11, 1), excelSheet.Application.Cells(11, 12)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(11, 1), excelSheet.Application.Cells(Row, 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(11, 1), excelSheet.Application.Cells(Row, 1)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(11, 1), excelSheet.Application.Cells(Row, 1)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row, 12)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row, 12)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row, 12)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(11, 12), excelSheet.Application.Cells(Row, 12)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(11, 12), excelSheet.Application.Cells(Row, 12)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(11, 12), excelSheet.Application.Cells(Row, 12)).Borders(xlEdgeRight).ColorIndex = xlAutomatic


    excelSheet.Application.Cells(1, 1).Select



excelApplication.Calculation = xlCalculationAutomatic



DSP_EXCEL = Right(Format(Now, "hh:mm:ss"), 5)
DoEvents



excelApplication.ScreenUpdating = True
'excelApplication.Visible = True
    
    

    excelApplication.displayalerts = False
    
    
ED_HIN_GAI = Text1(ptxHin_Gai).Text
    
If Right(RTrim(ED_HIN_GAI), 1) = "." Then
'    Right(RTrim(ED_HIN_GAI), 1) = "_"

    For ED_I = 20 To 0 Step -1
        If Mid(ED_HIN_GAI, ED_I, 1) = "." Then
            Mid(ED_HIN_GAI, ED_I, 1) = "_"
            Exit For
        End If
    Next ED_I
    

End If
    
    
    
    excelWorkBook.saveas FileName:=(Save_Dir & Trim(ED_HIN_GAI))






    Set excelSheet = Nothing
    
    Set excelWorkBook = Nothing
    
    excelApplication.quit

    
    Set excelApplication = Nothing

    
S_END = Right(Format(Now, "hh:mm:ss"), 5)
    
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "S=" & S_Start & _
    " S.CRE=" & CREATE_EXCEL & _
    " S.BODY1=" & BODY1_EXCEL & _
    " S.BODY2=" & BODY2_EXCEL & _
    " S.BODY3=" & BODY3_EXCEL & _
    " S.INS1=" & INS1_EXCEL & _
    " S.INS2=" & INS2_EXCEL & _
    " S.INS3=" & INS3_EXCEL & _
    " S.TOTAL=" & TOTAL_EXCEL & _
    " S.FOOT=" & FOOT_EXCEL & _
    " S.VISIBLE=" & DSP_EXCEL & _
    " E=" & S_END, Me.hwnd, 0)
    
    
    
Call LOG_OUT(LOG_F, S_TITLE & "Hin=" & StrConv(ITEMREC.HIN_GAI, vbUnicode) & "S=" & S_Start & _
    " S.CRE=" & CREATE_EXCEL & _
    " S.BODY1=" & BODY1_EXCEL & _
    " S.BODY2=" & BODY2_EXCEL & _
    " S.BODY3=" & BODY3_EXCEL & _
    " S.INS1=" & INS1_EXCEL & _
    " S.INS2=" & INS2_EXCEL & _
    " S.INS3=" & INS3_EXCEL & _
    " S.TOTAL=" & TOTAL_EXCEL & _
    " S.FOOT=" & FOOT_EXCEL & _
    " S.VISIBLE=" & DSP_EXCEL & _
    " E=" & S_END)
    
    Call Input_UnLock
    
Debug.Print "out Estimate_Proc=" & Format(Now, "hh:mm:ss")
    
    
    
    Estimate_Proc = False
End Function

Private Function Detail_Disp_Proc(Errflg As Integer) As Integer
'----------------------------------------------------------------------------
'                   ���ݒl��ʕ\��
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim j           As Integer
Dim wkint       As Integer
Dim wkDouble    As Double

Dim wkKUSATU    As Variant
Dim c           As String * 128

Dim wkBikou     As String

Dim INV_F       As Boolean

Dim CATE_ST_SEC As Long


    Detail_Disp_Proc = True
    
    '�i�ڃ}�X�^�ǂݍ���
    Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHin_Gai).Text)

    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound

            Text1(ptxHin_Name).Text = ""
            Text1(ptxST_SOKO).Text = ""
            Text1(ptxST_RETU).Text = ""
            Text1(ptxST_REN).Text = ""
            Text1(ptxST_DAN).Text = ""
'2018.05.17            MsgBox "���͂������ڂ̓G���[�ł��B(�i��)"
            Errflg = True
            Detail_Disp_Proc = False
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Function

    End Select
    
    
    
    For i = 2 To 6      '2013.01.16 5-->6
        Command1(i).Enabled = True
    Next i
    
    
    '�i��
    Text1(ptxHin_Name).Text = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
    '�W���I��
    Text1(ptxST_SOKO).Text = StrConv(ITEMREC.ST_SOKO, vbUnicode)
    Text1(ptxST_RETU).Text = StrConv(ITEMREC.ST_RETU, vbUnicode)
    Text1(ptxST_REN).Text = StrConv(ITEMREC.ST_REN, vbUnicode)
    Text1(ptxST_DAN).Text = StrConv(ITEMREC.ST_DAN, vbUnicode)
    
    '�i���J�e�S���B
    Text1(ptxCATEGORY_CODE).Text = Trim(StrConv(ITEMREC.CATEGORY_CODE, vbUnicode))
    For i = 0 To Combo1(pcmbCATEGORY_Name).ListCount - 1
        If Trim(Text1(ptxCATEGORY_CODE).Text) = Trim(Right(Combo1(pcmbCATEGORY_Name).List(i), 8)) Then
            Text1(ptxCATEGORY_CODE).Text = Trim(Right(Combo1(pcmbCATEGORY_Name).List(i), 8))
            Combo1(pcmbCATEGORY_Name).ListIndex = i
            Exit For
        End If
    Next i
    If i > Combo1(pcmbCATEGORY_Name).ListCount - 1 Then
        Combo1(pcmbCATEGORY_Name).ListIndex = 0
    End If
    
    
    
    
    
    
    
    
    
    '-----------------------------------------------------------    2009.06.02 ��
    '���Ϗ����l
    wkBikou = Replace(StrConv(ITEMREC.M_BIKOU, vbUnicode), Chr(0), " ")
    RichTextBox1(prchM_BIKOU).Text = RTrim(wkBikou)
    
    '�d�l����
    Text1(ptxSHIYOU_NO).Text = RTrim(StrConv(ITEMREC.SHIYOU_NO, vbUnicode))
    
    '���ϋ敪
    Text1(ptxMITSUMORI_KBN).Text = RTrim(StrConv(ITEMREC.MITSUMORI_KBN, vbUnicode))
    '�P���ؑ֓�
    Text1(ptxTANKA_KIRIKAE_DT).Text = RTrim(StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode))
    '�ؑ֋敪
    Text1(ptxKIRIKAE_KBN).Text = RTrim(StrConv(ITEMREC.KIRIKAE_KBN, vbUnicode))

    '-----------------------------------------------------------    2009.06.02 ��
    
    
    
    
    
    '-----------------------------------    ���P��  2009.07.24
    
    
    '(����)���i���H��
    If IsNumeric(StrConv(ITEMREC.BEF_S_KOUSU_BAIKA, vbUnicode)) Then
        Text1(ptxOLD_S_KOUSU_BAIKA).Text = Format(StrConv(ITEMREC.BEF_S_KOUSU_BAIKA, vbUnicode), "#0.00")
    Else
        Text1(ptxOLD_S_KOUSU_BAIKA).Text = "0.00"
    End If
    
    '(����)���i���H��
    If IsNumeric(StrConv(ITEMREC.BEF_S_SHIZAI_BAIKA, vbUnicode)) Then
        Text1(ptxOLD_S_SHIZAI_BAIKA).Text = Format(StrConv(ITEMREC.BEF_S_SHIZAI_BAIKA, vbUnicode), "#0.00")
    Else
        Text1(ptxOLD_S_SHIZAI_BAIKA).Text = "0.00"
    End If
    
    '�O���P��
    If IsNumeric(StrConv(ITEMREC.BEF_S_GAISO_TANKA, vbUnicode)) Then
        Text1(ptxOLD_S_GAISO_TANKA).Text = Format(StrConv(ITEMREC.BEF_S_GAISO_TANKA, vbUnicode), "#0.00")
    Else
        Text1(ptxOLD_S_GAISO_TANKA).Text = "0.00"
    End If
    
    'PPSC���H�P��
    If IsNumeric(StrConv(ITEMREC.BEF_S_PPSC_KAKO_KOSU, vbUnicode)) Then
        Text1(ptxOLD_S_PPSC_KAKO_KOSU).Text = Format(StrConv(ITEMREC.BEF_S_PPSC_KAKO_KOSU, vbUnicode), "#0.00")
    Else
        Text1(ptxOLD_S_PPSC_KAKO_KOSU).Text = "0.00"
    End If
    
    'BU���H�P��
    If IsNumeric(StrConv(ITEMREC.BEF_S_BU_KAKO_KOSU, vbUnicode)) Then
        Text1(ptxOLD_S_BU_KAKO_KOSU).Text = Format(StrConv(ITEMREC.BEF_S_BU_KAKO_KOSU, vbUnicode), "#0.00")
    Else
        Text1(ptxOLD_S_BU_KAKO_KOSU).Text = "0.00"
    End If
'------2009.07.24
    
    
    
    
    
    
    
    
    '-----------------------------------    ���P��  2009.07.24
    
    
    
    
    '-----------------------------------    �ύX�O
    
    
    
    If IsNumeric(StrConv(ITEMREC.SEI_LOT, vbUnicode)) Then
        Text1(ptxBEF_SEI_LOT).Text = Format(Val(StrConv(ITEMREC.SEI_LOT, vbUnicode)), "#0")
    Else
'        Text1(ptxBEF_SEI_LOT).Text = "1"
        Text1(ptxBEF_SEI_LOT).Text = ""
    End If
    
    
    '��ڰ�
    If IsNumeric(StrConv(ITEMREC.SEI_RATE, vbUnicode)) Then
        Text1(ptxBEF_SEI_RATE).Text = Format(Val(StrConv(ITEMREC.SEI_RATE, vbUnicode)), "#0.00")
    Else
        
        Text1(ptxBEF_SEI_RATE).Text = ""
    End If
    
    
    
    
    
    '�H��
    If IsNumeric(StrConv(ITEMREC.S_KOUSU, vbUnicode)) Then
        Text1(ptxBEF_S_KOUSU).Text = Format(CDbl(StrConv(ITEMREC.S_KOUSU, vbUnicode)), "#0.00")
    Else
'        Text1(ptxBEF_S_KOUSU).Text = "0.0"
        Text1(ptxBEF_S_KOUSU).Text = ""
    End If
    '(����)�H��
    If IsNumeric(StrConv(ITEMREC.S_KOUSU_GENKA, vbUnicode)) Then
        Text1(ptxBEF_S_KOUSU_GENKA).Text = Format(CDbl(StrConv(ITEMREC.S_KOUSU_GENKA, vbUnicode)), "#0.00")
    Else
'        Text1(ptxBEF_S_KOUSU_GENKA).Text = "0.00"
        Text1(ptxBEF_S_KOUSU_GENKA).Text = ""
    End If
    '�H��
    If IsNumeric(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)) Then
        Text1(ptxBEF_S_KOUSU_BAIKA).Text = Format(CDbl(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)), "#0.00")
    Else
'        Text1(ptxBEF_S_KOUSU_BAIKA).Text = "0.00"
        Text1(ptxBEF_S_KOUSU_BAIKA).Text = ""
    End If
    '(����)����
    If IsNumeric(StrConv(ITEMREC.S_SHIZAI_GENKA, vbUnicode)) Then
        Text1(ptxBEF_S_SHIZAI_GENKA).Text = Format(CDbl(StrConv(ITEMREC.S_SHIZAI_GENKA, vbUnicode)), "#0.00")
    Else
'        Text1(ptxBEF_S_SHIZAI_GENKA).Text = "0.00"
        Text1(ptxBEF_S_SHIZAI_GENKA).Text = ""
    End If
    '����
    If IsNumeric(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)) Then
        Text1(ptxBEF_S_SHIZAI_BAIKA).Text = Format(CDbl(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)), "#0.00")
    Else
'        Text1(ptxBEF_S_SHIZAI_BAIKA).Text = "0.00"
        Text1(ptxBEF_S_SHIZAI_BAIKA).Text = ""
    End If
    
    
    
    '�O����
    If IsNumeric(StrConv(ITEMREC.S_GAISO_TANKA, vbUnicode)) Then
        Text1(ptxBEF_S_GAISO_TANKA).Text = Format(CDbl(StrConv(ITEMREC.S_GAISO_TANKA, vbUnicode)), "#0.00")
    Else
'        Text1(ptxBEF_S_GAISO_TANKA).Text = "0.00"
        Text1(ptxBEF_S_GAISO_TANKA).Text = ""
    End If
    
    
    'PPSC���H�P��
    If IsNumeric(StrConv(ITEMREC.S_PPSC_KAKO_KOSU, vbUnicode)) Then
        Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text = Format(CDbl(StrConv(ITEMREC.S_PPSC_KAKO_KOSU, vbUnicode)), "#0.00")
    Else
'        Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text = "0.00"
        Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text = ""
    End If
    'BU���H�P��
    If IsNumeric(StrConv(ITEMREC.S_BU_KAKO_KOSU, vbUnicode)) Then
        Text1(ptxBEF_S_BU_KAKO_KOSU).Text = Format(CDbl(StrConv(ITEMREC.S_BU_KAKO_KOSU, vbUnicode)), "#0.00")
    Else
        Text1(ptxBEF_S_BU_KAKO_KOSU).Text = ""
    End If
    
    
    
    
    
    
    '�ݒ��
    Text1(ptxBEF_S_KOUSU_SET_DATE).Text = Trim(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode))
    '�S����
    Text1(ptxBEF_SEI_TANKA_TANTO).Text = Trim(StrConv(ITEMREC.SEI_TANKA_TANTO, vbUnicode))
    '����
    Text1(ptxBEF_SE_TANKA_MEMO).Text = Trim(StrConv(ITEMREC.SE_TANKA_MEMO, vbUnicode))


    '-----------------------------------    �ύX��
    
    
    If Trim(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode)) = "" Then
        'ۯĐ�
        Text1(ptxAFT_SEI_LOT).Text = "1"
    Else
        'ۯĐ�
        If IsNumeric(StrConv(ITEMREC.SEI_LOT, vbUnicode)) Then
            Text1(ptxAFT_SEI_LOT).Text = Format(Val(StrConv(ITEMREC.SEI_LOT, vbUnicode)), "#0")
        Else
            Text1(ptxAFT_SEI_LOT).Text = "1"
        End If
    End If
    
    Text1(ptxAFT_SEI_RATE).Text = Format(Val(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)), "#0.00")
    
    
    
    '�ݒ��
    Text1(ptxAFT_S_KOUSU_SET_DATE).Text = ""
    '�S����
    Text1(ptxAFT_SEI_TANKA_TANTO).Text = Text1(ptxTanto_Code).Text
    '����
    Text1(ptxAFT_SE_TANKA_MEMO).Text = ""
    
    '-----------------------------------    �����Ϗo�א�
    If MONTHLYQTY_Disp_Proc() Then
        Exit Function
    End If
    
    '-----------------------------------    �\���i�\��
    If P_COMPO_Disp_Proc() Then
        Exit Function
    End If
    
    '-----------------------------------    ��ƍH��
    '�@
    
    If IsNumeric(StrConv(P_KANRIREC02.MAIN_KOTEI(0).KOTEI, vbUnicode)) Then
        
        Text1(ptxMAIN_KOUTEI_TANI01).Text = Format(CInt(StrConv(P_KANRIREC02.MAIN_KOTEI(0).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxMAIN_KOUTEI_TANI01).Text = "0"
    End If
    
    
    
    
    If IsNumeric(StrConv(ITEMREC.SEI_LABEL_QTY, vbUnicode)) Then
        '2009.09.18
        If IsDate(Mid(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode), 1, 4) & "/" & Mid(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode), 5, 2) & "/" & Mid(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode), 7, 4)) Then
            Text1(ptxMAIN_KOUTEI_QTY01).Text = Format(CInt(StrConv(ITEMREC.SEI_LABEL_QTY, vbUnicode)), "#0")
        Else
            Text1(ptxMAIN_KOUTEI_QTY01).Text = "1"
        End If
    Else
        Text1(ptxMAIN_KOUTEI_QTY01).Text = "1"
    End If
    Text1(ptxMAIN_KOUTEI_KOUSU01).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI01).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY01).Text), "#0")
    
    
    
    
    
    '�A
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    
                    If IsNumeric(KOUSEI(i, ColS_KOUSU)) Then
                    
                        wkint = wkint + CInt(ToRoundUp(CCur(KOUSEI(i, ColS_KOUSU)) * CDbl(KOUSEI(i, ColKO_QTY)), 0))
                    
                    End If
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    Text1(ptxMAIN_KOUTEI_TANI02).Text = Format(wkint, "#0")
    Text1(ptxMAIN_KOUTEI_QTY02).Text = "1"
    Text1(ptxMAIN_KOUTEI_KOUSU02).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI02).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY02).Text), "#0")
    '�B
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(DOUKON_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = DOUKON_T(j) Then
                    
                    
                    If IsNumeric(KOUSEI(i, ColKO_QTY)) Then
                        wkint = wkint + CInt(ToRoundUp(CCur(KOUSEI(i, ColKO_QTY)), 0))
                    End If
                    
                    
                    
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    If IsNumeric(StrConv(P_KANRIREC02.MAIN_KOTEI(3).KOTEI, vbUnicode)) Then
        Text1(ptxMAIN_KOUTEI_TANI03).Text = Format(CInt(StrConv(P_KANRIREC02.MAIN_KOTEI(3).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxMAIN_KOUTEI_TANI03).Text = "0"
    End If
    Text1(ptxMAIN_KOUTEI_QTY03).Text = Format(wkint, "#0")
    Text1(ptxMAIN_KOUTEI_KOUSU03).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI03).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY03).Text), "#0")
    '�C
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(KAKOU_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = KAKOU_T(j) Then
                    
                    
                    If IsNumeric(KOUSEI(i, ColS_KOUSU)) Then
                        wkint = wkint + CInt(KOUSEI(i, ColS_KOUSU))
                    End If
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    Text1(ptxMAIN_KOUTEI_TANI04).Text = Format(wkint, "#0")
    Text1(ptxMAIN_KOUTEI_QTY04).Text = "1"
    Text1(ptxMAIN_KOUTEI_KOUSU04).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI04).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY04).Text), "#0")
    '�D
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
            
            
            For j = 0 To UBound(SHIZAI_T)
            
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    If IsNumeric(KOUSEI(i, ColSEI_SYU_KON)) Then
                        wkint = wkint + CInt(KOUSEI(i, ColSEI_SYU_KON))
                    End If
                End If
            
            Next j
            
            
            
            
        Next i
    End If
    Text1(ptxMAIN_KOUTEI_TANI05).Text = Format(wkint, "#0")
    Text1(ptxMAIN_KOUTEI_QTY05).Text = "1"
    Text1(ptxMAIN_KOUTEI_KOUSU05).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI05).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY05).Text), "#0")
    '�v
    wkint = 0
    For i = ptxMAIN_KOUTEI_KOUSU01 To ptxMAIN_KOUTEI_KOUSU09 Step 3
    
        If IsNumeric(Text1(i).Text) Then
            wkint = wkint + CInt(Text1(i).Text)
        End If
    Next i
    Text1(ptxMAIN_KOUTEI_KEI1).Text = Format(wkint, "#0")
    
    '-----------------------------------    �O��H��
    If CATEGORY_Disp_Proc() Then
        Exit Function
    End If
    
    '�i�ڃ}�X�^�ǂݍ���
    Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHin_Gai).Text)

    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound

            Text1(ptxHin_Name).Text = ""
            Text1(ptxST_SOKO).Text = ""
            Text1(ptxST_RETU).Text = ""
            Text1(ptxST_REN).Text = ""
            Text1(ptxST_DAN).Text = ""
            MsgBox "���͂������ڂ̓G���[�ł��B(�i��)"
            Detail_Disp_Proc = False
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Function

    End Select
    
    
    
'>>>>>>>>>>>>>>>    2012.01.24
    Text1(ptxCATE_ST_FUKA).Text = ""
    Text1(ptxCATE_AD_FUKA).Text = ""


    If IsNumeric(StrConv(ITEMREC.CATE_ST_FUKA, vbUnicode)) Then
'>>>>>>>>>>>>>>>    2012.01.24
'        Text1(ptxCATE_AD_FUKA).Text = Val(StrConv(ITEMREC.CATE_AD_FUKA, vbUnicode))
        Text1(ptxCATE_ST_FUKA).Text = Format(Val(StrConv(ITEMREC.CATE_ST_FUKA, vbUnicode)), "#")
    Else
        Text1(ptxCATE_ST_FUKA).Text = ""
    End If
    If IsNumeric(StrConv(ITEMREC.CATE_AD_FUKA, vbUnicode)) Then
        Text1(ptxCATE_AD_FUKA).Text = Format(Val(StrConv(ITEMREC.CATE_AD_FUKA, vbUnicode)), "#")
    Else
        If IsNumeric(StrConv(ITEMREC.CATE_ST_FUKA, vbUnicode)) Then
            Text1(ptxCATE_AD_FUKA).Text = Format(Val(StrConv(ITEMREC.CATE_ST_FUKA, vbUnicode)), "#")
        Else
            Text1(ptxCATE_AD_FUKA).Text = ""
        End If
    End If

'>>>>>>>>>>>>>>>    2012.01.24
    
    
    
    Call CATEGORY_KEISAN_PROC
    
    '�H��
    Text1(ptxAFT_S_KOUSU).Text = Text1(ptxCATE_AD_FUN).Text
    '�H��
    If IsNumeric(Text1(ptxCATE_AD_KOURYO).Text) Then
        Text1(ptxAFT_S_KOUSU_BAIKA).Text = Format(CDbl(Text1(ptxCATE_AD_KOURYO).Text), "#0.00")
    Else
        Text1(ptxAFT_S_KOUSU_BAIKA).Text = ""
    End If
    
    '-----------------------------------    �ύX�O�^�ύX��i�W�v�l�j
    
    
'    '�H��
    Text1(ptxAFT_S_KOUSU).Text = Text1(ptxCATE_AD_FUN).Text
'    '�H��
    If IsNumeric(Text1(ptxCATE_AD_KOURYO).Text) Then
        Text1(ptxAFT_S_KOUSU_BAIKA).Text = Format(CDbl(Text1(ptxCATE_AD_KOURYO).Text), "#0.00")
    Else
        Text1(ptxAFT_S_KOUSU_BAIKA).Text = ""
    End If
'
'    '����
    wkDouble = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
    
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    If IsNumeric(KOUSEI(i, ColG_ST_SHIKIN)) Then
                        wkDouble = wkDouble + CDbl(KOUSEI(i, ColG_ST_SHIKIN))
                    End If
                    Exit For
                End If
    
            Next j
    
        Next i
    End If
    Text1(ptxAFT_S_SHIZAI_GENKA).Text = Format(wkDouble, "#0.00")






    wkDouble = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
    
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    If IsNumeric(KOUSEI(i, ColG_ST_URIKIN)) Then
                        wkDouble = wkDouble + CDbl(KOUSEI(i, ColG_ST_URIKIN))
                    End If
                    Exit For
                End If
    
            Next j
    
        Next i
    End If
    Text1(ptxAFT_S_SHIZAI_BAIKA).Text = Format(wkDouble, "#0.00")

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  ������ʒP��    2012.01.04
    
    Call UniCode_Conv(K0_ITEM_CATEGORY.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_ITEM_CATEGORY.CATEGORY_CODE, Text1(ptxCATEGORY_CODE).Text)
        
    sts = BTRV(BtOpGetEqual, ITEM_CATEGORY_POS, ITEM_CATEGORYREC, Len(ITEM_CATEGORYREC), K0_ITEM_CATEGORY, Len(K0_ITEM_CATEGORY), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound

            Call UniCode_Conv(ITEM_CATEGORYREC.TOKU_TANKA_KOURYO, "")
            Call UniCode_Conv(ITEM_CATEGORYREC.TOKU_TANKA_HAKO, "")

        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i���J�e�S���}�X�^")
            Unload Me

    End Select
    
    
    
    If IsNumeric(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_HAKO, vbUnicode)) Then
        Text1(ptxAFT_S_SHIZAI_BAIKA).Text = Format(Val(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_HAKO, vbUnicode)), "0.00")
    Else
        If IsNumeric(Text1(ptxAFT_S_SHIZAI_BAIKA).Text) Then
            Text1(ptxAFT_S_KOUSU_BAIKA).Text = Format(CDbl(Text1(ptxCATE_AD_KOURYO).Text), "#0.00")
        Else
            Text1(ptxAFT_S_KOUSU_BAIKA).Text = ""
        End If
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  ������ʒP��    2012.01.04




    '�O������
    wkDouble = 0
    If KUSATU_F Then
        If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
        Else
            For i = 1 To KOUSEI.UpperBound(1)
        
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = GAISO_KBN Then
            
            
                    If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
            
                        wkDouble = wkDouble + CDbl(KOUSEI(i, ColG_ST_URIKIN_KUSATU))
                    End If
            
                End If
        
        
            Next i
        End If
    End If
    Text1(ptxAFT_S_GAISO_TANKA).Text = Format(wkDouble, "#0.00")






    'PPSC����
    wkDouble = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
    
            For j = 0 To UBound(PPSC_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = PPSC_T(j) Then
                    
                    If IsNumeric(KOUSEI(i, ColS_KOUSU)) Then
                        wkDouble = wkDouble + CDbl(ToRoundUp(CCur(CDbl(KOUSEI(i, ColS_KOUSU)) / 60 * CInt(Text1(ptxAFT_SEI_RATE).Text)), 2))
                    End If
                    Exit For
                End If
    
            Next j
    
        Next i
        Text1(ptxAFT_S_PPSC_KAKO_KOSU).Text = Format(wkDouble, "#0.00")
    End If



    'BU����
    wkDouble = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
    
            For j = 0 To UBound(BU_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = BU_T(j) Then
                    
                    If IsNumeric(KOUSEI(i, ColS_KOUSU)) Then
                        wkDouble = wkDouble + CDbl(ToRoundUp(CCur(CDbl(KOUSEI(i, ColS_KOUSU)) / 60 * CInt(Text1(ptxAFT_SEI_RATE).Text)), 2))
                    End If
                    Exit For
                End If
    
            Next j
    
        Next i
    
        Text1(ptxAFT_S_BU_KAKO_KOSU).Text = Format(wkDouble, "#0.00")
    
    End If





    Detail_Disp_Proc = False

End Function

Private Function MONTHLYQTY_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   �����Ϗo�א���ʕ\��
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim Total       As Long

Dim S_YM        As String * 6
Dim E_YM        As String * 6
Dim GET_YM      As String * 6


Dim NOW_YM      As String * 6

Dim cVer1       As String
Dim cVer2       As String

Dim cHEX        As String

Dim i           As Integer
Dim j           As Integer
Dim k           As Integer
Dim MONTH_Cnt   As Integer
Dim MONTH_QTY   As Long


    MONTHLYQTY_Disp_Proc = True
    
    
    
    NOW_YM = Left(Format(Now, "YYYYMMDD"), 6)
    
    
    '�O�N�x�Ώ۔N��
    If Right(NOW_YM, 2) < "04" Then
        S_YM = Format(CInt(Left(NOW_YM, 4) - 2), "0000") & "04"
    Else
        S_YM = Format(CInt(Left(NOW_YM, 4) - 1), "0000") & "04"
    End If
    
    
    '�����Ϗo�א� (���ʏW�v)�ǂݍ��݁��W�v
    Total = 0
    
    j = ptxZEN_SYUKAQTY04
    
    
    For i = 0 To 11
    
        DoEvents
    
            
            
    
        GET_YM = Left(S_YM, 4) + Format(CInt(Right(S_YM, 2)) + i, "00")
        If Right(GET_YM, 2) > "12" Then
            GET_YM = Format(CInt(Left(GET_YM, 4)) + 1, "0000") & Format(CInt(Right(GET_YM, 2)) - 12, "00")
        End If
    
    
        Call UniCode_Conv(K0_MONTHLYQTY.DT, GET_YM)
        Call UniCode_Conv(K0_MONTHLYQTY.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
        Call UniCode_Conv(K0_MONTHLYQTY.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
        Call UniCode_Conv(K0_MONTHLYQTY.HIN_GAI, Text1(ptxHin_Gai).Text)
        
        
    
        sts = BTRV(BtOpGetEqual, MONTHLYQTY_POS, MONTHLYQTYREC, Len(MONTHLYQTYREC), K0_MONTHLYQTY, Len(K0_MONTHLYQTY), 0)
        Select Case sts
            Case BtNoErr
            
            
                cVer1 = ""
                For k = 0 To UBound(MONTHLYQTYREC.SyukaQty)
                
                    cHEX = Hex(MONTHLYQTYREC.SyukaQty(k))
                    If Len(cHEX) < 2 Then
                        cHEX = "0" & cHEX
                    End If
                            
                    cVer1 = cVer1 & cHEX
                
                Next k
                MONTH_QTY = CLng(Left(cVer1, 9))
                    
                Text1(j).Text = Format(MONTH_QTY, "#,##0")
                Total = Total + MONTH_QTY
            
            
            
            Case BtErrKeyNotFound
                Text1(j).Text = "0"
Debug.Print j & " " & Text1(j).Text
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�����Ϗo�א� (���ʏW�v)")
                Exit Function
    
        End Select
        
    
        j = j + 1
    
    Next i
    
    
    Total = ToRoundUp(CCur(Total / 12), 0)
    Text1(ptxZEN_AVE).Text = Format(Total, "#,##0")
    
    
    
    
    
    
    
    '���N�x�Ώ۔N��
    If Right(NOW_YM, 2) < "04" Then
        S_YM = Format(CInt(Left(NOW_YM, 4) - 1), "0000") & "04"
    Else
        S_YM = Left(NOW_YM, 4) & "04"
    End If
    
    E_YM = Left(Format(DateAdd("m", -1, Left(Format(Now, "YYYY/MM/DD"), 7) & "/01"), "YYYYMMDD"), 6)
    
    
    
    
    
    '�����Ϗo�א� (���ʏW�v)�ǂݍ��݁��W�v
    Total = 0
    MONTH_Cnt = 0
    j = ptxTOU_SYUKAQTY04
    
    
    For i = 0 To 11
    
        DoEvents
    
            
            
    
        GET_YM = Left(S_YM, 4) + Format(CInt(Right(S_YM, 2)) + i, "00")
        If Right(GET_YM, 2) > "12" Then
            GET_YM = Format(CInt(Left(GET_YM, 4)) + 1, "0000") & Format(CInt(Right(GET_YM, 2)) - 12, "00")
        End If
    
        If GET_YM > E_YM Then
            Exit For
        End If
    
        Call UniCode_Conv(K0_MONTHLYQTY.DT, GET_YM)
        Call UniCode_Conv(K0_MONTHLYQTY.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
        Call UniCode_Conv(K0_MONTHLYQTY.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
        Call UniCode_Conv(K0_MONTHLYQTY.HIN_GAI, Text1(ptxHin_Gai).Text)
        
        
    
        sts = BTRV(BtOpGetEqual, MONTHLYQTY_POS, MONTHLYQTYREC, Len(MONTHLYQTYREC), K0_MONTHLYQTY, Len(K0_MONTHLYQTY), 0)
        Select Case sts
            Case BtNoErr
            
            
                cVer1 = ""
                For k = 0 To UBound(MONTHLYQTYREC.SyukaQty)
                
                    cHEX = Hex(MONTHLYQTYREC.SyukaQty(k))
                    If Len(cHEX) < 2 Then
                        cHEX = "0" & cHEX
                    End If
                            
                    cVer1 = cVer1 & cHEX
                
                Next k
                MONTH_QTY = CLng(Left(cVer1, 9))
                    
                Text1(j).Text = Format(MONTH_QTY, "#,##0")
                Total = Total + MONTH_QTY
            
            
            
            Case BtErrKeyNotFound
                Text1(j).Text = "0"
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�����Ϗo�א� (���ʏW�v)")
                Exit Function
    
        End Select
        
        MONTH_Cnt = MONTH_Cnt + 1
    
        j = j + 1
    
    Next i
    
    
    If MONTH_Cnt = 0 Then
        Total = 0
    Else
        Total = ToRoundUp(CCur(Total / MONTH_Cnt), 0)
        
    End If
    Text1(ptxTOU_AVE).Text = Format(Total, "#,##0")
    
    
    
    
    
    
    
    
    MONTHLYQTY_Disp_Proc = False

End Function
Private Function TANKA_KEISAN_Proc() As Integer
'----------------------------------------------------------------------------
'                   �P���v�Z����
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim j           As Integer
Dim wkint       As Integer
Dim wkDouble    As Double


Dim c           As String * 128
Dim wkKUSATU    As Variant
Dim INV_F       As Boolean


    TANKA_KEISAN_Proc = True
    
    '�i�ڃ}�X�^�ǂݍ���
    Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHin_Gai).Text)

    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound

            Text1(ptxHin_Name).Text = ""
            Text1(ptxST_SOKO).Text = ""
            Text1(ptxST_RETU).Text = ""
            Text1(ptxST_REN).Text = ""
            Text1(ptxST_DAN).Text = ""
            MsgBox "���͂������ڂ̓G���[�ł��B(�i��)"
            TANKA_KEISAN_Proc = False
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Function

    End Select


    '�ݒ��
    Text1(ptxAFT_S_KOUSU_SET_DATE).Text = Format(Now, "YYYYMMDD")
    '�S����
    Text1(ptxAFT_SEI_TANKA_TANTO).Text = Text1(ptxTanto_Code).Text
    
    
    '-----------------------------------    ��ƍH��
    '�@
    If IsNumeric(StrConv(P_KANRIREC02.MAIN_KOTEI(0).KOTEI, vbUnicode)) Then
        Text1(ptxMAIN_KOUTEI_TANI01).Text = Format(CInt(StrConv(P_KANRIREC02.MAIN_KOTEI(0).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxMAIN_KOUTEI_TANI01).Text = "0"
    End If
'    Text1(ptxMAIN_KOUTEI_QTY01).Text = "1"
    If Trim(Text1(ptxMAIN_KOUTEI_QTY01).Text) = "" Then
        If IsNumeric(StrConv(ITEMREC.SEI_LABEL_QTY, vbUnicode)) Then
                Text1(ptxMAIN_KOUTEI_QTY01).Text = Format(CInt(StrConv(ITEMREC.SEI_LABEL_QTY, vbUnicode)), "#0")
        Else
                Text1(ptxMAIN_KOUTEI_QTY01).Text = "1"
        End If
    End If
    Text1(ptxMAIN_KOUTEI_KOUSU01).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI01).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY01).Text), "#0")
    
    '�A
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    
                    If IsNumeric(KOUSEI(i, ColS_KOUSU)) Then
                        wkint = wkint + CInt(ToRoundUp(CCur(KOUSEI(i, ColS_KOUSU)) * CDbl(KOUSEI(i, ColKO_QTY)), 0))
                    End If
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    Text1(ptxMAIN_KOUTEI_TANI02).Text = Format(wkint, "#0")
    Text1(ptxMAIN_KOUTEI_QTY02).Text = "1"
    Text1(ptxMAIN_KOUTEI_KOUSU02).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI02).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY02).Text), "#0")
    '�B
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(DOUKON_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = DOUKON_T(j) Then
                    
                    If IsNumeric(KOUSEI(i, ColKO_QTY)) Then
                        wkint = wkint + CInt(ToRoundUp(CCur(KOUSEI(i, ColKO_QTY)), 0))
                    End If
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    If IsNumeric(StrConv(P_KANRIREC02.MAIN_KOTEI(3).KOTEI, vbUnicode)) Then
        Text1(ptxMAIN_KOUTEI_TANI03).Text = Format(CInt(StrConv(P_KANRIREC02.MAIN_KOTEI(3).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxMAIN_KOUTEI_TANI03).Text = "0"
    End If
    Text1(ptxMAIN_KOUTEI_QTY03).Text = Format(wkint, "#0")
    Text1(ptxMAIN_KOUTEI_KOUSU03).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI03).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY03).Text), "#0")
    '�C
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(KAKOU_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = KAKOU_T(j) Then
                    If IsNumeric(KOUSEI(i, ColS_KOUSU)) Then
                        wkint = wkint + CInt(KOUSEI(i, ColS_KOUSU))
                    End If
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    Text1(ptxMAIN_KOUTEI_TANI04).Text = Format(wkint, "#0")
    Text1(ptxMAIN_KOUTEI_QTY04).Text = "1"
    Text1(ptxMAIN_KOUTEI_KOUSU04).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI04).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY04).Text), "#0")
    '�D
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
            
            
            For j = 0 To UBound(SHIZAI_T)
            
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    If IsNumeric(KOUSEI(i, ColSEI_SYU_KON)) Then
                        wkint = wkint + CInt(KOUSEI(i, ColSEI_SYU_KON))
                    End If
                End If
            
            Next j
            
            
            
            
        Next i
    End If
    Text1(ptxMAIN_KOUTEI_TANI05).Text = Format(wkint, "#0")
    Text1(ptxMAIN_KOUTEI_QTY05).Text = "1"
    Text1(ptxMAIN_KOUTEI_KOUSU05).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI05).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY05).Text), "#0")
    '�v
    wkint = 0
    For i = ptxMAIN_KOUTEI_KOUSU01 To ptxMAIN_KOUTEI_KOUSU09 Step 3
    
        If IsNumeric(Text1(i).Text) Then
            wkint = wkint + CInt(Text1(i).Text)
        End If
    Next i
    Text1(ptxMAIN_KOUTEI_KEI1).Text = Format(wkint, "#0")
    
    '-----------------------------------    �i���J�e�S���B�v�Z
    
    Call CATEGORY_KEISAN_PROC
    '�H��
    Text1(ptxAFT_S_KOUSU).Text = Text1(ptxCATE_AD_FUN).Text
    '�H��
    If IsNumeric(Text1(ptxCATE_AD_KOURYO).Text) Then
        Text1(ptxAFT_S_KOUSU_BAIKA).Text = Format(CDbl(Text1(ptxCATE_AD_KOURYO).Text), "#0.00")
    Else
        Text1(ptxAFT_S_KOUSU_BAIKA).Text = ""
    End If
    
    
    
    '-----------------------------------    �ύX��
'    '����
    wkDouble = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
    
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    If IsNumeric(KOUSEI(i, ColG_ST_SHIKIN)) Then
                        wkDouble = wkDouble + CDbl(KOUSEI(i, ColG_ST_SHIKIN))
                    End If
                    Exit For
                End If
    
            Next j
    
        Next i
    End If
    Text1(ptxAFT_S_SHIZAI_GENKA).Text = Format(wkDouble, "#0.00")
'
    wkDouble = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
    
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    
                    If IsNumeric(KOUSEI(i, ColG_ST_URIKIN)) Then
                        wkDouble = wkDouble + CDbl(KOUSEI(i, ColG_ST_URIKIN))
                    End If
                    Exit For
                End If
    
            Next j
    
        Next i
    End If
    Text1(ptxAFT_S_SHIZAI_BAIKA).Text = Format(wkDouble, "#0.00")


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  ������ʒP��    2012.01.04
    If IsNumeric(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_HAKO, vbUnicode)) Then
        Text1(ptxAFT_S_SHIZAI_BAIKA).Text = Format(Val(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_HAKO, vbUnicode)), "0.00")
    Else
        If IsNumeric(Text1(ptxAFT_S_SHIZAI_BAIKA).Text) Then
            Text1(ptxAFT_S_KOUSU_BAIKA).Text = Format(CDbl(Text1(ptxCATE_AD_KOURYO).Text), "#0.00")
        Else
            Text1(ptxAFT_S_KOUSU_BAIKA).Text = ""
        End If
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  ������ʒP��    2012.01.04




    '�O������
    wkDouble = 0
    If KUSATU_F Then
        If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
        Else
            For i = 1 To KOUSEI.UpperBound(1)
        
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = GAISO_KBN Then
            
            
'                    If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
                    If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
            
                        wkDouble = wkDouble + CDbl(KOUSEI(i, ColG_ST_URIKIN_KUSATU))
                    End If
            
                End If
        
        
            Next i
        End If
    End If
    Text1(ptxAFT_S_GAISO_TANKA).Text = Format(wkDouble, "#0.00")




    'PPSC����   2011.06.23
    wkDouble = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
    
            For j = 0 To UBound(PPSC_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = PPSC_T(j) Then
                    
                    If IsNumeric(KOUSEI(i, ColS_KOUSU)) Then
                        wkDouble = wkDouble + CDbl(ToRoundUp(CCur(CDbl(KOUSEI(i, ColS_KOUSU)) / 60 * CInt(Text1(ptxAFT_SEI_RATE).Text)), 2))
                    End If
                    Exit For
                
                End If
    
            Next j
    
        Next i
        Text1(ptxAFT_S_PPSC_KAKO_KOSU).Text = Format(wkDouble, "#0.00")
    End If



    'BU����
    wkDouble = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
    
            For j = 0 To UBound(BU_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = BU_T(j) Then
                    
                    If IsNumeric(KOUSEI(i, ColS_KOUSU)) Then
                        wkDouble = wkDouble + CDbl(ToRoundUp(CCur(CDbl(KOUSEI(i, ColS_KOUSU)) / 60 * CInt(Text1(ptxAFT_SEI_RATE).Text)), 2))
                    End If
                    Exit For
                End If
    
            Next j
    
        Next i
    
        Text1(ptxAFT_S_BU_KAKO_KOSU).Text = Format(wkDouble, "#0.00")
    
    End If



    TANKA_KEISAN_Proc = False

End Function


Private Function KARI_TANKA_KEISAN_Proc() As Integer
'----------------------------------------------------------------------------
'                   ���@�P���v�Z����
'       2013.01.16
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim j           As Integer
Dim wkint       As Integer
Dim wkDouble    As Double


Dim c           As String * 128
Dim wkKUSATU    As Variant
Dim INV_F       As Boolean


    KARI_TANKA_KEISAN_Proc = True
    
    '�i�ڃ}�X�^�ǂݍ���
    Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHin_Gai).Text)

    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound

            Text1(ptxHin_Name).Text = ""
            Text1(ptxST_SOKO).Text = ""
            Text1(ptxST_RETU).Text = ""
            Text1(ptxST_REN).Text = ""
            Text1(ptxST_DAN).Text = ""
            MsgBox "���͂������ڂ̓G���[�ł��B(�i��)"
            KARI_TANKA_KEISAN_Proc = False
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Function

    End Select



    Set TDBGrid1(pGrdKOUSEI).Array = KOUSEI         '2013.03.27
    TDBGrid1(pGrdKOUSEI).Update                     '2013.03.27

    '-----------------------------------    ��ƍH��
    '�@
    If IsNumeric(StrConv(P_KANRIREC02.MAIN_KOTEI(0).KOTEI, vbUnicode)) Then
        Text1(ptxMAIN_KOUTEI_TANI01).Text = Format(CInt(StrConv(P_KANRIREC02.MAIN_KOTEI(0).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxMAIN_KOUTEI_TANI01).Text = "0"
    End If
'    Text1(ptxMAIN_KOUTEI_QTY01).Text = "1"
    If Trim(Text1(ptxMAIN_KOUTEI_QTY01).Text) = "" Then
        If IsNumeric(StrConv(ITEMREC.SEI_LABEL_QTY, vbUnicode)) Then
                Text1(ptxMAIN_KOUTEI_QTY01).Text = Format(CInt(StrConv(ITEMREC.SEI_LABEL_QTY, vbUnicode)), "#0")
        Else
                Text1(ptxMAIN_KOUTEI_QTY01).Text = "1"
        End If
    End If
    Text1(ptxMAIN_KOUTEI_KOUSU01).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI01).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY01).Text), "#0")
    
    '�A
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    
                    If IsNumeric(KOUSEI(i, ColS_KOUSU)) Then
                        wkint = wkint + CInt(ToRoundUp(CCur(KOUSEI(i, ColS_KOUSU)) * CDbl(KOUSEI(i, ColKO_QTY)), 0))
                    End If
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    Text1(ptxMAIN_KOUTEI_TANI02).Text = Format(wkint, "#0")
    Text1(ptxMAIN_KOUTEI_QTY02).Text = "1"
    Text1(ptxMAIN_KOUTEI_KOUSU02).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI02).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY02).Text), "#0")
    '�B
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(DOUKON_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = DOUKON_T(j) Then
                    
                    If IsNumeric(KOUSEI(i, ColKO_QTY)) Then
                        wkint = wkint + CInt(ToRoundUp(CCur(KOUSEI(i, ColKO_QTY)), 0))
                    End If
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    If IsNumeric(StrConv(P_KANRIREC02.MAIN_KOTEI(3).KOTEI, vbUnicode)) Then
        Text1(ptxMAIN_KOUTEI_TANI03).Text = Format(CInt(StrConv(P_KANRIREC02.MAIN_KOTEI(3).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxMAIN_KOUTEI_TANI03).Text = "0"
    End If
    Text1(ptxMAIN_KOUTEI_QTY03).Text = Format(wkint, "#0")
    Text1(ptxMAIN_KOUTEI_KOUSU03).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI03).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY03).Text), "#0")
    '�C
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(KAKOU_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = KAKOU_T(j) Then
                    If IsNumeric(KOUSEI(i, ColS_KOUSU)) Then
                        wkint = wkint + CInt(KOUSEI(i, ColS_KOUSU))
                    End If
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    Text1(ptxMAIN_KOUTEI_TANI04).Text = Format(wkint, "#0")
    Text1(ptxMAIN_KOUTEI_QTY04).Text = "1"
    Text1(ptxMAIN_KOUTEI_KOUSU04).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI04).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY04).Text), "#0")
    '�D
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
            
            
            For j = 0 To UBound(SHIZAI_T)
            
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    If IsNumeric(KOUSEI(i, ColSEI_SYU_KON)) Then
                        wkint = wkint + CInt(KOUSEI(i, ColSEI_SYU_KON))
                    End If
                End If
            
            Next j
            
            
            
            
        Next i
    End If
    Text1(ptxMAIN_KOUTEI_TANI05).Text = Format(wkint, "#0")
    Text1(ptxMAIN_KOUTEI_QTY05).Text = "1"
    Text1(ptxMAIN_KOUTEI_KOUSU05).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI05).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY05).Text), "#0")
    '�v
    wkint = 0
    For i = ptxMAIN_KOUTEI_KOUSU01 To ptxMAIN_KOUTEI_KOUSU09 Step 3
    
        If IsNumeric(Text1(i).Text) Then
            wkint = wkint + CInt(Text1(i).Text)
        End If
    Next i
    Text1(ptxMAIN_KOUTEI_KEI1).Text = Format(wkint, "#0")
    
    
    
    '����ƍH��1�@�b                                                        2013.03.27
    Text1(ptxCATE_ST_JITU1).Text = Val(Text1(ptxMAIN_KOUTEI_KEI1).Text)     '2013.03.27
    Text1(ptxCATE_AD_JITU1).Text = Val(Text1(ptxMAIN_KOUTEI_KEI1).Text)     '2013.03.27
    
    
    
    '����ƍH��2�@�b                                                        2013.03.27
    If IsNumeric(Text1(ptxMAIN_KOUTEI_KEI1).Text) And _
        IsNumeric(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) Then
    
        Text1(ptxCATE_ST_JITU2).Text = ToHalfAdjust(CCur(CCur(Text1(ptxMAIN_KOUTEI_KEI1).Text) * _
                                                    CCur(CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)))), 0)
        Text1(ptxCATE_AD_JITU2).Text = ToHalfAdjust(CCur(CCur(Text1(ptxMAIN_KOUTEI_KEI1).Text) * _
                                                    CCur(CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)))), 0)
    End If
    '����ƍH��2�@�b                                                        2013.03.27
    
    
    '-----------------------------------    �i���J�e�S���B�v�Z
    
    Call CATEGORY_KEISAN_PROC
    '�H��
    Text1(ptxAFT_S_KOUSU).Text = Text1(ptxCATE_AD_FUN).Text
    '�H��
    If IsNumeric(Text1(ptxCATE_AD_KOURYO).Text) Then
        Text1(ptxAFT_S_KOUSU_BAIKA).Text = Format(CDbl(Text1(ptxCATE_AD_KOURYO).Text), "#0.00")
    Else
        Text1(ptxAFT_S_KOUSU_BAIKA).Text = ""
    End If
    
    
    
    '-----------------------------------    �ύX��
'    '����
    wkDouble = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
    
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    If IsNumeric(KOUSEI(i, ColG_ST_SHIKIN)) Then
                        wkDouble = wkDouble + CDbl(KOUSEI(i, ColG_ST_SHIKIN))
                    End If
                    Exit For
                End If
    
            Next j
    
        Next i
    End If
    Text1(ptxAFT_S_SHIZAI_GENKA).Text = Format(wkDouble, "#0.00")
'
    wkDouble = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
    
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    
                    If IsNumeric(KOUSEI(i, ColG_ST_URIKIN)) Then
                        wkDouble = wkDouble + CDbl(KOUSEI(i, ColG_ST_URIKIN))
                    End If
                    Exit For
                End If
    
            Next j
    
        Next i
    End If
    Text1(ptxAFT_S_SHIZAI_BAIKA).Text = Format(wkDouble, "#0.00")


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  ������ʒP��    2012.01.04
    If IsNumeric(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_HAKO, vbUnicode)) Then
        Text1(ptxAFT_S_SHIZAI_BAIKA).Text = Format(Val(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_HAKO, vbUnicode)), "0.00")
    Else
        If IsNumeric(Text1(ptxAFT_S_SHIZAI_BAIKA).Text) Then
            Text1(ptxAFT_S_KOUSU_BAIKA).Text = Format(CDbl(Text1(ptxCATE_AD_KOURYO).Text), "#0.00")
        Else
            Text1(ptxAFT_S_KOUSU_BAIKA).Text = ""
        End If
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  ������ʒP��    2012.01.04




    '�O������
    wkDouble = 0
    If KUSATU_F Then
        If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
        Else
            For i = 1 To KOUSEI.UpperBound(1)
        
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = GAISO_KBN Then
            
            
'                    If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
                    If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
            
                        wkDouble = wkDouble + CDbl(KOUSEI(i, ColG_ST_URIKIN_KUSATU))
                    End If
            
                End If
        
        
            Next i
        End If
    End If
    Text1(ptxAFT_S_GAISO_TANKA).Text = Format(wkDouble, "#0.00")




    'PPSC����   2011.06.23
    wkDouble = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
    
            For j = 0 To UBound(PPSC_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = PPSC_T(j) Then
                    
                    If IsNumeric(KOUSEI(i, ColS_KOUSU)) Then
                        wkDouble = wkDouble + CDbl(ToRoundUp(CCur(CDbl(KOUSEI(i, ColS_KOUSU)) / 60 * CInt(Text1(ptxAFT_SEI_RATE).Text)), 2))
                    End If
                    Exit For
                End If
    
            Next j
    
        Next i
        Text1(ptxAFT_S_PPSC_KAKO_KOSU).Text = Format(wkDouble, "#0.00")
    End If



    'BU����
    wkDouble = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
    
            For j = 0 To UBound(BU_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = BU_T(j) Then
                    
                    If IsNumeric(KOUSEI(i, ColS_KOUSU)) Then
                        wkDouble = wkDouble + CDbl(ToRoundUp(CCur(CDbl(KOUSEI(i, ColS_KOUSU)) / 60 * CInt(Text1(ptxAFT_SEI_RATE).Text)), 2))
                    End If
                    Exit For
                End If
    
            Next j
    
        Next i
    
        Text1(ptxAFT_S_BU_KAKO_KOSU).Text = Format(wkDouble, "#0.00")
    
    End If



    KARI_TANKA_KEISAN_Proc = False

End Function



Private Function Tanka_Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   �P���o�^����
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim ans         As Integer

Dim wkGAISO     As Double
    
Dim i           As Integer
Dim j            As Integer
    
    
Dim wkint       As Integer
    
    Tanka_Update_Proc = True

    '�i�ڃ}�X�^�ǂݍ���
    Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHin_Gai).Text)


    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
            
                MsgBox "���[���Ńf�[�^���A�ύX����Ă��܂��B�P���o�^�����𒆎~���܂��B"
                Tanka_Update_Proc = False
                Exit Function
            
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Tanka_Update_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
                Exit Function
        
        End Select
    
    Loop


    '�V�P���|�|�����P�� 2009.06.02
    Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode))
    Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode))
    Call UniCode_Conv(ITEMREC.BEF_S_GAISO_TANKA, StrConv(ITEMREC.S_GAISO_TANKA, vbUnicode))
    Call UniCode_Conv(ITEMREC.BEF_S_PPSC_KAKO_KOSU, StrConv(ITEMREC.S_PPSC_KAKO_KOSU, vbUnicode))
    Call UniCode_Conv(ITEMREC.BEF_S_BU_KAKO_KOSU, StrConv(ITEMREC.S_BU_KAKO_KOSU, vbUnicode))



    '���b�g��
    Call UniCode_Conv(ITEMREC.SEI_LOT, Format(CLng(Text1(ptxAFT_SEI_LOT).Text), "00000000"))
    '�����[�g
    Call UniCode_Conv(ITEMREC.SEI_RATE, Format(CDbl(Text1(ptxAFT_SEI_RATE).Text), "0000.00"))
    '�H��
'2012.03.23    Call UniCode_Conv(ITEMREC.S_KOUSU, Format(CDbl(Text1(ptxAFT_S_KOUSU).Text), "0000.00"))
    '�H������
    Call UniCode_Conv(ITEMREC.S_KOUSU_GENKA, Format(CDbl(Text1(ptxAFT_S_KOUSU_GENKA).Text), "0000000.00"))
    '�H������
    Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, Format(CDbl(Text1(ptxAFT_S_KOUSU_BAIKA).Text), "0000000.00"))
    '�ݒ��
    Call UniCode_Conv(ITEMREC.S_KOUSU_SET_DATE, Format(Now, "YYYYMMDD"))
    
    
    '���㌴��
    Call UniCode_Conv(ITEMREC.S_SHIZAI_GENKA, Format(CDbl(Text1(ptxAFT_S_SHIZAI_GENKA).Text), "00000000.00"))
    '���㔄��
    Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, Format(CDbl(Text1(ptxAFT_S_SHIZAI_BAIKA).Text), "00000000.00"))
    
    
    
    '�O������
    If IsNumeric(Text1(ptxAFT_S_GAISO_TANKA).Text) Then
        Call UniCode_Conv(ITEMREC.S_GAISO_TANKA, Format(CDbl(Text1(ptxAFT_S_GAISO_TANKA).Text), "00000.00"))
    Else
        Call UniCode_Conv(ITEMREC.S_GAISO_TANKA, "00000.00")
    End If
    
    
    'PPSC�P��
    
    If IsNumeric(Text1(ptxAFT_S_PPSC_KAKO_KOSU).Text) Then
        Call UniCode_Conv(ITEMREC.S_PPSC_KAKO_KOSU, Format(CDbl(Text1(ptxAFT_S_PPSC_KAKO_KOSU).Text), "00000.00"))
    Else
        Call UniCode_Conv(ITEMREC.S_PPSC_KAKO_KOSU, "00000.00")
    End If
    'BU�P��
    If IsNumeric(Text1(ptxAFT_S_BU_KAKO_KOSU).Text) Then
       Call UniCode_Conv(ITEMREC.S_BU_KAKO_KOSU, Format(CDbl(Text1(ptxAFT_S_BU_KAKO_KOSU).Text), "00000.00"))
    Else
       Call UniCode_Conv(ITEMREC.S_BU_KAKO_KOSU, "00000.00")
    End If
    
    
    
    '�ݒ��
    Call UniCode_Conv(ITEMREC.S_SHIZAI_SET_DATE, Format(Now, "YYYYMMDD"))
    '�S����
    Call UniCode_Conv(ITEMREC.SEI_TANKA_TANTO, Text1(ptxTanto_Code).Text)
    '����
    Call UniCode_Conv(ITEMREC.SE_TANKA_MEMO, Text1(ptxAFT_SE_TANKA_MEMO).Text)
    
    '���x���\��t������
    Call UniCode_Conv(ITEMREC.SEI_LABEL_QTY, Format(CInt(Text1(ptxMAIN_KOUTEI_QTY01).Text), "00"))
    
    '�X�V�S����
    Call UniCode_Conv(ITEMREC.UPD_TANTO, Text1(ptxTanto_Code).Text)
    '�X�V ����
    Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
    
    
    '2008.09.03 �ǉ���
    
    '�d������
    Call UniCode_Conv(ITEMREC.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
    
        
    '���ތ���
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    wkint = wkint + 1
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    Call UniCode_Conv(ITEMREC.SEI_SZI_CNT, Format(wkint, "00"))
        
    '��������
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(DOUKON_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = DOUKON_T(j) Then
                    wkint = wkint + 1
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    Call UniCode_Conv(ITEMREC.SEI_DKN_CNT, Format(wkint, "00"))
        
        
        
    

    
    
    '2008.09.03 �ǉ���
    
    
    
    '2008.09.20 �ǉ���
    
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2011.12.12
'    '�O���
'    i = ptxBEF_KOUTEI_KOUSU01
'
'
'    For j = 0 To 9
'
'        If IsNumeric(Text1(i).Text) Then
'            Call UniCode_Conv(ITEMREC.BEF_KOUTEI(j).BEF_KOUTEI, Format(CDbl(Text1(i).Text), "000.00"))
'        Else
'            Call UniCode_Conv(ITEMREC.BEF_KOUTEI(j).BEF_KOUTEI, "000.00")
'        End If
'
'        i = i + 3
'
'
'
'    Next j

    For j = 0 To 9

        Call UniCode_Conv(ITEMREC.BEF_KOUTEI(j).BEF_KOUTEI, "000.00")
    
    Next j






    If IsNumeric(Text1(ptxCATE_ST_FUKA).Text) Then
        Call UniCode_Conv(ITEMREC.PLUS_KOUSU, Format(CDbl(Text1(ptxCATE_ST_FUKA).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.PLUS_KOUSU, "000.00")
    End If

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2011.12.12
    '����
    i = ptxMAIN_KOUTEI_KOUSU01
    
    
    For j = 0 To 8
    
        
        If IsNumeric(Text1(i).Text) Then
            Call UniCode_Conv(ITEMREC.MAIN_KOUTEI(j).MAIN_KOUTEI, Format(CDbl(Text1(i).Text), "000.00"))
        Else
            Call UniCode_Conv(ITEMREC.MAIN_KOUTEI(j).MAIN_KOUTEI, "000.00")
        End If
    
    
    
        i = i + 3
    
    
    Next j
    Call UniCode_Conv(ITEMREC.MAIN_KOUTEI(9).MAIN_KOUTEI, "000.00")
    
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2011.12.12
'    '����
'    i = ptxAFT_KOUTEI_KOUSU01
'
'
'    For j = 0 To 9
'
'        If IsNumeric(Text1(i).Text) Then
'            Call UniCode_Conv(ITEMREC.AFT_KOUTEI(j).AFT_KOUTEI, Format(CDbl(Text1(i).Text), "000.00"))
'        Else
'            Call UniCode_Conv(ITEMREC.AFT_KOUTEI(j).AFT_KOUTEI, "000.00")
'        End If
'
'
'
'        i = i + 3
'
'
'    Next j

    For j = 0 To 9

        Call UniCode_Conv(ITEMREC.AFT_KOUTEI(j).AFT_KOUTEI, "000.00")
    Next j


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2011.12.12
    
    
    
    '�q�ɋ敪
    Call UniCode_Conv(ITEMREC.SE_IO_TANKA_No, StrConv(SOKOREC.IO_TANKA_No, vbUnicode))
    '����
    Call UniCode_Conv(ITEMREC.SE_TANKA_MEMO, Text1(ptxAFT_SE_TANKA_MEMO).Text)
    '���Ϗ����l
    Call UniCode_Conv(ITEMREC.M_BIKOU, RichTextBox1(prchM_BIKOU).Text)
    '�d�l����
    Call UniCode_Conv(ITEMREC.SHIYOU_NO, Text1(ptxSHIYOU_NO).Text)
    '���ϋ敪
    Call UniCode_Conv(ITEMREC.MITSUMORI_KBN, Text1(ptxMITSUMORI_KBN).Text)
    '�P���ؑ֓�
    Call UniCode_Conv(ITEMREC.TANKA_KIRIKAE_DT, Text1(ptxTANKA_KIRIKAE_DT).Text)
    '�ؑ֋敪
    Call UniCode_Conv(ITEMREC.KIRIKAE_KBN, Text1(ptxKIRIKAE_KBN).Text)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    �i���J�e�S���B
    '���ʒP��
    Call UniCode_Conv(ITEMREC.G_SPTAN, "00000000.00")
    
    ' �O��H���i�b�j    �W��
    If IsNumeric(Text1(ptxCATE_ST_KOUTEI).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_ST_KOUTEI, Format(CDbl(Text1(ptxCATE_ST_KOUTEI).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_ST_KOUTEI, "000.00")
    End If
    
    ' �t���H���i�b�j    �W��
    If IsNumeric(Text1(ptxCATE_ST_FUKA).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_ST_FUKA, Format(CDbl(Text1(ptxCATE_ST_FUKA).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_ST_FUKA, "000.00")
    End If
    
    ' ����ƍH���i�b�j    �W��
    If IsNumeric(Text1(ptxCATE_ST_JITU1).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_ST_JITU1, Format(CDbl(Text1(ptxCATE_ST_JITU1).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_ST_JITU1, "000.00")
    End If
    
    ' �]�T���i���j    �W��
    If IsNumeric(Text1(ptxCATE_ST_YOYU_RITU).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_ST_YOYU_RITU, Format(CDbl(Text1(ptxCATE_ST_YOYU_RITU).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_ST_YOYU_RITU, "000.00")
    End If
    
    ' ����ƍH���i�b�j    �W��
    If IsNumeric(Text1(ptxCATE_ST_JITU2).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_ST_JITU2, Format(CDbl(Text1(ptxCATE_ST_JITU2).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_ST_JITU2, "000.00")
    End If
    
    ' ��Ǝ��Ԍv�i�b�j    �W��
    If IsNumeric(Text1(ptxCATE_ST_TOTAL).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_ST_TOTAL, Format(CDbl(Text1(ptxCATE_ST_TOTAL).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_ST_TOTAL, "000.00")
    End If
    ' ��/�i��/�j    �W��
    If IsNumeric(Text1(ptxCATE_ST_FUN).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_ST_FUN, Format(CDbl(Text1(ptxCATE_ST_FUN).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_ST_FUN, "000.00")
    End If
    
    ' ��ڰāi�~/���j    �W��
    If IsNumeric(Text1(ptxCATE_ST_FUN_RATE).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_ST_FUN_RATE, Format(CDbl(Text1(ptxCATE_ST_FUN_RATE).Text), "0000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_ST_FUN_RATE, "0000.00")
    End If
    
    ' �H�����i�~/�j    �W��
    If IsNumeric(Text1(ptxCATE_ST_KOURYO).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_ST_KOURYO, Format(CDbl(Text1(ptxCATE_ST_FUN_RATE).Text), "0000000000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_ST_KOURYO, "0000000000.00")
    End If
    
    
    
    ' �O��H���i�b�j    ����
    If IsNumeric(Text1(ptxCATE_AD_KOUTEI).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_AD_KOUTEI, Format(CDbl(Text1(ptxCATE_AD_KOUTEI).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_AD_KOUTEI, "000.00")
    End If
    
    ' �t���H���i�b�j    ����
    If IsNumeric(Text1(ptxCATE_AD_FUKA).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_AD_FUKA, Format(CDbl(Text1(ptxCATE_AD_FUKA).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_AD_FUKA, "000.00")
    End If
    
    ' ����ƍH���i�b�j    ����
    If IsNumeric(Text1(ptxCATE_AD_JITU1).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_AD_JITU1, Format(CDbl(Text1(ptxCATE_AD_JITU1).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_AD_JITU1, "000.00")
    End If
    
    ' �]�T���i���j    ����
    If IsNumeric(Text1(ptxCATE_AD_YOYU_RITU).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_AD_YOYU_RITU, Format(CDbl(Text1(ptxCATE_AD_YOYU_RITU).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_ST_YOYU_RITU, "000.00")
    End If
    
    ' ����ƍH���i�b�j    ����
    If IsNumeric(Text1(ptxCATE_AD_JITU2).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_AD_JITU2, Format(CDbl(Text1(ptxCATE_AD_JITU2).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_AD_JITU2, "000.00")
    End If
    
    ' ��Ǝ��Ԍv�i�b�j    ����
    If IsNumeric(Text1(ptxCATE_AD_TOTAL).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_AD_TOTAL, Format(CDbl(Text1(ptxCATE_AD_TOTAL).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_AD_TOTAL, "000.00")
    End If
    ' ��/�i��/�j    ����
    If IsNumeric(Text1(ptxCATE_AD_FUN).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_AD_FUN, Format(CDbl(Text1(ptxCATE_AD_FUN).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_AD_FUN, "000.00")
    End If
    
    ' ��ڰāi�~/���j    ����
    If IsNumeric(Text1(ptxCATE_AD_FUN_RATE).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_AD_FUN_RATE, Format(CDbl(Text1(ptxCATE_AD_FUN_RATE).Text), "0000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_AD_FUN_RATE, "0000.00")
    End If
    
    ' �H�����i�~/�j    ����
    If IsNumeric(Text1(ptxCATE_AD_KOURYO).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_AD_KOURYO, Format(CDbl(Text1(ptxCATE_AD_FUN_RATE).Text), "0000000000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_AD_KOURYO, "0000000000.00")
    End If
    
    ' �J�e�S���[�R�[�h
    Call UniCode_Conv(ITEMREC.CATEGORY_CODE, Trim(Text1(ptxCATEGORY_CODE).Text))
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �i�ڶú�؊֌W
    
    
Debug.Print StrConv(ITEMREC.CATE_ST_FUKA, vbUnicode)
    
    Do
        sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Tanka_Update_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpUpdate, "�i�ڃ}�X�^")
                Exit Function
        
        End Select
    
    Loop
    
    
    '�P���X�V�����o��
    Do
        sts = BTRV(BtOpInsert, ITEM_HST_POS, ITEMREC, Len(ITEMREC), K0_ITEM_HST, Len(K0_ITEM_HST), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM_HST.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Tanka_Update_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpUpdate, "�i�ڒP���X�V����")
                Exit Function
        
        End Select
    
    Loop
    

    Tanka_Update_Proc = False


End Function

Private Function Grid_Error_Check_Proc() As Integer
'----------------------------------------------------------------------------
'                   ��د�ޓ��e�̃G���[�`�F�b�N����
'----------------------------------------------------------------------------
Dim i   As Integer

Dim sts As Integer
    
    
Dim K_SEQNO As Integer
Dim G_SEQNO As Integer
Dim D_SEQNO As Integer
    
Dim SHIZAI_UMU  As Boolean  '2013.01.16
Dim SHIZAI_CNT  As Long     '2013.01.16
Dim j           As Long     '2013.01.16
    
    Grid_Error_Check_Proc = True
    
    
    SHIZAI_UMU = True
    For i = 0 To UBound(ITEM_CATEGORY_SUMI)
        If Trim(Text1(ptxCATEGORY_CODE).Text) = Trim(ITEM_CATEGORY_SUMI(i)) Then
            SHIZAI_UMU = False
            Exit For
        End If
    Next i
    
    
    Set TDBGrid1(pGrdKOUSEI).Array = KOUSEI
    
'    TDBGrid1.Refresh
    
    TDBGrid1(pGrdKOUSEI).Update
    
    If KOUSEI.Count(1) < 1 Then
        
        
        If SHIZAI_UMU Then                                          '2013.01.16
            MsgBox "���i�����K�v�ȕi�ڂŎ��ޕi�����o�^�ł��B"       '2013.01.16
            Exit Function                                           '2013.01.16
        End If                                                      '2013.01.16
        
        Grid_Error_Check_Proc = False
        Exit Function
    End If

    SHIZAI_CNT = 0                                                  '2013.01.16

    For i = 1 To KOUSEI.Count(1)
    
    
        If Trim(KOUSEI(i, ColKO_HIN_GAI)) = "" Then
            
            KOUSEI(i, ColKO_JGYOBU) = ""
            KOUSEI(i, ColKO_NAIGAI) = ""
            
            KOUSEI(i, ColKO_HIN_NAME) = ""
            KOUSEI(i, ColKO_QTY) = ""
            KOUSEI(i, ColG_ST_SHITAN) = ""
            KOUSEI(i, ColG_ST_URITAN) = ""
            KOUSEI(i, ColG_ST_SHIKIN) = ""
            KOUSEI(i, ColG_ST_URIKIN) = ""
            KOUSEI(i, ColS_KOUSU) = ""
            KOUSEI(i, ColSEI_SYU_KON) = ""
    
            KOUSEI(i, ColKO_BIKOU) = ""
    
        Else
    
    
    
            Select Case Right(KOUSEI(i, ColKO_SYUBETSU), 2)
            
                Case KOSOU_KBN          '��
                    K_SEQNO = K_SEQNO + 10
                    
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.01.25
                    'If K_SEQNO > 50 Then
                    '    MsgBox "�����ޓo�^�������I�[�o�[���Ă��܂��B"
                    '    Exit Function
                    'End If
                    
                    If K_SEQNO > 50 Then
                        D_SEQNO = D_SEQNO + 10
                        If D_SEQNO > 250 Then
                            MsgBox "�����o�^�������I�[�o�[���Ă��܂��B"
                            Exit Function
                        End If
                    End If
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.01.25
                
                Case GAISO_KBN          '�O��
                    G_SEQNO = G_SEQNO + 10
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.01.25
                    'If G_SEQNO > 30 Then
                    '    MsgBox "�O�����ޓo�^�������I�[�o�[���Ă��܂��B"
                    '    Exit Function
                    'End If
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.01.25
                    If G_SEQNO > 30 Then
                        D_SEQNO = D_SEQNO + 10
                        If D_SEQNO > 250 Then
                            MsgBox "�����o�^�������I�[�o�[���Ă��܂��B"
                            Exit Function
                        End If
                    End If
                Case Else               '����
                    D_SEQNO = D_SEQNO + 10
                    If D_SEQNO > 250 Then
                        MsgBox "�����o�^�������I�[�o�[���Ă��܂��B"
                        Exit Function
                    End If
            End Select
    
    
    
    
            '�i��
            If Trim(KOUSEI(i, ColKO_JGYOBU)) = "" And _
                Trim(KOUSEI(i, ColKO_NAIGAI)) = "" Then
                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
            Else
                Call UniCode_Conv(K0_ITEM.JGYOBU, KOUSEI(i, ColKO_JGYOBU))
                Call UniCode_Conv(K0_ITEM.NAIGAI, KOUSEI(i, ColKO_NAIGAI))
            End If
            
            Call UniCode_Conv(K0_ITEM.HIN_GAI, KOUSEI(i, ColKO_HIN_GAI))
        
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    '���ޕi�œǂݑւ�
                                    
                    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, KOUSEI(i, ColKO_HIN_GAI))
                    
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                            
                            If HIN_INV Then
                                '���o�^�i�ԁ@�@���ނƂ��Ă���
                                Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)
                                Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
                                Call UniCode_Conv(ITEMREC.HIN_NAME, "���o�^�i��")
                                Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                            Else
                                MsgBox "[" & Format(i, "0") & "]�s�� ���͂������ڂ̓G���[�ł��B(�i��)"
                                Exit Function
                            End If
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                            Exit Function
                    
                    End Select
            
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                    Exit Function
            
            End Select
    
            KOUSEI(i, ColKO_JGYOBU) = StrConv(ITEMREC.JGYOBU, vbUnicode)
            KOUSEI(i, ColKO_NAIGAI) = StrConv(ITEMREC.NAIGAI, vbUnicode)
            KOUSEI(i, ColKO_HIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
    
    
            '����
            If Trim(KOUSEI(i, ColKO_QTY)) = "" Then
                KOUSEI(i, ColKO_QTY) = "1.00"
            End If
            If IsNumeric(KOUSEI(i, ColKO_QTY)) Then
                KOUSEI(i, ColKO_QTY) = Format(CDbl(KOUSEI(i, ColKO_QTY)), "#0.00")
            Else
                MsgBox "[" & Format(i, "0") & "]�s�� ���͂������ڂ̓G���[�ł��B(����)"
    
            End If
    
    
            '�d����
            If Trim(KOUSEI(i, ColG_ST_SHITAN)) = "" Then
                KOUSEI(i, ColG_ST_SHITAN) = "0.00"
            End If
            If IsNumeric(KOUSEI(i, ColG_ST_SHITAN)) Then
                KOUSEI(i, ColG_ST_SHITAN) = Format(CDbl(KOUSEI(i, ColG_ST_SHITAN)), "#0.00")
            Else
                MsgBox "[" & Format(i, "0") & "]�s�� ���͂������ڂ̓G���[�ł��B(�d����)"
            End If
            '�̔���
            
            Select Case StrConv(ITEMREC.SEI_KBN, vbUnicode)
            
            
                Case "1"
            
                    KOUSEI(i, ColG_ST_URITAN) = "�ʔ�"
            
                Case "2"
            
                    KOUSEI(i, ColG_ST_URITAN) = "�x��"
            
            
                Case Else
                    If Trim(KOUSEI(i, ColG_ST_URITAN)) = "" Then
                        KOUSEI(i, ColG_ST_URITAN) = "0.00"
                    End If
                    
                    If IsNumeric(KOUSEI(i, ColG_ST_URITAN)) Then
                        KOUSEI(i, ColG_ST_URITAN) = Format(CDbl(KOUSEI(i, ColG_ST_URITAN)), "#0.00")
                    Else
                        MsgBox "[" & Format(i, "0") & "]�s�� ���͂������ڂ̓G���[�ł��B(�̔���)"
            
                    End If
            
            End Select
            
            '���ʒP����
            If Trim(KOUSEI(i, ColG_SPTAN)) = "" Then
            Else
                If IsNumeric(KOUSEI(i, ColG_SPTAN)) Then
                    KOUSEI(i, ColG_SPTAN) = Format(CDbl(KOUSEI(i, ColG_SPTAN)), "#0.00")
                Else
                    MsgBox "[" & Format(i, "0") & "]�s�� ���͂������ڂ̓G���[�ł��B(���ʒP����)"
                End If
            End If
            
            '�d�����z�v
            If Trim(KOUSEI(i, ColG_ST_SHIKIN)) = "" Then
                KOUSEI(i, ColG_ST_SHIKIN) = "0.00"
            End If
            If IsNumeric(KOUSEI(i, ColG_ST_SHIKIN)) Then
                KOUSEI(i, ColG_ST_SHIKIN) = Format(CDbl(KOUSEI(i, ColG_ST_SHIKIN)), "#0.00")
            Else
                MsgBox "[" & Format(i, "0") & "]�s�� ���͂������ڂ̓G���[�ł��B(�d�����z�v)"
    
            End If
            
            '�̔����z�v
            If StrConv(ITEMREC.SEI_KBN, vbUnicode) <> "1" And StrConv(ITEMREC.SEI_KBN, vbUnicode) <> "2" Then
            
                If Trim(KOUSEI(i, ColG_ST_URIKIN)) = "" Then
                    KOUSEI(i, ColG_ST_URIKIN) = "0.00"
                End If
                If IsNumeric(KOUSEI(i, ColG_ST_URIKIN)) Then
                    KOUSEI(i, ColG_ST_URIKIN) = Format(CDbl(KOUSEI(i, ColG_ST_URIKIN)), "#0.00")
                Else
                    MsgBox "[" & Format(i, "0") & "]�s�� ���͂������ڂ̓G���[�ł��B(�̔����z�v)"
                End If
            End If
            
            '��Ǝ���
            If Trim(KOUSEI(i, ColS_KOUSU)) = "" Then
                KOUSEI(i, ColS_KOUSU) = "0"
            End If
            If IsNumeric(KOUSEI(i, ColS_KOUSU)) Then
                KOUSEI(i, ColS_KOUSU) = Format(CDbl(KOUSEI(i, ColS_KOUSU)), "#0")
            Else
                MsgBox "[" & Format(i, "0") & "]�s�� ���͂������ڂ̓G���[�ł��B(��Ǝ���)"
            End If
            '�W�������
            If Trim(KOUSEI(i, ColSEI_SYU_KON)) = "" Then
                KOUSEI(i, ColSEI_SYU_KON) = "0"
            End If
            If IsNumeric(KOUSEI(i, ColSEI_SYU_KON)) Then
                KOUSEI(i, ColSEI_SYU_KON) = Format(CDbl(KOUSEI(i, ColSEI_SYU_KON)), "#0")
            Else
                MsgBox "[" & Format(i, "0") & "]�s�� ���͂������ڂ̓G���[�ł��B(�W�������)"
            End If
    
    
    
            '���ޗL��   2013.01.16
            For j = 0 To UBound(CHK_SHIZAI_T)
                If Trim(CHK_SHIZAI_T(j)) = Right(KOUSEI(i, ColKO_SYUBETSU), 2) Then
                    SHIZAI_CNT = SHIZAI_CNT + 1
                    Exit For
                End If
            Next j
            '���ޗL��   2013.01.16
    
    
        End If
    Next i


    '���ޗL��   2013.01.16
    If SHIZAI_UMU Then
        If SHIZAI_CNT = 0 Then
            MsgBox "���i�����K�v�ȕi�ڂŎ��ޕi�����o�^�ł��B"       '2013.01.16
            Exit Function                                           '2013.01.16
        End If
    End If

    Grid_Error_Check_Proc = False



End Function
Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   �\���}�X�^�o��
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim ans         As Integer
Dim com         As Integer

Dim K_SEQNO     As Integer
Dim G_SEQNO     As Integer
Dim D_SEQNO     As Integer


Dim i           As Integer
Dim j           As Integer

Dim MESG        As String


    Update_Proc = True
                                        
    Call Input_Lock
                                        
                                        '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If
    '---------------------------------------------------    '�\���}�X�^�X�V
    '�Y���f�[�^�S���폜
    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHin_Gai).Text)
       
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
                        Trim(StrConv(P_COMPO_O_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxHin_Gai).Text) Then
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
                                                                                '�d�����溰��
    Call UniCode_Conv(P_COMPO_O_REC.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
                                                                                '���ƕ�
    Call UniCode_Conv(P_COMPO_O_REC.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
                                                                                '�����O
    Call UniCode_Conv(P_COMPO_O_REC.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
    Call UniCode_Conv(P_COMPO_O_REC.HIN_GAI, Text1(ptxHin_Gai).Text)
    Call UniCode_Conv(P_COMPO_O_REC.DATA_KBN, P_HEAD)
    Call UniCode_Conv(P_COMPO_O_REC.SEQNO, "000")

    Call UniCode_Conv(P_COMPO_O_REC.CLASS_CODE, Text1(ptxS_CLASS_CODE).Text)    '�׽����
    Call UniCode_Conv(P_COMPO_O_REC.BIKOU, RichTextBox1(prchBIKOU).Text)        '���l
    
    Call UniCode_Conv(P_COMPO_O_REC.F_CLASS_CODE, Text1(ptxF_CLASS_CODE).Text)  '�t������
    
    Call UniCode_Conv(P_COMPO_O_REC.N_CLASS_CODE, Text1(ptxN_CLASS_CODE).Text)  '���E����
    
    Call UniCode_Conv(P_COMPO_O_REC.FILLER, "")

    Call UniCode_Conv(P_COMPO_O_REC.UPD_TANTO, Text1(ptxTanto_Code))            '�X�V�S���Һ���
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



    '�\���}�X�^(���ި)�o��
    
    Set TDBGrid1(pGrdKOUSEI).Array = KOUSEI
    
'    TDBGrid1.Refresh
    
    TDBGrid1(pGrdKOUSEI).Update


    K_SEQNO = 0
    G_SEQNO = 0
    D_SEQNO = 0


    '2009.03.24
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then

    Else


        For i = 1 To KOUSEI.UpperBound(1)
    
    
            If Trim(KOUSEI(i, ColKO_HIN_GAI)) = "" Then
            Else
                                                                                            '�d�����溰��
                Call UniCode_Conv(P_COMPO_K_REC.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
                                                                                            '���ƕ�
                Call UniCode_Conv(P_COMPO_K_REC.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
                                                                                            '�����O
                Call UniCode_Conv(P_COMPO_K_REC.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
                Call UniCode_Conv(P_COMPO_K_REC.HIN_GAI, Text1(ptxHin_Gai).Text)
            
            
            
                Select Case Right(KOUSEI(i, ColKO_SYUBETSU), 2)
                
                    Case KOSOU_KBN          '��
                    
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2012.01.25
                        If K_SEQNO > 40 Then
                    
                            K_SEQNO = K_SEQNO + 10
                            Call UniCode_Conv(P_COMPO_K_REC.DATA_KBN, P_DOUKON)             '�f�[�^�敪
                            D_SEQNO = D_SEQNO + 10
                            Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(D_SEQNO, "000"))  '�ǔ�
                                                                                            '���
                            Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, Right(KOUSEI(i, ColKO_SYUBETSU), 2))
                        Else
                            Call UniCode_Conv(P_COMPO_K_REC.DATA_KBN, P_KOSOU)              '�f�[�^�敪
                            K_SEQNO = K_SEQNO + 10
                            Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(K_SEQNO, "000"))  '�ǔ�
                            Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, "")                '���
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2012.01.25
                    
                    Case GAISO_KBN          '�O��
                
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2012.01.25
                        If G_SEQNO > 20 Then
                        
                            G_SEQNO = G_SEQNO + 10
                        
                        
                            Call UniCode_Conv(P_COMPO_K_REC.DATA_KBN, P_DOUKON)             '�f�[�^�敪
                            
                            D_SEQNO = D_SEQNO + 10
                            Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(D_SEQNO, "000"))  '�ǔ�
                                                                                            '���
                            Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, Right(KOUSEI(i, ColKO_SYUBETSU), 2))
                        Else
                            Call UniCode_Conv(P_COMPO_K_REC.DATA_KBN, P_GAISOU)             '�f�[�^�敪
                            G_SEQNO = G_SEQNO + 10
                            Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(G_SEQNO, "000"))  '�ǔ�
                            Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, "")                '���
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2012.01.25
                
                
                
                
                
                    Case Else               '����
                
                
                        Call UniCode_Conv(P_COMPO_K_REC.DATA_KBN, P_DOUKON)             '�f�[�^�敪
                        
                        D_SEQNO = D_SEQNO + 10
                        
                        Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(D_SEQNO, "000"))  '�ǔ�
                                                                                        '���
                        Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, Right(KOUSEI(i, ColKO_SYUBETSU), 2))
                
                End Select
            
            
                Call UniCode_Conv(P_COMPO_K_REC.KO_JGYOBU, KOUSEI(i, ColKO_JGYOBU))         '�q�@���ƕ�
                Call UniCode_Conv(P_COMPO_K_REC.KO_NAIGAI, KOUSEI(i, ColKO_NAIGAI))         '�q�@�����O
                Call UniCode_Conv(P_COMPO_K_REC.KO_HIN_GAI, KOUSEI(i, ColKO_HIN_GAI))       '�q�@�i��
                                                                                            '����
                Call UniCode_Conv(P_COMPO_K_REC.KO_QTY, Format(CDbl(KOUSEI(i, ColKO_QTY)), "000.00"))
                Call UniCode_Conv(P_COMPO_K_REC.KO_BIKOU, KOUSEI(i, ColKO_BIKOU))           '�q�@���l
            
                Call UniCode_Conv(P_COMPO_K_REC.FILLER, "")
            
                Call UniCode_Conv(P_COMPO_K_REC.UPD_TANTO, Text1(ptxTanto_Code).Text)       '�X�V�S���Һ���
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
    
    
                Call UniCode_Conv(K0_ITEM.JGYOBU, KOUSEI(i, ColKO_JGYOBU))         '�q�@���ƕ�
                Call UniCode_Conv(K0_ITEM.NAIGAI, KOUSEI(i, ColKO_NAIGAI))         '�q�@�����O
                Call UniCode_Conv(K0_ITEM.HIN_GAI, KOUSEI(i, ColKO_HIN_GAI))       '�q�@�i��
    
    
                Do
                    sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrKeyNotFound
                        
                            MsgBox "���[���Ńf�[�^���A�ύX����Ă��܂��B�\���|�ۑ������𒆎~���܂��B"
                            Update_Proc = False
                            GoTo Abort_Tran
                        
                        
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Update_Proc = False
                                GoTo Abort_Tran
                            End If
                        Case Else
                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
                            GoTo Abort_Tran
                    
                    End Select
                
                Loop
    
                '�H��
                Call UniCode_Conv(ITEMREC.S_KOUSU, Format(KOUSEI(i, ColS_KOUSU), "00000.00"))
                '�W������
                Call UniCode_Conv(ITEMREC.SEI_SYU_KON, Format(KOUSEI(i, ColSEI_SYU_KON), "000.00"))
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    �i���J�e�S���B
                '���ʒP��
                If IsNumeric(KOUSEI(i, ColG_SPTAN)) Then
                    Call UniCode_Conv(ITEMREC.G_SPTAN, Format(KOUSEI(i, ColG_SPTAN), "00000000.00"))
                Else
                    Call UniCode_Conv(ITEMREC.G_SPTAN, "")
                End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    �i���J�e�S���B
                Do
                    sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrKeyNotFound
                        
                            MsgBox "���[���Ńf�[�^���A�ύX����Ă��܂��B�\���|�ۑ������𒆎~���܂��B"
                            Update_Proc = False
                            GoTo Abort_Tran
                        
                        
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Update_Proc = False
                                GoTo Abort_Tran
                            End If
                        Case Else
                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
                            GoTo Abort_Tran
                    
                    End Select
                
                Loop
    
            End If
        Next i
    End If


    '---------------------------------------------------    '�i��Ͻ��@�e�i�ԍX�V    2009.06.02

    Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHin_Gai).Text)

    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
            
                MsgBox "���[���Ńf�[�^���A�ύX����Ă��܂��B�\���|�ۑ������𒆎~���܂��B"
                Update_Proc = False
                GoTo Abort_Tran
            
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Update_Proc = False
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
                GoTo Abort_Tran
        
        End Select
    Loop

    '���Ϗ����l
    Call UniCode_Conv(ITEMREC.M_BIKOU, RichTextBox1(prchM_BIKOU).Text)
    '�d�l����
    Call UniCode_Conv(ITEMREC.SHIYOU_NO, Text1(ptxSHIYOU_NO).Text)
    '���ϋ敪
    Call UniCode_Conv(ITEMREC.MITSUMORI_KBN, Text1(ptxMITSUMORI_KBN).Text)
    '�P���ؑ֓�
    Call UniCode_Conv(ITEMREC.TANKA_KIRIKAE_DT, Text1(ptxTANKA_KIRIKAE_DT).Text)
    '�ؑ֋敪
    Call UniCode_Conv(ITEMREC.KIRIKAE_KBN, Text1(ptxKIRIKAE_KBN).Text)




    '-----  �P���� 2009.07.24
    '���b�g��
    
    If IsNumeric(Text1(ptxBEF_SEI_LOT).Text) Then
        Call UniCode_Conv(ITEMREC.SEI_LOT, Format(CLng(Text1(ptxBEF_SEI_LOT).Text), "00000000"))
    Else
        Call UniCode_Conv(ITEMREC.SEI_LOT, "")
    End If
      '�����[�g
    If IsNumeric(Text1(ptxBEF_SEI_RATE).Text) Then
        Call UniCode_Conv(ITEMREC.SEI_RATE, Format(CDbl(Text1(ptxBEF_SEI_RATE).Text), "0000.00"))
    Else
        Call UniCode_Conv(ITEMREC.SEI_RATE, "")
    End If
    '�H��
    If IsNumeric(Text1(ptxBEF_S_KOUSU).Text) Then
        Call UniCode_Conv(ITEMREC.S_KOUSU, Format(CDbl(Text1(ptxBEF_S_KOUSU).Text), "00000.00"))
    Else
        Call UniCode_Conv(ITEMREC.S_KOUSU, "")
    End If
    '�H������
    If IsNumeric(Text1(ptxBEF_S_KOUSU_GENKA).Text) Then
        Call UniCode_Conv(ITEMREC.S_KOUSU_GENKA, Format(CDbl(Text1(ptxBEF_S_KOUSU_GENKA).Text), "00000000.00"))
    Else
        Call UniCode_Conv(ITEMREC.S_KOUSU_GENKA, "")
    End If
    '�H������
    If IsNumeric(Text1(ptxBEF_S_KOUSU_BAIKA).Text) Then
        Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, Format(CDbl(Text1(ptxBEF_S_KOUSU_BAIKA).Text), "00000000.00"))
    Else
        Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "")
    End If
    '�ݒ��
    If Trim(Text1(ptxBEF_S_KOUSU_SET_DATE).Text) = "" Then
        Call UniCode_Conv(ITEMREC.S_KOUSU_SET_DATE, "")
    Else
        Call UniCode_Conv(ITEMREC.S_KOUSU_SET_DATE, Format(Now, "YYYYMMDD"))
    End If
    '���㌴��
    If IsNumeric(Text1(ptxBEF_S_SHIZAI_GENKA).Text) Then
        Call UniCode_Conv(ITEMREC.S_SHIZAI_GENKA, Format(CDbl(Text1(ptxBEF_S_SHIZAI_GENKA).Text), "00000000.00"))
    Else
        Call UniCode_Conv(ITEMREC.S_SHIZAI_GENKA, "")
    End If
    '���㔄��
    If IsNumeric(Text1(ptxBEF_S_SHIZAI_BAIKA).Text) Then
        Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, Format(CDbl(Text1(ptxBEF_S_SHIZAI_BAIKA).Text), "00000000.00"))
    Else
        Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "")
    End If
    '�O������
    If IsNumeric(Text1(ptxBEF_S_GAISO_TANKA).Text) Then
        Call UniCode_Conv(ITEMREC.S_GAISO_TANKA, Format(CDbl(Text1(ptxBEF_S_GAISO_TANKA).Text), "00000.00"))
    Else
        Call UniCode_Conv(ITEMREC.S_GAISO_TANKA, "")
    End If
    'PPSC�P��
    If IsNumeric(Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text) Then
        Call UniCode_Conv(ITEMREC.S_PPSC_KAKO_KOSU, Format(CDbl(Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text), "00000.00"))
    Else
        Call UniCode_Conv(ITEMREC.S_PPSC_KAKO_KOSU, "")
    End If
    'BU�P��
    If IsNumeric(Text1(ptxBEF_S_BU_KAKO_KOSU).Text) Then
       Call UniCode_Conv(ITEMREC.S_BU_KAKO_KOSU, Format(CDbl(Text1(ptxBEF_S_BU_KAKO_KOSU).Text), "00000.00"))
    Else
       Call UniCode_Conv(ITEMREC.S_BU_KAKO_KOSU, "")
    End If
    '�ݒ��
    If Trim(Text1(ptxBEF_S_KOUSU_SET_DATE).Text) = "" Then
        Call UniCode_Conv(ITEMREC.S_SHIZAI_SET_DATE, "")
    Else
        Call UniCode_Conv(ITEMREC.S_SHIZAI_SET_DATE, Format(Now, "YYYYMMDD"))
    End If
    '�S����
    Call UniCode_Conv(ITEMREC.SEI_TANKA_TANTO, Text1(ptxTanto_Code).Text)
    '����
    Call UniCode_Conv(ITEMREC.SE_TANKA_MEMO, Text1(ptxBEF_SE_TANKA_MEMO).Text)
    '���x���\��t������
    Call UniCode_Conv(ITEMREC.SEI_LABEL_QTY, Format(CInt(Text1(ptxMAIN_KOUTEI_QTY01).Text), "00"))
    
    
    
    '�H������
    Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, Format(CDbl(Text1(ptxOLD_S_KOUSU_BAIKA).Text), "00000000.00"))
    '���㔄��
    Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, Format(CDbl(Text1(ptxOLD_S_SHIZAI_BAIKA).Text), "00000000.00"))
    '�O������
    If IsNumeric(Text1(ptxOLD_S_GAISO_TANKA).Text) Then
        Call UniCode_Conv(ITEMREC.BEF_S_GAISO_TANKA, Format(CDbl(Text1(ptxOLD_S_GAISO_TANKA).Text), "00000.00"))
    Else
        Call UniCode_Conv(ITEMREC.BEF_S_GAISO_TANKA, "00000.00")
    End If
    'PPSC�P��
    If IsNumeric(Text1(ptxOLD_S_PPSC_KAKO_KOSU).Text) Then
        Call UniCode_Conv(ITEMREC.BEF_S_PPSC_KAKO_KOSU, Format(CDbl(Text1(ptxOLD_S_PPSC_KAKO_KOSU).Text), "00000.00"))
    Else
        Call UniCode_Conv(ITEMREC.BEF_S_PPSC_KAKO_KOSU, "00000.00")
    End If
    'BU�P��
    If IsNumeric(Text1(ptxOLD_S_BU_KAKO_KOSU).Text) Then
       Call UniCode_Conv(ITEMREC.BEF_S_BU_KAKO_KOSU, Format(CDbl(Text1(ptxOLD_S_BU_KAKO_KOSU).Text), "00000.00"))
    Else
       Call UniCode_Conv(ITEMREC.BEF_S_BU_KAKO_KOSU, "00000.00")
    End If
    
    '�t���H��
    If IsNumeric(Text1(ptxCATE_ST_FUKA).Text) Then
        Call UniCode_Conv(ITEMREC.PLUS_KOUSU, Format(CDbl(Text1(ptxCATE_ST_FUKA).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.PLUS_KOUSU, "000.00")
    End If
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �i�ڶú�؊֌W
    
    '���ʒP��
    Call UniCode_Conv(ITEMREC.G_SPTAN, "00000000.00")
    
    ' �O��H���i�b�j    �W��
    If IsNumeric(Text1(ptxCATE_ST_KOUTEI).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_ST_KOUTEI, Format(CDbl(Text1(ptxCATE_ST_KOUTEI).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_ST_KOUTEI, "000.00")
    End If
    
    ' �t���H���i�b�j    �W��
    If IsNumeric(Text1(ptxCATE_ST_FUKA).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_ST_FUKA, Format(CDbl(Text1(ptxCATE_ST_FUKA).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_ST_FUKA, "000.00")
    End If
    
    ' ����ƍH���i�b�j    �W��
    If IsNumeric(Text1(ptxCATE_ST_JITU1).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_ST_JITU1, Format(CDbl(Text1(ptxCATE_ST_JITU1).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_ST_JITU1, "000.00")
    End If
    
    ' �]�T���i���j    �W��
    If IsNumeric(Text1(ptxCATE_ST_YOYU_RITU).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_ST_YOYU_RITU, Format(CDbl(Text1(ptxCATE_ST_YOYU_RITU).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_ST_YOYU_RITU, "000.00")
    End If
    
    ' ����ƍH���i�b�j    �W��
    If IsNumeric(Text1(ptxCATE_ST_JITU2).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_ST_JITU2, Format(CDbl(Text1(ptxCATE_ST_JITU2).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_ST_JITU2, "000.00")
    End If
    
    ' ��Ǝ��Ԍv�i�b�j    �W��
    If IsNumeric(Text1(ptxCATE_ST_TOTAL).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_ST_TOTAL, Format(CDbl(Text1(ptxCATE_ST_TOTAL).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_ST_TOTAL, "000.00")
    End If
    ' ��/�i��/�j    �W��
    If IsNumeric(Text1(ptxCATE_ST_FUN).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_ST_FUN, Format(CDbl(Text1(ptxCATE_ST_FUN).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_ST_FUN, "000.00")
    End If
    
    ' ��ڰāi�~/���j    �W��
    If IsNumeric(Text1(ptxCATE_ST_FUN_RATE).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_ST_FUN_RATE, Format(CDbl(Text1(ptxCATE_ST_FUN_RATE).Text), "0000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_ST_FUN_RATE, "0000.00")
    End If
    
    ' �H�����i�~/�j    �W��
    If IsNumeric(Text1(ptxCATE_ST_KOURYO).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_ST_KOURYO, Format(CDbl(Text1(ptxCATE_ST_FUN_RATE).Text), "0000000000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_ST_KOURYO, "0000000000.00")
    End If
    
    
    
    ' �O��H���i�b�j    ����
    If IsNumeric(Text1(ptxCATE_AD_KOUTEI).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_AD_KOUTEI, Format(CDbl(Text1(ptxCATE_AD_KOUTEI).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_AD_KOUTEI, "000.00")
    End If
    
    ' �t���H���i�b�j    ����
    If IsNumeric(Text1(ptxCATE_AD_FUKA).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_AD_FUKA, Format(CDbl(Text1(ptxCATE_AD_FUKA).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_AD_FUKA, "000.00")
    End If
    
    ' ����ƍH���i�b�j    ����
    If IsNumeric(Text1(ptxCATE_AD_JITU1).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_AD_JITU1, Format(CDbl(Text1(ptxCATE_AD_JITU1).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_AD_JITU1, "000.00")
    End If
    
    ' �]�T���i���j    ����
    If IsNumeric(Text1(ptxCATE_AD_YOYU_RITU).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_AD_YOYU_RITU, Format(CDbl(Text1(ptxCATE_AD_YOYU_RITU).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_ST_YOYU_RITU, "000.00")
    End If
    
    ' ����ƍH���i�b�j    ����
    If IsNumeric(Text1(ptxCATE_AD_JITU2).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_AD_JITU2, Format(CDbl(Text1(ptxCATE_AD_JITU2).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_AD_JITU2, "000.00")
    End If
    
    ' ��Ǝ��Ԍv�i�b�j    ����
    If IsNumeric(Text1(ptxCATE_AD_TOTAL).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_AD_TOTAL, Format(CDbl(Text1(ptxCATE_AD_TOTAL).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_AD_TOTAL, "000.00")
    End If
    ' ��/�i��/�j    ����
    If IsNumeric(Text1(ptxCATE_AD_FUN).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_AD_FUN, Format(CDbl(Text1(ptxCATE_AD_FUN).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_AD_FUN, "000.00")
    End If
    
    ' ��ڰāi�~/���j    ����
    If IsNumeric(Text1(ptxCATE_AD_FUN_RATE).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_AD_FUN_RATE, Format(CDbl(Text1(ptxCATE_AD_FUN_RATE).Text), "0000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_AD_FUN_RATE, "0000.00")
    End If
    
    ' �H�����i�~/�j    ����
    If IsNumeric(Text1(ptxCATE_AD_KOURYO).Text) Then
        Call UniCode_Conv(ITEMREC.CATE_AD_KOURYO, Format(CDbl(Text1(ptxCATE_AD_FUN_RATE).Text), "0000000000.00"))
    Else
        Call UniCode_Conv(ITEMREC.CATE_AD_KOURYO, "0000000000.00")
    End If
    
    ' �J�e�S���[�R�[�h
    Call UniCode_Conv(ITEMREC.CATEGORY_CODE, Trim(Text1(ptxCATEGORY_CODE).Text))
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �i�ڶú�؊֌W
    '-----  �P���� 2009.07.24

    '�X�V�S����
    Call UniCode_Conv(ITEMREC.UPD_TANTO, Text1(ptxTanto_Code).Text)
    '�X�V ����
    Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))


    Do
        sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
            
                MsgBox "���[���Ńf�[�^���A�ύX����Ă��܂��B�\���|�ۑ������𒆎~���܂��B"
                Update_Proc = False
                GoTo Abort_Tran
            
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Update_Proc = False
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
                GoTo Abort_Tran
        
        End Select
    
    Loop

End_Tran:
                                        '�g�����U�N�V�����I��
    sts = BTRV(BtOpEndTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
    
    Call Input_UnLock
    
    Update_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If
    
    Call Input_UnLock

End Function

Private Sub Text1_LostFocus(Index As Integer)
    
Dim i   As Integer
    
    
    Select Case Index
        Case ptxHin_Gai
            
            
            Text1(Index).Text = StrConv(RTrim(Text1(Index).Text), vbUpperCase)
            
            
            If Trim(svHin_Gai) = (Text1(ptxHin_Gai).Text) Then
            Else
                For i = 2 To 5
                    Command1(i).Enabled = False
                Next i
            
                Text1(ptxMAIN_KOUTEI_QTY01).Text = ""
            
            End If
    End Select
End Sub


Private Sub Estimate_Head_Proc(excelApplication As Object, excelWorkBook As Object, excelSheet As Object)
'----------------------------------------------------------------------------
'                   �d�w�b�d�k�i�䌩�Ϗ��w�b�_�[�j�o��
'       2009.06.02
'----------------------------------------------------------------------------
Dim i   As Integer
Debug.Print "in Estimate_head_Proc=" & Now
    
    excelApplication.StandardFontSize = 11
    
    excelApplication.StandardFont = "�l�r�@�o�S�V�b�N"
    
    '�y�[�W�ݒ�
    
    If Trim(EXCEL_TEMPLATE) = "" Then
    
        With excelSheet.Application.ActiveSheet.PageSetup
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
            .Orientation = xlPortrait
        End With
    
    Else
    
        With excelSheet.Application.ActiveSheet.PageSetup
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
        End With
    
    End If

'---    �P�s��
    '�Z���̌���
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 5), excelSheet.Application.Cells(1, 8)).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Range(excelSheet.Application.Cells(1, 5), excelSheet.Application.Cells(1, 8)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 5), excelSheet.Application.Cells(1, 8)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 5), excelSheet.Application.Cells(1, 8)).Font.FontStyle = "����"
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 5), excelSheet.Application.Cells(1, 8)).Font.Size = 24
    excelSheet.Application.Cells(1, 5).Value = "�@��@���@�ρ@���@"
'---    �Q�s��
    '�Z���̌���
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 10), excelSheet.Application.Cells(2, 12)).HorizontalAlignment = xlRight
'2011.11.21    excelSheet.Application.Range(excelSheet.Application.Cells(2, 10), excelSheet.Application.Cells(2, 12)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 10), excelSheet.Application.Cells(2, 12)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 10), excelSheet.Application.Cells(2, 12)).Font.Size = 11
    excelSheet.Application.Cells(2, 10).Value = Format(Now, "yyyy�Nm��d��")
'---    �R�s��
    excelSheet.Application.Cells(3, 1).Font.Size = 13
    excelSheet.Application.Cells(3, 1).Value = Trim(EX_NAME1)
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 1), excelSheet.Application.Cells(3, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 1), excelSheet.Application.Cells(3, 5)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 1), excelSheet.Application.Cells(3, 5)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
'---    �S�s��
    
    If Trim(EX_NAME2) <> "" Then
    
        excelSheet.Application.Cells(4, 1).Font.Size = 13
        excelSheet.Application.Cells(4, 1).Value = Trim(EX_NAME2)
        excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(4, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(4, 5)).Borders(xlEdgeBottom).Weight = xlThin
        excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(4, 5)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

    End If
'---    �T�s��
    excelSheet.Application.Cells(5, 1).Font.Size = 9
    excelSheet.Application.Cells(5, 1).Value = Trim(EX_BIKOU1)
    
    
    excelSheet.Application.Cells(5, 12).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(5, 12).Value = Trim(EX_SYAMEI)
'---    �U�s��
    excelSheet.Application.Cells(6, 1).Font.Size = 9
    excelSheet.Application.Cells(6, 1).Value = Trim(EX_BIKOU2)
        
    
    excelSheet.Application.Range(excelSheet.Application.Cells(6, 9), excelSheet.Application.Cells(6, 12)).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Range(excelSheet.Application.Cells(6, 9), excelSheet.Application.Cells(6, 12)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(6, 9), excelSheet.Application.Cells(6, 12)).MergeCells = True
    excelSheet.Application.Cells(6, 9).Font.Size = 9
    excelSheet.Application.Cells(6, 9).Value = Trim(EX_ADDR1)
'---    �V�s��
    excelSheet.Application.Range(excelSheet.Application.Cells(7, 9), excelSheet.Application.Cells(7, 12)).HorizontalAlignment = xlRight
'2011.11.21    excelSheet.Application.Range(excelSheet.Application.Cells(7, 9), excelSheet.Application.Cells(7, 12)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(7, 9), excelSheet.Application.Cells(7, 12)).MergeCells = True
    excelSheet.Application.Cells(7, 9).Font.Size = 9
    excelSheet.Application.Cells(7, 9).Value = Trim(EX_ADDR2)


'---    �W�s��
    excelSheet.Application.Cells(8, 10).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(8, 10).Value = Trim(EX_CENTER_NAME)
'---    �X�s��
    excelSheet.Application.Range(excelSheet.Application.Cells(9, 8), excelSheet.Application.Cells(9, 10)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(9, 8), excelSheet.Application.Cells(9, 10)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(9, 8), excelSheet.Application.Cells(9, 10)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(9, 8), excelSheet.Application.Cells(9, 10)).Font.Size = 9
    excelSheet.Application.Cells(9, 8).Value = Trim(EX_CENTER_ADDR1)
    excelSheet.Application.Cells(9, 8).ShrinkToFit = True
        
'---    10�s��
    excelSheet.Application.Range(excelSheet.Application.Cells(10, 8), excelSheet.Application.Cells(10, 10)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(10, 8), excelSheet.Application.Cells(10, 10)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(10, 8), excelSheet.Application.Cells(10, 10)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(10, 8), excelSheet.Application.Cells(10, 10)).Font.Size = 9
    excelSheet.Application.Cells(10, 8).Value = Trim(EX_CENTER_ADDR2)
    excelSheet.Application.Cells(10, 8).ShrinkToFit = True
        




Debug.Print "out Estimate_head_Proc=" & Now

End Sub


Private Function Estimate_SHIZAI_Proc(excelApplication As Object, excelWorkBook As Object, excelSheet As Object, Row As Integer) As Integer
'----------------------------------------------------------------------------
'                   �d�w�b�d�k�i�䌩�Ϗ� ���ށj�o��
'       2009.06.02
'----------------------------------------------------------------------------
Dim i       As Integer
Dim j       As Integer

Dim com     As Integer
Dim sts     As Integer


Dim wkNum1  As Currency
Dim wkNum2  As Currency

Dim SP_TANKA_F  As Boolean  '2012.01.05


Debug.Print "in Estimate_shizai_Proc=" & Now

    Estimate_SHIZAI_Proc = True
'---    14�s��
    excelSheet.Application.Rows(14).RowHeight = 13.5
    excelSheet.Application.Cells(14, 2).Font.Size = 10
    excelSheet.Application.Cells(14, 2).Value = "�y�����ޔ�z"
    
    
'---    15�s��
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Font.Size = 10
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).HorizontalAlignment = xlCenter
    excelSheet.Application.Cells(15, 2).Value = "���ޕi��"
    excelSheet.Application.Cells(15, 4).Value = "���"
    excelSheet.Application.Cells(15, 5).Value = "�`���E�T�C�Y��"
    excelSheet.Application.Cells(15, 8).Value = "����"
    excelSheet.Application.Cells(15, 9).Value = "�P��"
    excelSheet.Application.Cells(15, 10).Value = "�� �z"
    
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 3)).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 3)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 3)).MergeCells = True
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 5), excelSheet.Application.Cells(15, 7)).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Range(excelSheet.Application.Cells(15, 5), excelSheet.Application.Cells(15, 7)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 5), excelSheet.Application.Cells(15, 7)).MergeCells = True

    excelSheet.Application.Range(excelSheet.Application.Cells(15, 10), excelSheet.Application.Cells(15, 11)).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Range(excelSheet.Application.Cells(15, 10), excelSheet.Application.Cells(15, 11)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 10), excelSheet.Application.Cells(15, 11)).MergeCells = True
    
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlDiagonalUp).LineStyle = xlNone
    
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlInsideVertical).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlInsideVertical).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlInsideVertical).ColorIndex = xlAutomatic
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlInsideHorizontal).LineStyle = xlNone
    



'2010.05.13
    excelSheet.Application.Cells(15, 14).Font.Size = 12
    excelSheet.Application.Cells(15, 14).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Cells(15, 14).VerticalAlignment = xlBottom
    excelSheet.Application.Cells(15, 14).Value = "�P��"

    excelSheet.Application.Cells(15, 15).Font.Size = 12
    excelSheet.Application.Cells(15, 15).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Cells(15, 15).VerticalAlignment = xlBottom
    excelSheet.Application.Cells(15, 15).Value = "�`�F�b�N"


    excelSheet.Application.Cells(15, 17).Font.Size = 12
    excelSheet.Application.Cells(15, 17).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Cells(15, 17).VerticalAlignment = xlBottom
    excelSheet.Application.Cells(15, 17).Value = "�`�F�b�N"

'2011.11.21    excelSheet.Application.Cells(16, 17).VerticalAlignment = xlBottom
    
    
'2011.12.12    excelSheet.Application.Cells(16, 17).FormulaR1C1 = Text1(ptxPLUS_KOUSU).Text


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>2012.02.17
    excelSheet.Application.Cells(16, 17).FormulaR1C1 = Text1(ptxCATE_AD_FUKA).Text
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>2012.02.17


'2010.05.13

    
    
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �i���ú��   2012.01.05
    
    SP_TANKA_F = False
    Call UniCode_Conv(K0_ITEM_CATEGORY.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_ITEM_CATEGORY.CATEGORY_CODE, Text1(ptxCATEGORY_CODE).Text)
    sts = BTRV(BtOpGetEqual, ITEM_CATEGORY_POS, ITEM_CATEGORYREC, Len(ITEM_CATEGORYREC), K0_ITEM_CATEGORY, Len(K0_ITEM_CATEGORY), 0)
    Select Case sts
        Case BtNoErr
            If IsNumeric(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_HAKO, vbUnicode)) Then
              SP_TANKA_F = True
            End If
        Case BtErrKeyNotFound
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i���J�e�S���[�}�X�^")
            Exit Function
    
    End Select
        
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �i���ú��   2012.01.05
    
    
    
    
    
    
'---    16�`20�s��
    If EX_SHIZAI_F Then
        
            
        Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
        Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
        Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHin_Gai).Text)
           
        Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_KOSOU)
        Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")
        
        com = BtOpGetGreaterEqual
            
        Row = 15
        Do
            DoEvents
            sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                
                                
                    If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Or _
                        StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                        StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Or _
                        Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxHin_Gai).Text) Then
                    
                        Exit Do
                
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "�\���}�X�^")
                    Exit Function
            End Select
            
        
            If StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) = P_KOSOU Then
                Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, KOSOU_KBN)
            End If
            If StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) = P_GAISOU Then
                Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, GAISO_KBN)
            End If
        
        
            For j = 0 To UBound(EX_SHIZAI_T)
                If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = EX_SHIZAI_T(j) Then
                    
                    
                    If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = FUTAI_KBN Then   '2009.09.05
                    Else
                    
                    
                    
                        Row = Row + 1
                        excelSheet.Application.Cells(Row, 2).Value = Trim(StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                        
                        If IsNumeric(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) Then
                            
                            
                            If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN And CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) <> 0 Then
                                excelSheet.Application.Cells(Row, 8).Value = CDbl(ToHalfAdjust(1 / CCur(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), 3))
                            Else
                                excelSheet.Application.Cells(Row, 8).Value = CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode))
                            End If
                        End If
                    
                        Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN06_CD)
                        Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode))
                        sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                        Select Case sts
                            Case BtNoErr
                                excelSheet.Application.Cells(Row, 4).Value = Trim(StrConv(P_CODEREC.C_RNAME, vbUnicode))
                                
                            Case BtErrKeyNotFound
                            Case Else
                                Call File_Error(sts, com, "�R�[�h�}�X�^")
                                Exit Function
                        End Select
                    
                    
                    
                        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
                        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                        
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                                
                                
                                excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row, 7)).HorizontalAlignment = xlLeft
'2011.11.21                                excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row, 7)).VerticalAlignment = xlBottom
                                excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row, 7)).MergeCells = True
                                
                                
                                excelSheet.Application.Cells(Row, 5).Value = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                                '2009.07.06
                                excelSheet.Application.Cells(Row, 5).ShrinkToFit = True
                                
                                excelSheet.Application.Range(excelSheet.Application.Cells(Row, 8), excelSheet.Application.Cells(Row, 10)).HorizontalAlignment = xlCenter
 
                                
                                Select Case StrConv(ITEMREC.SEI_KBN, vbUnicode)
                                                    
                                    Case "1"
                                        excelSheet.Application.Cells(Row, 9).Value = "�ʔ�"
                                    Case "2"
                                        excelSheet.Application.Cells(Row, 9).Value = "�x��"
                                    Case Else
                                
                                        If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                                            excelSheet.Application.Cells(Row, 9).Value = CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode))
                                        Else
                                            excelSheet.Application.Cells(Row, 9).Value = "�ʔ�"
                                        End If
                                        
                                End Select
                                
                                
                                
                                
                                
                                
                                
                            Case BtErrKeyNotFound
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                Exit Function
                        End Select
                    
                        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).HorizontalAlignment = xlCenter
'2011.11.21                        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).VerticalAlignment = xlBottom
                        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).MergeCells = True
                    
                        If IsNumeric(excelSheet.Application.Cells(Row, 8).Value) And IsNumeric(excelSheet.Application.Cells(Row, 9).Value) Then
                                
                        
                            If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
                        
                        
                                excelSheet.Application.Cells(Row, 10).FormulaR1C1 = "=ROUNDUP(RC[-2]*RC[-1],2)"
                        
                                excelSheet.Application.Cells(Row, 10).NumberFormatLocal = "#,##0.00_ "
                                
                            Else
                                
                                If KUSATU_F Then
                            
                                    excelSheet.Application.Cells(Row, 10).FormulaR1C1 = "=ROUNDUP(RC[-2]*RC[-1],2)"
                                
                                    excelSheet.Application.Cells(Row, 10).NumberFormatLocal = "#,##0.00_ "
                                
                                
                                End If
                                
                                
                                
                            End If
                    
                        
                        
                        
                        
                        
                        End If
                    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �i���ú��   2012.01.05
                        If SP_TANKA_F Then
                            If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
                                excelSheet.Application.Cells(Row, 10).Value = 0
                        
                                excelSheet.Application.Cells(Row, 10).NumberFormatLocal = "#,##0.00_ "
                            End If
                        End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �i���ú��   2012.01.05
                    
                    
                    
                        '2010.05.13
                        excelSheet.Application.Cells(Row, 14).HorizontalAlignment = xlRight
'2011.11.21                        excelSheet.Application.Cells(Row, 14).VerticalAlignment = xlBottom
                        excelSheet.Application.Cells(Row, 14).NumberFormatLocal = "#,##0.00_ "
                        excelSheet.Application.Cells(Row, 14).FormulaR1C1 = "=RC[-4]"


                        excelSheet.Application.Cells(Row, 15).HorizontalAlignment = xlCenter
'2011.11.21                        excelSheet.Application.Cells(Row, 15).VerticalAlignment = xlBottom
                        excelSheet.Application.Cells(Row, 15).FormulaR1C1 = "=IF(RC[-5]=RC[-1],""��"",""�~"")"
'
                        '2010.05.13
                    
                    
                    
                    
                    
                    
                    End If  '2009.09.05
                
                
                
                
                
                
                
                
                
                
                End If
            
            
            Next j
        
            com = BtOpGetNext
        
        Loop
        '�ް��������
        If Trim(EX_BCR_CODE) <> "" Then
        
            If IsNumeric(Trim(Text1(ptxMAIN_KOUTEI_QTY01).Text)) Then
                If CDbl(Trim(Text1(ptxMAIN_KOUTEI_QTY01).Text)) > 0 Then
                    Row = Row + 1
                
                    excelSheet.Application.Cells(Row, 2).Value = Trim(EX_BCR_CODE)

                
                    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, EX_BCR_CODE)
                    
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                            excelSheet.Application.Cells(Row, 5).Value = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                        Case BtErrKeyNotFound
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                            Exit Function
                    End Select
                
                                    
                    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 8), excelSheet.Application.Cells(Row, 10)).HorizontalAlignment = xlCenter
        
        '            excelSheet.Application.Cells(row, 9).NumberFormatLocal = "#,##0_ "
                    excelSheet.Application.Cells(Row, 8).Value = CDbl(Trim(Text1(ptxMAIN_KOUTEI_QTY01).Text))
                    excelSheet.Application.Cells(Row, 9).Value = "�ʔ�"
                
                    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).HorizontalAlignment = xlCenter
'2011.11.21                    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).VerticalAlignment = xlBottom
                    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).MergeCells = True
                
                
                     '2010.05.13
                    excelSheet.Application.Cells(Row, 14).HorizontalAlignment = xlRight
'2011.11.21                    excelSheet.Application.Cells(Row, 14).VerticalAlignment = xlBottom
                    excelSheet.Application.Cells(Row, 14).NumberFormatLocal = "#,##0.00_ "
                    excelSheet.Application.Cells(Row, 14).FormulaR1C1 = "=RC[-4]"


                    excelSheet.Application.Cells(Row, 15).HorizontalAlignment = xlCenter
'2011.11.21                    excelSheet.Application.Cells(Row, 15).VerticalAlignment = xlBottom
                    excelSheet.Application.Cells(Row, 15).FormulaR1C1 = "=IF(RC[-5]=RC[-1],""��"",""�~"")"
                    
                    '2010.05.13
               
                
                
                End If
            End If
        End If
    
    
'---    ���׌r��
        
        If Row <> 15 Then
        
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlDiagonalUp).LineStyle = xlNone
        
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeLeft).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
        
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeTop).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
        
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeBottom).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeRight).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
        
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideVertical).LineStyle = xlNone
        
            If Row > 16 Then
                excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
                excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideHorizontal).Weight = xlThin
                excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
            End If
    
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 4), excelSheet.Application.Cells(Row, 4)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 4), excelSheet.Application.Cells(Row, 4)).Borders(xlEdgeLeft).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 4), excelSheet.Application.Cells(Row, 4)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 6), excelSheet.Application.Cells(Row, 5)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 6), excelSheet.Application.Cells(Row, 5)).Borders(xlEdgeLeft).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 6), excelSheet.Application.Cells(Row, 5)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 9), excelSheet.Application.Cells(Row, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 9), excelSheet.Application.Cells(Row, 8)).Borders(xlEdgeLeft).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 9), excelSheet.Application.Cells(Row, 8)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 10), excelSheet.Application.Cells(Row, 9)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 10), excelSheet.Application.Cells(Row, 9)).Borders(xlEdgeLeft).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 10), excelSheet.Application.Cells(Row, 9)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 11), excelSheet.Application.Cells(Row, 10)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 11), excelSheet.Application.Cells(Row, 10)).Borders(xlEdgeLeft).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 11), excelSheet.Application.Cells(Row, 10)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
        End If

'        If row <> 15 Or (IsNumeric(Text1(ptxBEF_S_SHIZAI_BAIKA).Text) And Val(Text1(ptxBEF_S_SHIZAI_BAIKA).Text) > 0) Then
'---    27�s��
            Row = Row + 1
            excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row, 9)).HorizontalAlignment = xlRight
            excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row, 9)).VerticalAlignment = xlCenter
            excelSheet.Application.Cells(Row, 9).Value = "�@�����ލ��v���z"
        
            excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).HorizontalAlignment = xlCenter
            excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).VerticalAlignment = xlCenter
            excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).MergeCells = True
            excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).Font.Size = 14
                
''2009.07.01            excelSheet.Application.Cells(row, 11).FormulaR1C1 = "=SUM(R[-1]C:R[" & -row + 15 & "]C)"
            
            
            '���v���z�G���[�`�F�b�N 2009.09.05
            excelSheet.Application.Cells(Row, 10).FormulaR1C1 = "=SUM(R[-1]C:R[" & -Row + 15 & "]C)"
            
            If IsNumeric(excelSheet.Application.Cells(Row, 10).Value) Then
                wkNum1 = CCur(excelSheet.Application.Cells(Row, 10).Value)
            Else
                wkNum1 = 0
            End If
            
            If IsNumeric(Text1(ptxBEF_S_SHIZAI_BAIKA).Text) Then
                wkNum2 = CCur(Text1(ptxBEF_S_SHIZAI_BAIKA).Text)
            Else
                wkNum2 = 0
            End If
                        
            
            If IsNumeric(Text1(ptxBEF_S_GAISO_TANKA).Text) Then
                wkNum2 = CCur(wkNum2 + CCur(Text1(ptxBEF_S_GAISO_TANKA).Text))
            End If
            
            
            
'Debug.Print wkNum1 - wkNum2
            
'            If CDbl(excelSheet.Application.Cells(row, 10).Value) <> (CDbl(Text1(ptxBEF_S_SHIZAI_BAIKA).Text) + CDbl(Text1(ptxBEF_S_GAISO_TANKA).Text)) Then
            If wkNum1 <> wkNum2 Then
                MsgBox "�@�����ލ��v���z�������ޖ��ׂ̍��v���z�ƈقȂ�܂��B"
                excelSheet.Application.Cells(Row, 13).Value = "�@�����ލ��v���z�������ޖ��ׂ̍��v���z�ƈقȂ�܂��B"
            End If
            
            
            If IsNumeric(Text1(ptxBEF_S_SHIZAI_BAIKA).Text) Then
                excelSheet.Application.Cells(Row, 10).Value = Val(Text1(ptxBEF_S_SHIZAI_BAIKA).Text)
            Else
                excelSheet.Application.Cells(Row, 10).Value = 0
            End If
'2009.07.06
            If IsNumeric(Text1(ptxAFT_S_GAISO_TANKA).Text) Then
                excelSheet.Application.Cells(Row, 10).Value = Val(excelSheet.Application.Cells(Row, 10).Value) + Val(Text1(ptxBEF_S_GAISO_TANKA).Text)
            End If
            
            
            excelSheet.Application.Cells(Row, 10).NumberFormatLocal = "#,##0.00_ "
            
            
            
            
            excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
            excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlDiagonalUp).LineStyle = xlNone
        
            excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeLeft).Weight = xlThick
            excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
        
            excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeTop).Weight = xlThick
            excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
        
            excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeBottom).Weight = xlThick
            excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        
            excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeRight).Weight = xlThick
            excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
        
            excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideVertical).LineStyle = xlNone
        
            excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideHorizontal).LineStyle = xlNone
        
        
            excelSheet.Application.Cells(Row, 10).Borders(xlEdgeLeft).LineStyle = xlContinuous
            excelSheet.Application.Cells(Row, 10).Borders(xlEdgeLeft).Weight = xlThick
            excelSheet.Application.Cells(Row, 10).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
        
'        End If
    
    
            '2010.05.13
            excelSheet.Application.Cells(Row, 14).HorizontalAlignment = xlRight
'2011.11.21            excelSheet.Application.Cells(Row, 14).VerticalAlignment = xlBottom
            excelSheet.Application.Cells(Row, 14).NumberFormatLocal = "#,##0.00_ "

            If (-Row + 16) = 0 Then
                excelSheet.Application.Cells(Row, 14).Value = 0
            Else
                excelSheet.Application.Cells(Row, 14).FormulaR1C1 = "=SUM(R[-1]C:R[" & -Row + 16 & "]C)"
            End If

            excelSheet.Application.Cells(Row, 15).HorizontalAlignment = xlCenter
'2011.11.21            excelSheet.Application.Cells(Row, 15).VerticalAlignment = xlBottom
            excelSheet.Application.Cells(Row, 15).FormulaR1C1 = "=IF(RC[-5]=RC[-1],""��"",""�~"")"
            
            '2010.05.13
    
    
    End If

    Estimate_SHIZAI_Proc = False

Debug.Print "out Estimate_shizai_Proc=" & Now

End Function


Private Function Estimate_DOUKON_Proc(excelApplication As Object, excelWorkBook As Object, excelSheet As Object, Row As Integer) As Integer
'----------------------------------------------------------------------------
'                   �d�w�b�d�k�i�䌩�Ϗ� �����j�o��
'       2009.06.02
'----------------------------------------------------------------------------
Dim i           As Integer
Dim j           As Integer
Dim k           As Integer
Dim l           As Integer


Dim com         As Integer
Dim sts         As Integer

Dim start_row   As Integer


    Estimate_DOUKON_Proc = True
    Row = Row + 2

'---    29�s��
    excelSheet.Application.Cells(Row, 2).Font.Size = 10
    excelSheet.Application.Cells(Row, 2).Value = "�y�������i���ׁz"
    
'---    �������i��
    Row = Row + 1
        
        
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 3)).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 3)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 3)).MergeCells = True
    

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row, 7)).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row, 7)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row, 7)).MergeCells = True

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).MergeCells = True
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlDiagonalUp).LineStyle = xlNone
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideVertical).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideVertical).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideVertical).ColorIndex = xlAutomatic
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideHorizontal).LineStyle = xlNone
    
    
    
    
    
    
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 8)).HorizontalAlignment = xlCenter
    excelSheet.Application.Cells(Row, 2).Value = "�����i��"
    excelSheet.Application.Cells(Row, 4).Value = "���"
    excelSheet.Application.Cells(Row, 5).Value = "�i��"
    excelSheet.Application.Cells(Row, 8).Value = "����"
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 8)).Font.Size = 10
    
    start_row = Row
'---    31�`40�s��
    If EX_DOUKON_F Then
        
            
        Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
        Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
        Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHin_Gai).Text)
           
        Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_DOUKON)
        Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")
        
        com = BtOpGetGreaterEqual
            
        Do
           
            DoEvents
           
            sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                
                                
                    If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Or _
                        StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                        StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Or _
                        Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxHin_Gai).Text) Then
                    
                        Exit Do
                
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "�\���}�X�^")
                    Exit Function
            End Select
            
        
            If StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) <> P_DOUKON Then
                Exit Do
            End If
        
        
            For j = 0 To UBound(EX_DOUKON_T)
                If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = EX_DOUKON_T(j) Then
                    
                    
                    
                    
                    Row = Row + 1

                    
                    
                    
                    excelSheet.Application.Cells(Row, 2).Value = Trim(StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                    
                    
                    
                    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN06_CD)
                    Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode))
                    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                    Select Case sts
                        Case BtNoErr
                            excelSheet.Application.Cells(Row, 4).Value = Trim(StrConv(P_CODEREC.C_RNAME, vbUnicode))
                            
                        Case BtErrKeyNotFound
                        Case Else
                            Call File_Error(sts, com, "�R�[�h�}�X�^")
                            Exit Function
                    End Select
                    
                    
                    
                    
                    If IsNumeric(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) Then
                        excelSheet.Application.Cells(Row, 8).NumberFormatLocal = "#,##0_ "
                        excelSheet.Application.Cells(Row, 8).HorizontalAlignment = xlCenter
                        excelSheet.Application.Cells(Row, 8).Value = CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode))
                    End If
                
                
                
                
                    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                    
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                            
                            
                            
                            excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row, 7)).HorizontalAlignment = xlLeft
'2011.11.21                            excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row, 7)).VerticalAlignment = xlBottom
                            excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row, 7)).MergeCells = True
                            
                            
                            excelSheet.Application.Cells(Row, 5).Value = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                            '2009.07.06
                            excelSheet.Application.Cells(Row, 5).ShrinkToFit = True
                            
                            
                        Case BtErrKeyNotFound
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                            Exit Function
                    End Select
                
                
                
                End If
            
                com = BtOpGetNext
            
            
            Next j
        
        
        
        
        
        
        
        
        
        Loop
    
    
    
    
    End If
    
    If Row <> start_row Then
    
    
    
    
            start_row = start_row + 1
    
    
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlDiagonalUp).LineStyle = xlNone
        
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeLeft).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
        
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeTop).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
        
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeBottom).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeRight).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
        
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideVertical).LineStyle = xlNone
        
            If start_row <> Row Then
                excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
                excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideHorizontal).Weight = xlThin
                excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
            End If

            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 4), excelSheet.Application.Cells(Row, 4)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 4), excelSheet.Application.Cells(Row, 4)).Borders(xlEdgeLeft).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 4), excelSheet.Application.Cells(Row, 4)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 5), excelSheet.Application.Cells(Row, 5)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 5), excelSheet.Application.Cells(Row, 5)).Borders(xlEdgeLeft).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 5), excelSheet.Application.Cells(Row, 5)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 8), excelSheet.Application.Cells(Row, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 8), excelSheet.Application.Cells(Row, 8)).Borders(xlEdgeLeft).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 8), excelSheet.Application.Cells(Row, 8)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 9), excelSheet.Application.Cells(Row, 9)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 9), excelSheet.Application.Cells(Row, 9)).Borders(xlEdgeLeft).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 9), excelSheet.Application.Cells(Row, 9)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 10), excelSheet.Application.Cells(Row, 10)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 10), excelSheet.Application.Cells(Row, 10)).Borders(xlEdgeLeft).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 10), excelSheet.Application.Cells(Row, 10)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
    
    
    
    
    
    
    End If
    
    
    
    
    Estimate_DOUKON_Proc = False
End Function



Private Function Estimate_FUKA_Proc(excelApplication As Object, excelWorkBook As Object, excelSheet As Object, Row As Integer) As Integer
'----------------------------------------------------------------------------
'                   �d�w�b�d�k�i�䌩�Ϗ� �t����Ɓj�o��
'       2009.06.02
'----------------------------------------------------------------------------
Dim i           As Integer
Dim j           As Integer
Dim k           As Integer
Dim l           As Integer


Dim com         As Integer
Dim sts         As Integer

Dim start_row   As Integer

Dim wkNum1      As Currency
Dim wkNum2      As Currency


Debug.Print "in Estimate_FUKA_Proc=" & Now

    Estimate_FUKA_Proc = True
    Row = Row + 2

'---    25�s��
    excelSheet.Application.Cells(Row, 2).Font.Size = 10
    excelSheet.Application.Cells(Row, 2).Value = "�y�t����Ɣ�z"
    
'---    �t����Ɨ�
    Row = Row + 1
        
        
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 9)).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 9)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 9)).MergeCells = True
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).MergeCells = True

    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlDiagonalUp).LineStyle = xlNone
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideVertical).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideVertical).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideVertical).ColorIndex = xlAutomatic
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideHorizontal).LineStyle = xlNone
    
    
    
    
    
    
    
    
    excelSheet.Application.Cells(Row, 2).Value = "��Ɠ��e"
    excelSheet.Application.Cells(Row, 10).Value = "�H��(�b)"
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 10)).Font.Size = 10
    
    start_row = Row

    
'---    26�`35�s��
    If EX_FUKA_F Then
        
            
        Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
        Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
        Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHin_Gai).Text)
           
        Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_DOUKON)
        Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")
        
        com = BtOpGetGreaterEqual
            
        Do
           
            DoEvents
           
            sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                
                                
                    If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Or _
                        StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                        StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Or _
                        Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxHin_Gai).Text) Then
                    
                        Exit Do
                
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "�\���}�X�^")
                    Exit Function
            End Select
            
        
            If StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) <> P_DOUKON Then
                Exit Do
            End If
        
        
            For j = 0 To UBound(EX_FUKA_T)
                If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = EX_FUKA_T(j) Then
                    
                    
                    
                    
                    Row = Row + 1

                    
                    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                    
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                            
                            If Not IsNumeric(StrConv(ITEMREC.S_KOUSU, vbUnicode)) Then
                                Call UniCode_Conv(ITEMREC.S_KOUSU, "00000000")
                            End If
                            
                            
                            
                        Case BtErrKeyNotFound
                        
                        
                            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                            Call UniCode_Conv(ITEMREC.S_KOUSU, "00000000")
                        
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                            Exit Function
                    End Select
                    
                    
                                        
                    
                    excelSheet.Application.Cells(Row, 2).Value = Trim(StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)) & " " & _
                                                                    Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)) & " " & _
                                                                    Trim(StrConv(P_COMPO_K_REC.KO_BIKOU, vbUnicode))

                    
                    If IsNumeric(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) Then
'                        excelSheet.Application.Cells(row, 11).NumberFormatLocal = "#,##0_ "
                        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).HorizontalAlignment = xlCenter
'2011.11.21                        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).VerticalAlignment = xlBottom
                        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).MergeCells = True
                        excelSheet.Application.Cells(Row, 10).Value = CDbl(StrConv(ITEMREC.S_KOUSU, vbUnicode))
' 2013.01.11 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                        excelSheet.Application.Cells(Row, 14).FormulaR1C1 = "=roundup(RC[-4]/60*" & Val(Text1(ptxBEF_SEI_RATE).Text) & ",2)"
' 2013.01.11 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                    End If
                
                
                
                    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 10)).Font.Size = 11

                
                
                
                End If
            
                com = BtOpGetNext
            
            
            Next j
        
        
        
        
        
        
        
        
        
        Loop
    
    
    
    
    End If
    
    If Row <> start_row Then
    
    
    
    
    
        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlDiagonalUp).LineStyle = xlNone
    
        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeLeft).Weight = xlThin
        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeTop).Weight = xlThin
        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
    
        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeBottom).Weight = xlThin
        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    
        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeRight).Weight = xlThin
        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
    
        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideVertical).LineStyle = xlNone
    
        If Row = start_row + 1 Then
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideHorizontal).LineStyle = xlNone
        Else
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideHorizontal).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
        End If


        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 10), excelSheet.Application.Cells(Row, 10)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 10), excelSheet.Application.Cells(Row, 10)).Borders(xlEdgeLeft).Weight = xlThin
        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 10), excelSheet.Application.Cells(Row, 10)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
    
    
    
    
    
    
    End If
    
'---    �t����Ɨ��i���o���j
    Row = Row + 1
        
        
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 6), excelSheet.Application.Cells(Row, 8)).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 6), excelSheet.Application.Cells(Row, 8)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 6), excelSheet.Application.Cells(Row, 8)).Font.Size = 10
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 6), excelSheet.Application.Cells(Row, 8)).MergeCells = True
    excelSheet.Application.Cells(Row, 6).Value = "��Ǝ��Ԍv(�b/��)"
    excelSheet.Application.Cells(Row, 6).ShrinkToFit = True
    
    
'    excelSheet.Application.Cells(row, 8).HorizontalAlignment = xlCenter
'    excelSheet.Application.Cells(row, 8).VerticalAlignment = xlBottom
'    excelSheet.Application.Cells(row, 8).Font.Size = 10
'    excelSheet.Application.Cells(row, 8).Value = "��/��"
    
    excelSheet.Application.Cells(Row, 9).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Cells(Row, 9).VerticalAlignment = xlBottom
    excelSheet.Application.Cells(Row, 9).Font.Size = 10
    excelSheet.Application.Cells(Row, 9).Value = "�����[�g"
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).Font.Size = 12
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).MergeCells = True
    excelSheet.Application.Cells(Row, 10).Value = "�A�t����Ɣ�"
    
'---    �t����Ɨ��i���e�j
    Row = Row + 1
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 6), excelSheet.Application.Cells(Row, 8)).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 6), excelSheet.Application.Cells(Row, 8)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 6), excelSheet.Application.Cells(Row, 8)).Font.Size = 12
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 6), excelSheet.Application.Cells(Row, 8)).MergeCells = True
    If (Row - 2) = start_row Then
        excelSheet.Application.Cells(Row, 6).Value = 0
    Else
        excelSheet.Application.Cells(Row, 6).FormulaR1C1 = "=SUM(R[-2]C[4]:R[" & start_row - Row + 1 & "]C[4]"
    End If
    
'    excelSheet.Application.Cells(row, 8).HorizontalAlignment = xlCenter
'    excelSheet.Application.Cells(row, 8).VerticalAlignment = xlBottom
'    excelSheet.Application.Cells(row, 8).Font.Size = 12
'    excelSheet.Application.Cells(row, 8).FormulaR1C1 = "=round(RC[-2]/60,2)"


    excelSheet.Application.Cells(Row, 9).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Cells(Row, 9).VerticalAlignment = xlBottom
    excelSheet.Application.Cells(Row, 9).Font.Size = 12
    excelSheet.Application.Cells(Row, 9).Value = Text1(ptxBEF_SEI_RATE).Text

' 2013.01.11 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'    excelSheet.Application.Cells(Row, 10).FormulaR1C1 = "=round(RC[-4]/60*RC[-1],2)"
    excelSheet.Application.Cells(Row, 10).FormulaR1C1 = "=SUM(R[" & start_row - Row & "]C[4]:R[-2]C[4]"
' 2013.01.11 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    
    If IsNumeric(excelSheet.Application.Cells(Row, 10).Value) Then
        wkNum1 = CCur(excelSheet.Application.Cells(Row, 10).Value)
    Else
        wkNum1 = 0
    End If
    
    
    If IsNumeric(Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text) Then
        wkNum2 = CCur(Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text)
    Else
        wkNum2 = 0
    End If
    
    If wkNum1 <> wkNum2 Then
        MsgBox "�A�t����Ɣ�v�Z�l(��/�~�����[�g)�ƈقȂ�܂��B"
        excelSheet.Application.Cells(Row, 13).Value = "�A�t����Ɣ�v�Z�l(��/�~�����[�g)�ƈقȂ�܂��B"
    End If
    
    
    
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).Font.Size = 14
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).MergeCells = True
    excelSheet.Application.Cells(Row, 10).Value = Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text
    excelSheet.Application.Cells(Row, 10).NumberFormatLocal = "#,##0.00_ "

    
    
'2010.05.13
    excelSheet.Application.Cells(Row - 1, 14).Font.Size = 12
    excelSheet.Application.Cells(Row - 1, 14).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Cells(Row - 1, 14).VerticalAlignment = xlBottom
    excelSheet.Application.Cells(Row - 1, 14).Value = "�P��"

' 2013.01.11 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'    excelSheet.Application.Cells(Row, 14).FormulaR1C1 = "=round(RC[-8]/60*RC[-5],2)"
    excelSheet.Application.Cells(Row, 14).FormulaR1C1 = "=SUM(R[" & start_row - Row & "]C:R[-2]C"
' 2013.01.11 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    excelSheet.Application.Cells(Row, 14).NumberFormatLocal = "#,##0.00_ "
    excelSheet.Application.Cells(Row, 14).HorizontalAlignment = xlRight
'2011.11.21    excelSheet.Application.Cells(Row, 14).VerticalAlignment = xlBottom


    excelSheet.Application.Cells(Row - 1, 15).Font.Size = 12
    excelSheet.Application.Cells(Row - 1, 15).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Cells(Row - 1, 15).VerticalAlignment = xlBottom
    excelSheet.Application.Cells(Row - 1, 15).Value = "�`�F�b�N"

    excelSheet.Application.Cells(Row, 15).HorizontalAlignment = xlCenter
'2011.11.21    excelSheet.Application.Cells(Row, 15).VerticalAlignment = xlBottom
    excelSheet.Application.Cells(Row, 15).FormulaR1C1 = "=IF(RC[-5]=RC[-1],""��"",""�~"")"


'2010.05.13
    
    
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 6), excelSheet.Application.Cells(Row, 10)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 6), excelSheet.Application.Cells(Row, 10)).Borders(xlDiagonalUp).LineStyle = xlNone
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 6), excelSheet.Application.Cells(Row, 10)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 6), excelSheet.Application.Cells(Row, 10)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 6), excelSheet.Application.Cells(Row, 10)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 6), excelSheet.Application.Cells(Row, 10)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 6), excelSheet.Application.Cells(Row, 10)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 6), excelSheet.Application.Cells(Row, 10)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 6), excelSheet.Application.Cells(Row, 10)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 6), excelSheet.Application.Cells(Row, 10)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 6), excelSheet.Application.Cells(Row, 10)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 6), excelSheet.Application.Cells(Row, 10)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 6), excelSheet.Application.Cells(Row, 10)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 6), excelSheet.Application.Cells(Row, 10)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 6), excelSheet.Application.Cells(Row, 10)).Borders(xlInsideVertical).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 6), excelSheet.Application.Cells(Row, 10)).Borders(xlInsideVertical).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 6), excelSheet.Application.Cells(Row, 10)).Borders(xlInsideVertical).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 6), excelSheet.Application.Cells(Row, 10)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 6), excelSheet.Application.Cells(Row, 10)).Borders(xlInsideHorizontal).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 6), excelSheet.Application.Cells(Row, 10)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
    
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlDiagonalUp).LineStyle = xlNone
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeLeft).Weight = xlThick
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeTop).Weight = xlThick
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeBottom).Weight = xlThick
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeRight).Weight = xlThick
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideVertical).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideVertical).Weight = xlThick
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideVertical).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideHorizontal).Weight = xlThick
    excelSheet.Application.Range(excelSheet.Application.Cells(Row - 1, 10), excelSheet.Application.Cells(Row, 11)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
    
    
    
    
    
    Estimate_FUKA_Proc = False

Debug.Print "out Estimate_FUKA_Proc=" & Now

End Function


Private Sub Estimate_Line11_13_Proc(excelApplication As Object, excelWorkBook As Object, excelSheet As Object)
'----------------------------------------------------------------------------
'                   �d�w�b�d�k�i�䌩�Ϗ� 11-13�s�ځj�o��
'----------------------------------------------------------------------------
    
    
    
    
    
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2011.12.15 �i���ú�ؑΉ�
    excelSheet.Application.Cells(11, 1).Font.Size = 10
    excelSheet.Application.Cells(11, 1).VerticalAlignment = xlCenter
    excelSheet.Application.Cells(11, 1).HorizontalAlignment = xlLeft
    excelSheet.Application.Cells(11, 1).Value = "�i���J�e�S���["
    
    
    
    excelSheet.Application.Cells(11, 3).Font.Size = 11
    excelSheet.Application.Cells(11, 3).VerticalAlignment = xlCenter
    excelSheet.Application.Cells(11, 3).HorizontalAlignment = xlLeft
    excelSheet.Application.Cells(11, 3).Value = Trim(Left(Combo1(pcmbCATEGORY_Name).Text, Len(Combo1(pcmbCATEGORY_Name).Text) - 8))
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2011.12.15 �i���ú�ؑΉ�
    
    
    excelSheet.Application.Rows(12).RowHeight = 23.25
    
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 1), excelSheet.Application.Cells(12, 2)).HorizontalAlignment = xlLeft
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 1), excelSheet.Application.Cells(12, 2)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 1), excelSheet.Application.Cells(12, 2)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 1), excelSheet.Application.Cells(12, 2)).Font.Size = 14
    excelSheet.Application.Cells(12, 1).Value = "���i�i��"

        
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 3)).HorizontalAlignment = xlLeft
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 3)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 3)).Font.Size = 16
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 3)).Font.NAME = "�l�r�@�S�V�b�N"
    excelSheet.Application.Cells(12, 3).Value = Trim(Text1(ptxHin_Gai).Text)
    
    
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlDiagonalUp).LineStyle = xlNone

    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlEdgeLeft).Weight = xlThick
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlEdgeTop).Weight = xlThick
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlEdgeBottom).Weight = xlThick
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlEdgeRight).Weight = xlThick
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlInsideVertical).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlInsideHorizontal).LineStyle = xlNone


    excelSheet.Application.Cells(12, 6).Font.Size = 10
    excelSheet.Application.Cells(12, 6).VerticalAlignment = xlCenter
    excelSheet.Application.Cells(12, 6).HorizontalAlignment = xlLeft
    excelSheet.Application.Cells(12, 6).Value = Trim(Text1(ptxHin_Name).Text)


'---    13�s��
    excelSheet.Application.Rows(11).RowHeight = 13.5

End Sub

Private Function CATEGORY_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   �i���J�e�S�����̕\��
'----------------------------------------------------------------------------
Dim i       As Integer
Dim Row     As Integer
    
Dim sts     As Integer
    
    CATEGORY_Disp_Proc = True
    
    
    

    
    
    
    '-----------------------------------    �O��H��
    If Trim(Text1(ptxCATEGORY_CODE).Text) = "" Then
        For i = ptxCATE_ST_KOUTEI To ptxCATE_AD_KOURYO
            '2012.01.28
            If i = ptxCATE_ST_FUKA Or i = ptxCATE_AD_FUKA Then
            Else
                Text1(i).Text = ""
            End If
        Next i
    Else
        '�O��H��
        Call UniCode_Conv(K0_ITEM_CATEGORY.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
        Call UniCode_Conv(K0_ITEM_CATEGORY.CATEGORY_CODE, Text1(ptxCATEGORY_CODE).Text)
            
        sts = BTRV(BtOpGetEqual, ITEM_CATEGORY_POS, ITEM_CATEGORYREC, Len(ITEM_CATEGORYREC), K0_ITEM_CATEGORY, Len(K0_ITEM_CATEGORY), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
    
                Call UniCode_Conv(ITEM_CATEGORYREC.SEI_LOT, "0000000000")
                Call UniCode_Conv(ITEM_CATEGORYREC.KOUSU_LOT, "0000000000")
                Call UniCode_Conv(ITEM_CATEGORYREC.KOUSU_QTY, "0000000000")
    
                Call UniCode_Conv(ITEM_CATEGORYREC.TOKU_TANKA_HAKO, "")
                Call UniCode_Conv(ITEM_CATEGORYREC.TOKU_TANKA_KOURYO, "")
    
    
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�i���J�e�S���}�X�^")
                Exit Function
    
        End Select
        'ۯĐ�
        If IsNumeric(StrConv(ITEM_CATEGORYREC.SEI_LOT, vbUnicode)) Then
            If Val(StrConv(ITEM_CATEGORYREC.SEI_LOT, vbUnicode)) <> 0 Then
                Text1(ptxAFT_SEI_LOT).Text = Val(StrConv(ITEM_CATEGORYREC.SEI_LOT, vbUnicode))
            End If
        End If
        '�O��H���@�b
        If IsNumeric(StrConv(ITEM_CATEGORYREC.KOUSU_LOT, vbUnicode)) Then
            Text1(ptxCATE_ST_KOUTEI).Text = Val(StrConv(ITEM_CATEGORYREC.KOUSU_QTY, vbUnicode))
        End If
        If IsNumeric(StrConv(ITEM_CATEGORYREC.KOUSU_LOT, vbUnicode)) Then
            Text1(ptxCATE_AD_KOUTEI).Text = Val(StrConv(ITEM_CATEGORYREC.KOUSU_QTY, vbUnicode))
        End If
        
    End If
        
    '�t���H���@�b
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.28
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.24
'    If IsNumeric(StrConv(ITEMREC.CATE_ST_FUKA, vbUnicode)) Then
'        Text1(ptxCATE_ST_FUKA).Text = Format(Val(StrConv(ITEMREC.CATE_ST_FUKA, vbUnicode)), "#")
'        Text1(ptxCATE_AD_FUKA).Text = Format(Val(StrConv(ITEMREC.CATE_ST_FUKA, vbUnicode)), "#")
'    Else
'        Text1(ptxCATE_ST_FUKA).Text = ""
'        Text1(ptxCATE_AD_FUKA).Text = ""
'    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.24
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.28
    
    
    
    '����ƍH��1�@�b
    Text1(ptxCATE_ST_JITU1).Text = Val(Text1(ptxMAIN_KOUTEI_KEI1).Text)
    Text1(ptxCATE_AD_JITU1).Text = Val(Text1(ptxMAIN_KOUTEI_KEI1).Text)
    '�]�T�� ��
    If IsNumeric(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) Then
        Text1(ptxCATE_ST_YOYU_RITU).Text = Format(CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)), "#0.00")
        Text1(ptxCATE_AD_YOYU_RITU).Text = Format(CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)), "#0.00")
    Else
        Text1(ptxCATE_ST_YOYU_RITU).Text = ""
        Text1(ptxCATE_AD_YOYU_RITU).Text = ""
    End If
    '����ƍH��2�@�b(�l�̌ܓ�)
    
    If IsNumeric(Text1(ptxMAIN_KOUTEI_KEI1).Text) And _
        IsNumeric(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) Then
    
        Text1(ptxCATE_ST_JITU2).Text = ToHalfAdjust(CCur(CCur(Text1(ptxMAIN_KOUTEI_KEI1).Text) * _
                                                    CCur(CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)))), 0)
        Text1(ptxCATE_AD_JITU2).Text = ToHalfAdjust(CCur(CCur(Text1(ptxMAIN_KOUTEI_KEI1).Text) * _
                                                    CCur(CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)))), 0)
    End If
    '��Ǝ��Ԍv
    Text1(ptxCATE_ST_TOTAL) = Val(Text1(ptxCATE_ST_KOUTEI).Text) + _
                                Val(Text1(ptxCATE_ST_FUKA).Text) + _
                                Val(Text1(ptxCATE_ST_JITU2).Text)
    Text1(ptxCATE_AD_TOTAL) = Val(Text1(ptxCATE_AD_KOUTEI).Text) + _
                                Val(Text1(ptxCATE_AD_FUKA).Text) + _
                                Val(Text1(ptxCATE_AD_JITU2).Text)


    '��/��
    Text1(ptxCATE_ST_FUN).Text = Format(ToHalfAdjust(CCur(Val(Text1(ptxCATE_ST_TOTAL)) / 60), 2), "#0.00")
    Text1(ptxCATE_AD_FUN).Text = Format(ToHalfAdjust(CCur(Val(Text1(ptxCATE_AD_TOTAL)) / 60), 2), "#0.00")
    
    '�H��
    Text1(ptxAFT_S_KOUSU).Text = Text1(ptxCATE_AD_FUN).Text
    
    
    '�����[�g (�~ / ��)
    If IsNumeric(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)) Then
        Text1(ptxCATE_ST_FUN_RATE).Text = Format(Val(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)), "#0.00")
        Text1(ptxCATE_AD_FUN_RATE).Text = Format(Val(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)), "#0.00")
    Else

        Text1(ptxCATE_ST_FUN_RATE).Text = ""
        Text1(ptxCATE_AD_FUN_RATE).Text = ""
    End If
    '(�~�^��)
    If IsNumeric(Text1(ptxCATE_ST_FUN_RATE)) Then
        Text1(ptxCATE_ST_KOURYO).Text = Format(ToRoundUp(CCur(Val(Text1(ptxCATE_ST_FUN).Text) * Val(Text1(ptxCATE_ST_FUN_RATE).Text)), 2), "#0.00")
        Text1(ptxCATE_AD_KOURYO).Text = Format(ToRoundUp(CCur(Val(Text1(ptxCATE_AD_FUN).Text) * Val(Text1(ptxCATE_AD_FUN_RATE).Text)), 2), "#0.00")

    Else
        Text1(ptxCATE_ST_KOURYO).Text = "0.00"
        Text1(ptxCATE_AD_KOURYO).Text = "0.00"
    End If
    '-----------------------------------    �ύX�O�^�ύX��i�W�v�l�j
    
    
'    '�H��
'    Text1(ptxAFT_S_KOUSU).Text = Text1(ptxCATE_AD_FUN).Text
'    '�H��
'    If IsNumeric(Text1(ptxCATE_AD_KOURYO).Text) Then
'        Text1(ptxAFT_S_KOUSU_BAIKA).Text = Format(CDbl(Text1(ptxCATE_AD_KOURYO).Text), "#0.00")
'    Else
'        Text1(ptxAFT_S_KOUSU_BAIKA).Text = ""
'    End If
    
    
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   ���ʒP��
'    If Trim(Text1(ptxCATEGORY_CODE).Text) = "" Then
'    Else
'        '���ʒP���ł̏���
'        For Row = 1 To KOUSEI.Count(1)
'            '��ƍH���@�b/��
'            For i = 0 To UBound(SP_KOUSU_T)
'                If SP_KOUSU_T(i) = Trim(Right(KOUSEI(Row, ColKO_SYUBETSU), 2)) Then
'
'
'                    If IsNumeric(StrConv(ITEMREC.G_SPTAN, vbUnicode)) Then
'                        KOUSEI(Row, ColG_ST_URIKIN) = Val(StrConv(ITEMREC.G_SPTAN, vbUnicode))
'                    Else
'
'                        Call UniCode_Conv(K0_ITEM_CATEGORY.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
'                        Call UniCode_Conv(K0_ITEM_CATEGORY.CATEGORY_CODE, Text1(ptxCATEGORY_CODE).Text)
'                        sts = BTRV(BtOpGetEqual, ITEM_CATEGORY_POS, ITEM_CATEGORYREC, Len(ITEM_CATEGORYREC), K0_ITEM_CATEGORY, Len(K0_ITEM_CATEGORY), 0)
'                        Select Case sts
'                            Case BtNoErr
'                                If IsNumeric(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_QTY, vbUnicode)) Then
'                                    KOUSEI(Row, ColG_SPTAN) = Val(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_QTY, vbUnicode))
'                                Else
'                                    KOUSEI(Row, ColG_SPTAN) = "0"
'                                End If
'                            Case BtErrKeyNotFound
'                                KOUSEI(Row, ColG_SPTAN) = "0"
'                            Case Else
'                                Call File_Error(sts, BtOpGetEqual, "�i���J�e�S���}�X�^")
'                                Exit Function
'                        End Select
'                        KOUSEI(Row, ColG_ST_URIKIN) = Val(KOUSEI(Row, ColG_SPTAN))
'                    End If
'                End If
'            Next i
'            '�H����
'            For i = 0 To UBound(SP_KOURYO_T)
'                If SP_KOURYO_T(i) = Trim(Right(KOUSEI(Row, ColKO_SYUBETSU), 2)) Then
'                    If IsNumeric(StrConv(ITEMREC.G_SPTAN, vbUnicode)) Then
'                        KOUSEI(Row, ColG_ST_URIKIN) = Val(StrConv(ITEMREC.G_SPTAN, vbUnicode))
'                    Else
'                        Call UniCode_Conv(K0_ITEM_CATEGORY.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
'                        Call UniCode_Conv(K0_ITEM_CATEGORY.CATEGORY_CODE, Text1(ptxCATEGORY_CODE).Text)
'                        sts = BTRV(BtOpGetEqual, ITEM_CATEGORY_POS, ITEM_CATEGORYREC, Len(ITEM_CATEGORYREC), K0_ITEM_CATEGORY, Len(K0_ITEM_CATEGORY), 0)
'                        Select Case sts
'                            Case BtNoErr
'                                If IsNumeric(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_KOURYO, vbUnicode)) Then
'                                    KOUSEI(Row, ColG_SPTAN) = Val(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_KOURYO, vbUnicode))
'                                Else
'                                    KOUSEI(Row, ColG_SPTAN) = "0"
'                                End If
'                            Case BtErrKeyNotFound
'                                KOUSEI(Row, ColG_SPTAN) = "0"
'                            Case Else
'                                Call File_Error(sts, BtOpGetEqual, "�i���J�e�S���}�X�^")
'                                Exit Function
'                        End Select
'                        KOUSEI(Row, ColG_ST_URIKIN) = Val(KOUSEI(Row, ColG_SPTAN))
'                    End If
'                End If
'            Next i
'            '���し
'            For i = 0 To UBound(SP_HAKO_T)
'                If SP_HAKO_T(i) = Trim(Right(KOUSEI(Row, ColKO_SYUBETSU), 2)) Then
'                    If IsNumeric(StrConv(ITEMREC.G_SPTAN, vbUnicode)) Then
'                        KOUSEI(Row, ColG_ST_URIKIN) = Val(StrConv(ITEMREC.G_SPTAN, vbUnicode))
'                    Else
'                        Call UniCode_Conv(K0_ITEM_CATEGORY.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
'                        Call UniCode_Conv(K0_ITEM_CATEGORY.CATEGORY_CODE, Text1(ptxCATEGORY_CODE).Text)
'                        sts = BTRV(BtOpGetEqual, ITEM_CATEGORY_POS, ITEM_CATEGORYREC, Len(ITEM_CATEGORYREC), K0_ITEM_CATEGORY, Len(K0_ITEM_CATEGORY), 0)
'                        Select Case sts
'                            Case BtNoErr
'                                If IsNumeric(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_HAKO, vbUnicode)) Then
'                                    KOUSEI(Row, ColG_SPTAN) = Val(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_HAKO, vbUnicode))
'                                Else
'                                    KOUSEI(Row, ColG_SPTAN) = "0"
'                                End If
'                            Case BtErrKeyNotFound
'                                KOUSEI(Row, ColG_SPTAN) = "0"
'                            Case Else
'                                Call File_Error(sts, BtOpGetEqual, "�i���J�e�S���}�X�^")
'                                Exit Function
'                        End Select
'                        KOUSEI(Row, ColG_ST_URIKIN) = Val(KOUSEI(Row, ColG_SPTAN))
'                    End If
'                End If
'            Next i
'        Next Row
'
'
'        Set TDBGrid1(pGrdKOUSEI).Array = KOUSEI
'
'
'        TDBGrid1(pGrdKOUSEI).Bookmark = Null
'
'        TDBGrid1(pGrdKOUSEI).ReBind
'        TDBGrid1(pGrdKOUSEI).Update
'        TDBGrid1(pGrdKOUSEI).ScrollBars = dbgAutomatic
'
'        If KOUSEI.Count(1) > 0 Then
'            TDBGrid1(pGrdKOUSEI).MoveFirst
'        End If
'
'    End If
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   ���ʒP��
    
    
    If IsNumeric(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_KOURYO, vbUnicode)) Then
        Text1(ptxCATE_AD_KOURYO).Text = Format(Val(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_KOURYO, vbUnicode)), "0.00")
    Else
        If IsNumeric(Text1(ptxCATE_AD_KOURYO).Text) Then
            Text1(ptxAFT_S_KOUSU_BAIKA).Text = Format(CDbl(Text1(ptxCATE_AD_KOURYO).Text), "#0.00")
        Else
            Text1(ptxAFT_S_KOUSU_BAIKA).Text = ""
        End If
    End If
        
        
        
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  ������ʒP��    2012.01.04
    
    If IsNumeric(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_HAKO, vbUnicode)) Then
        Text1(ptxAFT_S_SHIZAI_BAIKA).Text = Format(Val(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_HAKO, vbUnicode)), "0.00")
    Else
        If IsNumeric(Text1(ptxAFT_S_SHIZAI_BAIKA).Text) Then
            Text1(ptxAFT_S_KOUSU_BAIKA).Text = Format(CDbl(Text1(ptxCATE_AD_KOURYO).Text), "#0.00")
        Else
            Text1(ptxAFT_S_KOUSU_BAIKA).Text = ""
        End If
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  ������ʒP��    2012.01.04
        
    
    
    CATEGORY_Disp_Proc = False
End Function

Private Sub CATEGORY_KEISAN_PROC()
'----------------------------------------------------------------------------
'                   �i���J�e�S�����̍Čv�Z
'----------------------------------------------------------------------------
Dim sts As Integer
    
    '��Ǝ��Ԍv
    Text1(ptxCATE_ST_TOTAL).Text = Val(Text1(ptxCATE_ST_KOUTEI).Text) + _
                                Val(Text1(ptxCATE_ST_FUKA).Text) + _
                                Val(Text1(ptxCATE_ST_JITU2).Text)
    Text1(ptxCATE_AD_TOTAL).Text = Val(Text1(ptxCATE_AD_KOUTEI).Text) + _
                                Val(Text1(ptxCATE_AD_FUKA).Text) + _
                                Val(Text1(ptxCATE_AD_JITU2).Text)


    '��/��
    Text1(ptxCATE_ST_FUN).Text = Format(ToHalfAdjust(CCur(Val(Text1(ptxCATE_ST_TOTAL)) / 60), 2), "#0.00")
    Text1(ptxCATE_AD_FUN).Text = Format(ToHalfAdjust(CCur(Val(Text1(ptxCATE_AD_TOTAL)) / 60), 2), "#0.00")
    
    
    
    '�����[�g (�~ / ��)
    If IsNumeric(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)) Then
        Text1(ptxCATE_ST_FUN_RATE).Text = Format(Val(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)), "#0.00")
        Text1(ptxCATE_AD_FUN_RATE).Text = Format(Val(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)), "#0.00")
    Else

        Text1(ptxCATE_ST_FUN_RATE).Text = ""
        Text1(ptxCATE_AD_FUN_RATE).Text = ""
    End If
    '(�~�^��)
    If IsNumeric(Text1(ptxCATE_ST_FUN_RATE)) Then
        Text1(ptxCATE_ST_KOURYO).Text = Format(ToRoundUp(CCur(Val(Text1(ptxCATE_ST_FUN).Text) * Val(Text1(ptxCATE_ST_FUN_RATE).Text)), 2), "#0.00")
        Text1(ptxCATE_AD_KOURYO).Text = Format(ToRoundUp(CCur(Val(Text1(ptxCATE_AD_FUN).Text) * Val(Text1(ptxCATE_AD_FUN_RATE).Text)), 2), "#0.00")

    Else
        Text1(ptxCATE_ST_KOURYO).Text = "0.00"
        Text1(ptxCATE_AD_KOURYO).Text = "0.00"
    End If


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �H�����ʒP��
    
    Call UniCode_Conv(K0_ITEM_CATEGORY.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    
    Call UniCode_Conv(K0_ITEM_CATEGORY.CATEGORY_CODE, Text1(ptxCATEGORY_CODE).Text)
        
    sts = BTRV(BtOpGetEqual, ITEM_CATEGORY_POS, ITEM_CATEGORYREC, Len(ITEM_CATEGORYREC), K0_ITEM_CATEGORY, Len(K0_ITEM_CATEGORY), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound

            Call UniCode_Conv(ITEM_CATEGORYREC.TOKU_TANKA_KOURYO, "")
            Call UniCode_Conv(ITEM_CATEGORYREC.TOKU_TANKA_HAKO, "")

        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i���J�e�S���}�X�^")
            Unload Me

    End Select
    
    
    
    
    If IsNumeric(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_KOURYO, vbUnicode)) Then
        Text1(ptxCATE_AD_KOURYO).Text = Format(Val(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_KOURYO, vbUnicode)), "0.00")
    Else
        If IsNumeric(Text1(ptxCATE_AD_KOURYO).Text) Then
            Text1(ptxAFT_S_KOUSU_BAIKA).Text = Format(CDbl(Text1(ptxCATE_AD_KOURYO).Text), "#0.00")
        Else
            Text1(ptxAFT_S_KOUSU_BAIKA).Text = ""
        End If
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �H�����ʒP��



'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  ������ʒP��    2012.01.04
    If IsNumeric(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_HAKO, vbUnicode)) Then
        Text1(ptxAFT_S_SHIZAI_BAIKA).Text = Format(Val(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_HAKO, vbUnicode)), "0.00")
    Else
        If IsNumeric(Text1(ptxAFT_S_SHIZAI_BAIKA).Text) Then
            Text1(ptxAFT_S_KOUSU_BAIKA).Text = Format(CDbl(Text1(ptxCATE_AD_KOURYO).Text), "#0.00")
        Else
            Text1(ptxAFT_S_KOUSU_BAIKA).Text = ""
        End If
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  ������ʒP��    2012.01.04




End Sub

Private Function Main_Update_Proc() As Integer

Dim sts     As Integer
Dim com     As Integer
Dim Fsw     As Integer

Dim i       As Integer
Dim Errflg  As Integer



Dim ON_SW   As Boolean      '2018.05.16
Dim j       As Integer      '2018.05.16



    Main_Update_Proc = True
    
    
    
    
    Call UniCode_Conv(wK2_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
    Call UniCode_Conv(wK2_P_COMPO.KO_JGYOBU, SHIZAI)
    Call UniCode_Conv(wK2_P_COMPO.KO_NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(wK2_P_COMPO.KO_HIN_GAI, MAIN_HIN_GAI)
       
    Call UniCode_Conv(wK2_P_COMPO.DATA_KBN, P_HEAD)
    Call UniCode_Conv(wK2_P_COMPO.SEQNO, "000")
       
    Fsw = 0
    com = BtOpGetGreater
    
    Do
        DoEvents
        
        sts = BTRV(com, wP_COMPO_POS, wP_COMPO_K_REC, Len(wP_COMPO_K_REC), wK2_P_COMPO, Len(wK2_P_COMPO), 2)
        Select Case sts
            Case BtNoErr
            
                If StrConv(wP_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Or _
                    StrConv(wP_COMPO_K_REC.KO_JGYOBU, vbUnicode) <> SHIZAI Or _
                    StrConv(wP_COMPO_K_REC.KO_NAIGAI, vbUnicode) <> NAIGAI_NAI Or _
                    StrConv(wP_COMPO_K_REC.KO_HIN_GAI, vbUnicode) <> MAIN_HIN_GAI Then
                    
                    
                    If Fsw = 0 Then
                        
                        List2.AddItem MAIN_HIN_GAI & " " & Space(20) & "NG"
Call LOG_OUT(SEI0018_LOG, "[NG]" & "KO_HIN_GAI_= " & MAIN_HIN_GAI & "HIN_GAI_= " & "�Y���Ȃ�" & " " & Now)
                        
                        NG_cnt = NG_cnt + 1
                        txtNG_CNT.Text = Format(NG_cnt, "#,##0")
                        DoEvents
                    
                    End If

                    Exit Do
                
                End If

                If StrConv(wP_COMPO_K_REC.DATA_KBN, vbUnicode) = "0" Or _
                    StrConv(wP_COMPO_K_REC.SEQNO, vbUnicode) = "000" Then
                Else

'>>>>>>>>>>>>>  2018.05.16
                    ON_SW = False

                    For i = 0 To UBound(DATA_KBN_TBL)
                    
                        If StrConv(wP_COMPO_K_REC.DATA_KBN, vbUnicode) = DATA_KBN_TBL(i) Then
                    
                            ON_SW = True
                            Exit For
                        
                        End If
                    
                    Next i
                    
                    If ON_SW And StrConv(wP_COMPO_K_REC.DATA_KBN, vbUnicode) = "3" Then

                        ON_SW = False
                        
                        For i = 0 To UBound(DATA_KBN_TBL)
                        
                            If StrConv(wP_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = SYUBETSU_TBL(i) Then
                        
                                ON_SW = True
                                Exit For
                            
                            End If
                        
                        Next i
                    
                    End If

                    If ON_SW Then
'>>>>>>>>>>>>>  2018.05.16

                        Fsw = 1
                        
                        Text1(ptxTanto_Code).Text = txtTANTO_CODE.Text
                        Text1(ptxHin_Gai).Text = StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode)
                        
                        
                        Combo1(pcmbSHIMUKE).ListIndex = cmbSHIMUKE.ListIndex
    
                        '�X�e�[�^�X�E�B���h�E���쐬����
                        hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                            "[�����V�X�e��]���Ϗ��ꊇ�쐬�����@�q�i��= " & MAIN_HIN_GAI & " �e�i��= " & StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode), Me.hwnd, 0)
    
    
    
                        Errflg = False
                        If Detail_Disp_Proc(Errflg) Then
    Call LOG_OUT(SEI0018_LOG, "[NG]" & "KO_HIN_GAI_= " & MAIN_HIN_GAI & "HIN_GAI_= " & StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode) & " " & Now)
                            Exit Function
                        End If
                            
                        
                        Errflg = False
                        For i = ptxTanto_Code To ptxAFT_S_BU_KAKO_KOSU
                        
                            If Error_Check_Proc(i) Then
                                Errflg = True
                                Exit For
                            End If
                        
                        
                        Next i
                        
                            
                            
                        If Not Errflg Then
                        
'>>>>>>>    2018.06.05 ���̂ݒǉ�
                            If Text1(ptxBEF_SEI_LOT).Text = "" Then
                                                                                    
                                                                                    
                                If HAKO_TANKA_KEISAN_Proc() Then
        Call LOG_OUT(SEI0018_LOG, "[NG]" & "KO_HIN_GAI_= " & MAIN_HIN_GAI & "HIN_GAI_= " & StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode) & " " & Now)
        Call LOG_OUT(SEI0018_LOG, "���Ϗ��ꊇ�쐬�@�ُ�I��[" & Now & "]")
                                        
                                    Exit Function
                                End If
                                                                                    
                                                                                    
                            Else
                                                                                        
                                                                                        
                                                                                        
                                                                                        
                                
                                If TANKA_KEISAN_Proc() Then
        Call LOG_OUT(SEI0018_LOG, "[NG]" & "KO_HIN_GAI_= " & MAIN_HIN_GAI & "HIN_GAI_= " & StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode) & " " & Now)
        Call LOG_OUT(SEI0018_LOG, "���Ϗ��ꊇ�쐬�@�ُ�I��[" & Now & "]")
                                        
                                    Exit Function
                                End If
                    
                    
                            End If
                    
'>>>>>>>    2018.06.05 ���̂ݒǉ�
                    
                    
                    
                    
'>>>>>>>    2018.06.05 ���̂ݒǉ�
                            If Text1(ptxBEF_SEI_LOT).Text = "" Then
                    
                    
                                If HAKO_Tanka_Update_Proc() Then
        Call LOG_OUT(SEI0018_LOG, "[NG]" & "KO_HIN_GAI_= " & MAIN_HIN_GAI & "HIN_GAI_= " & StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode) & " " & Now)
        Call LOG_OUT(SEI0018_LOG, "���Ϗ��ꊇ�쐬�@�ُ�I��[" & Now & "]")
                                            
                                    Exit Function
                                End If
                        
                            Else
                        
                                If Tanka_Update_Proc() Then
        Call LOG_OUT(SEI0018_LOG, "[NG]" & "KO_HIN_GAI_= " & MAIN_HIN_GAI & "HIN_GAI_= " & StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode) & " " & Now)
        Call LOG_OUT(SEI0018_LOG, "���Ϗ��ꊇ�쐬�@�ُ�I��[" & Now & "]")
                                            
                                    Exit Function
                                End If
                            End If
'>>>>>>>    2018.06.05 ���̂ݒǉ�
                        
                        
                        
                            If Detail_Disp_Proc(Errflg) Then
    Call LOG_OUT(SEI0018_LOG, "���Ϗ��ꊇ�쐬�@�ُ�I��[" & Now & "]")
                                Unload Me
                            End If
                            
                            If Estimate_Proc() Then
    Call LOG_OUT(SEI0018_LOG, "[NG]" & "KO_HIN_GAI_= " & MAIN_HIN_GAI & "HIN_GAI_= " & StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode) & " " & Now)
    Call LOG_OUT(SEI0018_LOG, "���Ϗ��ꊇ�쐬�@�ُ�I��[" & Now & "]")
                                Exit Function
                            End If
                    
                            OK_cnt = OK_cnt + 1
                            txtOK_CNT.Text = Format(OK_cnt, "#,##0")
                            DoEvents
                        
                            List2.AddItem MAIN_HIN_GAI & " " & StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode) & "OK"
    Call LOG_OUT(SEI0018_LOG, "[OK]" & "KO_HIN_GAI_= " & MAIN_HIN_GAI & "HIN_GAI_= " & StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode) & " " & Now)
                        
                        Else
    Call LOG_OUT(SEI0018_LOG, "[NG]" & "KO_HIN_GAI_= " & MAIN_HIN_GAI & "HIN_GAI_= " & StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode) & " " & Now)
                            
                            List2.AddItem MAIN_HIN_GAI & " " & StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode) & "NG"
                            NG_cnt = NG_cnt + 1
                            txtNG_CNT.Text = Format(NG_cnt, "#,##0")
                        End If
                
                '>>>>>>>>>>>>>  2018.05.16
                    End If
                '>>>>>>>>>>>>>  2018.05.16
                
                End If
            
            Case BtErrEOF
                If Fsw = 0 Then
                    
Call LOG_OUT(SEI0018_LOG, "[NG]" & "KO_HIN_GAI_= " & MAIN_HIN_GAI & "HIN_GAI_= " & "�Y���Ȃ�" & " " & Now)
                    
                    NG_cnt = NG_cnt + 1
                    List2.AddItem MAIN_HIN_GAI & " " & Space(20) & "NG"
                    txtNG_CNT.Text = Format(NG_cnt, "#,##0")
                    
                    
                    DoEvents
                End If
            
            
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "�\���}�X�^")
Call LOG_OUT(SEI0018_LOG, "���Ϗ��ꊇ�쐬�@�ُ�I��[" & Now & "]")
                Exit Function
                
    
        End Select
    
    
    
        com = BtOpGetNext
    
    Loop
    
    
    
    
    Main_Update_Proc = False
    


End Function




Private Function Main_Update_OYA_Proc() As Integer

Dim sts     As Integer
Dim com     As Integer
Dim Fsw     As Integer

Dim i       As Integer
Dim Errflg  As Integer


    Main_Update_OYA_Proc = True
    
                    
    Text1(ptxTanto_Code).Text = txtTANTO_CODE.Text
    Text1(ptxHin_Gai).Text = MAIN_HIN_GAI
    
    Combo1(pcmbSHIMUKE).ListIndex = cmbSHIMUKE.ListIndex
    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[�����V�X�e��]���Ϗ��ꊇ�쐬�����@�e�i��= " & MAIN_HIN_GAI, Me.hwnd, 0)



    If Detail_Disp_Proc(Errflg) Then
        Call LOG_OUT(SEI0018_LOG, "[NG]" & "HIN_GAI_= " & MAIN_HIN_GAI & " " & Now)
        Exit Function
    End If
                        
                    
    Errflg = False
    For i = ptxTanto_Code To ptxAFT_S_BU_KAKO_KOSU
            
        If Error_Check_Proc(i) Then
            Errflg = True
            Exit For
        End If
            
            
    Next i
                    
                        
                        
    If Not Errflg Then
                    
'>>>>>>>    2018.06.05 ���̂ݒǉ�
        If Text1(ptxBEF_SEI_LOT).Text = "" Then
                                                                                    
                                                                                    
            If HAKO_TANKA_KEISAN_Proc() Then
                Exit Function
            End If
        Else
            If TANKA_KEISAN_Proc() Then
                    
                Exit Function
            End If


        End If
                    
'>>>>>>>    2018.06.05 ���̂ݒǉ�
                
'>>>>>>>    2018.06.05 ���̂ݒǉ�
            If Text1(ptxBEF_SEI_LOT).Text = "" Then
    
    
                If HAKO_Tanka_Update_Proc() Then
                            
                    Exit Function
                End If
        
            Else
                        
                If Tanka_Update_Proc() Then
                    Exit Function
                End If
            End If
'>>>>>>>    2018.06.05 ���̂ݒǉ�
                
                
                
                    
        If Detail_Disp_Proc(Errflg) Then
            Unload Me
        End If
                        
        If Estimate_Proc() Then
            Call LOG_OUT(SEI0018_LOG, "[NG]" & "HIN_GAI_= " & MAIN_HIN_GAI & " " & Now)
            Exit Function
        End If
                
        OK_cnt = OK_cnt + 1
        txtOK_CNT.Text = Format(OK_cnt, "#,##0")
        DoEvents
                    
        List3.AddItem MAIN_HIN_GAI & "OK"
        Call LOG_OUT(SEI0018_LOG, "[OK]" & "HIN_GAI_= " & MAIN_HIN_GAI & " " & Now)
                    
    Else
        Call LOG_OUT(SEI0018_LOG, "[NG]" & "HIN_GAI_= " & MAIN_HIN_GAI & " " & Now)
                        
        List3.AddItem MAIN_HIN_GAI & "NG"
        NG_cnt = NG_cnt + 1
        txtNG_CNT.Text = Format(NG_cnt, "#,##0")
    End If
                
    
    
    
    
    
    
    Main_Update_OYA_Proc = False
    


End Function


Private Function COUNT_Proc() As Integer

Dim sts     As Integer
Dim com     As Integer
Dim Fsw     As Integer

Dim i       As Integer
Dim Errflg  As Integer


Dim ON_SW   As Boolean      '2018.05.16
Dim j       As Integer      '2018.05.16


    COUNT_Proc = True
    
    Call UniCode_Conv(wK2_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
    Call UniCode_Conv(wK2_P_COMPO.KO_JGYOBU, SHIZAI)
    Call UniCode_Conv(wK2_P_COMPO.KO_NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(wK2_P_COMPO.KO_HIN_GAI, MAIN_HIN_GAI)
       
    Call UniCode_Conv(wK2_P_COMPO.DATA_KBN, P_HEAD)
    Call UniCode_Conv(wK2_P_COMPO.SEQNO, "000")
       
    Fsw = 0
    com = BtOpGetGreater
       
    
    Do
        DoEvents
        
        sts = BTRV(com, wP_COMPO_POS, wP_COMPO_K_REC, Len(wP_COMPO_K_REC), wK2_P_COMPO, Len(wK2_P_COMPO), 2)
        Select Case sts
            Case BtNoErr
            
                If StrConv(wP_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Or _
                    StrConv(wP_COMPO_K_REC.KO_JGYOBU, vbUnicode) <> SHIZAI Or _
                    StrConv(wP_COMPO_K_REC.KO_NAIGAI, vbUnicode) <> NAIGAI_NAI Or _
                    StrConv(wP_COMPO_K_REC.KO_HIN_GAI, vbUnicode) <> MAIN_HIN_GAI Then
                    

                    Exit Do
                
                End If

                If StrConv(wP_COMPO_K_REC.DATA_KBN, vbUnicode) = "0" Or _
                    StrConv(wP_COMPO_K_REC.SEQNO, vbUnicode) = "000" Then
                Else



'>>>>>>>>>>>>>  2018.05.16
                    ON_SW = False

                    For i = 0 To UBound(DATA_KBN_TBL)
                    
                        If StrConv(wP_COMPO_K_REC.DATA_KBN, vbUnicode) = DATA_KBN_TBL(i) Then
                    
                            ON_SW = True
                            Exit For
                        
                        End If
                    
                    Next i
                    
                    If ON_SW And StrConv(wP_COMPO_K_REC.DATA_KBN, vbUnicode) = "3" Then

                        ON_SW = False
                        
                        For i = 0 To UBound(DATA_KBN_TBL)
                        
                            If StrConv(wP_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = SYUBETSU_TBL(i) Then
                        
                                ON_SW = True
                                Exit For
                            
                            End If
                        
                        Next i
                    
                    End If


                    If ON_SW Then

                        
                        
                        Text1(ptxTanto_Code).Text = txtTANTO_CODE.Text
                        Text1(ptxHin_Gai).Text = StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode)
                        
                        
                        Combo1(pcmbSHIMUKE).ListIndex = cmbSHIMUKE.ListIndex
                        
                        
                        Errflg = False
                        If Detail_Disp_Proc(Errflg) Then
                            Exit Function
                        End If
                        
                        If Errflg Then
                        
                            ON_SW = False
                        End If
                        
                        
                        Errflg = False
                        For i = ptxTanto_Code To ptxAFT_S_BU_KAKO_KOSU
                        
                            If Error_Check_Proc(i) Then
                                Errflg = True
                                Exit For
                            End If
                        
                        
                        Next i


                        If Errflg Then
                        
                            ON_SW = False
                        End If

                    End If


'                    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(wP_COMPO_K_REC.JGYOBU, vbUnicode))
'                    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(wP_COMPO_K_REC.NAIGAI, vbUnicode))
'                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode))
'
'                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                    Select Case sts
'                        Case BtNoErr
'
'
'                        Case BtErrKeyNotFound
'
'                            ON_SW = False
'
'                        Case Else
'                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
'                            Exit Function
'
'
'                    End Select

                    If ON_SW Then
'>>>>>>>>>>>>>  2018.05.16








                        List2.AddItem MAIN_HIN_GAI & " " & StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode)
    
    
                        IN_cnt = IN_cnt + 1
                        txtOUT_CNT.Text = Format(IN_cnt, "#,##0")
                
'>>>>>>>>>>>>>  2018.05.16
                    End If
'>>>>>>>>>>>>>  2018.05.16
                                    
                End If
            
            Case BtErrEOF
            
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "�\���}�X�^")
                Exit Function
                
    
        End Select
    
    
        com = BtOpGetNext
    
    Loop
    
    
    
    
    COUNT_Proc = False

End Function



Private Function HAKO_TANKA_KEISAN_Proc() As Integer
'----------------------------------------------------------------------------
'                   �P���v�Z����(����̂�)
'               2018.06.05
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim j           As Integer
Dim wkint       As Integer
Dim wkDouble    As Double


Dim c           As String * 128
Dim wkKUSATU    As Variant
Dim INV_F       As Boolean


    HAKO_TANKA_KEISAN_Proc = True
    
    '�i�ڃ}�X�^�ǂݍ���
    Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHin_Gai).Text)

    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound

            Text1(ptxHin_Name).Text = ""
            Text1(ptxST_SOKO).Text = ""
            Text1(ptxST_RETU).Text = ""
            Text1(ptxST_REN).Text = ""
            Text1(ptxST_DAN).Text = ""
            MsgBox "���͂������ڂ̓G���[�ł��B(�i��)"
            HAKO_TANKA_KEISAN_Proc = False
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Function

    End Select


    '�ݒ��
    Text1(ptxAFT_S_KOUSU_SET_DATE).Text = Format(Now, "YYYYMMDD")
    '�S����
    Text1(ptxAFT_SEI_TANKA_TANTO).Text = Text1(ptxTanto_Code).Text
    
    
    '-----------------------------------    �ύX��
'    '����
    wkDouble = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
    
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    If IsNumeric(KOUSEI(i, ColG_ST_SHIKIN)) Then
                        wkDouble = wkDouble + CDbl(KOUSEI(i, ColG_ST_SHIKIN))
                    End If
                    Exit For
                End If
    
            Next j
    
        Next i
    End If
    Text1(ptxAFT_S_SHIZAI_GENKA).Text = Format(wkDouble, "#0.00")
'
    wkDouble = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
    
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    
                    If IsNumeric(KOUSEI(i, ColG_ST_URIKIN)) Then
                        wkDouble = wkDouble + CDbl(KOUSEI(i, ColG_ST_URIKIN))
                    End If
                    Exit For
                End If
    
            Next j
    
        Next i
    End If
    Text1(ptxAFT_S_SHIZAI_BAIKA).Text = Format(wkDouble, "#0.00")


    '�O������
    wkDouble = 0
    If KUSATU_F Then
        If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
        Else
            For i = 1 To KOUSEI.UpperBound(1)
        
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = GAISO_KBN Then
            
            
'                    If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
                    If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
            
                        wkDouble = wkDouble + CDbl(KOUSEI(i, ColG_ST_URIKIN_KUSATU))
                    End If
            
                End If
        
        
            Next i
        End If
    End If
    Text1(ptxAFT_S_GAISO_TANKA).Text = Format(wkDouble, "#0.00")







    HAKO_TANKA_KEISAN_Proc = False

End Function

Private Function HAKO_Tanka_Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   �P���o�^����(����̂�)
'               2018.06.05
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim ans         As Integer

Dim wkGAISO     As Double
    
Dim i           As Integer
Dim j            As Integer
    
    
Dim wkint       As Integer
    
    HAKO_Tanka_Update_Proc = True

    '�i�ڃ}�X�^�ǂݍ���
    Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHin_Gai).Text)


    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
            
                MsgBox "���[���Ńf�[�^���A�ύX����Ă��܂��B�P���o�^�����𒆎~���܂��B"
                HAKO_Tanka_Update_Proc = False
                Exit Function
            
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    HAKO_Tanka_Update_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
                Exit Function
        
        End Select
    
    Loop


    '�V�P���|�|�����P�� 2009.06.02
    Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode))
    Call UniCode_Conv(ITEMREC.BEF_S_GAISO_TANKA, StrConv(ITEMREC.S_GAISO_TANKA, vbUnicode))



    '�ݒ��
    Call UniCode_Conv(ITEMREC.S_KOUSU_SET_DATE, Format(Now, "YYYYMMDD"))
    
    
    '���㌴��
    Call UniCode_Conv(ITEMREC.S_SHIZAI_GENKA, Format(CDbl(Text1(ptxAFT_S_SHIZAI_GENKA).Text), "00000000.00"))
    '���㔄��
    Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, Format(CDbl(Text1(ptxAFT_S_SHIZAI_BAIKA).Text), "00000000.00"))
    
    
    
    '�O������
    If IsNumeric(Text1(ptxAFT_S_GAISO_TANKA).Text) Then
        Call UniCode_Conv(ITEMREC.S_GAISO_TANKA, Format(CDbl(Text1(ptxAFT_S_GAISO_TANKA).Text), "00000.00"))
    Else
        Call UniCode_Conv(ITEMREC.S_GAISO_TANKA, "00000.00")
    End If
    
    
    '�ݒ��
    Call UniCode_Conv(ITEMREC.S_SHIZAI_SET_DATE, Format(Now, "YYYYMMDD"))
    '�S����
    Call UniCode_Conv(ITEMREC.SEI_TANKA_TANTO, Text1(ptxTanto_Code).Text)
    '����
    Call UniCode_Conv(ITEMREC.SE_TANKA_MEMO, "")
    
    '�X�V�S����
    Call UniCode_Conv(ITEMREC.UPD_TANTO, Text1(ptxTanto_Code).Text)
    '�X�V ����
    Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
    
    
    
    Do
        sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    HAKO_Tanka_Update_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpUpdate, "�i�ڃ}�X�^")
                Exit Function
        
        End Select
    
    Loop
    
    
    '�P���X�V�����o��
    Do
        sts = BTRV(BtOpInsert, ITEM_HST_POS, ITEMREC, Len(ITEMREC), K0_ITEM_HST, Len(K0_ITEM_HST), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM_HST.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    HAKO_Tanka_Update_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpUpdate, "�i�ڒP���X�V����")
                Exit Function
        
        End Select
    
    Loop
    

    HAKO_Tanka_Update_Proc = False


End Function

