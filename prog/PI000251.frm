VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form PI000251 
   Caption         =   "���i�����у����e�i���X "
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16125
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
   ScaleHeight     =   11010
   ScaleWidth      =   16125
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox txtLOAD_LIMIT 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Left            =   12720
      MaxLength       =   5
      TabIndex        =   111
      Top             =   10320
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   13
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   12
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      Height          =   375
      Index           =   50
      Left            =   3120
      TabIndex        =   13
      Top             =   5160
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   360
      IMEMode         =   3  '�̌Œ�
      Index           =   49
      Left            =   14970
      Locked          =   -1  'True
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   6000
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  '�ׯ�
      Height          =   360
      Index           =   2
      Left            =   1680
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   106
      Top             =   5640
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '�Ȃ�
      Height          =   315
      Index           =   5
      Left            =   3240
      TabIndex        =   104
      TabStop         =   0   'False
      Top             =   5280
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '�Ȃ�
      Height          =   315
      Index           =   4
      Left            =   3240
      TabIndex        =   103
      TabStop         =   0   'False
      Top             =   4800
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '�Ȃ�
      Height          =   315
      Index           =   3
      Left            =   3240
      TabIndex        =   102
      TabStop         =   0   'False
      Top             =   4320
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '�Ȃ�
      Height          =   315
      Index           =   2
      Left            =   4440
      TabIndex        =   101
      TabStop         =   0   'False
      Top             =   3780
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   2955
      MaxLength       =   3
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   360
      IMEMode         =   3  '�̌Œ�
      Index           =   42
      Left            =   14970
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   41
      Left            =   14205
      MaxLength       =   6
      TabIndex        =   41
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   40
      Left            =   13440
      MaxLength       =   4
      TabIndex        =   40
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   360
      IMEMode         =   3  '�̌Œ�
      Index           =   39
      Left            =   14970
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   38
      Left            =   14205
      MaxLength       =   6
      TabIndex        =   38
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   37
      Left            =   13440
      MaxLength       =   4
      TabIndex        =   37
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   360
      IMEMode         =   3  '�̌Œ�
      Index           =   36
      Left            =   14970
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   35
      Left            =   14205
      MaxLength       =   6
      TabIndex        =   35
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   34
      Left            =   13440
      MaxLength       =   4
      TabIndex        =   34
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   360
      IMEMode         =   3  '�̌Œ�
      Index           =   48
      Left            =   14970
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   47
      Left            =   14205
      MaxLength       =   6
      TabIndex        =   49
      Top             =   5640
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   46
      Left            =   13440
      MaxLength       =   4
      TabIndex        =   48
      Top             =   5640
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   4
      Left            =   8940
      TabIndex        =   47
      Top             =   5640
      Width           =   4590
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   360
      IMEMode         =   3  '�̌Œ�
      Index           =   45
      Left            =   14970
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   44
      Left            =   14205
      MaxLength       =   6
      TabIndex        =   45
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   43
      Left            =   13440
      MaxLength       =   4
      TabIndex        =   44
      Top             =   5280
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   3
      Left            =   8940
      TabIndex        =   43
      Top             =   5280
      Width           =   4590
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   360
      IMEMode         =   3  '�̌Œ�
      Index           =   33
      Left            =   14970
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   32
      Left            =   14205
      MaxLength       =   6
      TabIndex        =   32
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   31
      Left            =   13440
      MaxLength       =   4
      TabIndex        =   31
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   360
      IMEMode         =   3  '�̌Œ�
      Index           =   30
      Left            =   14970
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   29
      Left            =   14205
      MaxLength       =   6
      TabIndex        =   29
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   28
      Left            =   13440
      MaxLength       =   4
      TabIndex        =   28
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   360
      IMEMode         =   3  '�̌Œ�
      Index           =   27
      Left            =   14970
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   26
      Left            =   14205
      MaxLength       =   6
      TabIndex        =   26
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   25
      Left            =   13440
      MaxLength       =   4
      TabIndex        =   25
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   360
      IMEMode         =   3  '�̌Œ�
      Index           =   24
      Left            =   14970
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   23
      Left            =   14205
      MaxLength       =   6
      TabIndex        =   23
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   22
      Left            =   13440
      MaxLength       =   4
      TabIndex        =   22
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   360
      IMEMode         =   3  '�̌Œ�
      Index           =   21
      Left            =   14970
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   20
      Left            =   14205
      MaxLength       =   6
      TabIndex        =   20
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   19
      Left            =   13440
      MaxLength       =   4
      TabIndex        =   19
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   18
      Left            =   14970
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   17
      Left            =   14205
      MaxLength       =   6
      TabIndex        =   17
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   16
      Left            =   13440
      MaxLength       =   4
      TabIndex        =   16
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   48
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      IMEMode         =   3  '�̌Œ�
      Index           =   15
      Left            =   11280
      MaxLength       =   8
      TabIndex        =   15
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   48
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      IMEMode         =   3  '�̌Œ�
      Index           =   14
      Left            =   8280
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   12
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   11
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   11
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   10
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   9
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3720
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  '�ׯ�
      Height          =   360
      Index           =   1
      Left            =   2280
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   8
      Top             =   3240
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   8
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   7
      Top             =   3240
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   360
      Index           =   0
      Left            =   1560
      Locked          =   -1  'True
      Style           =   1  '�W������
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "Combo1"
      Top             =   2760
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      Height          =   360
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   4
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   1560
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
      TabIndex        =   63
      Top             =   10320
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
      Left            =   9600
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   10320
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
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   10320
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
      Left            =   7920
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   10320
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
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   10320
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
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   10320
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
      Index           =   5
      Left            =   4920
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   10320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�� �V"
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
      Left            =   4080
      TabIndex        =   56
      Top             =   10320
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
      Left            =   2760
      TabIndex        =   55
      Top             =   10320
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
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   10320
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
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   10320
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
      TabIndex        =   53
      Top             =   10320
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '�Ȃ�
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   1860
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '�Ȃ�
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   100
      TabStop         =   0   'False
      Top             =   2340
      Width           =   2175
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   3615
      Left            =   630
      TabIndex        =   105
      Top             =   6600
      Width           =   14550
      _ExtentX        =   25665
      _ExtentY        =   6376
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "���s����"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "�w�}�[��"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "�d������"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "��z��"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "�i��"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "�i��"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "�\��"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "����"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "���P/�S����"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   9
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=9"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2090"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1984"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=512"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2011"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1905"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=512"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2778"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2672"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=512"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=3281"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=3175"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=512"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=2778"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2672"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=512"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=3493"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=3387"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=512"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=1826"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=1720"
      Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=514"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(36)=   "Column(7).Width=1826"
      Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=1720"
      Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=514"
      Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(41)=   "Column(8).Width=3810"
      Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=3704"
      Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=516"
      Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=�l�r �S�V�b�N"
      PrintInfos(0).PageFooterFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=�l�r �S�V�b�N"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFF&,.bold=0,.fontsize=1200"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=�l�r �S�V�b�N"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bold=0,.fontsize=1200"
      _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=�l�r �S�V�b�N"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=�l�r �S�V�b�N"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFF00&"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=43,.parent=1,.bold=0,.fontsize=975,.italic=0,.underline=0"
      _StyleDefs(25)  =   ":id=43,.strikethrough=0,.charset=128"
      _StyleDefs(26)  =   ":id=43,.fontname=�l�r �S�V�b�N"
      _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
      _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=45,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=47,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=58,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(39)  =   ":id=58,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(40)  =   ":id=58,.fontname=�l�r �S�V�b�N"
      _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=44"
      _StyleDefs(42)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=45"
      _StyleDefs(43)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=47"
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=62,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(45)  =   ":id=62,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(46)  =   ":id=62,.fontname=�l�r �S�V�b�N"
      _StyleDefs(47)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=44"
      _StyleDefs(48)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=45"
      _StyleDefs(49)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=47"
      _StyleDefs(50)  =   "Splits(0).Columns(2).Style:id=28,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(51)  =   ":id=28,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(52)  =   ":id=28,.fontname=�l�r �S�V�b�N"
      _StyleDefs(53)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=44"
      _StyleDefs(54)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=45"
      _StyleDefs(55)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=47"
      _StyleDefs(56)  =   "Splits(0).Columns(3).Style:id=66,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(57)  =   ":id=66,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(58)  =   ":id=66,.fontname=�l�r �S�V�b�N"
      _StyleDefs(59)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=44"
      _StyleDefs(60)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=45"
      _StyleDefs(61)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=47"
      _StyleDefs(62)  =   "Splits(0).Columns(4).Style:id=32,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(63)  =   ":id=32,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(64)  =   ":id=32,.fontname=�l�r �S�V�b�N"
      _StyleDefs(65)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=44"
      _StyleDefs(66)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=45"
      _StyleDefs(67)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=47"
      _StyleDefs(68)  =   "Splits(0).Columns(5).Style:id=70,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(69)  =   ":id=70,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(70)  =   ":id=70,.fontname=�l�r �S�V�b�N"
      _StyleDefs(71)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=44"
      _StyleDefs(72)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=45"
      _StyleDefs(73)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=47"
      _StyleDefs(74)  =   "Splits(0).Columns(6).Style:id=74,.parent=43,.alignment=1"
      _StyleDefs(75)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=44"
      _StyleDefs(76)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=45"
      _StyleDefs(77)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=47"
      _StyleDefs(78)  =   "Splits(0).Columns(7).Style:id=78,.parent=43,.alignment=1"
      _StyleDefs(79)  =   "Splits(0).Columns(7).HeadingStyle:id=75,.parent=44"
      _StyleDefs(80)  =   "Splits(0).Columns(7).FooterStyle:id=76,.parent=45"
      _StyleDefs(81)  =   "Splits(0).Columns(7).EditorStyle:id=77,.parent=47"
      _StyleDefs(82)  =   "Splits(0).Columns(8).Style:id=16,.parent=43"
      _StyleDefs(83)  =   "Splits(0).Columns(8).HeadingStyle:id=13,.parent=44"
      _StyleDefs(84)  =   "Splits(0).Columns(8).FooterStyle:id=14,.parent=45"
      _StyleDefs(85)  =   "Splits(0).Columns(8).EditorStyle:id=15,.parent=47"
      _StyleDefs(86)  =   "Named:id=33:Normal"
      _StyleDefs(87)  =   ":id=33,.parent=0"
      _StyleDefs(88)  =   "Named:id=34:Heading"
      _StyleDefs(89)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(90)  =   ":id=34,.wraptext=-1"
      _StyleDefs(91)  =   "Named:id=35:Footing"
      _StyleDefs(92)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(93)  =   "Named:id=36:Selected"
      _StyleDefs(94)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(95)  =   "Named:id=37:Caption"
      _StyleDefs(96)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(97)  =   "Named:id=38:HighlightRow"
      _StyleDefs(98)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(99)  =   "Named:id=39:EvenRow"
      _StyleDefs(100) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(101) =   "Named:id=40:OddRow"
      _StyleDefs(102) =   ":id=40,.parent=33"
      _StyleDefs(103) =   "Named:id=41:RecordSelector"
      _StyleDefs(104) =   ":id=41,.parent=34"
      _StyleDefs(105) =   "Named:id=42:FilterBar"
      _StyleDefs(106) =   ":id=42,.parent=33"
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�\������"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   22
      Left            =   11520
      TabIndex        =   110
      Top             =   10440
      Width           =   1035
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '����
      Caption         =   "���H��Ɓi�a�t�����j"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   47
      Left            =   8910
      TabIndex        =   92
      Top             =   4920
      Width           =   4545
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '����
      Caption         =   "��Еt��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   46
      Left            =   8910
      TabIndex        =   91
      Top             =   4560
      Width           =   4545
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '����
      Caption         =   "���i����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   45
      Left            =   8910
      TabIndex        =   90
      Top             =   4200
      Width           =   4545
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '����
      Caption         =   "�B"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   8910
      TabIndex        =   83
      Top             =   3840
      Width           =   4545
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "��Ǝ���"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   8910
      TabIndex        =   77
      Top             =   1680
      Width           =   4545
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '����
      Caption         =   "�A����Ɓi���{�m�F�܂ށj"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   8910
      TabIndex        =   82
      Top             =   3480
      Width           =   4545
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '����
      Caption         =   "�@���x���\��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   8910
      TabIndex        =   81
      Top             =   3120
      Width           =   4545
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '����
      Caption         =   "�������i����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   8910
      TabIndex        =   80
      Top             =   2760
      Width           =   4545
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '����
      Caption         =   "�����ޏ���"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   8910
      TabIndex        =   79
      Top             =   2400
      Width           =   4545
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '����
      Caption         =   "���i����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   8910
      TabIndex        =   78
      Top             =   2040
      Width           =   4545
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�H��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   8205
      TabIndex        =   109
      Top             =   1680
      Width           =   750
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "���H"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   8205
      TabIndex        =   108
      Top             =   4920
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���P/�S����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   56
      Left            =   0
      TabIndex        =   107
      Top             =   5640
      Width           =   1395
   End
   Begin VB.Label Label1 
      Height          =   255
      Index           =   55
      Left            =   4800
      TabIndex        =   98
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�����:"
      Height          =   255
      Index           =   54
      Left            =   3675
      TabIndex        =   97
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      Caption         =   "-"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   2595
      TabIndex        =   96
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   8205
      TabIndex        =   95
      Top             =   4200
      Width           =   750
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   " ���"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   1
      Left            =   8205
      TabIndex        =   94
      Top             =   3120
      Width           =   750
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�O"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   0
      Left            =   8205
      TabIndex        =   93
      Top             =   2040
      Width           =   750
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '����
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   42
      Left            =   8205
      TabIndex        =   89
      Top             =   5640
      Width           =   750
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '����
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   39
      Left            =   8205
      TabIndex        =   88
      Top             =   5280
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "���v"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   38
      Left            =   8205
      TabIndex        =   87
      Top             =   6000
      Width           =   6810
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "���v"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   27
      Left            =   14970
      TabIndex        =   86
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   14205
      TabIndex        =   85
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�l"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   13440
      TabIndex        =   84
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BorderStyle     =   1  '����
      Caption         =   "�����ϐ�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   12
      Left            =   11280
      TabIndex        =   76
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BorderStyle     =   1  '����
      Caption         =   "�\�萔"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   11
      Left            =   8280
      TabIndex        =   75
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "�@���E�N���X"
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
      Index           =   10
      Left            =   240
      TabIndex        =   74
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "�@�t���N���X"
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
      Index           =   9
      Left            =   240
      TabIndex        =   73
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "���i���N���X"
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
      Index           =   8
      Left            =   240
      TabIndex        =   72
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�i��"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   71
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "��z��"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   70
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�d������"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   69
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���F��"
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   68
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�S����"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   67
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���s��"
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   66
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "������"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   65
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�w�}�[��"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   64
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "PI000251"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'����



Private Const plblBUNNOU_CNT% = 55          '���[��

'���̕\���p�e�L�X�g�Y��
Private Const ptx2TANTO_NAME% = 0          '�S���Җ���
Private Const ptx2SHONIN_NAME% = 1         '���F�Җ���
Private Const ptx2HIN_NAME% = 2            '�i��
Private Const ptx2S_CLASS_NAME% = 3        '���i���N���X
Private Const ptx2F_CLASS_NAME% = 4        '�t���N���X
Private Const ptx2N_CLASS_NAME% = 5        '���E�N���X
    
    
    
'�e�L�X�g�p�Y��
Private Const ptxSHIJI_NO% = 0              '�w�}�[��
Private Const ptxSEQNO% = 1                 '�ǔ�
Private Const ptxUKEIRE_DT% = 2             '������t
Private Const ptxHAKKO_DT% = 3              '���s��
Private Const ptxTANTO_CODE% = 4            '�S���Һ���
'Private Const ptxTANTO_NAME% = 5            '�S���Җ���
Private Const ptxSHONIN_CODE% = 6           '���F�Һ���
'Private Const ptxSHONIN_NAME% = 7           '���F�Җ���
Private Const ptxUKEHARAI_CODE% = 8         '���F�Һ���
Private Const ptxHIN_GAI% = 9               '�i��
'Private Const ptxHIN_NAME% = 10             '�i��
Private Const ptxS_CLASS_CODE% = 11         '���i���׽
Private Const ptxF_CLASS_CODE% = 12         '�t���׽
Private Const ptxN_CLASS_CODE% = 13         '���E�׽

Private Const ptxSHIJI_QTY% = 14            '�\��
Private Const ptxUKEIRE_QTY% = 15           '���

Private Const ptxNIN01% = 16                '����1(�w�}�[/���ٔ��s)  �l
Private Const ptxTIMES01% = 17              '����1(�w�}�[/���ٔ��s)  ��
Private Const ptxTOTAL01% = 18              '����1(�w�}�[/���ٔ��s)  ���v

Private Const ptxNIN02% = 19                '����2(���i����)�@�l
Private Const ptxTIMES02% = 20              '����2(���i����)�@��
Private Const ptxTOTAL02% = 21              '����2(���i����)�@���v

Private Const ptxNIN03% = 22                '����3(���ޏo��)�@�l
Private Const ptxTIMES03% = 23              '����3(���ޏo��)�@��
Private Const ptxTOTAL03% = 24              '����3(���ޏo��)�@���v

Private Const ptxNIN04% = 25                '����4(��������o�ɂȂ�)�@�l
Private Const ptxTIMES04% = 26              '����4(��������o�ɂȂ�)�@��
Private Const ptxTOTAL04% = 27              '����4(��������o�ɂȂ�)�@���v

Private Const ptxNIN05% = 28                '���1(���ٓ\��)�@�l
Private Const ptxTIMES05% = 29              '���1(���ٓ\��)�@��
Private Const ptxTOTAL05% = 30              '���1(���ٓ\��)�@���v

Private Const ptxNIN06% = 31                '���2�@�l
Private Const ptxTIMES06% = 32              '���2�@��
Private Const ptxTOTAL06% = 33              '���2�@���v

Private Const ptxNIN07% = 34                '���3�@�l
Private Const ptxTIMES07% = 35              '���3�@��
Private Const ptxTOTAL07% = 36              '���3�@���v

Private Const ptxNIN08% = 37                '���1(���i����) �l
Private Const ptxTIMES08% = 38              '���1(���i����) ��
Private Const ptxTOTAL08% = 39              '���1(���i����) ���v

Private Const ptxNIN09% = 40                '���2(���i����) �l
Private Const ptxTIMES09% = 41              '���2(�w�}�[�����o�^) ��
Private Const ptxTOTAL09% = 42              '���2(�w�}�[�����o�^) ���v

Private Const ptxNIN10% = 43                '���Ӂ@�@�@�l
Private Const ptxTIMES10% = 44              '���Ӂ@�@�@��
Private Const ptxTOTAL10% = 45              '���Ӂ@�@���v

Private Const ptxNIN11% = 46                '���Ӂ@�@�@�l
Private Const ptxTIMES11% = 47              '���Ӂ@�@�@��
Private Const ptxTOTAL11% = 48              '���Ӂ@�@���v


Private Const ptxTOTAL% = 49                '���v

    '���l 2010.09.28
Private Const ptxBIKOU% = 50                '���l


'�R���{�p�Y��
Private Const pcmbSHIMUKE% = 0              '�d������
Private Const pcmbUKEHARAI% = 1             '��z��

Private Const pcmbS_TANTO% = 2              '���P�^�S����


Private Const pcmbJISEKI% = 3               '���ӗv��
Private Const pcmbTASEKI% = 4               '���ӗv��



'Glid�p��

Dim SSHIJI  As New XArrayDB

Private Const Min_Row% = 1                  '�ŏ��s��
Private Const Min_Col% = 0                  '�ŏ���
Private Const Max_Col% = 8                  '�ő��


Private Const colHAKKO_DT% = 0              '���s����
Private Const colSHIJI_NO% = 1              '�w�}�[��
Private Const colSHIMUKE% = 2               '�d������
Private Const colUKEHARAI% = 3              '��z��
Private Const colHIN_GAI% = 4               '�i��
Private Const colHIN_NAME% = 5              '�i��
Private Const colSHIJI_QTY% = 6             '�\��
Private Const colUKEIRE_QTY% = 7            '���

Private Const colS_TANTO% = 8               '���P�^�S����


Private Sort_Tbl(colHAKKO_DT To colS_TANTO) _
                As Integer                  '��Ă̐��� 0:���� 1:�~��
                                            
Private JISEKI_TITLE    As Variant      '���ӂ̖��̃^�C�g��
Private TASEKI_TITLE    As Variant      '���ӂ̖��̃^�C�g��
Public PRI_S_TANTO      As Boolean      '���x�^�S���҈�� OFF:����Ȃ� ON:�������

Private JISSEKI_DSP     As String * 1   '2008.08.19
Private Time_Input_F    As Boolean      '2008.08.19

'Private Const LAST_UPDATE_DAY$ = "[PI00025] 2017.09.27 15:45"
Private Const LAST_UPDATE_DAY$ = "[PI00025] 2019.04.05 11:15"




Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    PI000251.MousePointer = vbHourglass

    Call Ctrl_Lock(PI000251)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PI000251)


    PI000251.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   ���͍��ڂ̃G���[�`�F�b�N
'----------------------------------------------------------------------------
    
Dim sts     As Integer
Dim i       As Integer
    
    
    Error_Check_Proc = True
    
    
    
    
    Select Case Mode
    
        Case ptxSHIJI_NO    '�w�}�[��
        
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CLng(Text1(Mode).Text), "00000000")
            End If
            '�w�}�[�ް�������
            
            If Text1(Mode).Text = StrConv(P_SSHIJI_O_REC.SHIJI_No, vbUnicode) And _
                Text1(Mode + 1).Text = StrConv(P_SUKEIRE_REC.SEQNO, vbUnicode) Then
                sts = BtNoErr
            Else
                sts = P_SSHIJI_Read_Proc()
            End If
            Select Case sts
                Case False, BtNoErr
                            
                    If StrConv(P_SSHIJI_O_REC.KAN_F, vbUnicode) <> P_KAN_ON Then
                        'MsgBox "�������ł��B"              '2017.07.22
                        MsgBox "�������̎w�}�[�ł��B"       '2017.07.22
                        Text1(Mode).SetFocus
                        Exit Function
                    End If
                
                    If StrConv(P_SSHIJI_O_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
                        'MsgBox "�L�����Z�������ς݂ł��B"              '2017.07.22
                        MsgBox "�L�����Z�������ς݂̎w�}�[�ł��B"       '2017.07.22
                        Text1(Mode).SetFocus
                        Exit Function
                    End If
                
                
                Case BtErrKeyNotFound
                    'MsgBox "���͂������ڂ̓G���[�ł��B"                '2017.07.22
                    MsgBox "���͂����w�}�[���͓o�^����Ă��܂���B"     '2017.07.22
                    Text1(Mode).SetFocus
                    Exit Function
                Case Else
                    Exit Function
            End Select
        
        
        Case ptxSEQNO       '�ǔ�
        
            If Trim(Text1(Mode).Text) = "" Then                 '2017.07.21
                MsgBox "�ǔԂ𐳂������͂��ĉ�����"             '2017.07.21
                Text1(Mode).SetFocus                            '2017.07.21
                Exit Function                                   '2017.07.21
            End If                                              '2017.07.21
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CLng(Text1(Mode).Text), "000")
            
            Else                                                '2017.07.21
                MsgBox "�ǔԂ𐳂������͂��ĉ�����"             '2017.07.21
                Text1(Mode).SetFocus                            '2017.07.21
                Exit Function                                   '2017.07.21
            End If
            '�w�}�[�ް�������
            
            If Text1(Mode).Text = StrConv(P_SSHIJI_O_REC.SHIJI_No, vbUnicode) And _
                Text1(Mode + 1).Text = StrConv(P_SUKEIRE_REC.SEQNO, vbUnicode) Then
                sts = BtNoErr
            Else
                Call UniCode_Conv(K0_P_SUKEIRE.SHIJI_No, Text1(ptxSHIJI_NO).Text)
                Call UniCode_Conv(K0_P_SUKEIRE.SEQNO, Text1(ptxSEQNO).Text)
                sts = P_SUKEIRE_Read_Proc()
            End If
            Select Case sts
                Case False, BtNoErr
                            
                
                
                Case BtErrKeyNotFound
                    'MsgBox "���͂������ڂ̓G���[�ł��B"                        '2017.07.22
                    MsgBox "���͂����w�}�[��(�ǔ�)�͓o�^����Ă��܂���B"       '2017.07.22
                    Text1(Mode).SetFocus
                    Exit Function
                Case Else
                    Exit Function
            End Select
        
        
        
        Case ptxHAKKO_DT    '���s��
            
        Case ptxUKEIRE_DT   '������
            
            If Not IsDate(Text1(ptxUKEIRE_DT).Text) Then
                'MsgBox "���͂������ڂ̓G���[�ł��B"                        '2017.07.22
                MsgBox "�������𐳂������͂��ĉ������B(YYYY/MM/DD)"         '2017.07.22
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxUKEIRE_DT).Text = Format(CDate(Text1(ptxUKEIRE_DT).Text), "YYYY/MM/DD")
            End If
        
        
        Case ptxUKEHARAI_CODE   '��z��
            
            
            
            Text1(Mode).Text = StrConv(Text1(Mode).Text, vbUpperCase)       '2017.07.22
            
            
            
            Combo1(pcmbUKEHARAI).ListIndex = -1
            For i = 0 To Combo1(pcmbUKEHARAI).ListCount - 1
                If Text1(ptxUKEHARAI_CODE).Text = Trim(Right(Combo1(pcmbUKEHARAI).List(i), 5)) Then
                    Combo1(pcmbUKEHARAI).ListIndex = i
                    Exit For
                End If
            
            Next i
     
            If i > Combo1(pcmbUKEHARAI).ListCount - 1 Then
                'MsgBox "���͂������ڂ̓G���[�ł��B"                        '2017.07.22
                MsgBox "���͂�����z��͓o�^����Ă��܂���B"               '2017.07.22
                Text1(Mode).SetFocus
                Exit Function
            End If
        
        
        
    
        Case ptxS_CLASS_CODE    '���i���׽
            
            
            
            Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
            Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, Text1(ptxS_CLASS_CODE).Text)

            sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
            Select Case sts
                Case BtNoErr
                
                    Text2(ptx2S_CLASS_NAME).Text = StrConv(P_CLASSREC.CLASS_NAME, vbUnicode)
                
                Case BtErrKeyNotFound
                    Text2(ptx2S_CLASS_NAME).Text = ""
'                    MsgBox "���͂������ڂ̓G���[�ł��B"                    '2017.07.22
                    MsgBox "���͂������i���N���X�͓o�^����Ă��܂���B"     '2017.07.22
                    Text1(Mode).SetFocus
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "���i���׽")
                    Exit Function
            
            End Select
    
        Case ptxF_CLASS_CODE    '�t���׽
        
            If Trim(Text1(ptxF_CLASS_CODE).Text) = "" Then
            Else
                
                
                Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
                Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, Text1(ptxF_CLASS_CODE).Text)
    
                sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
                Select Case sts
                    Case BtNoErr
                        Text2(ptx2F_CLASS_NAME).Text = StrConv(P_CLASSREC.CLASS_NAME, vbUnicode)
                    Case BtErrKeyNotFound
                        Text2(ptx2F_CLASS_NAME).Text = ""
                
                        'MsgBox "���͂������ڂ̓G���[�ł��B"                    '2017.07.22
                        MsgBox "���͂����t���N���X�͓o�^����Ă��܂���B"       '2017.07.22
                        Text1(Mode).SetFocus
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "���i���׽")
                        Exit Function
                
                End Select
            End If
    
        Case ptxN_CLASS_CODE    '���E�׽
        
            If Trim(Text1(ptxN_CLASS_CODE).Text) = "" Then
            Else
                
                
                
                Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
                Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, Text1(ptxN_CLASS_CODE).Text)
    
                sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
                Select Case sts
                    Case BtNoErr
                        Text2(ptx2F_CLASS_NAME).Text = StrConv(P_CLASSREC.CLASS_NAME, vbUnicode)
                    Case BtErrKeyNotFound
                        Text2(ptx2F_CLASS_NAME).Text = ""
                
                        'MsgBox "���͂������ڂ̓G���[�ł��B"                '2017.07.22
                        MsgBox "���͂������E�N���X�͓o�^����Ă��܂���B"   '2017.07.22
                        Text1(Mode).SetFocus
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "���i���׽")
                        Exit Function
                
                End Select
            End If
        
        
        
        Case ptxUKEIRE_QTY       '�����
    
            If Not IsNumeric(Text1(ptxUKEIRE_QTY).Text) Then
'                MsgBox "���͂������ڂ̓G���[�ł��B"                '2017.07.22
                MsgBox "������͐��l����͂��ĉ������B"             '2017.07.22
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxUKEIRE_QTY).Text = Format(CLng(Text1(ptxUKEIRE_QTY).Text), "#0")
            
                If CLng(Text1(ptxUKEIRE_QTY).Text) > CLng(Text1(ptxSHIJI_QTY).Text) Then
                    'MsgBox "���͂������ڂ̓G���[�ł��B"            '2017.07.22
                    MsgBox "��������\�萔���I�[�o�[���Ă��܂��B"   '2017.07.22
                    Text1(Mode).SetFocus
                    Exit Function
                End If
                    
                    
                    
            End If
    
            
        Case ptxNIN01, ptxNIN02, ptxNIN03, ptxNIN04, ptxNIN05, ptxNIN06, ptxNIN07, ptxNIN08, ptxNIN09 '�l��
            If Text1(Mode).Text = "" Then
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
                    'MsgBox "���͂������ڂ̓G���[�ł��B"        '2017.07.22
                    MsgBox "�l���͐��l����͂��ĉ������B"       '2017.07.22
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    If InStr(1, Trim(Text1(Mode).Text), ".") > 2 Then
                        'MsgBox "���͂������ڂ̓G���[�ł��B(9.9)"   '2017.07.22
                        MsgBox "�l����9.9�ȉ��œ��͂��ĉ������B"    '2017.07.22
                        Text1(Mode).SetFocus
                        Exit Function
                    End If
                    
                    Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.0")
                    If IsNumeric(Text1(Mode + 1).Text) Then
                        Text1(Mode + 2).Text = Format(CDbl(Text1(Mode).Text) * CDbl(Text1(Mode + 1).Text), "#0.00")
                    Else
                        Text1(Mode + 2).Text = "0.00"
                    End If
                
                    Text1(ptxTOTAL).Text = "0"
                    For i = ptxTOTAL01 To ptxTOTAL11 Step 3
                        
                        If IsNumeric(Text1(i).Text) Then
                           Text1(ptxTOTAL).Text = Format(CDbl(Text1(ptxTOTAL).Text) + CDbl(Text1(i).Text), "#0.00")
                        End If
                    
                    Next i
                End If
            End If
    
        Case ptxTIMES01, ptxTIMES02, ptxTIMES03, ptxTIMES04, ptxTIMES05, ptxTIMES06, ptxTIMES07, ptxTIMES08, ptxTIMES09 '����
            If Text1(Mode).Text = "" Then
                If Text1(Mode - 1).Text = "" Then
                    Text1(Mode + 1).Text = ""
                End If
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
                    'MsgBox "���͂������ڂ̓G���[�ł��B"            '2017.07.22
                    MsgBox "���Ԃ͐��l�œ��͂��ĉ������B"           '2017.07.22
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.00")
                    If IsNumeric(Text1(Mode - 1).Text) Then
                        Text1(Mode + 1).Text = Format(CDbl(Text1(Mode - 1).Text) * CDbl(Text1(Mode).Text), "#0.00")
                    Else
                        Text1(Mode + 1).Text = "0.00"
                    End If
                
                    Text1(ptxTOTAL).Text = "0.00"
                    For i = ptxTOTAL01 To ptxTOTAL11 Step 3
                        
                        If IsNumeric(Text1(i).Text) Then
                           Text1(ptxTOTAL).Text = Format(CDbl(Text1(ptxTOTAL).Text) + CDbl(Text1(i).Text), "#0.00")
                        End If
                    
                    Next i
                End If
            End If
    
    
        Case ptxNIN10, ptxNIN11             '���Ӂ@�l��
            
            If Text1(Mode).Text = "" Then
            
            
            
            Else
                
                If Mode = ptxNIN10 Then
                    
                    If Combo1(pcmbJISEKI).Text = "" Then
                                
                        'MsgBox "���͂������ڂ̓G���[�ł��B"            '2017.07.22
                        MsgBox "���ӓ��e��I����ɓ��͂��ĉ������B"     '2017.07.22
                        Text1(Mode).SetFocus
                        Exit Function
                
                    End If
                
                Else
                
                    If Combo1(pcmbTASEKI).Text = "" Then
                                
                        'MsgBox "���͂������ڂ̓G���[�ł��B"            '2017.07.22
                        MsgBox "���ӓ��e��I����ɓ��͂��ĉ������B"     '2017.07.22
                        Text1(Mode).SetFocus
                        Exit Function
                
                    End If
                
                End If
                
                
                If Not IsNumeric(Text1(Mode).Text) Then
                    'MsgBox "���͂������ڂ̓G���[�ł��B"                '2017.07.22
                    MsgBox "�l���͐��l����͂��ĉ������B"               '2017.07.22
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    
                    Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.0")
                    If IsNumeric(Text1(Mode + 1).Text) Then
                        Text1(Mode + 2).Text = Format(CDbl(Text1(Mode).Text) * CDbl(Text1(Mode + 1).Text), "#0.00")
                    Else
                        Text1(Mode + 2).Text = "0.00"
                    End If
                
                    Text1(ptxTOTAL).Text = "0.00"
                    For i = ptxTOTAL01 To ptxTOTAL11 Step 3
                        
                        If IsNumeric(Text1(i).Text) Then
                           Text1(ptxTOTAL).Text = Format(CDbl(Text1(ptxTOTAL).Text) + CDbl(Text1(i).Text), "#0.00")
                        End If
                    
                    Next i
                End If
            End If
    
        Case ptxTIMES10, ptxTIMES11  '����
            
            
            If Text1(Mode).Text = "" Then
                If Text1(Mode - 1).Text = "" Then
                    Text1(Mode + 1).Text = ""
                End If
            Else
                
                If Mode = ptxTIMES10 Then
                
                    If Combo1(pcmbJISEKI).Text = "" Then
                                
                        'MsgBox "���͂������ڂ̓G���[�ł��B"            '2017.07.22
                        MsgBox "���ӓ��e��I����ɓ��͂��ĉ������B"     '2017.07.22
                        Text1(Mode).SetFocus
                        Exit Function
                
                    End If
            
                Else
            
                    If Combo1(pcmbTASEKI).Text = "" Then
                            
                        'MsgBox "���͂������ڂ̓G���[�ł��B"            '2017.07.22
                        MsgBox "���ӓ��e��I����ɓ��͂��ĉ������B"     '2017.07.22
                        Text1(Mode).SetFocus
                        Exit Function
                    End If
            
                End If
            
            
            
                If Not IsNumeric(Text1(Mode).Text) Then
                    'MsgBox "���͂������ڂ̓G���[�ł��B"            '2017.07.22
                    MsgBox "���Ԃ͐��l�œ��͂��ĉ������B"           '2017.07.22
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.00")
                    If IsNumeric(Text1(Mode - 1).Text) Then
                        Text1(Mode + 1).Text = Format(CDbl(Text1(Mode - 1).Text) * CDbl(Text1(Mode).Text), "#0.00")
                    Else
                        Text1(Mode + 1).Text = "0.00"
                    End If
                
                    Text1(ptxTOTAL).Text = "0"
                    For i = ptxTOTAL01 To ptxTOTAL11 Step 3
                        
                        If IsNumeric(Text1(i).Text) Then
                           Text1(ptxTOTAL).Text = Format(CDbl(Text1(ptxTOTAL).Text) + CDbl(Text1(i).Text), "#0.00")
                        End If
                    
                    Next i
                End If
            End If
    
    End Select
        
        
    Error_Check_Proc = False
    

End Function
Private Function Item_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   ��ʕ\��
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim i       As Integer

    Item_Disp_Proc = True
    
        
    
    Text1(ptxSHIJI_NO).Text = StrConv(P_SSHIJI_O_REC.SHIJI_No, vbUnicode)           '�w�}�[��
                                                                                    '�����
    If CInt(StrConv(P_SSHIJI_O_REC.BUNNOU_CNT, vbUnicode)) = 0 Then
        Label1(plblBUNNOU_CNT).Caption = Format(CInt(StrConv(P_SSHIJI_O_REC.BUNNOU_CNT, vbUnicode)) + 1, "#0")
    Else
        Label1(plblBUNNOU_CNT).Caption = Format(CInt(StrConv(P_SSHIJI_O_REC.BUNNOU_CNT, vbUnicode)), "#0")
    End If
    
'    Text1(ptxUKEIRE_DT).Text = Format(Now, "YYYY/MM/DD")                            '�����
    
    
    
    
    Text1(ptxHAKKO_DT).Text = Mid(StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 7, 2)  '���s��
    
    Text1(ptxTANTO_CODE).Text = StrConv(P_SSHIJI_O_REC.TANTO_CODE, vbUnicode)       '�S���Һ��ށ^����
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxTANTO_CODE).Text)

    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
            Text2(ptx2TANTO_NAME).Text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Text2(ptx2TANTO_NAME).Text = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
            Exit Function
    
    End Select
    
    Text1(ptxSHONIN_CODE).Text = StrConv(P_SSHIJI_O_REC.SHONIN_CODE, vbUnicode)     '���F�Һ��ށ^����
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxSHONIN_CODE).Text)

    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
            Text2(ptx2SHONIN_NAME).Text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Text2(ptx2SHONIN_NAME).Text = ""
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
    Text1(ptxUKEHARAI_CODE).Text = Trim(StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode))   '��z��
    For i = 0 To Combo1(pcmbUKEHARAI).ListCount - 1
    
        If Text1(ptxUKEHARAI_CODE).Text = Trim(Right(Combo1(pcmbUKEHARAI).List(i), 5)) Then
            Combo1(pcmbUKEHARAI).ListIndex = i
            Exit For
        End If
    
    Next i
    
    
    Text1(ptxHIN_GAI).Text = Trim(StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode))       '�i�ԁ^�i���^�W���I�ԁ^�����i�^���i����
        
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)

    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            Text2(ptx2HIN_NAME).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Text2(ptx2HIN_NAME).Text = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Function
    
    End Select
                                                                                    
    Text1(ptxS_CLASS_CODE).Text = Trim(StrConv(P_SSHIJI_O_REC.S_CLASS_CODE, vbUnicode)) '���i���׽
    
    Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
    Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, StrConv(P_SSHIJI_O_REC.S_CLASS_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
    Select Case sts
        Case BtNoErr
            Text2(ptx2S_CLASS_NAME).Text = StrConv(P_CLASSREC.CLASS_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Text2(ptx2S_CLASS_NAME).Text = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�N���X�}�X�^")
            Exit Function
    
    End Select
    
    
    
    
    Text1(ptxF_CLASS_CODE).Text = Trim(StrConv(P_SSHIJI_O_REC.F_CLASS_CODE, vbUnicode)) '�t���׽
    
    Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
    Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, StrConv(P_SSHIJI_O_REC.F_CLASS_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
    Select Case sts
        Case BtNoErr
            Text2(ptx2F_CLASS_NAME).Text = StrConv(P_CLASSREC.CLASS_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Text2(ptx2S_CLASS_NAME).Text = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�N���X�}�X�^")
            Exit Function
    
    End Select
    
    
    Text1(ptxN_CLASS_CODE).Text = Trim(StrConv(P_SSHIJI_O_REC.N_CLASS_CODE, vbUnicode)) '���E�׽
                                                                                        
    Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
    Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, StrConv(P_SSHIJI_O_REC.N_CLASS_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
    Select Case sts
        Case BtNoErr
            Text2(ptx2N_CLASS_NAME).Text = StrConv(P_CLASSREC.CLASS_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Text2(ptx2N_CLASS_NAME).Text = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�N���X�}�X�^")
            Exit Function
    
    End Select
                                                                                        
    If Combo1(pcmbS_TANTO).ListCount = 0 Then                                       '���P�^�S����
    Else
        For i = 0 To Combo1(pcmbS_TANTO).ListCount - 1
            If StrConv(P_SSHIJI_O_REC.S_TANTO, vbUnicode) = Right(Combo1(pcmbS_TANTO).List(i), 2) Then
                Combo1(pcmbS_TANTO).ListIndex = i
                Exit For
            End If
        Next i
    End If
                                                                                        
                                                                                        '�w������
    Text1(ptxSHIJI_QTY).Text = Format(CDbl(StrConv(P_SSHIJI_O_REC.SHIJI_QTY, vbUnicode)), "#0")
    
    
    
    
    
    Time_Input_F = False    '2008.08.19
    
    Item_Disp_Proc = False

End Function

Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                  ���i���w�}�[�ް�/��������X�V
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim ans         As Integer
Dim com         As Integer

Dim SEQNO       As Integer



Dim i           As Integer
Dim j           As Integer


    Update_Proc = True
                                        
    Call Input_Lock
                                        
                                        '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If

    
    
    
    
    
    
    
    
    '---------------------------------------------------    '�w�}�[�f�[�^�X�V
    
    '�w�}�[�f�[�^(ͯ�ް)����
    Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_No, Text1(ptxSHIJI_NO).Text)
    
    Do
    
        sts = BTRV(BtOpGetEqual + BtSNoWait, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
            
        Select Case sts
            Case BtNoErr
            
                
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
                                                                            '����溰��
    Call UniCode_Conv(P_SSHIJI_O_REC.UKEHARAI_CODE, Text1(ptxUKEHARAI_CODE).Text)
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, Text1(ptxUKEHARAI_CODE).Text)
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
                                                                                    
    Call UniCode_Conv(P_SSHIJI_O_REC.S_CLASS_CODE, Text1(ptxS_CLASS_CODE))  '���i���׽
    Call UniCode_Conv(P_SSHIJI_O_REC.F_CLASS_CODE, Text1(ptxF_CLASS_CODE))  '�t���׽
    Call UniCode_Conv(P_SSHIJI_O_REC.N_CLASS_CODE, Text1(ptxN_CLASS_CODE))  '���E�׽
    
                                                                            '���P�^�S���Ҹ׽
    Call UniCode_Conv(P_SSHIJI_O_REC.S_TANTO, Right(Combo1(pcmbS_TANTO).Text, 2))
    
    
    '�}�C�i�X�X�V
    
    
    
    Call UniCode_Conv(P_SSHIJI_O_REC.UKEIRE_QTY, Format(CLng(StrConv(P_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)) - _
                                                    Val(StrConv(P_SUKEIRE_REC.UKEIRE_QTY, vbUnicode)), "00000000.00"))
                                                    
                                                
    For i = 0 To 9
    
        
        
        
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(i).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(i).NIN, vbUnicode)) - _
                                                                    Val(StrConv(P_SUKEIRE_REC.GENKA_TBL(i).NIN, vbUnicode)), "0.0"))
                                                    
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(i).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(i).TIMES, vbUnicode)) - _
                                                                    Val(StrConv(P_SUKEIRE_REC.GENKA_TBL(i).TIMES, vbUnicode)), "000.00"))
                                                    
    
    Next i
                                                
    Call UniCode_Conv(P_SSHIJI_O_REC.JISEKI_NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.JISEKI_NIN, vbUnicode)) - _
                                                                Val(StrConv(P_SUKEIRE_REC.JISEKI_NIN, vbUnicode)), "0.0"))
                                                
    Call UniCode_Conv(P_SSHIJI_O_REC.JISEKI_TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)) - _
                                                                Val(StrConv(P_SUKEIRE_REC.JISEKI_TIMES, vbUnicode)), "000.00"))
                                                
    Call UniCode_Conv(P_SSHIJI_O_REC.TASEKI_NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.TASEKI_NIN, vbUnicode)) - _
                                                                Val(StrConv(P_SUKEIRE_REC.TASEKI_NIN, vbUnicode)), "0.0"))
                                                
    Call UniCode_Conv(P_SSHIJI_O_REC.TASEKI_TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)) - _
                                                                Val(StrConv(P_SUKEIRE_REC.TASEKI_TIMES, vbUnicode)), "000.00"))
                                                
                                                
                                                
                                                
    '�����敪
    Call UniCode_Conv(P_SSHIJI_O_REC.TORI_KBN, StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode))
                                                
                                                
    '�O1
    If IsNumeric(Text1(ptxNIN01).Text) Then        '�l
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(4).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(4).NIN, vbUnicode)) + _
                                                            CDbl(Text1(ptxNIN01).Text), "0.0"))
    End If

    If IsNumeric(Text1(ptxTIMES01).Text) Then      '����
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(4).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(4).TIMES, vbUnicode)) + _
                                                            CDbl(Text1(ptxTIMES01).Text), "000.00"))
    End If
    
    If Trim(Text1(ptxNIN01).Text) = "" And _
         Trim(Text1(ptxTIMES01).Text) = "" Then
    Else
        If Trim(Text1(ptxNIN01).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(4).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(4).NIN, vbUnicode)) + 1, "0.0"))
        End If
    
        If Trim(Text1(ptxTIMES01).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(4).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(4).NIN, vbUnicode)) + 1, "000.00"))
        End If
    End If
                                                
                                                
    '�O2
    If IsNumeric(Text1(ptxNIN02).Text) Then        '�l
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(5).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(5).NIN, vbUnicode)) + CDbl(Text1(ptxNIN02).Text), "0.0"))
    End If

    If IsNumeric(Text1(ptxTIMES02).Text) Then      '����
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(5).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(5).TIMES, vbUnicode)) + CDbl(Text1(ptxTIMES02).Text), "000.00"))
    End If
    
    If Trim(Text1(ptxNIN02).Text) = "" And _
         Trim(Text1(ptxTIMES02).Text) = "" Then
    Else
        If Trim(Text1(ptxNIN02).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(5).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(5).NIN, vbUnicode)) + 1, "0.0"))
        End If
    
        If Trim(Text1(ptxTIMES02).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(5).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(5).NIN, vbUnicode)) + 1, "000.00"))
        End If
    End If
                                                
    '�O3
    If IsNumeric(Text1(ptxNIN03).Text) Then        '�l
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(6).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(6).NIN, vbUnicode)) + CDbl(Text1(ptxNIN03).Text), "0.0"))
    End If

    If IsNumeric(Text1(ptxTIMES03).Text) Then      '����
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(6).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(6).TIMES, vbUnicode)) + CDbl(Text1(ptxTIMES03).Text), "000.00"))
    End If
    
    If Trim(Text1(ptxNIN03).Text) = "" And _
         Trim(Text1(ptxTIMES03).Text) = "" Then
    Else
        If Trim(Text1(ptxNIN03).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(6).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(6).NIN, vbUnicode)) + 1, "0.0"))
        End If
    
        If Trim(Text1(ptxTIMES03).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(6).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(6).NIN, vbUnicode)) + 1, "000.00"))
        End If
    End If
    
    '���1
    If IsNumeric(Text1(ptxNIN04).Text) Then        '�l
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(0).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(0).NIN, vbUnicode)) + CDbl(Text1(ptxNIN04).Text), "0.0"))
    End If

    If IsNumeric(Text1(ptxTIMES04).Text) Then      '����
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(0).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(0).TIMES, vbUnicode)) + CDbl(Text1(ptxTIMES04).Text), "000.00"))
    End If
    
    If Trim(Text1(ptxNIN04).Text) = "" And _
         Trim(Text1(ptxTIMES04).Text) = "" Then
    Else
        If Trim(Text1(ptxNIN04).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(0).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(0).NIN, vbUnicode)) + 1, "0.0"))
        End If
    
        If Trim(Text1(ptxTIMES04).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(0).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(0).NIN, vbUnicode)) + 1, "000.00"))
        End If
    End If
                                                
    '���2
    If IsNumeric(Text1(ptxNIN05).Text) Then        '�l
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(1).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(1).NIN, vbUnicode)) + CDbl(Text1(ptxNIN05).Text), "0.0"))
    End If

    If IsNumeric(Text1(ptxTIMES05).Text) Then      '����
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(1).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(1).TIMES, vbUnicode)) + CDbl(Text1(ptxTIMES05).Text), "000.00"))
    End If
    
    If Trim(Text1(ptxNIN05).Text) = "" And _
         Trim(Text1(ptxTIMES05).Text) = "" Then
    Else
        If Trim(Text1(ptxNIN05).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(1).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(1).NIN, vbUnicode)) + 1, "0.0"))
        End If
    
        If Trim(Text1(ptxTIMES05).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(1).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(1).NIN, vbUnicode)) + 1, "000.00"))
        End If
    End If
    
    '���3
    If IsNumeric(Text1(ptxNIN06).Text) Then        '�l
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(2).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(2).NIN, vbUnicode)) + CDbl(Text1(ptxNIN06).Text), "0.0"))
    End If

    If IsNumeric(Text1(ptxTIMES06).Text) Then      '����
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(2).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(2).TIMES, vbUnicode)) + CDbl(Text1(ptxTIMES06).Text), "000.00"))
    End If
    
    If Trim(Text1(ptxNIN06).Text) = "" And _
         Trim(Text1(ptxTIMES06).Text) = "" Then
    Else
        If Trim(Text1(ptxNIN06).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(2).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(2).NIN, vbUnicode)) + 1, "0.0"))
        End If
    
        If Trim(Text1(ptxTIMES06).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(2).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(2).NIN, vbUnicode)) + 1, "000.00"))
        End If
    End If
    
    '��1
    If IsNumeric(Text1(ptxNIN07).Text) Then        '�l
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(7).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(7).NIN, vbUnicode)) + CDbl(Text1(ptxNIN07).Text), "0.0"))
    End If

    If IsNumeric(Text1(ptxTIMES07).Text) Then      '����
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(7).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(7).TIMES, vbUnicode)) + CDbl(Text1(ptxTIMES07).Text), "000.00"))
    End If
    
    If Trim(Text1(ptxNIN07).Text) = "" And _
         Trim(Text1(ptxTIMES07).Text) = "" Then
    Else
        If Trim(Text1(ptxNIN07).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(7).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(7).NIN, vbUnicode)) + 1, "0.0"))
        End If
    
        If Trim(Text1(ptxTIMES07).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(7).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(7).NIN, vbUnicode)) + 1, "000.00"))
        End If
    End If
                                                
    '��2
    If IsNumeric(Text1(ptxNIN08).Text) Then        '�l
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(8).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(8).NIN, vbUnicode)) + CDbl(Text1(ptxNIN08).Text), "0.0"))
    End If

    If IsNumeric(Text1(ptxTIMES08).Text) Then      '����
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(8).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(8).TIMES, vbUnicode)) + CDbl(Text1(ptxTIMES08).Text), "000.00"))
    End If
    
    If Trim(Text1(ptxNIN08).Text) = "" And _
         Trim(Text1(ptxTIMES08).Text) = "" Then
    Else
        If Trim(Text1(ptxNIN08).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(8).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(8).NIN, vbUnicode)) + 1, "0.0"))
        End If
    
        If Trim(Text1(ptxTIMES08).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(8).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(8).NIN, vbUnicode)) + 1, "000.00"))
        End If
    End If
                                                
    '���H1
    If IsNumeric(Text1(ptxNIN09).Text) Then        '�l
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(9).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(9).NIN, vbUnicode)) + CDbl(Text1(ptxNIN09).Text), "0.0"))
    End If

    If IsNumeric(Text1(ptxTIMES09).Text) Then      '����
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(9).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(9).TIMES, vbUnicode)) + CDbl(Text1(ptxTIMES09).Text), "000.00"))
    End If
    
    If Trim(Text1(ptxNIN09).Text) = "" And _
         Trim(Text1(ptxTIMES09).Text) = "" Then
    Else
        If Trim(Text1(ptxNIN09).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(8).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(8).NIN, vbUnicode)) + 1, "0.0"))
        End If
    
        If Trim(Text1(ptxTIMES09).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(8).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(8).NIN, vbUnicode)) + 1, "000.00"))
        End If
    End If
                                                
                                                
                                                
                                                '���Ӂ@����
    Call UniCode_Conv(P_SSHIJI_O_REC.JISEKI_NAME, Combo1(pcmbJISEKI).Text)
    If IsNumeric(Text1(ptxNIN10).Text) Then        '�l
        Call UniCode_Conv(P_SSHIJI_O_REC.JISEKI_NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.JISEKI_NIN, vbUnicode)) + CDbl(Text1(ptxNIN10).Text), "0.0"))
    End If

    If IsNumeric(Text1(ptxTIMES10).Text) Then      '����
        Call UniCode_Conv(P_SSHIJI_O_REC.JISEKI_TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)) + CDbl(Text1(ptxTIMES10).Text), "000.00"))
    End If
    
    
    If Trim(Text1(ptxNIN10).Text) = "" And _
         Trim(Text1(ptxTIMES10).Text) = "" Then
    Else
        If Trim(Text1(ptxNIN10).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.JISEKI_NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.JISEKI_NIN, vbUnicode)) + 1, "0.0"))
        End If
    
        If Trim(Text1(ptxTIMES10).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.JISEKI_TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)) + 1, "000.00"))
        End If
    End If
    
                                                '���Ӂ@����
    Call UniCode_Conv(P_SSHIJI_O_REC.TASEKI_NAME, Combo1(pcmbTASEKI).Text)
    If IsNumeric(Text1(ptxNIN11).Text) Then        '�l
        Call UniCode_Conv(P_SSHIJI_O_REC.TASEKI_NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.TASEKI_NIN, vbUnicode)) + CDbl(Text1(ptxNIN11).Text), "0.0"))
    End If

    If IsNumeric(Text1(ptxTIMES11).Text) Then      '����
        Call UniCode_Conv(P_SSHIJI_O_REC.TASEKI_TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)) + CDbl(Text1(ptxTIMES11).Text), "000.00"))
    End If
                                                        
    If Trim(Text1(ptxNIN11).Text) = "" And _
         Trim(Text1(ptxTIMES11).Text) = "" Then
    Else
        If Trim(Text1(ptxNIN11).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.TASEKI_NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.TASEKI_NIN, vbUnicode)) + 1, "0.0"))
        End If
    
        If Trim(Text1(ptxTIMES11).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.TASEKI_TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)) + 1, "000.00"))
        End If
    End If
                                                        '�������
    Call UniCode_Conv(P_SSHIJI_O_REC.UKEIRE_QTY, Format(CDbl(CDbl(StrConv(P_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)) + CDbl(Text1(ptxUKEIRE_QTY).Text)), "00000000.00"))
                                                        '�X�V����
    Call UniCode_Conv(P_SSHIJI_O_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    Do
        
        DoEvents
        
        sts = BTRV(BtOpUpdate, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_SSHIJI_O.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
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
                Call File_Error(sts, BtOpUpdate, "���i���w�}�ް�(�e)")
                GoTo Abort_Tran
        End Select
    
    Loop
    
    
    SEQNO = 0
    
    
    
    '��������ް�����
    Call UniCode_Conv(K0_P_SUKEIRE.SHIJI_No, Text1(ptxSHIJI_NO).Text)
    Call UniCode_Conv(K0_P_SUKEIRE.SEQNO, Text1(ptxSEQNO).Text)
    
    Do
    
        sts = BTRV(BtOpGetEqual + BtSNoWait, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K0_P_SUKEIRE, Len(K0_P_SUKEIRE), 0)
            
        Select Case sts
            Case BtNoErr
                Exit Do
            
            Case BtErrKeyNotFound
                MsgBox "��������Ȃ��I�I"
                GoTo Abort_Tran
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                MsgBox "�ް��g�p���I�I"
                GoTo Abort_Tran
            Case Else
                Call File_Error(sts, BtOpGetEqual, "���i���������")
                GoTo Abort_Tran
        End Select
        
        
    Loop
        
        
    Call UniCode_Conv(P_SUKEIRE_REC.SHIJI_No, Text1(ptxSHIJI_NO).Text)          '�w�}�[��
                                                                                '�d�����溰��
    Call UniCode_Conv(P_SUKEIRE_REC.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
                                                                                '�����
    Call UniCode_Conv(P_SUKEIRE_REC.UKEIRE_DT, Format(Text1(ptxUKEIRE_DT).Text, "YYYYMMDD"))
                                                                                '�������
    Call UniCode_Conv(P_SUKEIRE_REC.UKEIRE_QTY, Format(CDbl(Text1(ptxUKEIRE_QTY).Text), "00000000.00"))
        
    '���l 2010.09.28
    Call UniCode_Conv(P_SUKEIRE_REC.BIKOU, Text1(ptxBIKOU).Text)
        
    '�O1
    If IsNumeric(Text1(ptxNIN01).Text) Then        '�l
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(4).NIN, Format(CDbl(Text1(ptxNIN01).Text), "0.0"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(4).NIN, "0.0")
    End If

    If IsNumeric(Text1(ptxTIMES01).Text) Then      '����
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(4).TIMES, Format(CDbl(Text1(ptxTIMES01).Text), "000.00"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(4).TIMES, "000.00")
    End If
    
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(4).NIN, vbUnicode)) = 0 And _
         CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(4).TIMES, vbUnicode)) = 0 Then
    Else
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(4).NIN, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(4).TIMES, "1.0")
        End If
    
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(4).TIMES, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(4).TIMES, "001.00")
        End If
    End If
        
        
    '�O2
    If IsNumeric(Text1(ptxNIN02).Text) Then        '�l
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(5).NIN, Format(CDbl(Text1(ptxNIN02).Text), "0.0"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(5).NIN, "0.0")
    End If

    If IsNumeric(Text1(ptxTIMES02).Text) Then      '����
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(5).TIMES, Format(CDbl(Text1(ptxTIMES02).Text), "000.00"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(5).TIMES, "000.00")
    End If
    
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(5).NIN, vbUnicode)) = 0 And _
         CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(5).TIMES, vbUnicode)) = 0 Then
    Else
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(5).NIN, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(5).TIMES, "1.0")
        End If
    
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(5).TIMES, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(5).TIMES, "001.00")
        End If
    End If
        
        
    '�O3
    If IsNumeric(Text1(ptxNIN03).Text) Then        '�l
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(6).NIN, Format(CDbl(Text1(ptxNIN03).Text), "0.0"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(6).NIN, "0.0")
    End If

    If IsNumeric(Text1(ptxTIMES03).Text) Then      '����
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(6).TIMES, Format(CDbl(Text1(ptxTIMES03).Text), "000.00"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(6).TIMES, "000.00")
    End If
    
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(6).NIN, vbUnicode)) = 0 And _
         CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(6).TIMES, vbUnicode)) = 0 Then
    Else
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(6).NIN, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(6).TIMES, "1.0")
        End If
    
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(6).TIMES, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(6).TIMES, "001.00")
        End If
    End If
        
    '���1
    If IsNumeric(Text1(ptxNIN04).Text) Then        '�l
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(0).NIN, Format(CDbl(Text1(ptxNIN04).Text), "0.0"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(0).NIN, "0.0")
    End If

    If IsNumeric(Text1(ptxTIMES04).Text) Then      '����
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(0).TIMES, Format(CDbl(Text1(ptxTIMES04).Text), "000.00"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(0).TIMES, "000.00")
    End If
    
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(0).NIN, vbUnicode)) = 0 And _
         CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(0).TIMES, vbUnicode)) = 0 Then
    Else
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(0).NIN, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(0).TIMES, "1.0")
        End If
    
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(0).TIMES, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(0).TIMES, "001.00")
        End If
    End If
        
    '���2
    If IsNumeric(Text1(ptxNIN05).Text) Then        '�l
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(1).NIN, Format(CDbl(Text1(ptxNIN05).Text), "0.0"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(1).NIN, "0.0")
    End If

    If IsNumeric(Text1(ptxTIMES05).Text) Then      '����
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(1).TIMES, Format(CDbl(Text1(ptxTIMES05).Text), "000.00"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(1).TIMES, "000.00")
    End If
    
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(1).NIN, vbUnicode)) = 0 And _
         CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(1).TIMES, vbUnicode)) = 0 Then
    Else
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(1).NIN, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(1).TIMES, "1.0")
        End If
    
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(1).TIMES, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(1).TIMES, "001.00")
        End If
    End If
        
    '���3
    If IsNumeric(Text1(ptxNIN06).Text) Then        '�l
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(2).NIN, Format(CDbl(Text1(ptxNIN06).Text), "0.0"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(2).NIN, "0.0")
    End If

    If IsNumeric(Text1(ptxTIMES06).Text) Then      '����
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(2).TIMES, Format(CDbl(Text1(ptxTIMES06).Text), "000.00"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(2).TIMES, "000.00")
    End If
    
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(2).NIN, vbUnicode)) = 0 And _
         CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(2).TIMES, vbUnicode)) = 0 Then
    Else
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(2).NIN, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(2).TIMES, "1.0")
        End If
    
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(2).TIMES, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(2).TIMES, "001.00")
        End If
    End If
        
    '��1
    If IsNumeric(Text1(ptxNIN07).Text) Then        '�l
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(7).NIN, Format(CDbl(Text1(ptxNIN07).Text), "0.0"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(7).NIN, "0.0")
    End If

    If IsNumeric(Text1(ptxTIMES07).Text) Then      '����
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(7).TIMES, Format(CDbl(Text1(ptxTIMES07).Text), "000.00"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(7).TIMES, "000.00")
    End If
    
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(7).NIN, vbUnicode)) = 0 And _
         CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(7).TIMES, vbUnicode)) = 0 Then
    Else
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(7).NIN, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(7).TIMES, "1.0")
        End If
    
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(7).TIMES, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(7).TIMES, "001.00")
        End If
    End If
        
    '��2
    If IsNumeric(Text1(ptxNIN08).Text) Then        '�l
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(8).NIN, Format(CDbl(Text1(ptxNIN08).Text), "0.0"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(8).NIN, "0.0")
    End If

    If IsNumeric(Text1(ptxTIMES08).Text) Then      '����
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(8).TIMES, Format(CDbl(Text1(ptxTIMES08).Text), "000.00"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(8).TIMES, "000.00")
    End If
    
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(8).NIN, vbUnicode)) = 0 And _
         CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(8).TIMES, vbUnicode)) = 0 Then
    Else
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(8).NIN, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(8).TIMES, "1.0")
        End If
    
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(8).TIMES, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(8).TIMES, "001.00")
        End If
    End If
        
    '���H1
    If IsNumeric(Text1(ptxNIN09).Text) Then        '�l
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(9).NIN, Format(CDbl(Text1(ptxNIN09).Text), "0.0"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(9).NIN, "0.0")
    End If

    If IsNumeric(Text1(ptxTIMES09).Text) Then      '����
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(9).TIMES, Format(CDbl(Text1(ptxTIMES09).Text), "000.00"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(9).TIMES, "000.00")
    End If
    
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(9).NIN, vbUnicode)) = 0 And _
         CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(9).TIMES, vbUnicode)) = 0 Then
    Else
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(9).NIN, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(9).TIMES, "1.0")
        End If
    
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(9).TIMES, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(9).TIMES, "001.00")
        End If
    End If
        
        
        
        
        
        
        
        
        
        
        
        
        
        
    Call UniCode_Conv(P_SUKEIRE_REC.JISEKI_NAME, Combo1(pcmbJISEKI).Text)
    If IsNumeric(Text1(ptxNIN10).Text) Then        '�l
        Call UniCode_Conv(P_SUKEIRE_REC.JISEKI_NIN, Format(CDbl(Text1(ptxNIN10).Text), "0.0"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.JISEKI_NIN, "0.0")
    End If

    If IsNumeric(Text1(ptxTIMES10).Text) Then      '����
        Call UniCode_Conv(P_SUKEIRE_REC.JISEKI_TIMES, Format(CDbl(Text1(ptxTIMES10).Text), "000.00"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.JISEKI_TIMES, "000.00")
    End If
    
    
    If CDbl(StrConv(P_SUKEIRE_REC.JISEKI_NIN, vbUnicode)) = 0 And _
         CDbl(StrConv(P_SUKEIRE_REC.JISEKI_TIMES, vbUnicode)) = 0 Then
    Else
        If CDbl(StrConv(P_SUKEIRE_REC.JISEKI_NIN, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.JISEKI_NIN, "1.0")
        End If
    
        If CDbl(StrConv(P_SUKEIRE_REC.JISEKI_TIMES, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.JISEKI_TIMES, "001.00")
        End If
    End If
    
    
    
    
                                                '���Ӂ@����
    Call UniCode_Conv(P_SUKEIRE_REC.TASEKI_NAME, Combo1(pcmbTASEKI).Text)
    If IsNumeric(Text1(ptxNIN11).Text) Then        '�l
        Call UniCode_Conv(P_SUKEIRE_REC.TASEKI_NIN, Format(CDbl(Text1(ptxNIN11).Text), "0.0"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.TASEKI_NIN, "0.0")
    End If

    If IsNumeric(Text1(ptxTIMES11).Text) Then      '����
        Call UniCode_Conv(P_SUKEIRE_REC.TASEKI_TIMES, Format(CDbl(Text1(ptxTIMES11).Text), "000.00"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.TASEKI_TIMES, "000.00")
    End If
        
    If CDbl(StrConv(P_SUKEIRE_REC.TASEKI_NIN, vbUnicode)) = 0 And _
         CDbl(StrConv(P_SUKEIRE_REC.TASEKI_TIMES, vbUnicode)) = 0 Then
    Else
        If CDbl(StrConv(P_SUKEIRE_REC.TASEKI_NIN, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.TASEKI_NIN, "1.0")
        End If
    
        If CDbl(StrConv(P_SUKEIRE_REC.TASEKI_TIMES, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.TASEKI_TIMES, "001.00")
        End If
    End If
        
                                                '�����
    Call UniCode_Conv(P_SUKEIRE_REC.TORI_CODE, StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode))
            
            
            
                
            
        
    Call UniCode_Conv(P_SUKEIRE_REC.FILLER, "")
                                                        '�X�V����
    Call UniCode_Conv(P_SUKEIRE_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    
    Do
            
        sts = BTRV(BtOpUpdate, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K0_P_SUKEIRE, Len(K0_P_SUKEIRE), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_SSHIJI_O.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, BtOpInsert, "���i���������")
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


Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
        
    Select Case Index
        Case pcmbSHIMUKE        '�d������
        Case pcmbUKEHARAI       '��z��
            Text1(ptxUKEHARAI_CODE).Text = Trim(Right(Combo1(pcmbUKEHARAI).Text, 5))
    End Select
    
    Call Tab_Ctrl(Shift)        '�ړ�

End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    
    Select Case Index
        Case pcmbSHIMUKE        '�d������
        Case pcmbUKEHARAI       '��z��
            Text1(ptxUKEHARAI_CODE).Text = Trim(Right(Combo1(pcmbUKEHARAI).Text, 5))
    End Select

End Sub

Private Sub Command1_Click(Index As Integer)

Dim ans         As Integer
Dim i           As Integer


Dim sts         As Integer

    Select Case Index
        Case P_CMD_Upd        '�X�V
            
            
            
            
            
            
            
            
            
            
            
            For i = ptxUKEIRE_DT To ptxTOTAL
            
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
                
                
                
                
                If Init_Proc() Then
                    Unload Me
                End If
            
                Text1(ptxSHIJI_NO).SetFocus
            
            
            Else
                Text1(ptxUKEIRE_DT).SetFocus
            End If
        
        Case P_CMD_DSP                      '����/�\��
        
            If List_Disp_Proc() Then
                Exit Sub
            End If
        
            '��ď��̏�����
            For i = 0 To UBound(Sort_Tbl)
                Sort_Tbl(i) = 0             '��̫�ď���
            Next i
        
            Sort_Tbl(colHIN_NAME) = 9       '��ď��O
        
        Case P_CMD_OUT                      '�ް��o��
        Case P_CMD_PRT                      '���
            
            
            
            
        Case P_CMD_End                      '�I��
            Unload Me
    End Select

End Sub


Private Sub Form_DblClick()
'    PrintForm                  '2017.07.22
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

Dim c               As String * 128
Dim sts             As Integer
Dim i               As Integer



'    If App.PrevInstance Then                       '2017.07.22
'        Beep                                       '2017.07.22
'        MsgBox "����v���O�������s���ł��B"        '2017.07.22
'        End                                        '2017.07.22
'    End If                                         '2017.07.22
                                '����
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "JISEKI", "P_SYS", c) Then            '2017.07.22
    If GetIni(StrConv(App.EXEName, vbUpperCase), "JISEKI", App.EXEName, c) Then         '2017.07.22
        JISEKI_TITLE = ""
    Else
        JISEKI_TITLE = Split(Trim(c), ",", -1)
    End If
    
                                '����
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "TASEKI", "P_SYS", c) Then            '2017.07.22
    If GetIni(StrConv(App.EXEName, vbUpperCase), "TASEKI", App.EXEName, c) Then         '2017.07.22
        TASEKI_TITLE = ""
    Else
        TASEKI_TITLE = Split(Trim(c), ",", -1)
    End If

                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
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
                                '�󕥐�}�X�^�n�o�d�m
    If P_UKEHARAI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���i���w�}��������ް��n�o�d�m
    If P_SUKEIRE_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    '����Ͻ���`
    Call P_CODE_TBL_Proc
    
    
    PI000251.Caption = PI000251.Caption & LAST_UPDATE_DAY           '2017.07.21
    
    
                                '���P�^�S���҂̎�荞��
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "S_TANTO", "P_SYS", c) Then                               '2017.07.22
    If GetIni(StrConv(App.EXEName, vbUpperCase), "S_TANTO", StrConv(App.EXEName, vbUpperCase), c) Then      '2017.07.22
        PRI_S_TANTO = False
    Else
        If RTrim(c) = "0" Then
            PRI_S_TANTO = False
        Else
            PRI_S_TANTO = True
        End If
    End If
                                
    Label1(56).Visible = PRI_S_TANTO
    Combo1(pcmbS_TANTO).Visible = PRI_S_TANTO
    
    TDBGrid1.Columns(colS_TANTO).Visible = PRI_S_TANTO
    
    
    
    
    
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
    If Code_Set_Proc(pcmbS_TANTO, P_KBN05_CD, 1) Then
        Unload Me
    End If
    
    '�󕥐�
    If Ukeharai_Set_Proc() Then
        Unload Me
    End If
    
                                '����
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "JISEKI", "P_SYS", c) Then                            '2017.07.22
    If GetIni(StrConv(App.EXEName, vbUpperCase), "JISEKI", StrConv(App.EXEName, vbUpperCase), c) Then   '2017.07.22
        JISEKI_TITLE = ""
    Else
        JISEKI_TITLE = Split(Trim(c), ",", -1)
    End If
    
    
                                '����
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "TASEKI", "P_SYS", c) Then                            '2017.07.22
    If GetIni(StrConv(App.EXEName, vbUpperCase), "TASEKI", StrConv(App.EXEName, vbUpperCase), c) Then   '2017.07.22
        TASEKI_TITLE = ""
    Else
        TASEKI_TITLE = Split(Trim(c), ",", -1)
    End If
    
    '�b�^����荞�� 2008.08.19
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "JISSEKI_DSP", "P_SYS", c) Then                           '2017.07.22
    If GetIni(StrConv(App.EXEName, vbUpperCase), "JISSEKI_DSP", StrConv(App.EXEName, vbUpperCase), c) Then  '2017.07.22
        JISSEKI_DSP = "m"
    Else
        JISSEKI_DSP = Trim(c)
        If JISSEKI_DSP <> "m" And JISSEKI_DSP <> "s" Then
            JISSEKI_DSP = "m"
        End If
    End If
    
    
    
'�\������   2017.08.05
    If GetIni(StrConv(App.EXEName, vbUpperCase), "LOAD_LIMIT", StrConv(App.EXEName, vbUpperCase), c) Then
        txtLOAD_LIMIT.Text = "0"
    Else
        txtLOAD_LIMIT.Text = Val(Trim(c))
    End If
    
    
    '��ʏ����ݒ�
    If Init_Proc() Then
        Unload Me
    End If



    If Val(txtLOAD_LIMIT.Text) > 0 Then     '2017.08.05

        If SSHIJI.UpperBound(1) > 0 Then
            'SSHIJI.QuickSort Min_Row, SSHIJI.UpperBound(1), 0, 1, XTYPE_STRING
            SSHIJI.QuickSort Min_Row, SSHIJI.UpperBound(1), 1, 1, XTYPE_STRING      '�~���Ƃ��� 2017.09.27
            
            Set TDBGrid1.Array = SSHIJI
            
            TDBGrid1.ReBind
            TDBGrid1.Update
            TDBGrid1.MoveFirst
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
                                            
                                            
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
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
    
                                            '�󕥐�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�󕥐�}�X�^")
        End If
    End If
                                            '��������b�k�n�r�d
    sts = BTRV(BtOpClose, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K0_P_SUKEIRE, Len(K0_P_SUKEIRE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�������")
        End If
    End If
    
    sts = BTRV(BtOpReset, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    
    sts = WriteIni(App.EXEName, "LOAD_LIMIT", App.EXEName, txtLOAD_LIMIT.Text)
    If sts Then
        MsgBox "�u�\�������v�������݃G���[ PI00025.INI"
    End If
    
    
    Set PI000251 = Nothing

    End
End Sub





Private Sub TDBGrid1_DblClick()
Dim sts As Integer
    
    Text1(ptxSHIJI_NO).Text = SSHIJI(TDBGrid1.Bookmark, colSHIJI_NO)
    '�w�}�[�ް�������
    sts = P_SSHIJI_Read_Proc()
    Select Case sts
        Case False, BtNoErr
                    
            If StrConv(P_SSHIJI_O_REC.KAN_F, vbUnicode) <> P_KAN_ON Then
                MsgBox "�������ł��B"
                TDBGrid1.SetFocus
                Exit Sub
            End If
        
            Call UniCode_Conv(P_SSHIJI_O_REC.SHIJI_No, "")
        
        Case BtErrKeyNotFound
            MsgBox "���͂������ڂ̓G���[�ł��B"
            TDBGrid1.SetFocus
            Exit Sub
        Case Else
            Exit Sub
    End Select
    
        
    Call UniCode_Conv(K0_P_SUKEIRE.SHIJI_No, Text1(ptxSHIJI_NO).Text)
    Call UniCode_Conv(K0_P_SUKEIRE.SEQNO, "001")
    Text1(ptxSEQNO).Text = "001"
    
    sts = P_SUKEIRE_Read_Proc()
    Select Case sts
        Case False, BtNoErr
        
        Case BtErrKeyNotFound
            MsgBox "���͂������ڂ̓G���[�ł��B"
            TDBGrid1.SetFocus
            Exit Sub
        Case Else
            Exit Sub
    End Select
    
    

End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)


    If Sort_Tbl(ColIndex) = 0 Then
        Sort_Tbl(ColIndex) = 1
    Else
        If Sort_Tbl(ColIndex) = 1 Then
            Sort_Tbl(ColIndex) = 0
        End If
    
    End If
    
    If Sort_Tbl(ColIndex) = 0 Or Sort_Tbl(ColIndex) = 1 Then
                    
        SSHIJI.QuickSort Min_Row, SSHIJI.UpperBound(1), ColIndex, Sort_Tbl(ColIndex), XTYPE_STRING
        
        Set TDBGrid1.Array = SSHIJI
        
        TDBGrid1.ReBind
        TDBGrid1.Update
        TDBGrid1.MoveFirst


    End If



End Sub



Private Sub Text1_Change(Index As Integer)
        Select Case Index
        
        
            Case ptxTIMES01, ptxTIMES02, ptxTIMES03, ptxTIMES04, ptxTIMES05, ptxTIMES06, ptxTIMES07, ptxTIMES08, ptxTIMES09, ptxTIMES10, ptxTIMES11 '����
        
        
                Time_Input_F = True
        
        End Select

End Sub

Private Sub Text1_GotFocus(Index As Integer)
    
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    
    
    
    
    
    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
        
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2017.07.22
    Select Case Index
        Case ptxUKEHARAI_CODE
            Text1(Index).Text = StrConv(Text1(Index).Text, vbUpperCase)
    
    
    

        
        Case ptxS_CLASS_CODE To ptxN_CLASS_CODE                             '2017.08.05
            Text1(Index).Text = StrConv(Text1(Index).Text, vbUpperCase)     '2017.08.05
    
    
    End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2017.07.22
        
        
        
        
    If JISSEKI_DSP = "s" Then   '2008.08.19

        If Time_Input_F Then

            Select Case Index
            
            
                Case ptxTIMES01, ptxTIMES02, ptxTIMES03, ptxTIMES04, ptxTIMES05, ptxTIMES06, ptxTIMES07, ptxTIMES08, ptxTIMES09, ptxTIMES10, ptxTIMES11 '����
            
            
                    If IsNumeric(Text1(Index).Text) Then
                        Text1(Index).Text = Format(ToHalfAdjust(CCur(CInt(Text1(Index).Text) / 60), 0), "#0")
                    End If
            
                    Time_Input_F = False
            
            
            End Select
        
    
        End If
    
    End If
        
        
        
        
    If Error_Check_Proc(Index) Then     '�G���[�`�F�b�N
        Exit Sub
    End If
        
    If JISSEKI_DSP = "s" Then   '2008.08.19

        If Time_Input_F Then

            Select Case Index
            
            
                Case ptxTIMES01, ptxTIMES02, ptxTIMES03, ptxTIMES04, ptxTIMES05, ptxTIMES06, ptxTIMES07, ptxTIMES08, ptxTIMES09, ptxTIMES10, ptxTIMES11 '����
            
            
                    Time_Input_F = False
            
            End Select
        
    
        End If
    
    End If
        
        
        
    Call Tab_Ctrl(Shift)        '�ړ�
End Sub
Private Function Init_Proc() As Integer
'----------------------------------------------------------------------------
'                   ���͉�ʂ̏����ݒ�
'----------------------------------------------------------------------------
Dim i       As Integer
Dim sts     As Integer


    Init_Proc = True
    
    
    For i = ptxSHIJI_NO To ptxTOTAL
        
        If i = 5 Or i = 7 Or i = 10 Then
        Else
            Text1(i).Text = ""
        End If
    
    Next i
    
    '���l 2010.09.28
    Text1(ptxBIKOU).Text = ""
    
    For i = ptx2TANTO_NAME To ptx2N_CLASS_NAME
        Text2(i).Text = ""
    Next i
    
    
    For i = pcmbSHIMUKE To pcmbS_TANTO
        
        Combo1(i).ListIndex = -1
    
    Next i

    If JISSEKI_DSP = "s" Then           '2008.08.19
    
        Label1(24).Caption = "�b"
    
    Else
    
        Label1(24).Caption = "��"
    
    
    End If

    '����
    
    Combo1(pcmbJISEKI).Clear
    Combo1(pcmbJISEKI).AddItem CStr(JISEKI_TITLE(0))
    Combo1(pcmbJISEKI).AddItem CStr(JISEKI_TITLE(1))
    Combo1(pcmbJISEKI).ListIndex = -1
    '����
    Combo1(pcmbTASEKI).Clear
    Combo1(pcmbTASEKI).AddItem CStr(TASEKI_TITLE(0))
    Combo1(pcmbTASEKI).AddItem CStr(TASEKI_TITLE(1))
    Combo1(pcmbTASEKI).ListIndex = -1


    If List_Disp_Proc() Then
        Exit Function
    End If

    '�ǂݍ����ޯ̧��ر�
    Call UniCode_Conv(P_SSHIJI_O_REC.SHIJI_No, "")

    '��ď��̏�����
    For i = 0 To UBound(Sort_Tbl)
        Sort_Tbl(i) = 0             '��̫�ď���
    Next i

    Sort_Tbl(colHIN_NAME) = 9       '��ď��O

    
    Set TDBGrid1.Array = SSHIJI
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst



    Call UniCode_Conv(P_SSHIJI_O_REC.SHIJI_No, "")
    
    
    Time_Input_F = False
    
    Init_Proc = False

End Function
Private Function Code_Set_Proc(Index As Integer, KBN As String, Mode As Integer) As Integer
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



Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'           �w�}�[�ް��̕\��
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim com     As Integer

Dim Row     As Long

    List_Disp_Proc = True
    PI000251.MousePointer = vbHourglass
    
    
    
    Set SSHIJI = Nothing
    
    
    
    com = BtOpGetLast

    Row = Min_Row - 1
       
    Do
    
        DoEvents
    
'        sts = BTRV(com, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)     '2019.04.05
        sts = BTRV(com, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K3_P_SSHIJI_O, Len(K3_P_SSHIJI_O), 3)      '2019.04.05
            
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���i���w�}�ް�(�e)")
                Exit Function
        End Select
    
    
        If StrConv(P_SSHIJI_O_REC.KAN_F, vbUnicode) = P_KAN_ON Then
            Row = Row + 1
            
'100
'            If Row > 500 Then                          '2017.08.05
            If Row > Val(txtLOAD_LIMIT.Text) Then       '2017.08.05
                Exit Do
            End If
            If Grid_Set_Proc(Row) Then
                Exit Function
            End If
        End If
        
        com = BtOpGetPrev
    
    Loop
    
    Set TDBGrid1.Array = SSHIJI
            
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    
    
    PI000251.MousePointer = vbDefault
    List_Disp_Proc = False
    


End Function

Private Function Grid_Set_Proc(Row As Long) As Integer
'----------------------------------------------------------------------------
'           �w�}�[�ް��̓��e���د�ނɾ�Ă���
'----------------------------------------------------------------------------
Dim sts As Integer

    Grid_Set_Proc = True
    
    SSHIJI.ReDim Min_Row, Row, Min_Col, Max_Col


    '���s����
    If Trim(StrConv(P_SSHIJI_O_REC.Print_datetime, vbUnicode)) = "" Then
        SSHIJI(Row, colHAKKO_DT) = Mid(StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 7, 2)
    Else
        SSHIJI(Row, colHAKKO_DT) = Mid(StrConv(P_SSHIJI_O_REC.Print_datetime, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(P_SSHIJI_O_REC.Print_datetime, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(P_SSHIJI_O_REC.Print_datetime, vbUnicode), 7, 2) & " " & _
                                    Mid(StrConv(P_SSHIJI_O_REC.Print_datetime, vbUnicode), 9, 2) & ":" & _
                                    Mid(StrConv(P_SSHIJI_O_REC.Print_datetime, vbUnicode), 11, 2)
    End If
    '�w�}�[��
    SSHIJI(Row, colSHIJI_NO) = StrConv(P_SSHIJI_O_REC.SHIJI_No, vbUnicode)
    '�d������
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_CODEREC.C_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�R�[�h�}�X�^")
            Exit Function
    End Select
    '�d������
    SSHIJI(Row, colSHIMUKE) = StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode) & " " & _
                                StrConv(P_CODEREC.C_RNAME, vbUnicode)
        
    '��z��
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�󕥐�}�X�^")
            Exit Function
    End Select
    '��z��
    SSHIJI(Row, colUKEHARAI) = StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode) & " " & _
                                StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode)
    '�i��
    SSHIJI(Row, colHIN_GAI) = StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode)
    '�i��
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Function
    End Select
    '�i��
    SSHIJI(Row, colHIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
    '�\�萔
    SSHIJI(Row, colSHIJI_QTY) = Format(CLng(StrConv(P_SSHIJI_O_REC.SHIJI_QTY, vbUnicode)), "#0")
    '������
    SSHIJI(Row, colUKEIRE_QTY) = Format(CLng(StrConv(P_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)), "#0")
        
    '���P�^�S����
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN05_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SSHIJI_O_REC.S_TANTO, vbUnicode))
    
    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_CODEREC.C_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�R�[�h�}�X�^")
            Exit Function
    End Select
    
    
    
    SSHIJI(Row, colS_TANTO) = Trim(StrConv(P_SSHIJI_O_REC.S_TANTO, vbUnicode)) & " " & Trim(StrConv(P_CODEREC.C_NAME, vbUnicode))

    Grid_Set_Proc = False

End Function

Private Function P_SUKEIRE_Read_Proc() As Integer
'----------------------------------------------------------------------------
'                   ���i�����ю�������ް��̓ǂݍ���
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
    
    
    P_SUKEIRE_Read_Proc = True
    
    
    sts = BTRV(BtOpGetEqual, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K0_P_SUKEIRE, Len(K0_P_SUKEIRE), 0)
        
    Select Case sts
        Case BtNoErr
        
        
        
        
        Case Else
            P_SUKEIRE_Read_Proc = sts
            Exit Function
    
    End Select
    
    
    If JISSEKI_Disp_Proc() Then
        Exit Function
    End If
    
    P_SUKEIRE_Read_Proc = False
        
    

End Function


Private Function P_SUKEIRE_Read_only_Proc() As Integer
'----------------------------------------------------------------------------
'                   ���i�����ю�������ް��̓ǂݍ��� 2018.03.03
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
    
    
    P_SUKEIRE_Read_only_Proc = True
    
    
    sts = BTRV(BtOpGetEqual, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K0_P_SUKEIRE, Len(K0_P_SUKEIRE), 0)
        
    Select Case sts
        Case BtNoErr
        
        
        
        
        Case Else
            P_SUKEIRE_Read_only_Proc = sts
            Exit Function
    
    End Select
    
    
    
    P_SUKEIRE_Read_only_Proc = False
        
    

End Function


Private Function JISSEKI_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   ������ѕ\��
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim i       As Integer
Dim TOTAL   As Double
    JISSEKI_Disp_Proc = True
    
    
    Text1(ptxUKEIRE_DT).Text = Mid(StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode), 7, 2)  '�����
    
    
    
    Text1(ptxUKEIRE_QTY).Text = Format(StrConv(P_SUKEIRE_REC.UKEIRE_QTY, vbUnicode), "#0")
            
        
    '���l 2010.09.28
    Text1(ptxBIKOU).Text = Trim(StrConv(P_SUKEIRE_REC.BIKOU, vbUnicode))
        
        
    For i = 0 To 9
        If Not IsNumeric(StrConv(P_SUKEIRE_REC.GENKA_TBL(i).NIN, vbUnicode)) Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(i).NIN, "0.0")
        End If
    
        If Not IsNumeric(StrConv(P_SUKEIRE_REC.GENKA_TBL(i).TIMES, vbUnicode)) Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(i).TIMES, "000.00")
        End If
    
    Next i
    
    
    '�O�P
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(4).NIN, vbUnicode)) = 0 Then
        Text1(ptxNIN01).Text = ""
    Else
        Text1(ptxNIN01).Text = Format(CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(4).NIN, vbUnicode)), "#0.0")
    End If
    
    
    If Not IsNumeric(StrConv(P_SUKEIRE_REC.GENKA_TBL(4).NIN, vbUnicode)) Then
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(4).TIMES, "0.0")
    End If
    
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(4).TIMES, vbUnicode)) = 0 Then
        Text1(ptxTIMES01).Text = ""
    Else
        Text1(ptxTIMES01).Text = Format(CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(4).TIMES, vbUnicode)), "#0.00")
    End If
    
    If IsNumeric(Text1(ptxNIN01).Text) And IsNumeric(Text1(ptxTIMES01).Text) Then
        TOTAL = CDbl(Text1(ptxNIN01).Text) * CDbl(Text1(ptxTIMES01).Text)
        If TOTAL = 0 Then
            Text1(ptxTOTAL01).Text = ""
        Else
            Text1(ptxTOTAL01).Text = Format(TOTAL, "#0.00")
        End If
    Else
        Text1(ptxTOTAL01).Text = ""
    End If
    '�O�Q
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(5).NIN, vbUnicode)) = 0 Then
        Text1(ptxNIN02).Text = ""
    Else
        Text1(ptxNIN02).Text = Format(CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(5).NIN, vbUnicode)), "#0.0")
    End If
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(5).TIMES, vbUnicode)) = 0 Then
        Text1(ptxTIMES02).Text = ""
    Else
        Text1(ptxTIMES02).Text = Format(CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(5).TIMES, vbUnicode)), "#0.00")
    End If
    
    
    If IsNumeric(Text1(ptxNIN02).Text) And IsNumeric(Text1(ptxTIMES02).Text) Then
        TOTAL = CDbl(Text1(ptxNIN02).Text) * CDbl(Text1(ptxTIMES02).Text)
        If TOTAL = 0 Then
            Text1(ptxTOTAL02).Text = ""
        Else
            Text1(ptxTOTAL02).Text = Format(TOTAL, "#0.00")
        End If
    Else
        Text1(ptxTOTAL02).Text = ""
    End If
    '�O�R
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(6).NIN, vbUnicode)) = 0 Then
        Text1(ptxNIN03).Text = ""
    Else
        Text1(ptxNIN03).Text = Format(CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(6).NIN, vbUnicode)), "#0.0")
    End If
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(5).TIMES, vbUnicode)) = 0 Then
        Text1(ptxTIMES03).Text = ""
    Else
        Text1(ptxTIMES03).Text = Format(CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(6).TIMES, vbUnicode)), "#0.00")
    End If
    If IsNumeric(Text1(ptxNIN03).Text) And IsNumeric(Text1(ptxTIMES03).Text) Then
        TOTAL = CDbl(Text1(ptxNIN03).Text) * CDbl(Text1(ptxTIMES03).Text)
        If TOTAL = 0 Then
            Text1(ptxTOTAL03).Text = ""
        Else
            Text1(ptxTOTAL03).Text = Format(TOTAL, "#0.00")
        End If
    Else
        Text1(ptxTOTAL03).Text = ""
    End If
    '��ƂP
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(0).NIN, vbUnicode)) = 0 Then
        Text1(ptxNIN04).Text = ""
    Else
        Text1(ptxNIN04).Text = Format(CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(0).NIN, vbUnicode)), "#0.0")
    End If
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(0).TIMES, vbUnicode)) = 0 Then
        Text1(ptxTIMES04).Text = ""
    Else
        Text1(ptxTIMES04).Text = Format(CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(0).TIMES, vbUnicode)), "#0.00")
    End If
    
    If IsNumeric(Text1(ptxNIN04).Text) And IsNumeric(Text1(ptxTIMES04).Text) Then
        TOTAL = CDbl(Text1(ptxNIN04).Text) * CDbl(Text1(ptxTIMES04).Text)
        If TOTAL = 0 Then
            Text1(ptxTOTAL04).Text = ""
        Else
            Text1(ptxTOTAL04).Text = Format(TOTAL, "#0.00")
        End If
    Else
        Text1(ptxTOTAL04).Text = ""
    End If
    '���2
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(1).NIN, vbUnicode)) = 0 Then
        Text1(ptxNIN05).Text = ""
    Else
        Text1(ptxNIN05).Text = Format(CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(1).NIN, vbUnicode)), "#0.0")
    End If
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(1).TIMES, vbUnicode)) = 0 Then
        Text1(ptxTIMES05).Text = ""
    Else
        Text1(ptxTIMES05).Text = Format(CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(1).TIMES, vbUnicode)), "#0.00")
    End If
    
    If IsNumeric(Text1(ptxNIN05).Text) And IsNumeric(Text1(ptxTIMES05).Text) Then
        TOTAL = CDbl(Text1(ptxNIN05).Text) * CDbl(Text1(ptxTIMES05).Text)
        If TOTAL = 0 Then
            Text1(ptxTOTAL05).Text = ""
        Else
            Text1(ptxTOTAL05).Text = Format(TOTAL, "#0.00")
        End If
    Else
        Text1(ptxTOTAL05).Text = ""
    End If
    
    '���3
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(2).NIN, vbUnicode)) = 0 Then
        Text1(ptxNIN06).Text = ""
    Else
        Text1(ptxNIN06).Text = Format(CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(2).NIN, vbUnicode)), "#0.0")
    End If
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(2).TIMES, vbUnicode)) = 0 Then
        Text1(ptxTIMES06).Text = ""
    Else
        Text1(ptxTIMES06).Text = Format(CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(2).TIMES, vbUnicode)), "#0.00")
    End If
    
    If IsNumeric(Text1(ptxNIN06).Text) And IsNumeric(Text1(ptxTIMES06).Text) Then
        TOTAL = CDbl(Text1(ptxNIN06).Text) * CDbl(Text1(ptxTIMES06).Text)
        If TOTAL = 0 Then
            Text1(ptxTOTAL06).Text = ""
        Else
            Text1(ptxTOTAL06).Text = Format(TOTAL, "#0.00")
        End If
    Else
        Text1(ptxTOTAL06).Text = ""
    End If
    
    '��1
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(7).NIN, vbUnicode)) = 0 Then
        Text1(ptxNIN07).Text = ""
    Else
        Text1(ptxNIN07).Text = Format(CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(7).NIN, vbUnicode)), "#0.0")
    End If
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(7).TIMES, vbUnicode)) = 0 Then
        Text1(ptxTIMES07).Text = ""
    Else
        Text1(ptxTIMES07).Text = Format(CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(7).TIMES, vbUnicode)), "#0.00")
    End If
    
    If IsNumeric(Text1(ptxNIN07).Text) And IsNumeric(Text1(ptxTIMES07).Text) Then
        TOTAL = CDbl(Text1(ptxNIN07).Text) * CDbl(Text1(ptxTIMES07).Text)
        If TOTAL = 0 Then
            Text1(ptxTOTAL07).Text = ""
        Else
            Text1(ptxTOTAL07).Text = Format(TOTAL, "#0.00")
        End If
    Else
        Text1(ptxTOTAL07).Text = ""
    End If
    
    '��2
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(8).NIN, vbUnicode)) = 0 Then
        Text1(ptxNIN08).Text = ""
    Else
        Text1(ptxNIN08).Text = Format(CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(8).NIN, vbUnicode)), "#0.0")
    End If
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(8).TIMES, vbUnicode)) = 0 Then
        Text1(ptxTIMES08).Text = ""
    Else
        Text1(ptxTIMES08).Text = Format(CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(8).TIMES, vbUnicode)), "#0.00")
    End If
    
    If IsNumeric(Text1(ptxNIN08).Text) And IsNumeric(Text1(ptxTIMES08).Text) Then
        TOTAL = CDbl(Text1(ptxNIN08).Text) * CDbl(Text1(ptxTIMES08).Text)
        If TOTAL = 0 Then
            Text1(ptxTOTAL08).Text = ""
        Else
            Text1(ptxTOTAL08).Text = Format(TOTAL, "#0.00")
        End If
    Else
        Text1(ptxTOTAL08).Text = ""
    End If
    
    '���H1
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(9).NIN, vbUnicode)) = 0 Then
        Text1(ptxNIN09).Text = ""
    Else
        Text1(ptxNIN09).Text = Format(CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(9).NIN, vbUnicode)), "#0.0")
    End If
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(9).TIMES, vbUnicode)) = 0 Then
        Text1(ptxTIMES09).Text = ""
    Else
        Text1(ptxTIMES09).Text = Format(CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(9).TIMES, vbUnicode)), "#0.00")
    End If
    If IsNumeric(Text1(ptxNIN09).Text) And IsNumeric(Text1(ptxTIMES09).Text) Then
        TOTAL = CDbl(Text1(ptxNIN09).Text) * CDbl(Text1(ptxTIMES09).Text)
        If TOTAL = 0 Then
            Text1(ptxTOTAL09).Text = ""
        Else
            Text1(ptxTOTAL09).Text = Format(TOTAL, "#0.00")
        End If
    Else
        Text1(ptxTOTAL09).Text = ""
    End If
    '����
    
    Combo1(pcmbJISEKI).ListIndex = -1
    For i = 0 To Combo1(pcmbJISEKI).ListCount - 1
        If Trim(StrConv(P_SUKEIRE_REC.JISEKI_NAME, vbUnicode)) = Trim(Combo1(pcmbJISEKI).List(i)) Then
            Combo1(pcmbJISEKI).ListIndex = i
            Exit For
        End If
    
    Next i
    
    If i > Combo1(pcmbJISEKI).ListCount - 1 Then
        Combo1(pcmbJISEKI).Text = Trim(StrConv(P_SUKEIRE_REC.JISEKI_NAME, vbUnicode))
    End If
    
    
    
    If CDbl(StrConv(P_SUKEIRE_REC.JISEKI_NIN, vbUnicode)) = 0 Then
        Text1(ptxNIN10).Text = ""
    Else
        Text1(ptxNIN10).Text = Format(CDbl(StrConv(P_SUKEIRE_REC.JISEKI_NIN, vbUnicode)), "#0.0")
    End If
    If CDbl(StrConv(P_SUKEIRE_REC.JISEKI_TIMES, vbUnicode)) = 0 Then
        Text1(ptxTIMES10).Text = ""
    Else
        Text1(ptxTIMES10).Text = Format(CDbl(StrConv(P_SUKEIRE_REC.JISEKI_TIMES, vbUnicode)), "#0.00")
    End If
    
    If IsNumeric(Text1(ptxNIN10).Text) And IsNumeric(Text1(ptxTIMES10).Text) Then
        TOTAL = CDbl(Text1(ptxNIN10).Text) * CDbl(Text1(ptxTIMES10).Text)
        If TOTAL = 0 Then
            Text1(ptxTOTAL10).Text = ""
        Else
            Text1(ptxTOTAL10).Text = Format(TOTAL, "#0.00")
        
        End If
    Else
        Text1(ptxTOTAL10).Text = ""
    
    End If
    
    '����
    Combo1(pcmbTASEKI).ListIndex = -1
    For i = 0 To Combo1(pcmbTASEKI).ListCount - 1
        If Trim(StrConv(P_SUKEIRE_REC.TASEKI_NAME, vbUnicode)) = Trim(Combo1(pcmbTASEKI).List(i)) Then
            Combo1(pcmbTASEKI).ListIndex = i
            Exit For
        End If
    
    Next i
    
    If i > Combo1(pcmbTASEKI).ListCount - 1 Then
        Combo1(pcmbTASEKI).Text = Trim(StrConv(P_SUKEIRE_REC.TASEKI_NAME, vbUnicode))
    End If
    
    
    If CDbl(StrConv(P_SUKEIRE_REC.TASEKI_NIN, vbUnicode)) = 0 Then
        Text1(ptxNIN11).Text = ""
    Else
        Text1(ptxNIN11).Text = Format(CDbl(StrConv(P_SUKEIRE_REC.TASEKI_NIN, vbUnicode)), "#0.0")
    End If
    If CDbl(StrConv(P_SUKEIRE_REC.TASEKI_TIMES, vbUnicode)) = 0 Then
        Text1(ptxTIMES11).Text = ""
    Else
        Text1(ptxTIMES11).Text = Format(CDbl(StrConv(P_SUKEIRE_REC.TASEKI_TIMES, vbUnicode)), "#0.00")
    End If
    If IsNumeric(Text1(ptxNIN11).Text) And IsNumeric(Text1(ptxTIMES11).Text) Then
        TOTAL = CDbl(Text1(ptxNIN11).Text) * CDbl(Text1(ptxTIMES11).Text)
        If TOTAL = 0 Then
            Text1(ptxTOTAL11).Text = ""
        Else
            Text1(ptxTOTAL11).Text = Format(TOTAL, "#0.00")
        End If
    Else
        Text1(ptxTOTAL11).Text = ""
    End If
    
    TOTAL = 0
    For i = ptxNIN01 To ptxNIN11 Step 3
        
        If IsNumeric(Text1(i + 2).Text) Then
        
            TOTAL = TOTAL + CDbl(Text1(i + 2).Text)
        End If
    
    Next i
    Text1(ptxTOTAL).Text = Format(TOTAL, "#0.00")
    
    
    Time_Input_F = False
    
    JISSEKI_Disp_Proc = False

End Function


Private Sub Text1_LostFocus(Index As Integer)
    
Dim sts     As Integer
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2017.07.22
    Select Case Index
        
        
'>      2018.03.30
        Case ptxSHIJI_NO
            
            
            If Trim(Text1(Index).Text) = "" Then    '2019.04.05
                Exit Sub                            '2019.04.05
            End If                                  '2019.04.05
            
            If Trim(Text1(ptxSEQNO).Text) = "" Then
                Text1(ptxSEQNO).Text = "001"
            End If
            
            If Text1(ptxSHIJI_NO).Text = StrConv(P_SUKEIRE_REC.SHIJI_No, vbUnicode) And _
                Text1(ptxSEQNO).Text = StrConv(P_SUKEIRE_REC.SEQNO, vbUnicode) Then
                
            Else
                Call UniCode_Conv(K0_P_SUKEIRE.SHIJI_No, Text1(ptxSHIJI_NO).Text)
                Call UniCode_Conv(K0_P_SUKEIRE.SEQNO, Text1(ptxSEQNO).Text)
                sts = P_SUKEIRE_Read_Proc()
                Select Case sts
                    Case False, BtNoErr
                                
                    
                    
                    Case BtErrKeyNotFound
                        MsgBox "���͂����w�}�[��(�ǔ�)�͓o�^����Ă��܂���B"
                        Text1(ptxSEQNO).SetFocus
                        Exit Sub
                    Case Else
                        Unload Me
                End Select
            End If
'>      2018.03.30
        
        
        
        
        
        Case ptxUKEHARAI_CODE
            Text1(Index).Text = StrConv(Text1(Index).Text, vbUpperCase)
    
    
    

        
        Case ptxS_CLASS_CODE To ptxN_CLASS_CODE                             '2017.08.05
            Text1(Index).Text = StrConv(Text1(Index).Text, vbUpperCase)     '2017.08.05
    
    
    End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2017.07.22
    
    
    
    
    
    If JISSEKI_DSP <> "s" Then
        Exit Sub
    End If

    If Time_Input_F Then

        Select Case Index
        
        
            Case ptxTIMES01, ptxTIMES02, ptxTIMES03, ptxTIMES04, ptxTIMES05, ptxTIMES06, ptxTIMES07, ptxTIMES08, ptxTIMES09, ptxTIMES10, ptxTIMES11 '����
        
        
                If IsNumeric(Text1(Index).Text) Then
                    Text1(Index).Text = Format(ToHalfAdjust(CCur(CInt(Text1(Index).Text) / 60), 0), "#0")
                End If
            
        End Select
        
        Time_Input_F = False
        
    End If




End Sub
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


Private Sub txtLOAD_LIMIT_KeyDown(KeyCode As Integer, Shift As Integer)
'2017.08.05
Dim sts As Integer
    
    If KeyCode <> vbKeyReturn Then Exit Sub


    If Not IsNumeric(txtLOAD_LIMIT) Then
    
        MsgBox "�\�������͐��l�œ��͂��ĉ�����"
        txtLOAD_LIMIT.SetFocus
        Exit Sub
    
    End If

    sts = WriteIni(App.EXEName, "LOAD_LIMIT", App.EXEName, txtLOAD_LIMIT.Text)
    If sts Then
        MsgBox "�u�\�������v�������݃G���[ PI00025.INI"
    End If


End Sub
