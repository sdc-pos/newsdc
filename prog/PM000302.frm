VERSION 5.00
Begin VB.Form PM000302 
   Caption         =   "���ރ}�X�^�����e�i���X"
   ClientHeight    =   11205
   ClientLeft      =   1920
   ClientTop       =   2430
   ClientWidth     =   16320
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
   ScaleHeight     =   11205
   ScaleWidth      =   16320
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   14880
      ScaleHeight     =   315
      ScaleWidth      =   195
      TabIndex        =   169
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��ʈ��"
      Height          =   495
      Left            =   14760
      TabIndex        =   168
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   63
      Left            =   11280
      TabIndex        =   16
      Top             =   1680
      Width           =   4935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "�݌ɐ���\��(�I���\)"
      Height          =   255
      Index           =   3
      Left            =   7920
      TabIndex        =   15
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   62
      Left            =   5130
      MaxLength       =   1
      TabIndex        =   36
      Top             =   3900
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   61
      Left            =   12600
      MaxLength       =   4
      TabIndex        =   24
      Top             =   2520
      Width           =   645
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   60
      Left            =   3150
      MaxLength       =   2
      TabIndex        =   35
      Top             =   3900
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   59
      Left            =   2520
      MaxLength       =   2
      TabIndex        =   34
      Top             =   3900
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   58
      Left            =   1890
      MaxLength       =   2
      TabIndex        =   33
      Top             =   3900
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   57
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   32
      Top             =   3900
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   56
      Left            =   7560
      MaxLength       =   2
      TabIndex        =   31
      Top             =   3480
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      Index           =   55
      Left            =   7560
      MaxLength       =   9
      TabIndex        =   27
      Top             =   3000
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   23
      Left            =   5175
      MaxLength       =   8
      TabIndex        =   30
      Top             =   3480
      Width           =   510
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   22
      Left            =   1320
      MaxLength       =   20
      TabIndex        =   29
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   15
      Left            =   8505
      MaxLength       =   10
      TabIndex        =   19
      Top             =   2160
      Width           =   1365
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   14
      Left            =   4635
      MaxLength       =   20
      TabIndex        =   18
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   13
      Left            =   1320
      MaxLength       =   20
      TabIndex        =   17
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   21
      Left            =   11250
      MaxLength       =   2
      TabIndex        =   28
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      Index           =   20
      Left            =   4635
      MaxLength       =   9
      TabIndex        =   26
      Top             =   3000
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   19
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   25
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   18
      Left            =   4680
      MaxLength       =   8
      TabIndex        =   22
      Top             =   2580
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   17
      Left            =   3285
      MaxLength       =   8
      TabIndex        =   21
      Top             =   2580
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   16
      Left            =   1935
      MaxLength       =   8
      TabIndex        =   20
      Top             =   2580
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   12
      Left            =   10620
      MaxLength       =   5
      TabIndex        =   23
      Top             =   2520
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   54
      Left            =   14025
      MaxLength       =   11
      TabIndex        =   71
      Top             =   8820
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   8
      ItemData        =   "PM000302.frx":0000
      Left            =   8190
      List            =   "PM000302.frx":0002
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   70
      Top             =   8820
      Width           =   4110
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   53
      Left            =   7470
      MaxLength       =   5
      TabIndex        =   69
      Top             =   8820
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   52
      Left            =   4815
      MaxLength       =   8
      TabIndex        =   68
      Top             =   8820
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   51
      Left            =   1800
      MaxLength       =   11
      TabIndex        =   67
      Top             =   8820
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   13185
      MaxLength       =   8
      TabIndex        =   11
      Top             =   720
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "�݌ɊǗ��ΏۊO"
      Height          =   255
      Index           =   2
      Left            =   5535
      TabIndex        =   14
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "���ٓ\��v��Ȃ�"
      Height          =   255
      Index           =   1
      Left            =   3015
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   4
      ItemData        =   "PM000302.frx":0004
      Left            =   1320
      List            =   "PM000302.frx":0006
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   12
      Top             =   1680
      Width           =   2835
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   0
      Left            =   1440
      Locked          =   -1  'True
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   32
      Left            =   12495
      MaxLength       =   8
      TabIndex        =   46
      Top             =   5340
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   31
      Left            =   12495
      MaxLength       =   11
      TabIndex        =   45
      Top             =   4860
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   40
      Left            =   12495
      MaxLength       =   11
      TabIndex        =   55
      Top             =   6300
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   41
      Left            =   12495
      MaxLength       =   8
      TabIndex        =   56
      Top             =   6780
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   50
      Left            =   12495
      MaxLength       =   8
      TabIndex        =   66
      Top             =   8280
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   49
      Left            =   12495
      MaxLength       =   11
      TabIndex        =   65
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   10
      Left            =   9000
      MaxLength       =   11
      TabIndex        =   75
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "�g�����i"
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
      Index           =   0
      Left            =   9720
      TabIndex        =   9
      Top             =   840
      Width           =   1365
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   7425
      MaxLength       =   3
      TabIndex        =   7
      Top             =   720
      Width           =   510
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   4320
      MaxLength       =   2
      TabIndex        =   5
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   1305
      MaxLength       =   2
      TabIndex        =   3
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   2  '��
      Index           =   1
      Left            =   4080
      MaxLength       =   40
      TabIndex        =   2
      Top             =   120
      Width           =   4965
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   1
      ItemData        =   "PM000302.frx":0008
      Left            =   1665
      List            =   "PM000302.frx":000A
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   4
      Top             =   720
      Width           =   1635
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   2
      Left            =   4680
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   6
      Top             =   720
      Width           =   1635
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   3
      Left            =   7920
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   8
      Top             =   720
      Width           =   1635
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   11520
      MaxLength       =   11
      TabIndex        =   10
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   7
      Left            =   1320
      MaxLength       =   11
      TabIndex        =   72
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   8
      Left            =   3825
      MaxLength       =   11
      TabIndex        =   73
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   9
      Left            =   6615
      MaxLength       =   11
      TabIndex        =   74
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   11
      Left            =   11745
      MaxLength       =   8
      TabIndex        =   76
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   24
      Left            =   1665
      MaxLength       =   5
      TabIndex        =   37
      Top             =   4380
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   5
      ItemData        =   "PM000302.frx":000C
      Left            =   2385
      List            =   "PM000302.frx":000E
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   38
      Top             =   4380
      Width           =   4110
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   25
      Left            =   7815
      MaxLength       =   11
      TabIndex        =   39
      Top             =   4380
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   26
      Left            =   10335
      MaxLength       =   11
      TabIndex        =   40
      Top             =   4380
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   29
      Left            =   7815
      MaxLength       =   8
      TabIndex        =   43
      Top             =   4860
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   30
      Left            =   10335
      MaxLength       =   3
      TabIndex        =   44
      Top             =   4860
      Width           =   510
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   27
      Left            =   12495
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   4380
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   28
      Left            =   13800
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   4380
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   33
      Left            =   1665
      MaxLength       =   5
      TabIndex        =   47
      Top             =   5820
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   6
      ItemData        =   "PM000302.frx":0010
      Left            =   2385
      List            =   "PM000302.frx":0012
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   48
      Top             =   5820
      Width           =   4110
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   34
      Left            =   7815
      MaxLength       =   11
      TabIndex        =   49
      Top             =   5820
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   35
      Left            =   10335
      MaxLength       =   11
      TabIndex        =   50
      Top             =   5820
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   38
      Left            =   7815
      MaxLength       =   8
      TabIndex        =   53
      Top             =   6300
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   39
      Left            =   10335
      MaxLength       =   3
      TabIndex        =   54
      Top             =   6300
      Width           =   510
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   36
      Left            =   12450
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   5820
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   37
      Left            =   13800
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   5820
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   42
      Left            =   1665
      MaxLength       =   5
      TabIndex        =   57
      Top             =   7320
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   7
      ItemData        =   "PM000302.frx":0014
      Left            =   2385
      List            =   "PM000302.frx":0016
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   58
      Top             =   7320
      Width           =   4110
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   43
      Left            =   7815
      MaxLength       =   11
      TabIndex        =   59
      Top             =   7320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   44
      Left            =   10335
      MaxLength       =   11
      TabIndex        =   60
      Top             =   7320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   47
      Left            =   7815
      MaxLength       =   8
      TabIndex        =   63
      Top             =   7800
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   48
      Left            =   10335
      MaxLength       =   3
      TabIndex        =   64
      Top             =   7800
      Width           =   510
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   45
      Left            =   12495
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   46
      Left            =   13800
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   7320
      Width           =   960
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
      Index           =   11
      Left            =   10305
      TabIndex        =   88
      Top             =   9960
      Width           =   870
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
      Left            =   9495
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   9960
      Width           =   870
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
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   9960
      Width           =   870
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
      Left            =   7785
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   9960
      Width           =   870
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
      Left            =   6480
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   9960
      Width           =   870
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
      Left            =   5625
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   9960
      Width           =   870
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
      Index           =   5
      Left            =   4815
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   9960
      Width           =   870
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
      Index           =   4
      Left            =   3960
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   9960
      Width           =   870
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
      Index           =   3
      Left            =   2655
      TabIndex        =   80
      Top             =   9960
      Width           =   870
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
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   9960
      Width           =   870
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
      Left            =   945
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   9960
      Width           =   870
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
      Left            =   135
      TabIndex        =   77
      Top             =   9960
      Width           =   870
   End
   Begin VB.Label Label 
      Caption         =   "���l"
      Height          =   255
      Index           =   75
      Left            =   10680
      TabIndex        =   167
      Top             =   1800
      Width           =   525
   End
   Begin VB.Label lblUpd_DateTime 
      Caption         =   "99999 99999999-999999"
      Height          =   315
      Left            =   8145
      TabIndex        =   166
      Top             =   9420
      Width           =   2670
   End
   Begin VB.Label Label 
      Caption         =   "�X�V�h�c�^�����F"
      Height          =   255
      Index           =   74
      Left            =   6075
      TabIndex        =   165
      Top             =   9420
      Width           =   2040
   End
   Begin VB.Label lblIns_DateTime 
      Caption         =   "99999 99999999-999999"
      Height          =   315
      Left            =   2385
      TabIndex        =   164
      Top             =   9360
      Width           =   2670
   End
   Begin VB.Label Label 
      Caption         =   "�o�^�h�c�^�����F"
      Height          =   255
      Index           =   73
      Left            =   315
      TabIndex        =   163
      Top             =   9360
      Width           =   2040
   End
   Begin VB.Label Label 
      Caption         =   "�� or 0�F�W������^�P�F�P�̍���"
      Height          =   255
      Index           =   72
      Left            =   5625
      TabIndex        =   162
      Top             =   4020
      Width           =   4155
   End
   Begin VB.Label Label 
      Caption         =   "����敪"
      Height          =   255
      Index           =   71
      Left            =   3960
      TabIndex        =   161
      Top             =   4020
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "����"
      Height          =   255
      Index           =   70
      Left            =   11970
      TabIndex        =   160
      Top             =   2640
      Width           =   555
   End
   Begin VB.Label Label 
      Caption         =   "m3"
      Height          =   255
      Index           =   69
      Left            =   7245
      TabIndex        =   159
      Top             =   2700
      Width           =   240
   End
   Begin VB.Label lblSIZE 
      Alignment       =   1  '�E����
      Height          =   315
      Left            =   6075
      TabIndex        =   158
      Top             =   2700
      Width           =   1140
   End
   Begin VB.Label Label 
      Caption         =   "="
      Height          =   255
      Index           =   68
      Left            =   5850
      TabIndex        =   157
      Top             =   2700
      Width           =   150
   End
   Begin VB.Label Label 
      Caption         =   "�|"
      Height          =   255
      Index           =   67
      Left            =   2925
      TabIndex        =   156
      Top             =   4020
      Width           =   150
   End
   Begin VB.Label Label 
      Caption         =   "�|"
      Height          =   255
      Index           =   66
      Left            =   2295
      TabIndex        =   155
      Top             =   4020
      Width           =   150
   End
   Begin VB.Label Label 
      Caption         =   "�|"
      Height          =   255
      Index           =   65
      Left            =   1665
      TabIndex        =   154
      Top             =   4020
      Width           =   150
   End
   Begin VB.Label Label 
      Caption         =   "�W���I��"
      Height          =   255
      Index           =   64
      Left            =   240
      TabIndex        =   153
      Top             =   4020
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   $"PM000302.frx":0018
      Height          =   255
      Index           =   63
      Left            =   7965
      TabIndex        =   152
      Top             =   3600
      Width           =   4155
   End
   Begin VB.Label Label 
      Caption         =   "���i������"
      Height          =   255
      Index           =   62
      Left            =   6210
      TabIndex        =   151
      Top             =   3600
      Width           =   1320
   End
   Begin VB.Label Label 
      Caption         =   "�b"
      Height          =   255
      Index           =   61
      Left            =   8820
      TabIndex        =   150
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label 
      Caption         =   "�W������"
      Height          =   255
      Index           =   60
      Left            =   6525
      TabIndex        =   149
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "�b"
      Height          =   255
      Index           =   59
      Left            =   5895
      TabIndex        =   148
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label 
      Caption         =   "�i�󔒁F�o�͂��Ȃ��j"
      Height          =   255
      Index           =   58
      Left            =   11655
      TabIndex        =   147
      Top             =   3120
      Width           =   2445
   End
   Begin VB.Label Label 
      Caption         =   "ð�ߒ�"
      Height          =   255
      Index           =   57
      Left            =   4005
      TabIndex        =   146
      Top             =   3600
      Width           =   780
   End
   Begin VB.Label Label 
      Caption         =   "ð�ߎ��"
      Height          =   255
      Index           =   56
      Left            =   240
      TabIndex        =   145
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "(1:����)"
      Height          =   255
      Index           =   55
      Left            =   1800
      TabIndex        =   144
      Top             =   3120
      Width           =   1005
   End
   Begin VB.Label Label 
      Caption         =   "���"
      Height          =   255
      Index           =   54
      Left            =   720
      TabIndex        =   143
      Top             =   3120
      Width           =   555
   End
   Begin VB.Label Label 
      Caption         =   "�ގ�"
      Height          =   255
      Index           =   53
      Left            =   4005
      TabIndex        =   142
      Top             =   2280
      Width           =   555
   End
   Begin VB.Label Label 
      Caption         =   "�`��"
      Height          =   255
      Index           =   52
      Left            =   720
      TabIndex        =   141
      Top             =   2280
      Width           =   555
   End
   Begin VB.Label Label 
      Caption         =   "�A�����e"
      Height          =   255
      Index           =   51
      Left            =   10080
      TabIndex        =   140
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "��ƍH��"
      Height          =   255
      Index           =   46
      Left            =   3465
      TabIndex        =   139
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "���x�^����"
      Height          =   255
      Index           =   50
      Left            =   7245
      TabIndex        =   138
      Top             =   2280
      Width           =   1320
   End
   Begin VB.Label Label 
      Caption         =   "X"
      Height          =   255
      Index           =   49
      Left            =   4455
      TabIndex        =   137
      Top             =   2700
      Width           =   150
   End
   Begin VB.Label Label 
      Caption         =   "X"
      Height          =   255
      Index           =   48
      Left            =   3105
      TabIndex        =   136
      Top             =   2700
      Width           =   150
   End
   Begin VB.Label Label 
      Caption         =   "�ː�"
      Height          =   255
      Index           =   45
      Left            =   9990
      TabIndex        =   134
      Top             =   2640
      Width           =   555
   End
   Begin VB.Label Label 
      Caption         =   "�ŐV�d���P��"
      Height          =   255
      Index           =   44
      Left            =   12450
      TabIndex        =   133
      Top             =   8940
      Width           =   1635
   End
   Begin VB.Label Label 
      Caption         =   "�ŐV�d����"
      Height          =   255
      Index           =   43
      Left            =   6120
      TabIndex        =   132
      Top             =   8940
      Width           =   1365
   End
   Begin VB.Label Label 
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
      Height          =   255
      Index           =   42
      Left            =   13815
      TabIndex        =   131
      Top             =   480
      Width           =   510
   End
   Begin VB.Label Label 
      Caption         =   "�ŏI�o�ɐ�"
      Height          =   255
      Index           =   41
      Left            =   3510
      TabIndex        =   130
      Top             =   8940
      Width           =   1365
   End
   Begin VB.Label Label 
      Caption         =   "�ŏI�o�ד�"
      Height          =   255
      Index           =   40
      Left            =   360
      TabIndex        =   129
      Top             =   8940
      Width           =   1365
   End
   Begin VB.Label Label 
      Caption         =   "/"
      Height          =   255
      Index           =   39
      Left            =   13095
      TabIndex        =   128
      Top             =   840
      Width           =   150
   End
   Begin VB.Label Label 
      Caption         =   "���ދ敪"
      Height          =   255
      Index           =   38
      Left            =   240
      TabIndex        =   127
      Top             =   1800
      Width           =   1005
   End
   Begin VB.Label Label 
      Caption         =   "�O�񒍕���"
      Height          =   255
      Index           =   37
      Left            =   11145
      TabIndex        =   126
      Top             =   4980
      Width           =   1230
   End
   Begin VB.Label Label 
      Caption         =   "������"
      Height          =   255
      Index           =   36
      Left            =   11640
      TabIndex        =   125
      Top             =   5460
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "������"
      Height          =   255
      Index           =   35
      Left            =   11640
      TabIndex        =   124
      Top             =   6900
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "�O�񒍕���"
      Height          =   255
      Index           =   34
      Left            =   11145
      TabIndex        =   123
      Top             =   6420
      Width           =   1230
   End
   Begin VB.Label Label 
      Caption         =   "������"
      Height          =   255
      Index           =   33
      Left            =   11640
      TabIndex        =   122
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "�O�񒍕���"
      Height          =   255
      Index           =   32
      Left            =   11145
      TabIndex        =   121
      Top             =   7920
      Width           =   1230
   End
   Begin VB.Label Label 
      Caption         =   "�ݒ��"
      Height          =   255
      Index           =   31
      Left            =   8145
      TabIndex        =   120
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "�ݒ��"
      Height          =   255
      Index           =   7
      Left            =   3015
      TabIndex        =   119
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "%"
      Height          =   255
      Index           =   30
      Left            =   14835
      TabIndex        =   118
      Top             =   7680
      Width           =   150
   End
   Begin VB.Label Label 
      Caption         =   "%"
      Height          =   255
      Index           =   23
      Left            =   14835
      TabIndex        =   117
      Top             =   6120
      Width           =   150
   End
   Begin VB.Label Label 
      Caption         =   "%"
      Height          =   255
      Index           =   16
      Left            =   14835
      TabIndex        =   116
      Top             =   4560
      Width           =   150
   End
   Begin VB.Label Label 
      Caption         =   "���ޕi��"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   115
      Top             =   240
      Width           =   1005
   End
   Begin VB.Label Label 
      Caption         =   "�i��"
      Height          =   255
      Index           =   2
      Left            =   2745
      TabIndex        =   114
      Top             =   240
      Width           =   510
   End
   Begin VB.Label Label 
      Caption         =   "�d���敪"
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
      Index           =   1
      Left            =   240
      TabIndex        =   113
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label Label 
      Caption         =   "�̔��敪"
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
      Index           =   3
      Left            =   3375
      TabIndex        =   112
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label Label 
      Caption         =   "���x�P��"
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
      Index           =   4
      Left            =   6345
      TabIndex        =   111
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label Label 
      Caption         =   "�O���݌ɋ��z"
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
      Index           =   5
      Left            =   11385
      TabIndex        =   110
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label 
      Caption         =   "�W������"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   109
      Top             =   1320
      Width           =   1005
   End
   Begin VB.Label Label 
      Caption         =   "�W������"
      Height          =   255
      Index           =   8
      Left            =   5400
      TabIndex        =   108
      Top             =   1320
      Width           =   1005
   End
   Begin VB.Label Label 
      Caption         =   "�댯�݌�"
      Height          =   255
      Index           =   9
      Left            =   10665
      TabIndex        =   107
      Top             =   1320
      Width           =   1005
   End
   Begin VB.Label Label 
      Caption         =   "�d����(1)"
      Height          =   255
      Index           =   10
      Left            =   495
      TabIndex        =   106
      Top             =   4500
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "�d���P��"
      Height          =   255
      Index           =   11
      Left            =   6735
      TabIndex        =   105
      Top             =   4500
      Width           =   1005
   End
   Begin VB.Label Label 
      Caption         =   "�ݒ��"
      Height          =   255
      Index           =   12
      Left            =   9480
      TabIndex        =   104
      Top             =   4500
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "ۯĐ�"
      Height          =   255
      Index           =   13
      Left            =   7095
      TabIndex        =   103
      Top             =   4980
      Width           =   645
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   360
      X2              =   14280
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label 
      Caption         =   "ذ�����"
      Height          =   255
      Index           =   14
      Left            =   9345
      TabIndex        =   102
      Top             =   4980
      Width           =   870
   End
   Begin VB.Label Label 
      Caption         =   "�e��"
      Height          =   255
      Index           =   15
      Left            =   11865
      TabIndex        =   101
      Top             =   4500
      Width           =   510
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   360
      X2              =   14280
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label Label 
      Caption         =   "�d����(2)"
      Height          =   255
      Index           =   17
      Left            =   495
      TabIndex        =   100
      Top             =   5940
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "�d���P��"
      Height          =   255
      Index           =   18
      Left            =   6735
      TabIndex        =   99
      Top             =   5940
      Width           =   1005
   End
   Begin VB.Label Label 
      Caption         =   "�ݒ��"
      Height          =   255
      Index           =   19
      Left            =   9480
      TabIndex        =   98
      Top             =   5940
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "ۯĐ�"
      Height          =   255
      Index           =   20
      Left            =   7095
      TabIndex        =   97
      Top             =   6420
      Width           =   645
   End
   Begin VB.Label Label 
      Caption         =   "ذ�����"
      Height          =   255
      Index           =   21
      Left            =   9345
      TabIndex        =   96
      Top             =   6420
      Width           =   870
   End
   Begin VB.Label Label 
      Caption         =   "�e��"
      Height          =   255
      Index           =   22
      Left            =   11865
      TabIndex        =   95
      Top             =   5940
      Width           =   510
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   360
      X2              =   14160
      Y1              =   7260
      Y2              =   7260
   End
   Begin VB.Label Label 
      Caption         =   "�d����(3)"
      Height          =   255
      Index           =   24
      Left            =   495
      TabIndex        =   94
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "�d���P��"
      Height          =   255
      Index           =   25
      Left            =   6735
      TabIndex        =   93
      Top             =   7440
      Width           =   1005
   End
   Begin VB.Label Label 
      Caption         =   "�ݒ��"
      Height          =   255
      Index           =   26
      Left            =   9480
      TabIndex        =   92
      Top             =   7440
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "ۯĐ�"
      Height          =   255
      Index           =   27
      Left            =   7095
      TabIndex        =   91
      Top             =   7920
      Width           =   645
   End
   Begin VB.Label Label 
      Caption         =   "ذ�����"
      Height          =   255
      Index           =   28
      Left            =   9345
      TabIndex        =   90
      Top             =   7920
      Width           =   870
   End
   Begin VB.Label Label 
      Caption         =   "�e��"
      Height          =   255
      Index           =   29
      Left            =   11865
      TabIndex        =   89
      Top             =   7440
      Width           =   510
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   630
      X2              =   14310
      Y1              =   8700
      Y2              =   8700
   End
   Begin VB.Label Label 
      Caption         =   "����(WxDxH)mm"
      Height          =   255
      Index           =   47
      Left            =   45
      TabIndex        =   135
      Top             =   2700
      Width           =   1635
   End
End
Attribute VB_Name = "PM000302"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�e�L�X�g�p�Y��
Private Const ptxHIN_GAI% = 0               '�i��
Private Const ptxHIN_NAME% = 1              '�i��
Private Const ptxG_SHIIRE_KBN% = 2          '�d���敪
Private Const ptxG_HANBAI_KBN% = 3          '�̔��敪
Private Const ptxG_SYUSHI% = 4              '���x�P��
Private Const ptxG_ZEN_ZAIKO_KIN% = 5       '�O���݌ɋ��z
Private Const ptxG_ZEN_ZAIKO_QTY% = 6       '�O���݌ɐ���
Private Const ptxG_ST_URITAN% = 7           '�W���e�������P��
Private Const ptxG_ST_URITAN_DT% = 8        '�W���e�������P���ݒ��
Private Const ptxG_ST_SHITAN% = 9           '�W���e�������P��
Private Const ptxG_ST_SHITAN_DT% = 10       '�W���e�������P���ݒ��
Private Const ptxHOJYU_P% = 11              '��[�_�i�댯�݌Ɂj

Private Const ptxSAI_SU% = 12               '�ː�               2008.02.14


Private Const ptxD_KEISHIKI% = 13           '�`��               2008.02.14
Private Const ptxD_MATERIAL% = 14           '�ގ�               2008.02.14
Private Const ptxD_THICKNESS% = 15          '����ްف@����      2008.02.14
    
    
Private Const ptxD_SIZE_W% = 16             '����ްٻ��ށiW�j   2008.02.14
Private Const ptxD_SIZE_D% = 17             '����ްٻ��ށiD�j   2008.02.14
Private Const ptxD_SIZE_H% = 18             '����ްٻ��ށiH�j   2008.02.14
        
Private Const ptxD_PRINT% = 19              '�������^���Ȃ�   2008.02.14
            
        
Private Const ptxS_KOUSU% = 20              '���i���@�H��       2008.02.14


Private Const ptxSE_USOU_F% = 21            '�A�����o��      2008.02.14


    
Private Const ptxUSE_TAPE_KIND% = 22        '�g�p�e�[�v���     2008.02.14
Private Const ptxUSE_TAPE_LNG% = 23         '�g�p�e�[�v��     2008.02.14



Private Const ptxSHI_CODE1% = 24            '�d���溰��(1)
Private Const ptxSHI_TANKA1% = 25           '�d���P��(1)
Private Const ptxSHI_TANKA_DT1% = 26        '�d���P���ݒ��(1)
Private Const ptxSHI_ARARI1% = 27           '�e���z(1)
Private Const ptxSHI_ARARI_RITU1% = 28      '�e����(1)
Private Const ptxSHI_LOT1% = 29             'ۯĐ�(1)
Private Const ptxSHI_LEAD_TIME1% = 30       'ذ�����(1)
Private Const ptxSHI_LAST_ORDER_DT1% = 31   '�O�񒍕���(1)
Private Const ptxSHI_LAST_ORDER_QTY1% = 32  '�O�񒍕���(1)

Private Const ptxSHI_CODE2% = 33            '�d���溰��(2)
Private Const ptxSHI_TANKA2% = 34           '�d���P��(2)
Private Const ptxSHI_TANKA_DT2% = 35        '�d���P���ݒ��(2)
Private Const ptxSHI_ARARI2% = 36           '�e���z(2)
Private Const ptxSHI_ARARI_RITU2% = 37      '�e����(2)
Private Const ptxSHI_LOT2% = 38             'ۯĐ�(2)
Private Const ptxSHI_LEAD_TIME2% = 39       'ذ�����(2)
Private Const ptxSHI_LAST_ORDER_DT2% = 40   '�O�񒍕���(2)
Private Const ptxSHI_LAST_ORDER_QTY2% = 41  '�O�񒍕���(2)

Private Const ptxSHI_CODE3% = 42            '�d���溰��(3)
Private Const ptxSHI_TANKA3% = 43           '�d���P��(3)
Private Const ptxSHI_TANKA_DT3% = 44        '�d���P���ݒ��(3)
Private Const ptxSHI_ARARI3% = 45           '�e���z(3)
Private Const ptxSHI_ARARI_RITU3% = 46      '�e����(3)
Private Const ptxSHI_LOT3% = 47             'ۯĐ�(3)
Private Const ptxSHI_LEAD_TIME3% = 48       'ذ�����(3)
Private Const ptxSHI_LAST_ORDER_DT3% = 49   '�O�񒍕���(3)
Private Const ptxSHI_LAST_ORDER_QTY3% = 50  '�O�񒍕���(3)

Private Const ptxLAST_SYU_DT = 51           '�ŏI�o�ɓ�
Private Const ptxG_LAST_SYUKA_QTY = 52      '�ŏI�o�ɐ�

Private Const ptxLAST_CODE = 53             '�ŐV�d����R�[�h   2007.05.28
Private Const ptxLAST_TANKA = 54            '�ŐV�d���P��       2007.05.28

Private Const ptxSEI_SYU_KON = 55           '�W������           2008.07.16
Private Const ptxSEI_KBN = 56               '�����敪           2008.07.16

Private Const ptxST_SOKO% = 57              '�W���I�ԁ@�q��     2009.09.01
Private Const ptxST_RETU% = 58              '�W���I�ԁ@��       2009.09.01
Private Const ptxST_REN% = 59               '�W���I�ԁ@�A       2009.09.01
Private Const ptxST_DAN% = 60               '�W���I�ԁ@�i       2009.09.01

Private Const ptxKUTI_SU% = 61              '����               2010.01.18
Private Const ptxKONPOU_F% = 62             '����敪           2010.01.18

Private Const ptxSHIIRE_BIKOU% = 63        '�d�����l           2018.04.19


'�R���{�p�Y��
Private Const pcmbNAIGAI% = 0               '�����O
Private Const pcmbG_SHIIRE% = 1             '�d���敪
Private Const pcmbG_HANBAI% = 2             '�̔��敪
Private Const pcmbG_SYUSHI% = 3             '���x�P��
Private Const pcmbG_SHIZAI_KBN% = 4         '���ދ敪
Private Const pcmbSHIIRE1% = 5              '�d����(1)
Private Const pcmbSHIIRE2% = 6              '�d����(2)
Private Const pcmbSHIIRE3% = 7              '�d����(3)
Private Const pcmbLAST_CODE% = 8            '�ŐV�d����         2007.05.28
'�`�F�b�N�p�Y��
Private Const pchkG_KUMITATE% = 0           '�g�����i
Private Const pchkG_LABEL_NON% = 1          '���ٓ\��v��Ȃ�
Private Const pchkZAIKO_F% = 2              '�݌ɊǗ��ΏۊO

Private Const pchkZAIKO_CLR_F% = 3          '�I���\�@�݌ɐ���\��   2012.12.13

Private INIT_FLG    As Boolean

Private svTANKA     As String               '2018.04.09

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    PM000302.MousePointer = vbHourglass

    Call Ctrl_Lock(PM000302)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PM000302)


    PM000302.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   ���͍��ڂ̃G���[�`�F�b�N
'----------------------------------------------------------------------------
Dim com     As Integer
Dim ans     As Integer
Dim sts     As Integer

Dim i       As Integer
    
Dim w_CODE  As String * 22      '2018.04.19
    
    
    Error_Check_Proc = True
    
    
    
    
    Select Case Mode
        
        Case ptxHIN_GAI      '�i��
            
            If Trim(Text1(ptxHIN_GAI).Text) = "" Then
                MsgBox "���͂������ڂ̓G���[�ł��B(�i��)"
                Text1(ptxHIN_GAI).SetFocus
                Exit Function
            End If
            
        
            If G_SCREEN_FLG = G_SCREEN_INS Then
                '�V�K���͏d���`�F�b�N
                Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
                Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo1(pcmbNAIGAI), 1))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)
            
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        ans = MsgBox("���͂����R�[�h�́A�o�^�ςł��B�X�V�����Ƃ��Čp�����܂����H", vbYesNo, "�m�F����")
                        If ans = vbNo Then
                            Text1(ptxHIN_GAI).SetFocus
                            Exit Function
                        End If
                                    
                        w_CODE = Last_JGYOBU & Right(Combo1(pcmbNAIGAI), 1) & Text1(ptxHIN_GAI).Text                '2018.04.19
'                        Call Item_Disp_Proc(Last_JGYOBU & Right(Combo1(pcmbNAIGAI), 1) & Text1(ptxHIN_GAI).Text)   '2018.04.19
                        Call Item_Disp_Proc(w_CODE)                                                                 '2018.04.19
                    
                    Case BtErrKeyNotFound
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                        Exit Function
                End Select
            
            
                Text1(ptxHIN_GAI).BackColor = G_INPUT_NG
                Text1(ptxHIN_GAI).Locked = True
                Text1(ptxHIN_GAI).TabStop = False
            
'                Text1(ptxHIN_NAME).BackColor = G_INPUT_NG
'                Text1(ptxHIN_NAME).Locked = True
'                Text1(ptxHIN_NAME).TabStop = False
            
            End If
        
        Case ptxG_SHIIRE_KBN       '�d���敪
        
    
                If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                    Error_Check_Proc = False        '2016.05.18
                    Exit Function                   '2016.05.18
                End If                              '2016.05.18
            
'2015.09.16            If Last_JGYOBU = SHIZAI Then
                If Trim(Text1(ptxG_SHIIRE_KBN).Text) = "" Then
'2016.06.18                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(�d���敪)"
                    Text1(ptxG_SHIIRE_KBN).SetFocus
                    Exit Function
                End If
            
            
                For i = 0 To Combo1(pcmbG_SHIIRE).ListCount - 1
                    If Text1(ptxG_SHIIRE_KBN).Text = Left(Right(Combo1(pcmbG_SHIIRE).List(i), 3), 2) Then
                        Combo1(pcmbG_SHIIRE).ListIndex = i
                        Exit For
                    End If
                Next i
                
                If i > (Combo1(pcmbG_SHIIRE).ListCount - 1) Then
'2016.05.18                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(�d���敪)"
                    Text1(ptxG_SHIIRE_KBN).SetFocus
                    Exit Function
                End If
                        
'2015.09.16            End If
        Case ptxG_HANBAI_KBN       '�̔��敪
        

                If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                    Error_Check_Proc = False        '2016.05.18
                    Exit Function                   '2016.05.18
                End If                              '2016.05.18


'2015.09.16            If Last_JGYOBU = SHIZAI Then
                If Trim(Text1(ptxG_HANBAI_KBN).Text) = "" Then
'                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(�̔��敪)"
                    Text1(ptxG_HANBAI_KBN).SetFocus
                    Exit Function
                End If
            
            
                For i = 0 To Combo1(pcmbG_HANBAI).ListCount - 1
                    If Text1(ptxG_HANBAI_KBN).Text = Left(Right(Combo1(pcmbG_HANBAI).List(i), 3), 2) Then
                        Combo1(pcmbG_HANBAI).ListIndex = i
                        Exit For
                    End If
                Next i
                
                If i > (Combo1(pcmbG_HANBAI).ListCount - 1) Then
'2016.05.18                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(�̔��敪)"
                    Text1(ptxG_HANBAI_KBN).SetFocus
                    Exit Function
                End If
'2015.09.16            End If
        Case ptxG_SYUSHI           '���x�P��
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
        
        
            If Trim(Text1(ptxG_SYUSHI).Text) = "" Then
''                MsgBox "���͂������ڂ̓G���[�ł��B"
''                Text1(ptxG_SYUSHI).SetFocus
''                Exit Function
            Else
        
                For i = 0 To Combo1(pcmbG_SYUSHI).ListCount - 1
                    If Text1(ptxG_SYUSHI).Text = Right(Combo1(pcmbG_SYUSHI).List(i), 3) Then
                        Combo1(pcmbG_SYUSHI).ListIndex = i
                        Exit For
                    End If
                Next i
                
                If i > (Combo1(pcmbG_SYUSHI).ListCount - 1) Then
'2016.05.18                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(���x�P��)"
                    Text1(ptxG_SYUSHI).SetFocus
                    Exit Function
                End If
            End If
        
        Case ptxG_ZEN_ZAIKO_KIN    '�O���݌ɋ��z
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
            
            If Trim(Text1(ptxG_ZEN_ZAIKO_KIN).Text) = "" Then
                Text1(ptxG_ZEN_ZAIKO_KIN).Text = "0"
            End If
        
            If Not IsNumeric(Text1(ptxG_ZEN_ZAIKO_KIN).Text) Then
'2016.05.18                MsgBox "���͂������ڂ̓G���[�ł��B"
                MsgBox "���͂������ڂ̓G���[�ł��B(�O���݌ɋ��z)"
                Text1(ptxG_ZEN_ZAIKO_KIN).SetFocus
                Exit Function
            End If
        
            Text1(ptxG_ZEN_ZAIKO_KIN).Text = Format(CLng(Text1(ptxG_ZEN_ZAIKO_KIN).Text), "#0")
        
        Case ptxG_ZEN_ZAIKO_QTY    '�O���݌ɐ���
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
            
            If Trim(Text1(ptxG_ZEN_ZAIKO_QTY).Text) = "" Then
                Text1(ptxG_ZEN_ZAIKO_QTY).Text = "0"
            End If
        
            If Not IsNumeric(Text1(ptxG_ZEN_ZAIKO_QTY).Text) Then
'2016.05.18                MsgBox "���͂������ڂ̓G���[�ł��B"
                MsgBox "���͂������ڂ̓G���[�ł��B(�O���݌ɐ���)"
                Text1(ptxG_ZEN_ZAIKO_QTY).SetFocus
                Exit Function
            End If
        
            Text1(ptxG_ZEN_ZAIKO_QTY).Text = Format(CLng(Text1(ptxG_ZEN_ZAIKO_QTY).Text), "#0")
        
        
        Case ptxG_ST_URITAN        '�W������
            
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
        
            If Trim(Text1(ptxG_ST_URITAN).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxG_ST_URITAN).Text) Then
'2016.05.18                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(�W������)"
                    Text1(ptxG_ST_URITAN).SetFocus
                    Exit Function
                End If
            
                Text1(ptxG_ST_URITAN).Text = Format(CDbl(Text1(ptxG_ST_URITAN).Text), "#0.00")
            End If
        
        Case ptxG_ST_URITAN_DT     '�W�������ݒ��
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
            
            If Trim(Text1(ptxG_ST_URITAN_DT).Text) = "" Then
                If Trim(Text1(ptxG_ST_URITAN).Text) <> "" Then
                    Text1(ptxG_ST_URITAN_DT).Text = Format(Now, "YYYY/MM/DD")
                End If
            Else
                If Not IsDate(Text1(ptxG_ST_URITAN_DT).Text) Then
'2016.05.18                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(�W�������ݒ��)"
                    Text1(ptxG_ST_URITAN_DT).SetFocus
                    Exit Function
                End If
            
                Text1(ptxG_ST_URITAN_DT).Text = Format(Text1(ptxG_ST_URITAN_DT).Text, "YYYY/MM/DD")
            End If
        
        Case ptxG_ST_SHITAN       '�W������
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
            
            
            If Trim(Text1(ptxG_ST_SHITAN).Text) = "" Then
            Else
        
                If Not IsNumeric(Text1(ptxG_ST_SHITAN).Text) Then
'2016.05.18                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(�W������)"
                    Text1(ptxG_ST_URITAN).SetFocus
                    Exit Function
                End If
            
                Text1(ptxG_ST_SHITAN).Text = Format(CDbl(Text1(ptxG_ST_SHITAN).Text), "#0.00")
            End If
        
        Case ptxG_ST_SHITAN_DT     '�W�������ݒ��
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
            
            
            If Trim(Text1(ptxG_ST_SHITAN_DT).Text) = "" Then
                If Trim(Text1(ptxG_ST_SHITAN).Text) <> "" Then
                    Text1(ptxG_ST_SHITAN_DT).Text = Format(Now, "YYYY/MM/DD")
                End If
            Else
                If Not IsDate(Text1(ptxG_ST_SHITAN_DT).Text) Then
'2016.05.18                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(�W�������ݒ��)"
                    Text1(ptxG_ST_SHITAN_DT).SetFocus
                    Exit Function
                End If
            
                Text1(ptxG_ST_SHITAN_DT).Text = Format(Text1(ptxG_ST_SHITAN_DT).Text, "YYYY/MM/DD")
            End If
        
        
        Case ptxHOJYU_P            '�댯�݌�
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
        
            If Trim(Text1(ptxHOJYU_P).Text) = "" Then
                Text1(ptxHOJYU_P).Text = "0"
            End If
            
            If Not IsNumeric(Text1(ptxHOJYU_P).Text) Then
'2016.05.18                MsgBox "���͂������ڂ̓G���[�ł��B"
                MsgBox "���͂������ڂ̓G���[�ł��B(�댯�݌�)"
                Text1(ptxG_ST_URITAN).SetFocus
                Exit Function
            End If
        
            Text1(ptxHOJYU_P).Text = Format(CLng(Text1(ptxHOJYU_P).Text), "#0")
              
              
              
        Case ptxSAI_SU              '�ː�   2008.02.14
              
            
            If Trim(Text1(ptxSAI_SU).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxSAI_SU).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(�ː�)"
                    Text1(ptxSAI_SU).SetFocus
                    Exit Function
                Else
                    Text1(ptxSAI_SU).Text = Format(CCur(Text1(ptxSAI_SU).Text), "#0.00")
                End If
            End If
              
              
              
              
              
              
        Case ptxD_KEISHIKI          '�`��               2008.02.14
        Case ptxD_MATERIAL          '�ގ�               2008.02.14
        Case ptxD_THICKNESS         '����ްف@����      2008.02.14
    
    
        Case ptxD_SIZE_W            '����ްٻ��ށiW�j   2008.02.14
        
            If Text1(ptxD_SIZE_W).Text = "" Then
            Else
                If Not IsNumeric(Text1(ptxD_SIZE_W).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(���ށiW�j)"
                    Text1(ptxD_SIZE_W).SetFocus
                    Exit Function
                Else
                    
On Error GoTo Error_Proc
                    
                    Text1(ptxD_SIZE_W).Text = Format(Val(Text1(ptxD_SIZE_W).Text), "#")
                
                
                    If IsNumeric(Text1(ptxD_SIZE_W).Text) And IsNumeric(Text1(ptxD_SIZE_D).Text) And IsNumeric(Text1(ptxD_SIZE_H).Text) Then
                    
                        lblSIZE.Caption = Format(ToHalfAdjust(CCur(Val(Text1(ptxD_SIZE_W).Text) / 1000 * Val(Text1(ptxD_SIZE_D).Text) / 1000 * Val(Text1(ptxD_SIZE_H).Text) / 1000), 4), "#0.000")
                    
                        If Trim(Text1(ptxSAI_SU).Text) = "" Then
                        
'                            Text1(ptxSAI_SU).Text = Format(ToHalfAdjust(CCur(CCur(lblSIZE.Caption) / 0.028), 1), "0.00")       '2016.04.22
                            Text1(ptxSAI_SU).Text = Format(ToHalfAdjust(CCur(CCur(lblSIZE.Caption) / 0.028), 2), "0.00")        '2016.04.22
                        
                        
                        End If
                    
                    End If
                
                
On Error GoTo 0
                
                
                
                End If
            End If
        
        Case ptxD_SIZE_D            '����ްٻ��ށiD�j   2008.02.14
        
            If Text1(ptxD_SIZE_D).Text = "" Then
            Else
                If Not IsNumeric(Text1(ptxD_SIZE_D).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(���ށiD�j)"
                    Text1(ptxD_SIZE_D).SetFocus
                    Exit Function
                Else
                    
On Error GoTo Error_Proc
                    
                    Text1(ptxD_SIZE_D).Text = Format(Val(Text1(ptxD_SIZE_D).Text), "#")
                
                
                    If IsNumeric(Text1(ptxD_SIZE_W).Text) And IsNumeric(Text1(ptxD_SIZE_D).Text) And IsNumeric(Text1(ptxD_SIZE_H).Text) Then
                    
                        
                        lblSIZE.Caption = Format(ToHalfAdjust(CCur(Val(Text1(ptxD_SIZE_W).Text) / 1000 * Val(Text1(ptxD_SIZE_D).Text) / 1000 * Val(Text1(ptxD_SIZE_H).Text) / 1000), 4), "#0.000")
                    
                        If Trim(Text1(ptxSAI_SU).Text) = "" Then
                        
'                            Text1(ptxSAI_SU).Text = Format(ToHalfAdjust(CCur(CCur(lblSIZE.Caption) / 0.028), 1), "0.00")       '2016.04.22
                            Text1(ptxSAI_SU).Text = Format(ToHalfAdjust(CCur(CCur(lblSIZE.Caption) / 0.028), 2), "0.00")        '2016.04.22
                        
                        
                        End If
                    
                    End If
                
On Error GoTo 0
                
                End If
            
            End If
        
        
        
        
        Case ptxD_SIZE_H            '����ްٻ��ށiH�j   2008.02.14
            
        
            If Text1(ptxD_SIZE_H).Text = "" Then
            Else
                If Not IsNumeric(Text1(ptxD_SIZE_H).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(���ށiH�j)"
                    Text1(ptxD_SIZE_H).SetFocus
                    Exit Function
                Else
                    
On Error GoTo Error_Proc
                    
                    Text1(ptxD_SIZE_H).Text = Format(Val(Text1(ptxD_SIZE_H).Text), "#")
                
                    If IsNumeric(Text1(ptxD_SIZE_W).Text) And IsNumeric(Text1(ptxD_SIZE_D).Text) And IsNumeric(Text1(ptxD_SIZE_H).Text) Then
                    
                        lblSIZE.Caption = Format(ToHalfAdjust(CCur(Val(Text1(ptxD_SIZE_W).Text) / 1000 * Val(Text1(ptxD_SIZE_D).Text) / 1000 * Val(Text1(ptxD_SIZE_H).Text) / 1000), 4), "#0.000")
                    
                        If Trim(Text1(ptxSAI_SU).Text) = "" Then
                        
'                            Text1(ptxSAI_SU).Text = Format(ToHalfAdjust(CCur(CCur(lblSIZE.Caption) / 0.028), 1), "0.00")       '2016.04.22
                            Text1(ptxSAI_SU).Text = Format(ToHalfAdjust(CCur(CCur(lblSIZE.Caption) / 0.028), 2), "0.00")        '2016.04.22
                        
                        
                        End If
                    
                    End If
On Error GoTo 0
                
                End If
            End If
        
        
        
        Case ptxD_PRINT             '�������^���Ȃ�   2008.02.14
        Case ptxS_KOUSU             '���i���@�H��       2008.07.16
            
        
            If Trim(Text1(Mode).Text) = "" Then
            Else
                If IsNumeric(Text1(Mode).Text) Then
                Else
                    MsgBox "���͂������ڂ̓G���[�ł��B(��ƍH��)"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If
        
        Case ptxSEI_SYU_KON         '�W������@�H��     2008.07.16
            
        
            If Trim(Text1(Mode).Text) = "" Then
            Else
                If IsNumeric(Text1(Mode).Text) Then
                Else
                    MsgBox "���͂������ڂ̓G���[�ł��B(�W������H��)"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If
        
        
        Case ptxSE_USOU_F           '�A�����@�o���׸�   2008.02.14
        Case ptxUSE_TAPE_KIND       '�g�p�e�[�v���     2008.02.14
        Case ptxUSE_TAPE_LNG        '�g�p�e�[�v��       2008.02.14
              
        Case ptxSEI_KBN        '�����敪           2008.07.16
            
            
            
            If Trim(Text1(Mode).Text) = "" Or _
                Trim(Text1(Mode).Text) = "1" Or _
                Trim(Text1(Mode).Text) = "2" Then
            Else
                MsgBox "���͂������ڂ̓G���[�ł��B(�����敪)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
        
        Case ptxST_SOKO      '�W���I��    2009.09.01
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(Val(Text1(Mode).Text), "00")
            End If
        
        Case ptxST_RETU      '�W���I��    2009.09.01
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(Val(Text1(Mode).Text), "00")
            End If
        
        Case ptxST_REN      '�W���I��    2009.09.01
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(Val(Text1(Mode).Text), "00")
            End If
        
        Case ptxST_DAN      '�W���I��    2009.09.01
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(Val(Text1(Mode).Text), "00")
            End If
        
        
            If Trim(Text1(ptxST_SOKO).Text) = "" And Trim(Text1(ptxST_RETU).Text) = "" And Trim(Text1(ptxST_REN).Text) = "" And Trim(Text1(ptxST_DAN).Text) = "" Then
            Else
                If Trim(Text1(ptxST_SOKO).Text) = "**" And Trim(Text1(ptxST_RETU).Text) = "**" And Trim(Text1(ptxST_REN).Text) = "**" And Trim(Text1(ptxST_DAN).Text) = "**" Then
                Else
                    Call UniCode_Conv(K0_TANA.SOKO_NO, Text1(ptxST_SOKO).Text)
                    Call UniCode_Conv(K0_TANA.Retu, Text1(ptxST_RETU).Text)
                    Call UniCode_Conv(K0_TANA.Ren, Text1(ptxST_REN).Text)
                    Call UniCode_Conv(K0_TANA.Dan, Text1(ptxST_DAN).Text)
            
            
                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                    Select Case sts
                        Case BtNoErr
                        
                        Case BtErrKeyNotFound
                            MsgBox "���͂������ڂ̓G���[�ł��B(�W���I��)"
                            Text1(ptxST_SOKO).SetFocus
                            Exit Function
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�I�}�X�^")
                            Exit Function
                    End Select
                End If
        
        
        
        
        
            End If
        
        
        
        Case ptxSHI_CODE1           '�d����(1)
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
            
            Text1(Mode).Text = StrConv(Text1(Mode).Text, vbUpperCase) '2016.01.26
            If Trim(Text1(ptxSHI_CODE1).Text) = "" Then
                For i = ptxSHI_CODE1 To ptxSHI_LAST_ORDER_QTY1
                    Text1(i).Text = ""
                Next i
                Combo1(pcmbSHIIRE1).ListIndex = -1
            Else
                For i = 0 To Combo1(pcmbSHIIRE1).ListCount - 1
                    If Trim(Text1(ptxSHI_CODE1).Text) = Trim(Right(Combo1(pcmbSHIIRE1).List(i), 5)) Then
                        Combo1(pcmbSHIIRE1).ListIndex = i
                        Exit For
                    End If
                Next i
                
                If i > (Combo1(pcmbSHIIRE1).ListCount - 1) Then
'2016.05.18                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(�d����(1))"
                    Text1(ptxSHI_CODE1).SetFocus
                    Exit Function
                End If
                        
            
            
            
            End If
        
        Case ptxSHI_TANKA1         '�d���P��(1)
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_TANKA1).Text) = "" Then
            
                Text1(ptxSHI_ARARI1).Text = ""
                Text1(ptxSHI_ARARI_RITU1).Text = ""
            
            Else
                
                If Not IsNumeric(Text1(ptxSHI_TANKA1).Text) Then
'2016.05.18                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(�d���P��(1))"
                    Text1(ptxSHI_TANKA1).SetFocus
                    Exit Function
                End If
            
                Text1(ptxSHI_TANKA1).Text = Format(CDbl(Text1(ptxSHI_TANKA1).Text), "#0.00")
            
                If Not IsNumeric(Text1(ptxSHI_TANKA1).Text) Or _
                    Not IsNumeric(Text1(ptxG_ST_URITAN).Text) Then

                    Text1(ptxSHI_ARARI1).Text = ""
                    Text1(ptxSHI_ARARI_RITU1).Text = ""
                Else
                    Text1(ptxSHI_ARARI1).Text = Format(CDbl(Text1(ptxG_ST_URITAN).Text) - CDbl(Text1(ptxSHI_TANKA1).Text), "#0.00")
                    
                    If CDbl(Text1(ptxG_ST_URITAN).Text) = 0 Then
                        Text1(ptxSHI_ARARI_RITU1).Text = "0.00"
                    Else
                        Text1(ptxSHI_ARARI_RITU1).Text = Format(CDbl(Text1(ptxSHI_ARARI1).Text) / CDbl(Text1(ptxG_ST_URITAN).Text) * 100, "#0.00")
                    End If
                End If
            
            End If
                
        Case ptxSHI_TANKA_DT1     '�d���P���ݒ��(1)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_TANKA_DT1).Text) = "" Then
                If Trim(Text1(ptxSHI_TANKA1).Text) <> "" Then
                    Text1(ptxSHI_TANKA_DT1).Text = Format(Now, "YYYY/MM/DD")
                End If
            Else
                If Not IsDate(Text1(ptxSHI_TANKA_DT1).Text) Then
'2016.05.18                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(�d���P���ݒ��(1))"
                    Text1(ptxSHI_TANKA_DT1).SetFocus
                    Exit Function
                End If
            
                Text1(ptxSHI_TANKA_DT1).Text = Format(Text1(ptxSHI_TANKA_DT1).Text, "YYYY/MM/DD")
            End If
                
        
        Case ptxSHI_LOT1           'ۯĐ�(1)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_LOT1).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxSHI_LOT1).Text) Then
'2016.05.18                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(ۯĐ�(1))"
                    Text1(ptxSHI_LOT1).SetFocus
                    Exit Function
                End If
        
                Text1(ptxSHI_LOT1).Text = Format(CLng(Text1(ptxSHI_LOT1).Text), "#0")
        
            End If
        
        Case ptxSHI_LEAD_TIME1     'ذ�����(1)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_LEAD_TIME1).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxSHI_LEAD_TIME1).Text) Then
'2016.05.18                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(ذ�����(1))"
                    Text1(ptxSHI_LEAD_TIME1).SetFocus
                    Exit Function
                End If
        
                Text1(ptxSHI_LEAD_TIME1).Text = Format(CLng(Text1(ptxSHI_LEAD_TIME1).Text), "#0")
            
            End If
        
        Case ptxSHI_LAST_ORDER_DT1     '�O�񒍕���(1)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_LAST_ORDER_DT1).Text) = "" Then
            Else
                If Not IsDate(Text1(ptxSHI_LAST_ORDER_DT1).Text) Then
'2016.05.18                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(�O�񒍕���(1))"
                    Text1(ptxSHI_LAST_ORDER_DT1).SetFocus
                    Exit Function
                End If
            
                Text1(ptxSHI_LAST_ORDER_DT1).Text = Format(Text1(ptxSHI_LAST_ORDER_DT1).Text, "YYYY/MM/DD")
            End If
        
        Case ptxSHI_LAST_ORDER_QTY1    '�O�񒍕���(1)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_LAST_ORDER_QTY1).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxSHI_LAST_ORDER_QTY1).Text) Then
'2016.05.18                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B('�O�񒍕���(1))"
                    Text1(ptxSHI_LAST_ORDER_QTY1).SetFocus
                    Exit Function
                End If
        
                Text1(ptxSHI_LAST_ORDER_QTY1).Text = Format(CLng(Text1(ptxSHI_LAST_ORDER_QTY1).Text), "#0")
        
            End If
        
        Case ptxSHI_CODE2          '�d����(2)
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
            Text1(Mode).Text = StrConv(Text1(Mode).Text, vbUpperCase) '2016.01.26
            If Trim(Text1(ptxSHI_CODE2).Text) = "" Then
                For i = ptxSHI_CODE2 To ptxSHI_LAST_ORDER_QTY2
                    Text1(i).Text = ""
                Next i
                Combo1(pcmbSHIIRE2).ListIndex = -1
            Else
                For i = 0 To Combo1(pcmbSHIIRE2).ListCount - 1
                    If Trim(Text1(ptxSHI_CODE2).Text) = Trim(Right(Combo1(pcmbSHIIRE2).List(i), 5)) Then
                        Combo1(pcmbSHIIRE2).ListIndex = i
                        Exit For
                    End If
                Next i
                
                If i > (Combo1(pcmbSHIIRE2).ListCount - 1) Then
'2016.05.18                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(�d����(2))"
                    Text1(ptxSHI_CODE2).SetFocus
                    Exit Function
                End If
                        
            
            
            
            End If
        
        Case ptxSHI_TANKA2         '�d���P��(2)
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_TANKA2).Text) = "" Then
            
                Text1(ptxSHI_ARARI2).Text = ""
                Text1(ptxSHI_ARARI_RITU2).Text = ""
            
            Else
                
                If Not IsNumeric(Text1(ptxSHI_TANKA2).Text) Then
'2016.05.18                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B�d���P��(2)"
                    Text1(ptxSHI_TANKA2).SetFocus
                    Exit Function
                End If
            
                Text1(ptxSHI_TANKA2).Text = Format(CDbl(Text1(ptxSHI_TANKA2).Text), "#0.00")
            
                If Not IsNumeric(Text1(ptxG_ST_URITAN).Text) Or _
                    Not IsNumeric(Text1(ptxG_ST_URITAN).Text) Then
                                        
                    Text1(ptxSHI_ARARI2).Text = ""
                    Text1(ptxSHI_ARARI_RITU2).Text = ""
                Else
                    Text1(ptxSHI_ARARI2).Text = Format(CDbl(Text1(ptxG_ST_URITAN).Text) - CDbl(Text1(ptxSHI_TANKA2).Text), "#0.00")
                    If CDbl(Text1(ptxG_ST_URITAN).Text) = 0 Then
                        Text1(ptxSHI_ARARI_RITU2).Text = "0.00"
                    Else
                        Text1(ptxSHI_ARARI_RITU2).Text = Format(CDbl(Text1(ptxSHI_ARARI2).Text) / CDbl(Text1(ptxG_ST_URITAN).Text) * 100, "#0.00")
                    End If
                End If
            
            End If
        
        
        
        
        
                
        Case ptxSHI_TANKA_DT2     '�d���P���ݒ��(2)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_TANKA_DT2).Text) = "" Then
                If Trim(Text1(ptxSHI_TANKA2).Text) <> "" Then
                    Text1(ptxSHI_TANKA_DT2).Text = Format(Now, "YYYY/MM/DD")
                End If
            Else
                If Not IsDate(Text1(ptxSHI_TANKA_DT2).Text) Then
'                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B�d���P���ݒ��(2)"
                    Text1(ptxSHI_TANKA_DT2).SetFocus
                    Exit Function
                End If
            
                Text1(ptxSHI_TANKA_DT2).Text = Format(Text1(ptxSHI_TANKA_DT2).Text, "YYYY/MM/DD")
            End If
                
        
        Case ptxSHI_LOT2           'ۯĐ�(2)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_LOT2).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxSHI_LOT2).Text) Then
'                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��BۯĐ�(2)"
                    Text1(ptxSHI_LOT2).SetFocus
                    Exit Function
                End If
        
                Text1(ptxSHI_LOT2).Text = Format(CLng(Text1(ptxSHI_LOT2).Text), "#0")
        
            End If
        
        Case ptxSHI_LEAD_TIME2     'ذ�����(2)
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
            
            If Trim(Text1(ptxSHI_LEAD_TIME2).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxSHI_LEAD_TIME2).Text) Then
'                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(ذ�����(2))"
                    Text1(ptxSHI_LEAD_TIME2).SetFocus
                    Exit Function
                End If
        
                Text1(ptxSHI_LEAD_TIME2).Text = Format(CLng(Text1(ptxSHI_LEAD_TIME2).Text), "#0")
            
            End If
        
        Case ptxSHI_LAST_ORDER_DT2     '�O�񒍕���(2)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_LAST_ORDER_DT2).Text) = "" Then
            Else
                If Not IsDate(Text1(ptxSHI_LAST_ORDER_DT2).Text) Then
'                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(�O�񒍕���(2))"
                    Text1(ptxSHI_LAST_ORDER_DT2).SetFocus
                    Exit Function
                End If
            
                Text1(ptxSHI_LAST_ORDER_DT2).Text = Format(Text1(ptxSHI_LAST_ORDER_DT2).Text, "YYYY/MM/DD")
            End If
        
        Case ptxSHI_LAST_ORDER_QTY2    '�O�񒍕���(2)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_LAST_ORDER_QTY2).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxSHI_LAST_ORDER_QTY2).Text) Then
'                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(�O�񒍕���(2))"
                    Text1(ptxSHI_LAST_ORDER_QTY2).SetFocus
                    Exit Function
                End If
        
                Text1(ptxSHI_LAST_ORDER_QTY2).Text = Format(CLng(Text1(ptxSHI_LAST_ORDER_QTY2).Text), "#0")
        
            End If
        Case ptxSHI_CODE3          '�d����(3)
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
            Text1(Mode).Text = StrConv(Text1(Mode).Text, vbUpperCase) '2016.01.26
            If Trim(Text1(ptxSHI_CODE3).Text) = "" Then
                For i = ptxSHI_CODE3 To ptxSHI_LAST_ORDER_QTY3
                    Text1(i).Text = ""
                Next i
                Combo1(pcmbSHIIRE3).ListIndex = -1
            Else
                For i = 0 To Combo1(pcmbSHIIRE3).ListCount - 1
                    If Trim(Text1(ptxSHI_CODE3).Text) = Trim(Right(Combo1(pcmbSHIIRE3).List(i), 5)) Then
                        Combo1(pcmbSHIIRE3).ListIndex = i
                        Exit For
                    End If
                Next i
                
                If i > (Combo1(pcmbSHIIRE3).ListCount - 1) Then
'                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(�d����(3))"
                    Text1(ptxSHI_CODE3).SetFocus
                    Exit Function
                End If
                        
            
            
            
            End If
        
        Case ptxSHI_TANKA3         '�d���P��(3)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
        
            If Trim(Text1(ptxSHI_TANKA3).Text) = "" Then
            
                Text1(ptxSHI_ARARI3).Text = ""
                Text1(ptxSHI_ARARI_RITU3).Text = ""
            
            Else
                
                If Not IsNumeric(Text1(ptxSHI_TANKA3).Text) Then
'                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(�d���P��(3))"
                    Text1(ptxSHI_TANKA3).SetFocus
                    Exit Function
                End If
            
                Text1(ptxSHI_TANKA3).Text = Format(CDbl(Text1(ptxSHI_TANKA3).Text), "#0.00")
            
                If Not IsNumeric(Text1(ptxG_ST_URITAN).Text) Or _
                    Not IsNumeric(Text1(ptxG_ST_URITAN).Text) Then
                                        
                    Text1(ptxSHI_ARARI3).Text = ""
                    Text1(ptxSHI_ARARI_RITU3).Text = ""
                Else
                    Text1(ptxSHI_ARARI3).Text = Format(CDbl(Text1(ptxG_ST_URITAN).Text) - CDbl(Text1(ptxSHI_TANKA3).Text), "#0.00")
                                    
                    If CDbl(Text1(ptxG_ST_URITAN).Text) = 0 Then
                        Text1(ptxSHI_ARARI_RITU3).Text = "0.00"
                    Else
                        Text1(ptxSHI_ARARI_RITU3).Text = Format(CDbl(Text1(ptxSHI_ARARI3).Text) / CDbl(Text1(ptxG_ST_URITAN).Text) * 100, "#0.00")
                    End If
                End If
            
            End If
        
                
        Case ptxSHI_TANKA_DT3     '�d���P���ݒ��(3)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_TANKA_DT3).Text) = "" Then
                If Trim(Text1(ptxSHI_TANKA3).Text) <> "" Then
                    Text1(ptxSHI_TANKA_DT3).Text = Format(Now, "YYYY/MM/DD")
                End If
            Else
                If Not IsDate(Text1(ptxSHI_TANKA_DT3).Text) Then
'                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(�d���P���ݒ��(3))"
                    Text1(ptxSHI_TANKA_DT3).SetFocus
                    Exit Function
                End If
            
                Text1(ptxSHI_TANKA_DT3).Text = Format(Text1(ptxSHI_TANKA_DT3).Text, "YYYY/MM/DD")
            End If
                
        
        Case ptxSHI_LOT3           'ۯĐ�(3)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_LOT3).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxSHI_LOT3).Text) Then
'                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(ۯĐ�(3))"
                    Text1(ptxSHI_LOT3).SetFocus
                    Exit Function
                End If
        
                Text1(ptxSHI_LOT3).Text = Format(CLng(Text1(ptxSHI_LOT3).Text), "#0")
        
            End If
        
        Case ptxSHI_LEAD_TIME3     'ذ�����(3)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_LEAD_TIME3).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxSHI_LEAD_TIME3).Text) Then
'                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(ذ�����(3))"
                    Text1(ptxSHI_LEAD_TIME3).SetFocus
                    Exit Function
                End If
        
                Text1(ptxSHI_LEAD_TIME3).Text = Format(CLng(Text1(ptxSHI_LEAD_TIME3).Text), "#0")
            
            End If
        
        Case ptxSHI_TANKA_DT3     '�O�񒍕���(3)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_LAST_ORDER_DT3).Text) = "" Then
            Else
                If Not IsDate(Text1(ptxSHI_LAST_ORDER_DT3).Text) Then
'                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(�O�񒍕���(3))"
                    Text1(ptxSHI_LAST_ORDER_DT3).SetFocus
                    Exit Function
                End If
            
                Text1(ptxSHI_LAST_ORDER_DT3).Text = Format(Text1(ptxSHI_LAST_ORDER_DT2).Text, "YYYY/MM/DD")
            End If
        
        Case ptxSHI_LAST_ORDER_QTY3    '�O�񒍕���(3)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_LAST_ORDER_QTY3).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxSHI_LAST_ORDER_QTY3).Text) Then
'                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(�O�񒍕���(3))"
                    Text1(ptxSHI_LAST_ORDER_QTY3).SetFocus
                    Exit Function
                End If
        
                Text1(ptxSHI_LAST_ORDER_QTY3).Text = Format(CLng(Text1(ptxSHI_LAST_ORDER_QTY3).Text), "#0")
        
            End If
        
        
        Case ptxLAST_SYU_DT     '�ŏI�o�ɓ�
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
            If Trim(Text1(ptxLAST_SYU_DT).Text) = "" Then
            Else
                If Not IsDate(Text1(ptxLAST_SYU_DT).Text) Then
'                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(�ŏI�o�ɓ�)"
                    Text1(ptxLAST_SYU_DT).SetFocus
                    Exit Function
                End If
            
                Text1(ptxLAST_SYU_DT).Text = Format(Text1(ptxLAST_SYU_DT).Text, "YYYY/MM/DD")
            End If
        
        Case ptxG_LAST_SYUKA_QTY    '�ŏI�o�ɐ�
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
            If Trim(Text1(ptxG_LAST_SYUKA_QTY).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxG_LAST_SYUKA_QTY).Text) Then
'                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(�ŏI�o�ɐ�)"
                    Text1(ptxG_LAST_SYUKA_QTY).SetFocus
                    Exit Function
                End If
            
                Text1(ptxG_LAST_SYUKA_QTY).Text = Format(CLng(Text1(ptxG_LAST_SYUKA_QTY).Text), "#0")
            End If
        
        
        
        
        
        
        Case ptxLAST_CODE          '�ŐV�d����      2007.05.28
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            Text1(Mode).Text = StrConv(Text1(Mode).Text, vbUpperCase) '2016.01.26
            If Trim(Text1(ptxLAST_CODE).Text) = "" Then
                Combo1(pcmbLAST_CODE).ListIndex = -1
            Else
                For i = 0 To Combo1(pcmbLAST_CODE).ListCount - 1
                    If Trim(Text1(ptxLAST_CODE).Text) = Trim(Right(Combo1(pcmbLAST_CODE).List(i), 5)) Then
                        Combo1(pcmbLAST_CODE).ListIndex = i
                        Exit For
                    End If
                Next i
                
                If i > (Combo1(pcmbLAST_CODE).ListCount - 1) Then
'                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(�ŐV�d����)"
                    Text1(ptxLAST_CODE).SetFocus
                    Exit Function
                End If
                        
            
            
            
            End If
        
        Case ptxLAST_TANKA          '�ŐV�d���P��    2007.05.28
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxLAST_TANKA).Text) = "" Then
            
            
            Else
                
                If Not IsNumeric(Text1(ptxLAST_TANKA).Text) Then
'                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "���͂������ڂ̓G���[�ł��B(�ŐV�d���P��)"
                    Text1(ptxLAST_TANKA).SetFocus
                    Exit Function
                End If
            
                Text1(ptxLAST_TANKA).Text = Format(CDbl(Text1(ptxLAST_TANKA).Text), "#0.00")
            
            
            End If
        
        
        
        
        
        Case ptxKUTI_SU             '����    2010.01.18
        
        
            If Trim(Text1(ptxKUTI_SU).Text) = "" Then
            
            
            Else
                
                If Not IsNumeric(Text1(ptxKUTI_SU).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(����)"
                    Text1(ptxKUTI_SU).SetFocus
                    Exit Function
                End If
            
                Text1(ptxKUTI_SU).Text = Format(CCur(Text1(ptxKUTI_SU).Text), "#0.0")
            
            
            End If
        
        
        
        Case ptxKONPOU_F             '����敪    2010.01.18
        
        
            If Trim(Text1(ptxKONPOU_F).Text) = "" Or Trim(Text1(ptxKONPOU_F).Text) = "0" Or Trim(Text1(ptxKONPOU_F).Text) = "1" Then
            
            
            Else
                
                MsgBox "���͂������ڂ̓G���[�ł��B(����敪)"
                Text1(ptxKONPOU_F).SetFocus
                Exit Function
            
            
            
            End If
        
        
        
    End Select
        
    Error_Check_Proc = False
    Exit Function


Error_Proc:
    
    If Err.Number = 6 Then
        MsgBox "�����I�[�o�[�ł��B���͓��e���m�F���Ă��������B"
    
    
    
    Else
        MsgBox "���ُ͈�ł��B���͓��e���m�F���Ă��������B"
    End If
    Text1(Mode).SetFocus

End Function
Private Function Item_Disp_Proc(CODE As String) As Integer
'----------------------------------------------------------------------------
'                   ��ʕ\��
'----------------------------------------------------------------------------
Dim sts As Integer
Dim i   As Integer

    Item_Disp_Proc = True
    
    '�i��Ͻ��ǂݍ���
    Call UniCode_Conv(K0_ITEM.JGYOBU, Left(CODE, 1))
    Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(CODE, 2, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Right(CODE, 20))
    
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            'ں��ޓ��e�̕\��
                                            '�����O
            For i = 0 To Combo1(pcmbNAIGAI).ListCount - 1
                If Right(Combo1(pcmbNAIGAI).List(i), 1) = StrConv(ITEMREC.JGYOBU, vbUnicode) Then
                    Combo1(pcmbNAIGAI).ListIndex = i
                    Exit For
                End If
            Next i
                                            '�i�ں���
            Text1(ptxHIN_GAI).Text = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                            '�i��
            Text1(ptxHIN_NAME).Text = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                                            '�d���敪
            Text1(ptxG_SHIIRE_KBN).Text = StrConv(ITEMREC.G_SHIIRE_KBN, vbUnicode)
                                            '�d���敪����
            Combo1(pcmbG_SHIIRE).ListIndex = -1
            For i = 0 To Combo1(pcmbG_SHIIRE).ListCount - 1
                If Left(Right(Combo1(pcmbG_SHIIRE).List(i), 3), 2) = Text1(ptxG_SHIIRE_KBN).Text Then
                    Combo1(pcmbG_SHIIRE).ListIndex = i
                    Exit For
                End If
            Next i
                                            '�̔��敪
            Text1(ptxG_HANBAI_KBN).Text = StrConv(ITEMREC.G_HANBAI_KBN, vbUnicode)
                                            '�̔��敪����
            Combo1(pcmbG_HANBAI).ListIndex = -1
            For i = 0 To Combo1(pcmbG_HANBAI).ListCount - 1
                If Left(Right(Combo1(pcmbG_HANBAI).List(i), 3), 2) = Text1(ptxG_HANBAI_KBN).Text Then
                    Combo1(pcmbG_HANBAI).ListIndex = i
                    Exit For
                End If
            Next i
                                            '���x�P��
            Text1(ptxG_SYUSHI).Text = StrConv(ITEMREC.G_SYUSHI, vbUnicode)
                                            '���x�P�ʌ���
            Combo1(pcmbG_SYUSHI).ListIndex = -1
            For i = 0 To Combo1(pcmbG_SYUSHI).ListCount - 1
                If Right(Combo1(pcmbG_SYUSHI).List(i), 3) = Text1(ptxG_SYUSHI).Text Then
                    Combo1(pcmbG_SYUSHI).ListIndex = i
                    Exit For
                End If
            Next i
                                            '�g�����i
            If StrConv(ITEMREC.G_KUMITATE, vbUnicode) = P_ASSEMBLY_ON Then
                Check1(pchkG_KUMITATE).Value = vbChecked
            Else
                Check1(pchkG_KUMITATE).Value = vbUnchecked
            End If
                                            '�O���݌ɋ��z
            If Trim(StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode)) = "" Then
                Text1(ptxG_ZEN_ZAIKO_KIN).Text = ""
            Else
                Text1(ptxG_ZEN_ZAIKO_KIN).Text = Format(CLng(StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode)), "#0")
            End If
                                            '�O���݌ɐ���
            If Not IsNumeric(StrConv(ITEMREC.G_ZEN_ZAIKO_QTY, vbUnicode)) Then
                Text1(ptxG_ZEN_ZAIKO_QTY).Text = ""
            Else
                Text1(ptxG_ZEN_ZAIKO_QTY).Text = Format(CLng(StrConv(ITEMREC.G_ZEN_ZAIKO_QTY, vbUnicode)), "#0")
            End If
                                            '�W������
            If Not IsNumeric(Trim(StrConv(ITEMREC.G_ST_URITAN, vbUnicode))) Then
                Text1(ptxG_ST_URITAN).Text = ""
            Else
                Text1(ptxG_ST_URITAN).Text = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), "#0.00")
            End If
                                            '�W�������ݒ��
            If Trim(StrConv(ITEMREC.G_ST_URITAN_DT, vbUnicode)) = "" Then
                Text1(ptxG_ST_URITAN_DT).Text = ""
            Else
                Text1(ptxG_ST_URITAN_DT).Text = Left(StrConv(ITEMREC.G_ST_URITAN_DT, vbUnicode), 4) & "/" & _
                                                Mid(StrConv(ITEMREC.G_ST_URITAN_DT, vbUnicode), 5, 2) & "/" & _
                                                Right(StrConv(ITEMREC.G_ST_URITAN_DT, vbUnicode), 2)
            End If
                                            '�W������
            If Not IsNumeric(Trim(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode))) Then
                Text1(ptxG_ST_SHITAN).Text = ""
            Else
                Text1(ptxG_ST_SHITAN).Text = Format(CDbl(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)), "#0.00")
            End If
                                            '�W�������ݒ��
            If Trim(StrConv(ITEMREC.G_ST_SHITAN_DT, vbUnicode)) = "" Then
                Text1(ptxG_ST_SHITAN_DT).Text = ""
            Else
                Text1(ptxG_ST_SHITAN_DT).Text = Left(StrConv(ITEMREC.G_ST_SHITAN_DT, vbUnicode), 4) & "/" & _
                                                Mid(StrConv(ITEMREC.G_ST_SHITAN_DT, vbUnicode), 5, 2) & "/" & _
                                                Right(StrConv(ITEMREC.G_ST_SHITAN_DT, vbUnicode), 2)
            End If
                                            
                                            '�댯�݌�
            If Not IsNumeric(Trim(StrConv(ITEMREC.HOJYU_P, vbUnicode))) Then
                Text1(ptxHOJYU_P).Text = ""
            Else
                Text1(ptxHOJYU_P).Text = Format(CDbl(StrConv(ITEMREC.HOJYU_P, vbUnicode)), "#0")
            End If
                                            '���ދ敪
            Combo1(pcmbG_SHIZAI_KBN).ListIndex = -1
            For i = 0 To Combo1(pcmbG_SHIZAI_KBN).ListCount - 1
                If Right(Combo1(pcmbG_SHIZAI_KBN).List(i), 1) = StrConv(ITEMREC.G_SHIZAI_KBN, vbUnicode) Then
                    Combo1(pcmbG_SHIZAI_KBN).ListIndex = i
                    Exit For
                End If
            Next i
                                            '���ٓ\��v��Ȃ�
            If StrConv(ITEMREC.G_LABEL_NON, vbUnicode) = P_G_LABEL_ON Then
                Check1(pchkG_LABEL_NON).Value = vbUnchecked
            Else
                Check1(pchkG_LABEL_NON).Value = vbChecked
            End If
                                            
                                            '�݌ɊǗ��ΏۊO
            If StrConv(ITEMREC.ZAIKO_F, vbUnicode) = P_ZAIKO_F_ON Then
                Check1(pchkZAIKO_F).Value = vbUnchecked
            Else
                Check1(pchkZAIKO_F).Value = vbChecked
            End If
                                            '�I���\�@�݌ɐ���\��   2012.12.13
            If StrConv(ITEMREC.ZAIKO_CLR_F, vbUnicode) = "1" Then
                Check1(pchkZAIKO_CLR_F).Value = vbChecked
            Else
                Check1(pchkZAIKO_CLR_F).Value = vbUnchecked
            End If
                                            '�ː�   2008.02.14
            If IsNumeric(StrConv(ITEMREC.SAI_SU, vbUnicode)) Then
                Text1(ptxSAI_SU).Text = Format(CDbl(StrConv(ITEMREC.SAI_SU, vbUnicode)), "#0.00")
            Else
                Text1(ptxSAI_SU).Text = ""
            End If
                                            
                                            '�`�� 2008.02.14
            Text1(ptxD_KEISHIKI).Text = Trim(StrConv(ITEMREC.D_KEISHIKI, vbUnicode))
                                            '����ްٌ��� 2008.02.14
            Text1(ptxD_THICKNESS).Text = Trim(StrConv(ITEMREC.D_THICKNESS, vbUnicode))
                                            '����ްٍގ� 2008.02.14
            Text1(ptxD_MATERIAL).Text = Trim(StrConv(ITEMREC.D_MATERIAL, vbUnicode))
                                            
                                            
                                            '����ްٻ���(W) 2008.02.14
            If IsNumeric(StrConv(ITEMREC.D_SIZE_W, vbUnicode)) Then
                Text1(ptxD_SIZE_W).Text = Format(Val(Trim(StrConv(ITEMREC.D_SIZE_W, vbUnicode))), "#0")
            Else
                Text1(ptxD_SIZE_W).Text = Trim(StrConv(ITEMREC.D_SIZE_W, vbUnicode))
            End If
                                            
                                            
                                            '����ްٻ���(D) 2008.02.14
                                            
            If IsNumeric(StrConv(ITEMREC.D_SIZE_D, vbUnicode)) Then
                Text1(ptxD_SIZE_D).Text = Format(Val(Trim(StrConv(ITEMREC.D_SIZE_D, vbUnicode))), "#0")
            Else
                Text1(ptxD_SIZE_D).Text = Trim(StrConv(ITEMREC.D_SIZE_D, vbUnicode))
            End If
                                            
                                            
                                            '����ްٻ���(H) 2008.02.14
            If IsNumeric(StrConv(ITEMREC.D_SIZE_H, vbUnicode)) Then
                Text1(ptxD_SIZE_H).Text = Format(Val(Trim(StrConv(ITEMREC.D_SIZE_H, vbUnicode))), "#0")
            Else
                Text1(ptxD_SIZE_H).Text = Trim(StrConv(ITEMREC.D_SIZE_H, vbUnicode))
            End If
                                            
                                            
            If IsNumeric(Text1(ptxD_SIZE_W).Text) And IsNumeric(Text1(ptxD_SIZE_D).Text) And IsNumeric(Text1(ptxD_SIZE_H).Text) Then
            
                
            
                lblSIZE.Caption = Format(ToHalfAdjust(CCur(Val(Text1(ptxD_SIZE_W).Text) / 1000 * Val(Text1(ptxD_SIZE_D).Text) / 1000 * Val(Text1(ptxD_SIZE_H).Text) / 1000), 4), "#0.000")
                                                
            End If
'            lblSIZE.Caption = (Val(Text1(ptxD_SIZE_W).Text) * Val(Text1(ptxD_SIZE_D).Text) * Val(Text1(ptxD_SIZE_H).Text)) / 1000000000
                                            
                                            
                                            
                                            '�������^���Ȃ�  2008.02.14
            Text1(ptxD_PRINT).Text = Trim(StrConv(ITEMREC.D_PRINT, vbUnicode))
                                            
                                            '���i���@�H��   2008.02.14
            If IsNumeric(StrConv(ITEMREC.S_KOUSU, vbUnicode)) Then
                Text1(ptxS_KOUSU).Text = Format(CDbl(StrConv(ITEMREC.S_KOUSU, vbUnicode)), "#")
            Else
                Text1(ptxS_KOUSU).Text = ""
            End If
                                            '�W������       2008.07.16
            If IsNumeric(StrConv(ITEMREC.SEI_SYU_KON, vbUnicode)) Then
                Text1(ptxSEI_SYU_KON).Text = Format(CDbl(StrConv(ITEMREC.SEI_SYU_KON, vbUnicode)), "#")
            Else
                Text1(ptxSEI_SYU_KON).Text = ""
            End If
                                            
                                            
                                            '�A�����o���׸� 2008.02.14
            Text1(ptxSE_USOU_F).Text = Trim(StrConv(ITEMREC.SE_USOU_F, vbUnicode))
                                            '�g�p�e�[�v��� 2008.02.14
            Text1(ptxUSE_TAPE_KIND).Text = Trim(StrConv(ITEMREC.USE_TAPE_KIND, vbUnicode))
                                            '�g�p�e�[�v���� 2008.02.14
            Text1(ptxUSE_TAPE_LNG).Text = Trim(StrConv(ITEMREC.USE_TAPE_LNG, vbUnicode))
                                            '�����敪       2008.07.16
            Text1(ptxSEI_KBN).Text = Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode))
                                            
                                            
                                            '�W���I��       2009.09.01
            Text1(ptxST_SOKO).Text = Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode))
                                            '�W���I��       2009.09.01
            Text1(ptxST_RETU).Text = Trim(StrConv(ITEMREC.ST_RETU, vbUnicode))
                                            '�W���I��       2009.09.01
            Text1(ptxST_REN).Text = Trim(StrConv(ITEMREC.ST_REN, vbUnicode))
                                            '�W���I��       2009.09.01
            Text1(ptxST_DAN).Text = Trim(StrConv(ITEMREC.ST_DAN, vbUnicode))
                                            
                                            
                                            
                                            '����           2010.01.18
            Text1(ptxKUTI_SU).Text = Trim(StrConv(ITEMREC.KUTI_SU, vbUnicode))
                                            '����敪           2010.01.18
            Text1(ptxKONPOU_F).Text = Trim(StrConv(ITEMREC.KONPOU_F, vbUnicode))
                                            
                                            
                                            
                                            
                                            '�d����-��(1)
            Text1(ptxSHI_CODE1).Text = Trim(StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode))
                                            '�d���於��(1)
            Combo1(pcmbSHIIRE1).ListIndex = -1
            For i = 0 To Combo1(pcmbSHIIRE1).ListCount - 1
                If Trim(Right(Combo1(pcmbSHIIRE1).List(i), 5)) = Text1(ptxSHI_CODE1).Text Then
                    Combo1(pcmbSHIIRE1).ListIndex = i
                    Exit For
                End If
            Next i
                                            '�d���P��(1)
            If Not IsNumeric(Trim(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode))) Then
                Text1(ptxSHI_TANKA1).Text = ""
            Else
                Text1(ptxSHI_TANKA1).Text = Format(CDbl(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode)), "#0.00")
            End If
                                            '�d���P���ݒ��(1)
            If Trim(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA_DT, vbUnicode)) = "" Then
                Text1(ptxSHI_TANKA_DT1).Text = ""
            Else
                Text1(ptxSHI_TANKA_DT1).Text = Left(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA_DT, vbUnicode), 4) & "/" & _
                                                Mid(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA_DT, vbUnicode), 5, 2) & "/" & _
                                                Right(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA_DT, vbUnicode), 2)
            End If
                                            '�e���^�e����(1)
            If Not IsNumeric(Text1(ptxG_ST_URITAN).Text) Or _
                Not IsNumeric(Text1(ptxSHI_TANKA1).Text) Then
                                    
                Text1(ptxSHI_ARARI1).Text = ""
                Text1(ptxSHI_ARARI_RITU1).Text = ""
            Else
                Text1(ptxSHI_ARARI1).Text = Format(CDbl(Text1(ptxG_ST_URITAN).Text) - CDbl(Text1(ptxSHI_TANKA1).Text), "#0.00")
                If CDbl(Text1(ptxG_ST_URITAN).Text) = 0 Then
                    Text1(ptxSHI_ARARI_RITU1).Text = "0.00"
                Else
                    Text1(ptxSHI_ARARI_RITU1).Text = Format(CDbl(Text1(ptxSHI_ARARI1).Text) / CDbl(Text1(ptxG_ST_URITAN).Text) * 100, "#0.00")
                End If
            End If
                                            'ۯĐ�(1)
            If Not IsNumeric(Trim(StrConv(ITEMREC.G_SHIIRE_TBL(0).LOT, vbUnicode))) Then
                Text1(ptxSHI_LOT1).Text = ""
            Else
                Text1(ptxSHI_LOT1).Text = Format(CDbl(StrConv(ITEMREC.G_SHIIRE_TBL(0).LOT, vbUnicode)), "#0")
            End If
                                            'ذ�����(1)
            If Not IsNumeric(Trim(StrConv(ITEMREC.G_SHIIRE_TBL(0).LEAD_TIME, vbUnicode))) Then
                Text1(ptxSHI_LEAD_TIME1).Text = ""
            Else
                Text1(ptxSHI_LEAD_TIME1).Text = Format(CLng(StrConv(ITEMREC.G_SHIIRE_TBL(0).LEAD_TIME, vbUnicode)), "#0")
            End If
                                            '�O�񒍕���(1)
            If Trim(StrConv(ITEMREC.G_SHIIRE_TBL(0).LAST_ORDER_DT, vbUnicode)) = "" Then
                Text1(ptxSHI_LAST_ORDER_DT1).Text = ""
            Else
                Text1(ptxSHI_LAST_ORDER_DT1).Text = Left(StrConv(ITEMREC.G_SHIIRE_TBL(0).LAST_ORDER_DT, vbUnicode), 4) & "/" & _
                                                Mid(StrConv(ITEMREC.G_SHIIRE_TBL(0).LAST_ORDER_DT, vbUnicode), 5, 2) & "/" & _
                                                Right(StrConv(ITEMREC.G_SHIIRE_TBL(0).LAST_ORDER_DT, vbUnicode), 2)
            End If
                                            '�O�񒍕���(1)
            If Not IsNumeric(Trim(StrConv(ITEMREC.G_SHIIRE_TBL(0).LAST_ORDER_QTY, vbUnicode))) Then
                Text1(ptxSHI_LAST_ORDER_QTY1).Text = ""
            Else
                Text1(ptxSHI_LAST_ORDER_QTY1).Text = Format(CLng(StrConv(ITEMREC.G_SHIIRE_TBL(0).LAST_ORDER_QTY, vbUnicode)), "#0")
            End If
            
            
            
            
            
            
                                            '�d����-��(2)
            Text1(ptxSHI_CODE2).Text = Trim(StrConv(ITEMREC.G_SHIIRE_TBL(1).CODE, vbUnicode))
                                            '�d���於��(2)
            Combo1(pcmbSHIIRE2).ListIndex = -1
            For i = 0 To Combo1(pcmbSHIIRE2).ListCount - 1
                If Trim(Right(Combo1(pcmbSHIIRE2).List(i), 5)) = Text1(ptxSHI_CODE2).Text Then
                    Combo1(pcmbSHIIRE2).ListIndex = i
                    Exit For
                End If
            Next i
                                            '�d���P��(2)
            If Not IsNumeric(Trim(StrConv(ITEMREC.G_SHIIRE_TBL(1).TANKA, vbUnicode))) Then
                Text1(ptxSHI_TANKA2).Text = ""
            Else
                Text1(ptxSHI_TANKA2).Text = Format(CDbl(StrConv(ITEMREC.G_SHIIRE_TBL(1).TANKA, vbUnicode)), "#0.00")
            End If
                                            '�d���P���ݒ��(2)
            If Trim(StrConv(ITEMREC.G_SHIIRE_TBL(1).TANKA_DT, vbUnicode)) = "" Then
                Text1(ptxSHI_TANKA_DT2).Text = ""
            Else
                Text1(ptxSHI_TANKA_DT2).Text = Left(StrConv(ITEMREC.G_SHIIRE_TBL(1).TANKA_DT, vbUnicode), 4) & "/" & _
                                                Mid(StrConv(ITEMREC.G_SHIIRE_TBL(1).TANKA_DT, vbUnicode), 5, 2) & "/" & _
                                                Right(StrConv(ITEMREC.G_SHIIRE_TBL(1).TANKA_DT, vbUnicode), 2)
            End If
                                            '�e���^�e����(2)
            If Not IsNumeric(Text1(ptxG_ST_URITAN).Text) Or _
                Not IsNumeric(Text1(ptxSHI_TANKA2).Text) Then
                Text1(ptxSHI_ARARI2).Text = ""
                Text1(ptxSHI_ARARI_RITU2).Text = ""
            Else
                Text1(ptxSHI_ARARI2).Text = Format(CDbl(Text1(ptxG_ST_URITAN).Text) - CDbl(Text1(ptxSHI_TANKA2).Text), "#0.00")
                If CDbl(Text1(ptxG_ST_URITAN).Text) = 0 Then
                    Text1(ptxSHI_ARARI_RITU2).Text = "0.00"
                Else
                    Text1(ptxSHI_ARARI_RITU2).Text = Format(CDbl(Text1(ptxSHI_ARARI2).Text) / CDbl(Text1(ptxG_ST_URITAN).Text) * 100, "#0.00")
                End If
            End If
                                            'ۯĐ�(2)
            If Not IsNumeric(Trim(StrConv(ITEMREC.G_SHIIRE_TBL(1).LOT, vbUnicode))) Then
                Text1(ptxSHI_LOT2).Text = ""
            Else
                Text1(ptxSHI_LOT2).Text = Format(CDbl(StrConv(ITEMREC.G_SHIIRE_TBL(1).LOT, vbUnicode)), "#0")
            End If
                                            'ذ�����(2)
            If Not IsNumeric(Trim(StrConv(ITEMREC.G_SHIIRE_TBL(1).LEAD_TIME, vbUnicode))) Then
                Text1(ptxSHI_LEAD_TIME2).Text = ""
            Else
                Text1(ptxSHI_LEAD_TIME2).Text = Format(CDbl(StrConv(ITEMREC.G_SHIIRE_TBL(1).LEAD_TIME, vbUnicode)), "#0")
            End If
                                            '�O�񒍕���(2)
            If Trim(StrConv(ITEMREC.G_SHIIRE_TBL(1).LAST_ORDER_DT, vbUnicode)) = "" Then
                Text1(ptxSHI_LAST_ORDER_DT2).Text = ""
            Else
                Text1(ptxSHI_LAST_ORDER_DT2).Text = Left(StrConv(ITEMREC.G_SHIIRE_TBL(1).LAST_ORDER_DT, vbUnicode), 4) & "/" & _
                                                Mid(StrConv(ITEMREC.G_SHIIRE_TBL(1).LAST_ORDER_DT, vbUnicode), 5, 2) & "/" & _
                                                Right(StrConv(ITEMREC.G_SHIIRE_TBL(1).LAST_ORDER_DT, vbUnicode), 2)
            End If
                                            '�O�񒍕���(2)
            If Not IsNumeric(Trim(StrConv(ITEMREC.G_SHIIRE_TBL(1).LAST_ORDER_QTY, vbUnicode))) Then
                Text1(ptxSHI_LAST_ORDER_QTY2).Text = ""
            Else
                Text1(ptxSHI_LAST_ORDER_QTY2).Text = Format(CLng(StrConv(ITEMREC.G_SHIIRE_TBL(1).LAST_ORDER_QTY, vbUnicode)), "#0")
            End If
            
            
            
            
                                            '�d����-��(3)
            Text1(ptxSHI_CODE3).Text = Trim(StrConv(ITEMREC.G_SHIIRE_TBL(2).CODE, vbUnicode))
                                            '�d���於��(3)
            Combo1(pcmbSHIIRE3).ListIndex = -1
            For i = 0 To Combo1(pcmbSHIIRE3).ListCount - 1
                If Trim(Right(Combo1(pcmbSHIIRE3).List(i), 5)) = Text1(ptxSHI_CODE3).Text Then
                    Combo1(pcmbSHIIRE3).ListIndex = i
                    Exit For
                End If
            Next i
                                            '�d���P��(3)
            If Not IsNumeric(Trim(StrConv(ITEMREC.G_SHIIRE_TBL(2).TANKA, vbUnicode))) Then
                Text1(ptxSHI_TANKA3).Text = ""
            Else
                Text1(ptxSHI_TANKA3).Text = Format(CDbl(StrConv(ITEMREC.G_SHIIRE_TBL(2).TANKA, vbUnicode)), "#0.00")
            End If
                                            '�d���P���ݒ��(3)
            If Trim(StrConv(ITEMREC.G_SHIIRE_TBL(2).TANKA_DT, vbUnicode)) = "" Then
                Text1(ptxSHI_TANKA_DT3).Text = ""
            Else
                Text1(ptxSHI_TANKA_DT3).Text = Left(StrConv(ITEMREC.G_SHIIRE_TBL(2).TANKA_DT, vbUnicode), 4) & "/" & _
                                                Mid(StrConv(ITEMREC.G_SHIIRE_TBL(2).TANKA_DT, vbUnicode), 5, 2) & "/" & _
                                                Right(StrConv(ITEMREC.G_SHIIRE_TBL(2).TANKA_DT, vbUnicode), 2)
            End If
                                            '�e���^�e����(3)
            If Not IsNumeric(Text1(ptxG_ST_URITAN).Text) Or _
                Not IsNumeric(Text1(ptxSHI_TANKA3).Text) Then
                Text1(ptxSHI_ARARI3).Text = ""
                Text1(ptxSHI_ARARI_RITU3).Text = ""
            Else
                Text1(ptxSHI_ARARI3).Text = Format(CDbl(Text1(ptxG_ST_URITAN).Text) - CDbl(Text1(ptxSHI_TANKA3).Text), "#0.00")
                If CDbl(Text1(ptxG_ST_URITAN).Text) = 0 Then
                    Text1(ptxSHI_ARARI_RITU3).Text = "0.00"
                Else
                    Text1(ptxSHI_ARARI_RITU3).Text = Format(CDbl(Text1(ptxSHI_ARARI3).Text) / CDbl(Text1(ptxG_ST_URITAN).Text) * 100, "#0.00")
                End If
            End If
                                            'ۯĐ�(3)
            If Not IsNumeric(Trim(StrConv(ITEMREC.G_SHIIRE_TBL(2).LOT, vbUnicode))) Then
                Text1(ptxSHI_LOT3).Text = ""
            Else
                Text1(ptxSHI_LOT3).Text = Format(CDbl(StrConv(ITEMREC.G_SHIIRE_TBL(2).LOT, vbUnicode)), "#0")
            End If
                                            'ذ�����(3)
            If Not IsNumeric(Trim(StrConv(ITEMREC.G_SHIIRE_TBL(2).LEAD_TIME, vbUnicode))) Then
                Text1(ptxSHI_LEAD_TIME3).Text = ""
            Else
                Text1(ptxSHI_LEAD_TIME3).Text = Format(CDbl(StrConv(ITEMREC.G_SHIIRE_TBL(2).LEAD_TIME, vbUnicode)), "#0")
            End If
                                            '�O�񒍕���(3)
            If Trim(StrConv(ITEMREC.G_SHIIRE_TBL(2).LAST_ORDER_DT, vbUnicode)) = "" Then
                Text1(ptxSHI_LAST_ORDER_DT3).Text = ""
            Else
                Text1(ptxSHI_LAST_ORDER_DT3).Text = Left(StrConv(ITEMREC.G_SHIIRE_TBL(2).LAST_ORDER_DT, vbUnicode), 4) & "/" & _
                                                Mid(StrConv(ITEMREC.G_SHIIRE_TBL(2).LAST_ORDER_DT, vbUnicode), 5, 2) & "/" & _
                                                Right(StrConv(ITEMREC.G_SHIIRE_TBL(2).LAST_ORDER_DT, vbUnicode), 2)
            End If
                                            '�O�񒍕���(3)
            If Not IsNumeric(Trim(StrConv(ITEMREC.G_SHIIRE_TBL(2).LAST_ORDER_QTY, vbUnicode))) Then
                Text1(ptxSHI_LAST_ORDER_QTY3).Text = ""
            Else
                Text1(ptxSHI_LAST_ORDER_QTY3).Text = Format(CLng(StrConv(ITEMREC.G_SHIIRE_TBL(2).LAST_ORDER_QTY, vbUnicode)), "#0")
            End If
            



                                            '�ŏI�o�ɓ�
            If Trim(StrConv(ITEMREC.LAST_SYU_DT, vbUnicode)) = "" Then
                Text1(ptxLAST_SYU_DT).Text = ""
            Else
                Text1(ptxLAST_SYU_DT).Text = Left(StrConv(ITEMREC.LAST_SYU_DT, vbUnicode), 4) & "/" & _
                                                Mid(StrConv(ITEMREC.LAST_SYU_DT, vbUnicode), 5, 2) & "/" & _
                                                Right(StrConv(ITEMREC.LAST_SYU_DT, vbUnicode), 2)
            End If
                                            
                                            '�ŏI�o�ɐ�
            If Not IsNumeric(StrConv(ITEMREC.G_LAST_SYUKA_QTY, vbUnicode)) Then
                Text1(ptxG_LAST_SYUKA_QTY).Text = ""
            Else
                Text1(ptxG_LAST_SYUKA_QTY).Text = Format(CDbl(StrConv(ITEMREC.G_LAST_SYUKA_QTY, vbUnicode)), "#0")
            End If


                                            '�ŐV�d���� 2007.05.28
            Text1(ptxLAST_CODE).Text = Trim(StrConv(ITEMREC.LAST_CODE, vbUnicode))
            Combo1(pcmbLAST_CODE).ListIndex = -1
            
            For i = 0 To Combo1(pcmbLAST_CODE).ListCount - 1
                If Trim(Right(Combo1(pcmbLAST_CODE).List(i), 5)) = Text1(ptxLAST_CODE).Text Then
                    Combo1(pcmbLAST_CODE).ListIndex = i
                    Exit For
                End If
            Next i
                                            '�ŐV�P��   2007.05.28
            If Not IsNumeric(Trim(StrConv(ITEMREC.LAST_TANKA, vbUnicode))) Then
                Text1(ptxLAST_TANKA).Text = ""
            Else
                Text1(ptxLAST_TANKA).Text = Format(CDbl(StrConv(ITEMREC.LAST_TANKA, vbUnicode)), "#0.00")
            End If

        
        
        
        
        
        
                                            '�d�����l   2018.04.19
            Text1(ptxSHIIRE_BIKOU).Text = StrConv(ITEMREC.SHIIRE_BIKOU, vbUnicode)
        
        
        
                                            '�ǉ��S����/����    2010.01.18
            lblIns_DateTime = StrConv(ITEMREC.INS_TANTO, vbUnicode) & " " & Mid(StrConv(ITEMREC.Ins_DateTime, vbUnicode), 1, 8) & "-" & Mid(StrConv(ITEMREC.Ins_DateTime, vbUnicode), 9, 4)
                                            
                                            
                                            
                                            '�X�V�S����/����    2010.01.18
            lblUpd_DateTime = StrConv(ITEMREC.UPD_TANTO, vbUnicode) & " " & Mid(StrConv(ITEMREC.UPD_DATETIME, vbUnicode), 1, 8) & "-" & Mid(StrConv(ITEMREC.UPD_DATETIME, vbUnicode), 9, 4)
        
        
        
        
        
        
        
        
        
        
        Case BtErrKeyNotFound
        
            MsgBox "���[���ŕύX����Ă��܂��B�O��ʂɖ߂�܂��B"
            PM000302.Visible = False
            INIT_FLG = False
            
        
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            PM000302.Visible = False
            INIT_FLG = False
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

    Update_Proc = True
    
    '�i�ڃ}�X�^�@�ǂݍ���
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo1(pcmbNAIGAI).Text, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)
    
    
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
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
        
        
        Rclr_ITEMREC    '2018.04.19
        
        Call UniCode_Conv(ITEMREC.JGYOBU, Last_JGYOBU)              '���ƕ�=����
                                                                    '�����O
        Call UniCode_Conv(ITEMREC.NAIGAI, Right(Combo1(pcmbNAIGAI).Text, 1))
        Call UniCode_Conv(ITEMREC.HIN_GAI, Text1(ptxHIN_GAI).Text)  '�i�ں���
        Call UniCode_Conv(ITEMREC.HIN_NAME, Text1(ptxHIN_NAME))     '�i��
        Call UniCode_Conv(ITEMREC.ST_SET_DT, "")                    '�W���I�Ԑݒ���t
        Call UniCode_Conv(ITEMREC.ST_SOKO, "")                      '�W�����Ɂ@�q��
        Call UniCode_Conv(ITEMREC.ST_RETU, "")                      '�W�����Ɂ@��
        Call UniCode_Conv(ITEMREC.ST_REN, "")                       '�W�����Ɂ@�A
        Call UniCode_Conv(ITEMREC.ST_DAN, "")                       '�W�����Ɂ@�i
        Call UniCode_Conv(ITEMREC.BEF_SOKO, "")                     '�O����Ɂ@�q��
        Call UniCode_Conv(ITEMREC.BEF_RETU, "")                     '�O����Ɂ@��
        Call UniCode_Conv(ITEMREC.BEF_REN, "")                      '�O����Ɂ@�A
        Call UniCode_Conv(ITEMREC.BEF_DAN, "")                      '�O����Ɂ@�i
        Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")                  '�ŏI���ɓ�
        Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")                  '�ŏI�o�ɓ�
        Call UniCode_Conv(ITEMREC.HIN_NAI, "")                      '�i�ԁi���j
        Call UniCode_Conv(ITEMREC.BIKOU_SOKO, "")                   'νđq��
        Call UniCode_Conv(ITEMREC.BIKOU_TANA, "")                   'νĒI��
        Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")              '��[�_
        Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")            '�����Ϗo�א�
        Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "0")                  '����ِ�
        Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "0")                  '����ِ�
        Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")                  '�ŏI���ד��t
        Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")                  '�ŏI�ƍ����t
        Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "00000000")         '�ƍ����݌ɐ�
        Call UniCode_Conv(ITEMREC.BIKOU, "")                        '������l
        Call UniCode_Conv(ITEMREC.IRI_QTY, "")                      '������萔
        Call UniCode_Conv(ITEMREC.JAN_CODE, "")                     'JAN����
        Call UniCode_Conv(ITEMREC.HIN_CHANGE, "")                   '�i�ԓǂݑւ�����
        Call UniCode_Conv(ITEMREC.GOODS_KBN, "1")                   '���i���L��
        Call UniCode_Conv(ITEMREC.PACKING_NO, "")                   '������
        Call UniCode_Conv(ITEMREC.RANK, "")                         '�����ݸ
        Call UniCode_Conv(ITEMREC.NEW_RANK, "")                     '�V�ݸ
        Call UniCode_Conv(ITEMREC.GLICS1_TANA, "")                  '��د���I��1
        Call UniCode_Conv(ITEMREC.GLICS2_TANA, "")                  '��د���I��2
        Call UniCode_Conv(ITEMREC.GLICS3_TANA, "")                  '��د���I��3
    
        Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, "")                 '�i��E
        Call UniCode_Conv(ITEMREC.L_BIKOU, "")                      '���l
        Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, "")                '��Ж�
        Call UniCode_Conv(ITEMREC.L_KISHU1, "")                     '�@��(1)
        Call UniCode_Conv(ITEMREC.L_KISHU2, "")                     '�@��(2)
        Call UniCode_Conv(ITEMREC.L_KISHU3, "")                     '�@��(3)
        Call UniCode_Conv(ITEMREC.L_PAPER, "")                      '��
        Call UniCode_Conv(ITEMREC.L_PLASTIC, "")                    '��׽���
        Call UniCode_Conv(ITEMREC.L_URIKIN1, "")                    '���i(1)
        Call UniCode_Conv(ITEMREC.L_URIKIN2, "")                    '���i(2)
        Call UniCode_Conv(ITEMREC.L_URIKIN3, "")                    '���i(3)
        Call UniCode_Conv(ITEMREC.L_LABEL, "")                      '�K�p�@������
        Call UniCode_Conv(ITEMREC.L_MAISU, "")                      '���ٖ���
        Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, "")                '�K�p�@����l
        Call UniCode_Conv(ITEMREC.L_SAGYO_SHIJI, "")                '��Ǝw��
        Call UniCode_Conv(ITEMREC.L_BIKOU3, "")                     '���l(3)
        Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, "")                '���ƕ���
        Call UniCode_Conv(ITEMREC.L_IRI_QTY, "")                    '���萔
        Call UniCode_Conv(ITEMREC.L_TANA1, "")                      '�I��(1)
        Call UniCode_Conv(ITEMREC.L_TANA2, "")                      '�I��(2)
        
        Call UniCode_Conv(ITEMREC.S_TANTO, "")                      '���P�^�S����
        Call UniCode_Conv(ITEMREC.G_LABEL_NON, P_G_LABEL_ON)        '���ٓ\��t��
        
        Call UniCode_Conv(ITEMREC.S_KOUSU_GENKA, "")                '�H������   2008.02.14
        Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "")                '�H������   2008.02.14
        Call UniCode_Conv(ITEMREC.S_KOUSU_SET_DATE, "")             '�H���ݒ�� 2008.02.14
        
        Call UniCode_Conv(ITEMREC.S_SHIZAI_GENKA, "")               '���ތ���   2008.02.14
        Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "")               '���ޔ���   2008.02.14
        Call UniCode_Conv(ITEMREC.S_SHIZAI_SET_DATE, "")            '���ސݒ�� 2008.02.14
        
        
        
        Call UniCode_Conv(ITEMREC.SE_USOU_F, "")                    '�A�����@�o���׸�   2008.02.14

        Call UniCode_Conv(ITEMREC.USE_TAPE_KIND, "")                '�g�p�e�[�v���     2008.02.14
        Call UniCode_Conv(ITEMREC.USE_TAPE_LNG, "")                 '�g�p�e�[�v��       2008.02.14

        Call UniCode_Conv(ITEMREC.H_TANA_MAKE, "")                  '�I�ԃ}�[�N         2008.04.02

        Call UniCode_Conv(ITEMREC.SE_TANKA_MEMO, "")                '�����P���@����     2008.04.15

        Call UniCode_Conv(ITEMREC.xGENSANKOKU, "")                  '���Y��             2008.06.11-->2009.07.16 ���g�p

        Call UniCode_Conv(ITEMREC.S_GAISO_TANKA, "")                '�O���P�� 9(8)V99   2008.06.12
        Call UniCode_Conv(ITEMREC.S_PPSC_KAKO_KOSU, "")             'PPSC���H�P��9(8)   2008.06.12
        Call UniCode_Conv(ITEMREC.S_BU_KAKO_KOSU, "")               'BU���H�P��9(8)     2008.06.12

        Call UniCode_Conv(ITEMREC.SEI_LOT, "")                      '���Y���b�g         2008.07.07
        Call UniCode_Conv(ITEMREC.SEI_RATE, "")                     '�����[�g           2008.07.07
        Call UniCode_Conv(ITEMREC.SEI_SYU_KON, "")                  '�W������           2008.07.07

        Call UniCode_Conv(ITEMREC.SEI_TANKA_TANTO, "")              '�P���ݒ�S����     2008.07.09

        Call UniCode_Conv(ITEMREC.SHIMUKE_CODE, "")                 '�d������           2008.07.09

        Call UniCode_Conv(ITEMREC.SEI_KBN, "")                      '�����敪           2008.07.16

        Call UniCode_Conv(ITEMREC.SEI_LABEL_QTY, "")                '���x���\�薇��     2008.07.19

        Call UniCode_Conv(ITEMREC.SEI_SZI_CNT, "")                  '���ތ���     �@    2008.08.20�ǉ�
        Call UniCode_Conv(ITEMREC.SEI_DKN_CNT, "")                  '��������           2008.08.20�ǉ�

        For i = 0 To 9                                              '2008.09.19
            Call UniCode_Conv(ITEMREC.BEF_KOUTEI(i).BEF_KOUTEI, "")
            Call UniCode_Conv(ITEMREC.MAIN_KOUTEI(i).MAIN_KOUTEI, "")
            Call UniCode_Conv(ITEMREC.AFT_KOUTEI(i).AFT_KOUTEI, "")
        Next i
    


        Call UniCode_Conv(ITEMREC.SE_IO_TANKA_No, "")               '�I�敪             200.09.19

        Call UniCode_Conv(ITEMREC.STAT, "")                         '��ԋ敪           2009.01.21

        Call UniCode_Conv(ITEMREC.INSP_MESSAGE, "")                 '�o�׌��iү����     2009.04.17

        Call UniCode_Conv(ITEMREC.S_SEIKYU_F, "")                   '���i�������׸�     2009.04.28
    
        Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "")            '���i���@�H������   2009.06.02
        Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "")           '���i���@���ޔ���   2009.06.02
        Call UniCode_Conv(ITEMREC.BEF_S_GAISO_TANKA, "")            '�O���P�� 9(8)V99   2009.06.02
        Call UniCode_Conv(ITEMREC.BEF_S_PPSC_KAKO_KOSU, "")         'PPSC���H�P��9(8)   2009.06.02
        Call UniCode_Conv(ITEMREC.BEF_S_BU_KAKO_KOSU, "")           'BU���H�P��9(8)     2009.06.02
    
        Call UniCode_Conv(ITEMREC.M_BIKOU, "")                      '���Ϗ����l         2009.06.02
        Call UniCode_Conv(ITEMREC.SHIYOU_NO, "")                    '�d�l����           2009.06.02
        Call UniCode_Conv(ITEMREC.MITSUMORI_KBN, "")                '���ς�敪         2009.06.02
        Call UniCode_Conv(ITEMREC.TANKA_KIRIKAE_DT, "")             '�P���ؑ֓��t       2009.06.02
        Call UniCode_Conv(ITEMREC.KIRIKAE_KBN, "")                  '�ؑ֋敪           2009.06.02
    
        Call UniCode_Conv(ITEMREC.GENSANKOKU, "")                   '���Y��             '2009.07.16
        
        
        
        
        
        
'-------    2010.10.04
        Call UniCode_Conv(ITEMREC.PLUS_KOUSU, "")                   '�v���X���H��       2009.09.17
        Call UniCode_Conv(ITEMREC.KUTI_SU, "")                      '����               2010.01.18
        Call UniCode_Conv(ITEMREC.KONPOU_F, "")                     '����敪           2010.01.18
    
        Call UniCode_Conv(ITEMREC.SAI_SU, "")                       '�ː�               2010.01.18
    
    
    
        Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, "")              '�捞�ݎ����Y��     2010.07.20
        Call UniCode_Conv(ITEMREC.TORI_GEN_GENSANKOKU, "")          '�捞�ݎ����Y���\�� 2010.07.20
        Call UniCode_Conv(ITEMREC.TORI_SHIIRE_WORK_CENTER, "")      '�d��ܰ��Z���^�[    2010.07.20
        
    
    
        Call UniCode_Conv(ITEMREC.KANKYO_KBN, "")                   '����ދ敪       2010.07.27
        Call UniCode_Conv(ITEMREC.KANKYO_KBN_ST, "")                '����ދ敪�K�p�J�n 2010.07.27
        Call UniCode_Conv(ITEMREC.KANKYO_KBN_SURYO, "")             '����ދ敪����   2010.07.27
    
        Call UniCode_Conv(ITEMREC.BEF_L_LABEL, "")                       '''''
    
        Call UniCode_Conv(ITEMREC.BEF_1_L_PAPER, "")                '           ��
        Call UniCode_Conv(ITEMREC.BEF_1_L_PLASTIC, "")              '           �v���X�`�b�N
        Call UniCode_Conv(ITEMREC.BEF_2_L_PAPER, "")                '           ��
        Call UniCode_Conv(ITEMREC.BEF_2_L_PLASTIC, "")              '           �v���X�`�b�N
        Call UniCode_Conv(ITEMREC.BEF_3_L_PAPER, "")                '           ��
        Call UniCode_Conv(ITEMREC.BEF_3_L_PLASTIC, "")              '           �v���X�`�b�N
        Call UniCode_Conv(ITEMREC.BEF_4_L_PAPER, "")                '           ��
        Call UniCode_Conv(ITEMREC.BEF_4_L_PLASTIC, "")              '           �v���X�`�b�N
        Call UniCode_Conv(ITEMREC.BEF_LAST_L_PAPER, "")             '           ��
        Call UniCode_Conv(ITEMREC.BEF_LAST_L_PLASTIC, "")           '           �v���X�`�b�N


        Call UniCode_Conv(ITEMREC.BIKOU20, "")
'-------    2010.09.04
        
        Call UniCode_Conv(ITEMREC.SHIIRE_BIKOU, "")                 '�d�����l   2018.04.19
        
        
        
        Call UniCode_Conv(ITEMREC.INS_TANTO, "PM030")               '�ǉ��@�S���ҁ@     2009.01.21
        Call UniCode_Conv(ITEMREC.Ins_DateTime, Format(Now, "YYYYMMDDHHMMSS"))  '�ǉ��@����         2009.01.21

        Call UniCode_Conv(ITEMREC.UPD_TANTO, "")                    '�X�V�@�S���ҁ@     2005.11.15
        Call UniCode_Conv(ITEMREC.UPD_DATETIME, "")                 '�X�V�@����         2005.11.15
        
        
        
        Call UniCode_Conv(ITEMREC.FILLER, "")                       'Filler
    
    End If
    
    
    Call UniCode_Conv(ITEMREC.HIN_NAME, Text1(ptxHIN_NAME).Text)
    
    
    Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, Text1(ptxG_SHIIRE_KBN).Text)                '�d���敪
    Call UniCode_Conv(ITEMREC.G_HANBAI_KBN, Text1(ptxG_HANBAI_KBN).Text)                '�̔��敪
    Call UniCode_Conv(ITEMREC.G_SYUSHI, Text1(ptxG_SYUSHI).Text)                        '���x�P��
    If Check1(pchkG_KUMITATE).Value = vbChecked Then                                    '�g�����i
        Call UniCode_Conv(ITEMREC.G_KUMITATE, P_ASSEMBLY_ON)
    Else
        Call UniCode_Conv(ITEMREC.G_KUMITATE, P_ASSEMBLY_OFF)
    End If
        
    If Trim(Text1(ptxG_ZEN_ZAIKO_KIN).Text) = "" Then                                   '�O���݌ɋ��z
        Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, "")
    Else
        If CLng(Text1(ptxG_ZEN_ZAIKO_KIN).Text) < 0 Then
            Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, Format(CLng(Text1(ptxG_ZEN_ZAIKO_KIN).Text), "0000000"))
        Else
            Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, Format(CLng(Text1(ptxG_ZEN_ZAIKO_KIN).Text), "00000000"))
        End If
    End If
    
    
    If Trim(Text1(ptxG_ZEN_ZAIKO_QTY).Text) = "" Then                                   '�O���݌ɐ���
        Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, "")
    Else
        If CLng(Text1(ptxG_ZEN_ZAIKO_QTY).Text) < 0 Then
        
            Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, Format(CLng(Text1(ptxG_ZEN_ZAIKO_QTY).Text), "0000000"))
        Else
            Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, Format(CLng(Text1(ptxG_ZEN_ZAIKO_QTY).Text), "00000000"))
        End If
    End If
    
    
    
    If Trim(Text1(ptxG_ST_URITAN).Text) = "" Then                                       '�W������
        Call UniCode_Conv(ITEMREC.G_ST_URITAN, "")
    Else
        Call UniCode_Conv(ITEMREC.G_ST_URITAN, Format(CDbl(Text1(ptxG_ST_URITAN).Text), "00000000.00"))
    End If
    If Trim(Text1(ptxG_ST_URITAN_DT).Text) = "" Then                                   '�W�������ݒ��
        Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, "")
    Else
        Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, Format(Text1(ptxG_ST_URITAN_DT).Text, "YYYYMMDD"))
    End If
    
    If Trim(Text1(ptxG_ST_SHITAN).Text) = "" Then                                       '�W������
        Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "")
    Else
        Call UniCode_Conv(ITEMREC.G_ST_SHITAN, Format(CDbl(Text1(ptxG_ST_SHITAN).Text), "00000000.00"))
    End If
    If Trim(Text1(ptxG_ST_SHITAN_DT).Text) = "" Then                                   '�W�������ݒ��
        Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, "")
    Else
        Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, Format(Text1(ptxG_ST_SHITAN_DT).Text, "YYYYMMDD"))
    End If
    
    If Trim(Text1(ptxHOJYU_P).Text) = "" Then                                           '�댯�݌�
        Call UniCode_Conv(ITEMREC.HOJYU_P, "")
    Else
        Call UniCode_Conv(ITEMREC.HOJYU_P, Format(CLng(Text1(ptxHOJYU_P).Text), "00000000"))
    End If
        
        
    Call UniCode_Conv(ITEMREC.G_SHIZAI_KBN, Right(Combo1(pcmbG_SHIZAI_KBN).Text, 1))    '���ދ敪
        
    If Check1(pchkG_LABEL_NON).Value = vbChecked Then                                   '���x���\��
        Call UniCode_Conv(ITEMREC.G_LABEL_NON, P_G_LABEL_OFF)
    Else
        Call UniCode_Conv(ITEMREC.G_LABEL_NON, P_G_LABEL_ON)
    End If
        
    If Check1(pchkZAIKO_F).Value = vbUnchecked Then                                     '�݌ɊǗ��Ώ�
        Call UniCode_Conv(ITEMREC.ZAIKO_F, P_ZAIKO_F_ON)
    Else
        Call UniCode_Conv(ITEMREC.ZAIKO_F, P_ZAIKO_F_OFF)
    End If
        
        
        
    If Check1(pchkZAIKO_CLR_F).Value = vbChecked Then                                   '�I���\�@�݌ɐ���\��   2012.12.13
        Call UniCode_Conv(ITEMREC.ZAIKO_CLR_F, "1")
    Else
        Call UniCode_Conv(ITEMREC.ZAIKO_CLR_F, "")
    End If
        
        
        
        
                                    '�ː�   2008.02.14
    If IsNumeric(Text1(ptxSAI_SU).Text) Then
'        Call UniCode_Conv(ITEMREC.SAI_SU, Format(CDbl(Text1(ptxSAI_SU).Text), "0.0"))
        Call UniCode_Conv(ITEMREC.SAI_SU, Format(CDbl(Text1(ptxSAI_SU).Text), "0.00"))
    Else
        Call UniCode_Conv(ITEMREC.SAI_SU, "")
    End If
                                    '�`��    2008.02.14
    Call UniCode_Conv(ITEMREC.D_KEISHIKI, Text1(ptxD_KEISHIKI).Text)
                                    '�ގ�   2008.02.14
    Call UniCode_Conv(ITEMREC.D_MATERIAL, Text1(ptxD_MATERIAL).Text)
                                    '����ްٌ���    2008.02.14
    Call UniCode_Conv(ITEMREC.D_THICKNESS, Text1(ptxD_THICKNESS).Text)
                                    
                                    '����ްٻ���(W) 2008.02.14
    Call UniCode_Conv(ITEMREC.D_SIZE_W, Text1(ptxD_SIZE_W).Text)
                                    '����ްٻ���(D) 2008.02.14
    Call UniCode_Conv(ITEMREC.D_SIZE_D, Text1(ptxD_SIZE_D).Text)
                                    '����ްٻ���(H) 2008.02.14
    Call UniCode_Conv(ITEMREC.D_SIZE_H, Text1(ptxD_SIZE_H).Text)
                                    
                                    '�������/���Ȃ� 2008.02.14
    Call UniCode_Conv(ITEMREC.D_PRINT, Text1(ptxD_PRINT).Text)
                                    '���i���H�� 2008.02.14
    If Trim(Text1(ptxS_KOUSU).Text) <> "" Then
        Call UniCode_Conv(ITEMREC.S_KOUSU, Format(CDbl(Text1(ptxS_KOUSU).Text), "00000000"))
    Else
        Call UniCode_Conv(ITEMREC.S_KOUSU, "")
    End If
                                    
                                    '�W������ 2008.07.16
    If Trim(Text1(ptxSEI_SYU_KON).Text) <> "" Then
        Call UniCode_Conv(ITEMREC.SEI_SYU_KON, Format(CDbl(Text1(ptxSEI_SYU_KON).Text), "000000"))
    Else
        Call UniCode_Conv(ITEMREC.SEI_SYU_KON, "")
    End If
                                    '�A���� 2008.02.14
    Call UniCode_Conv(ITEMREC.SE_USOU_F, Text1(ptxSE_USOU_F).Text)
        
                                    '�g�pð�ߎ�� 2008.02.14
    Call UniCode_Conv(ITEMREC.USE_TAPE_KIND, Text1(ptxUSE_TAPE_KIND).Text)
                                    '�g�pð�ߒ� 2008.02.14
    Call UniCode_Conv(ITEMREC.USE_TAPE_LNG, Text1(ptxUSE_TAPE_LNG).Text)
                                    '�����敪 2008.07.16
    Call UniCode_Conv(ITEMREC.SEI_KBN, Text1(ptxSEI_KBN).Text)
                                    
                                    
    If Trim(Text1(ptxST_SOKO).Text) = "" And Trim(Text1(ptxST_RETU).Text) = "" And Trim(Text1(ptxST_REN).Text) = "" And Trim(Text1(ptxST_DAN).Text) = "" Then
    
    
        Call UniCode_Conv(ITEMREC.ST_SET_DT, "")
    
    
    Else
        If Trim(Text1(ptxST_SOKO).Text) = Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) And Trim(Text1(ptxST_RETU).Text) = Trim(StrConv(ITEMREC.ST_RETU, vbUnicode)) And Trim(Text1(ptxST_REN).Text) = Trim(StrConv(ITEMREC.ST_REN, vbUnicode)) And Trim(Text1(ptxST_DAN).Text) = Trim(StrConv(ITEMREC.ST_DAN, vbUnicode)) Then
        Else
            Call UniCode_Conv(ITEMREC.ST_SET_DT, Format(Now, "YYYYMMDD"))
        End If
    End If
                                    
                                    
                                    
                                    
                                    
                                    
                                    
                                    '�W���I�� 2009.09.01
    Call UniCode_Conv(ITEMREC.ST_SOKO, Text1(ptxST_SOKO).Text)
                                    '�W���I�� 2009.09.01
    Call UniCode_Conv(ITEMREC.ST_RETU, Text1(ptxST_RETU).Text)
                                    '�W���I�� 2009.09.01
    Call UniCode_Conv(ITEMREC.ST_REN, Text1(ptxST_REN).Text)
                                    '�W���I�� 2009.09.01
    Call UniCode_Conv(ITEMREC.ST_DAN, Text1(ptxST_DAN).Text)
        
        
        
        
        
                                    '���� 2010.01.18
    Call UniCode_Conv(ITEMREC.KUTI_SU, Text1(ptxKUTI_SU).Text)
                                    '����F 2010.01.18
    Call UniCode_Conv(ITEMREC.KONPOU_F, Text1(ptxKONPOU_F).Text)
        
        
        
        
        
        
    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).CODE, Text1(ptxSHI_CODE1).Text)           '�d����(1)
    If Trim(Text1(ptxSHI_TANKA1).Text) = "" Then                                        '�d���P��(1)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).TANKA, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).TANKA, Format(CDbl(Text1(ptxSHI_TANKA1).Text), "00000000.00"))
    End If
    If Trim(Text1(ptxSHI_TANKA_DT1).Text) = "" Then                                     '�d���P���ݒ��(1)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).TANKA_DT, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).TANKA_DT, Format(Text1(ptxSHI_TANKA_DT1).Text, "YYYYMMDD"))
    End If
    If Trim(Text1(ptxSHI_LOT1).Text) = "" Then                                          'ۯ�(1)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).LOT, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).LOT, Format(CLng(Text1(ptxSHI_LOT1).Text), "00000000"))
    End If
    If Trim(Text1(ptxSHI_LEAD_TIME1).Text) = "" Then                                    'ذ�����(1)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).LEAD_TIME, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).LEAD_TIME, Format(CLng(Text1(ptxSHI_LEAD_TIME1).Text), "000"))
    End If
    If Trim(Text1(ptxSHI_LAST_ORDER_DT1).Text) = "" Then                                '�O�񒍕���(1)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).LAST_ORDER_DT, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).LAST_ORDER_DT, Format(Text1(ptxSHI_LAST_ORDER_DT1).Text, "YYYYMMDD"))
    End If
    If Trim(Text1(ptxSHI_LAST_ORDER_QTY1).Text) = "" Then                               '�O�񒍕���(1)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).LAST_ORDER_QTY, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).LAST_ORDER_QTY, Format(CLng(Text1(ptxSHI_LAST_ORDER_QTY1).Text), "00000000"))
    End If
                
    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(1).CODE, Text1(ptxSHI_CODE2).Text)           '�d����(2)
    If Trim(Text1(ptxSHI_TANKA2).Text) = "" Then                                        '�d���P��(2)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(1).TANKA, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(1).TANKA, Format(CDbl(Text1(ptxSHI_TANKA2).Text), "00000000.00"))
    End If
    If Trim(Text1(ptxSHI_TANKA_DT2).Text) = "" Then                                     '�d���P���ݒ��(2)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(1).TANKA_DT, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(1).TANKA_DT, Format(Text1(ptxSHI_TANKA_DT2).Text, "YYYYMMDD"))
    End If
    If Trim(Text1(ptxSHI_LOT2).Text) = "" Then                                          'ۯ�(2)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(1).LOT, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(1).LOT, Format(CLng(Text1(ptxSHI_LOT2).Text), "00000000"))
    End If
    If Trim(Text1(ptxSHI_LEAD_TIME2).Text) = "" Then                                    'ذ�����(2)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(1).LEAD_TIME, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(1).LEAD_TIME, Format(CLng(Text1(ptxSHI_LEAD_TIME2).Text), "000"))
    End If
    If Trim(Text1(ptxSHI_LAST_ORDER_DT2).Text) = "" Then                                '�O�񒍕���(2)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(1).LAST_ORDER_DT, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(1).LAST_ORDER_DT, Format(Text1(ptxSHI_LAST_ORDER_DT2).Text, "YYYYMMDD"))
    End If
    If Trim(Text1(ptxSHI_LAST_ORDER_QTY2).Text) = "" Then                               '�O�񒍕���(2)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(1).LAST_ORDER_QTY, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(1).LAST_ORDER_QTY, Format(CLng(Text1(ptxSHI_LAST_ORDER_QTY2).Text), "00000000"))
    End If
                
    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(2).CODE, Text1(ptxSHI_CODE3).Text)           '�d����(3)
    If Trim(Text1(ptxSHI_TANKA3).Text) = "" Then                                        '�d���P��(3)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(2).TANKA, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(2).TANKA, Format(CDbl(Text1(ptxSHI_TANKA3).Text), "00000000.00"))
    End If
    If Trim(Text1(ptxSHI_TANKA_DT3).Text) = "" Then                                     '�d���P���ݒ��(3)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(2).TANKA_DT, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(2).TANKA_DT, Format(Text1(ptxSHI_TANKA_DT3).Text, "YYYYMMDD"))
    End If
    If Trim(Text1(ptxSHI_LOT3).Text) = "" Then                                           'ۯ�(3)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(2).LOT, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(2).LOT, Format(CLng(Text1(ptxSHI_LOT3).Text), "00000000"))
    End If
    If Trim(Text1(ptxSHI_LEAD_TIME3).Text) = "" Then                                    'ذ�����(3)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(2).LEAD_TIME, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(2).LEAD_TIME, Format(CLng(Text1(ptxSHI_LEAD_TIME3).Text), "000"))
    End If
    If Trim(Text1(ptxSHI_LAST_ORDER_DT3).Text) = "" Then                                '�O�񒍕���(3)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(2).LAST_ORDER_DT, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(2).LAST_ORDER_DT, Format(Text1(ptxSHI_LAST_ORDER_DT3).Text, "YYYYMMDD"))
    End If
    If Trim(Text1(ptxSHI_LAST_ORDER_QTY3).Text) = "" Then                               '�O�񒍕���(3)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(2).LAST_ORDER_QTY, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(2).LAST_ORDER_QTY, Format(CLng(Text1(ptxSHI_LAST_ORDER_QTY3).Text), "00000000"))
    End If
    
        
    If Trim(Text1(ptxLAST_SYU_DT).Text) = "" Then                                       '�ŏI�o�ד�
        Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")
    Else
        Call UniCode_Conv(ITEMREC.LAST_SYU_DT, Format(Text1(ptxLAST_SYU_DT).Text, "YYYYMMDD"))
    End If
    
    If Trim(Text1(ptxG_LAST_SYUKA_QTY).Text) = "" Then                                  '�ŏI�o�א�
        Call UniCode_Conv(ITEMREC.G_LAST_SYUKA_QTY, "00000000")
    Else
        Call UniCode_Conv(ITEMREC.G_LAST_SYUKA_QTY, Format(CLng(Text1(ptxG_LAST_SYUKA_QTY).Text), "00000000"))
    End If
    
    
                                                                                        '�ŐV�d����     2007.05.29
    Call UniCode_Conv(ITEMREC.LAST_CODE, Text1(ptxLAST_CODE).Text)
    
    If Trim(Text1(ptxLAST_TANKA).Text) = "" Then                                        '�ŐV�d���P��   2007.05.29
        Call UniCode_Conv(ITEMREC.LAST_TANKA, "00000000.00")
    Else
        Call UniCode_Conv(ITEMREC.LAST_TANKA, Format(CDbl(Text1(ptxLAST_TANKA).Text), "00000000.00"))
    End If
    
    Call UniCode_Conv(ITEMREC.SHIIRE_BIKOU, Text1(ptxSHIIRE_BIKOU).Text)                '�d�����l�@2018.04.19
    
    
    
    
    Call UniCode_Conv(ITEMREC.UPD_TANTO, App.EXEName)                                   '�X�V�S���Һ���
                                                                                        '�X�V����
    Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    
    Do
        
        DoEvents
        
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
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
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo1(pcmbNAIGAI).Text, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)
    
    
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
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
        
        sts = BTRV(BtOpDelete, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
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
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
        
    Select Case Index
        Case pcmbG_SHIIRE       '�d���敪
            Text1(ptxG_SHIIRE_KBN).Text = Left(Right(Combo1(pcmbG_SHIIRE).Text, 3), 2)
        Case pcmbG_HANBAI       '�̔��敪
            Text1(ptxG_HANBAI_KBN).Text = Left(Right(Combo1(pcmbG_HANBAI).Text, 3), 2)
        Case pcmbG_SYUSHI       '���x�P��
            Text1(ptxG_SYUSHI).Text = Right(Combo1(pcmbG_SYUSHI).Text, 3)
        Case pcmbSHIIRE1        '�d����(1)
            Text1(ptxSHI_CODE1).Text = Right(Combo1(pcmbSHIIRE1).Text, 5)
        Case pcmbSHIIRE2        '�d����(2)
            Text1(ptxSHI_CODE2).Text = Right(Combo1(pcmbSHIIRE2).Text, 5)
        Case pcmbSHIIRE3        '�d����(3)
            Text1(ptxSHI_CODE3).Text = Right(Combo1(pcmbSHIIRE3).Text, 5)
        Case pcmbLAST_CODE      '�ŐV�d����     2007.05.28
            Text1(ptxLAST_CODE).Text = Right(Combo1(pcmbLAST_CODE).Text, 5)
    
    End Select
    
    Call Tab_Ctrl(Shift)        '�ړ�

End Sub


Private Sub Command1_Click(Index As Integer)

Dim ans As Integer
Dim i   As Integer


    Select Case Index
        Case P_CMD_Upd                      '�X�V
            
'2010.01.18            For i = ptxHIN_GAI To ptxST_DAN
            For i = ptxHIN_GAI To ptxKONPOU_F
            
                If Error_Check_Proc(i) Then     '�G���[�`�F�b�N
                    Exit Sub
                End If
            
            Next i
            
            Beep
            ans = MsgBox("�X�V���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Update_Proc() Then
                    PM000302.Visible = False
                    INIT_FLG = False
                End If
            Else
                Exit Sub
            End If
            PM000302.Visible = False
            INIT_FLG = False
                    
        
        
        Case P_CMD_DEL                      '�폜
            ans = MsgBox("�폜���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Delete_Proc() Then
                    PM000302.Visible = False
                    INIT_FLG = False
                End If
            Else
                Exit Sub
            End If
            PM000302.Visible = False
            INIT_FLG = False
        Case P_CMD_DSP                      '����/�\��
        Case P_CMD_OUT                      '�ް��o��
        Case P_CMD_PRT                      '���
        
        Case P_CMD_End                      '�I��
            PM000302.Visible = False
            INIT_FLG = False
    End Select

End Sub

Private Sub Command2_Click()
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "���ރ}�X�^�����e�i���X ��ʈ�����J�n���܂��� ", Me.hwnd, 0)


Call Form_HCopy_Win7_NEW(Picture1, vbPRPSA4, vbPRORLandscape)       '2018.11.21

    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "���ރ}�X�^�����e�i���X ��ʈ�����I�����܂��� ", Me.hwnd, 0)


End Sub

Private Sub Form_Activate()
    
Dim i       As Integer
Dim CODE    As String
    
    If INIT_FLG Then
        Exit Sub
    End If


    Select Case G_SCREEN_FLG
        Case G_SCREEN_INS       '�V�K
                
            Text1(ptxHIN_GAI).BackColor = G_INPUT_OK
            Text1(ptxHIN_GAI).TabStop = True
            Text1(ptxHIN_GAI).Locked = False
                
'            Text1(ptxHIN_NAME).BackColor = G_INPUT_OK
'            Text1(ptxHIN_NAME).TabStop = True
'            Text1(ptxHIN_NAME).Locked = False
                
                
            For i = ptxHIN_GAI To ptxKONPOU_F
                Text1(i).Text = ""
            Next i
                
            Text1(ptxSHIIRE_BIKOU).Text = ""    '2018.04.19
                
            For i = pcmbG_SHIIRE To pcmbLAST_CODE
            
                Combo1(i).ListIndex = -1
            Next i
                
'2012.12.13            For i = pchkG_KUMITATE To pchkZAIKO_F
            For i = pchkG_KUMITATE To pchkZAIKO_CLR_F   '2012.12.13
                Check1(i).Value = vbUnchecked
            Next i
                
                
            lblSIZE.Caption = ""
            lblIns_DateTime.Caption = ""
            lblUpd_DateTime.Caption = ""
                
                
                
                
            Text1(ptxHIN_GAI).SetFocus
                
                
                
        
        Case G_SCREEN_UPD       '�X�V
    
                
    
    
            Text1(ptxHIN_GAI).BackColor = G_INPUT_NG
            Text1(ptxHIN_GAI).TabStop = False
            Text1(ptxHIN_GAI).Locked = True
    
'            Text1(ptxHIN_NAME).BackColor = G_INPUT_NG
'            Text1(ptxHIN_NAME).TabStop = False
'            Text1(ptxHIN_NAME).Locked = True
    
            
            CODE = PM000301.txSEL_KEY
            
            If Item_Disp_Proc(CODE) Then
                Unload Me
            End If
    
            Text1(ptxHIN_NAME).SetFocus
    
    End Select


    INIT_FLG = True

End Sub

Private Sub Form_DblClick()
'    PrintForm
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


    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "���ރ}�X�^�����e�i���X", Me.hwnd, 0)
    '�Ō�̗v�f��-1�ɂ����
    '�e�E�B���h�E�̑S�̂̕��̎c��̕���
    '�����I�Ɋ��蓖�Ă�
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)


    Combo1(pcmbNAIGAI).AddItem NAIGAI1 & "   " & NAIGAI_NAI
    Combo1(pcmbNAIGAI).AddItem NAIGAI2 & "   " & NAIGAI_GAI
    Combo1(pcmbNAIGAI).ListIndex = 0
    
    '�d���敪�̃Z�b�g
    If Code_Set_Proc(pcmbG_SHIIRE, P_KBN01_CD, 0) Then
        Unload Me
    End If
    
    
    '�̔��敪�̃Z�b�g
    If Code_Set_Proc(pcmbG_HANBAI, P_KBN02_CD, 0) Then
        Unload Me
    End If
    
    
    
    '���x���e�̃Z�b�g
    If Code_Set_Proc(pcmbG_SYUSHI, P_KBN03_CD, 1) Then
        Unload Me
    End If
    
    '���ދ敪�̃Z�b�g
    If Code_Set_Proc(pcmbG_SHIZAI_KBN, P_KBN08_CD, 1) Then
        Unload Me
    End If
    
    
    
    '�d����̃Z�b�g
    Combo1(pcmbSHIIRE1).Clear
    Combo1(pcmbSHIIRE2).Clear
    Combo1(pcmbSHIIRE3).Clear
    
    Combo1(pcmbLAST_CODE).Clear
    
    
    com = BtOpGetFirst
    
    Do
        DoEvents
    
        sts = BTRV(com, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
            
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�󕥐�}�X�^")
                Unload Me
        End Select
        
        Combo1(pcmbSHIIRE1).AddItem StrConv(P_UKEHARAIREC.UKEHARAI_NAME, vbUnicode) & " " & _
                                    StrConv(P_UKEHARAIREC.UKEHARAI_CODE, vbUnicode)
        Combo1(pcmbSHIIRE2).AddItem StrConv(P_UKEHARAIREC.UKEHARAI_NAME, vbUnicode) & " " & _
                                    StrConv(P_UKEHARAIREC.UKEHARAI_CODE, vbUnicode)
        Combo1(pcmbSHIIRE3).AddItem StrConv(P_UKEHARAIREC.UKEHARAI_NAME, vbUnicode) & " " & _
                                    StrConv(P_UKEHARAIREC.UKEHARAI_CODE, vbUnicode)
        
        '�ŐV�d���� 2007.05.28
        Combo1(pcmbLAST_CODE).AddItem StrConv(P_UKEHARAIREC.UKEHARAI_NAME, vbUnicode) & " " & _
                                    StrConv(P_UKEHARAIREC.UKEHARAI_CODE, vbUnicode)
        
        
        com = BtOpGetNext
    
    Loop
    
    
    INIT_FLG = False
    
    
    
End Sub

Private Sub Form_Unload(CANCEL As Integer)

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �O��ʂɖ߂�    2016.01.27
'Dim sts As Integer
'
'                                            '�i�ڃ}�X�^�b�k�n�r�d
'    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'    If sts Then
'        If sts <> BtErrNoOpen Then
'            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
'        End If
'    End If
'
'
'                                            '�󕥐�}�X�^�b�k�n�r�d
'    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
'    If sts Then
'        If sts <> BtErrNoOpen Then
'            Call File_Error(sts, BtOpClose, "�󕥐�}�X�^")
'        End If
'    End If
'
'                                            '�R�[�h�}�X�^�b�k�n�r�d
'    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
'    If sts Then
'        If sts <> BtErrNoOpen Then
'            Call File_Error(sts, BtOpClose, "�R�[�h�}�X�^")
'        End If
'    End If
'    sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
'    If sts Then
'        Call File_Error(sts, BtOpReset, "")
'    End If
'    Set PM000301 = Nothing
'    Set PM000302 = Nothing
'
'    End
'

    PM000302.Visible = False
    INIT_FLG = False



'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �O��ʂɖ߂�    2016.01.27

End Sub


Private Sub Text1_GotFocus(Index As Integer)
    
    
'>>>>>  2018.04.09
    Select Case Index
    
        Case ptxG_ST_URITAN, ptxG_ST_SHITAN, ptxSHI_TANKA1, ptxSHI_TANKA2, ptxSHI_TANKA3
            svTANKA = Text1(Index).Text
    
    End Select
'>>>>>  2018.04.09
    
    
    
    
    
    
    
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
        
    If Error_Check_Proc(Index) Then     '�G���[�`�F�b�N
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '�ړ�

End Sub

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

Private Sub Text1_LostFocus(Index As Integer)

'>>>>>>>>>>>>>>>>>>>>>>>>>  2016.01.26
    Select Case Index
        Case ptxST_SOKO
            Text1(Index).Text = StrConv(Text1(Index).Text, vbUpperCase)
                    
        Case ptxSHI_CODE1
            Text1(Index).Text = StrConv(Text1(Index).Text, vbUpperCase)
        Case ptxSHI_CODE2
            Text1(Index).Text = StrConv(Text1(Index).Text, vbUpperCase)
        Case ptxSHI_CODE3
            Text1(Index).Text = StrConv(Text1(Index).Text, vbUpperCase)
    
        Case ptxLAST_CODE
            Text1(Index).Text = StrConv(Text1(Index).Text, vbUpperCase)
    
    
'>>>>>  2018.04.09
        Case ptxG_ST_URITAN, ptxG_ST_SHITAN, ptxSHI_TANKA1, ptxSHI_TANKA2, ptxSHI_TANKA3
            If svTANKA <> Text1(Index).Text Then
                If Trim(Text1(Index).Text) = "" Then
                    Text1(Index + 1).Text = ""
                Else
                    Text1(Index + 1).Text = Format(Now, "YYYY/MM/DD")
                End If
            End If
'>>>>>  2018.04.09
    
    
    
    End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>  2016.01.26


End Sub
