VERSION 5.00
Begin VB.Form F1040101 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�݌ɖ⍇�킹�i�i�ԕʁj"
   ClientHeight    =   11205
   ClientLeft      =   2025
   ClientTop       =   2550
   ClientWidth     =   15825
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
   ScaleHeight     =   11205
   ScaleWidth      =   15825
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CommandButton Command2 
      Caption         =   "ini�\��"
      Height          =   495
      Left            =   12240
      TabIndex        =   73
      Top             =   0
      Width           =   1695
   End
   Begin VB.ListBox List2 
      Height          =   6780
      Left            =   14640
      TabIndex        =   71
      Top             =   1680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʈ��"
      Height          =   495
      Left            =   14040
      TabIndex        =   70
      Top             =   0
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   135
      Left            =   240
      ScaleHeight     =   75
      ScaleWidth      =   195
      TabIndex        =   57
      Top             =   7320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   15
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2445
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   14
      Left            =   5280
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1485
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   13
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1485
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   12
      Left            =   5280
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1485
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   7
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1485
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   5280
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1485
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1485
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   11
      Left            =   3240
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   6840
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   10
      Left            =   2280
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   5640
      Width           =   330
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   9
      Left            =   2280
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   5160
      Width           =   330
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   8
      Left            =   2280
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   4680
      Width           =   330
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1485
   End
   Begin VB.TextBox Text 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1485
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   2400
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   0
      Top             =   8280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   6780
      Left            =   7320
      TabIndex        =   4
      Top             =   1680
      Width           =   7215
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   6000
      MaxLength       =   20
      TabIndex        =   2
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox Text 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   9240
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   600
      Width           =   6525
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   1
      Top             =   600
      Width           =   2535
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
      Left            =   10560
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   9840
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
      Left            =   9720
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   9840
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
      Left            =   8880
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   9840
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
      Index           =   8
      Left            =   8040
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "��  �V"
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
      Left            =   6720
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   9840
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
      Left            =   5880
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   9840
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
      Left            =   5040
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   9840
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
      Left            =   4200
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   9840
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
      Index           =   3
      Left            =   2880
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   9840
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
      Left            =   2040
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9840
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
      Left            =   1200
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   9840
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
      Index           =   0
      Left            =   360
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.Label lblNOW 
      Alignment       =   2  '��������
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   72
      Top             =   10320
      Width           =   3015
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��(*)"
      Height          =   255
      Index           =   34
      Left            =   7320
      TabIndex        =   69
      Top             =   1440
      Width           =   780
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�W���I��"
      Height          =   240
      Index           =   5
      Left            =   645
      TabIndex        =   68
      Top             =   1320
      Width           =   960
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "���_�݌�"
      Height          =   240
      Index           =   2
      Left            =   645
      TabIndex        =   67
      Top             =   1800
      Width           =   960
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����q�ɍ݌�"
      Height          =   240
      Index           =   11
      Left            =   165
      TabIndex        =   66
      Top             =   2400
      Width           =   1440
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����Ϗo�א�"
      Height          =   240
      Index           =   13
      Left            =   165
      TabIndex        =   65
      Top             =   3000
      Width           =   1440
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "���Y�v��p"
      Height          =   240
      Index           =   22
      Left            =   405
      TabIndex        =   64
      Top             =   3720
      Width           =   1200
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "BU�݌�"
      Height          =   240
      Index           =   12
      Left            =   4440
      TabIndex        =   63
      Top             =   2400
      Width           =   720
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����Ϗo�׌���"
      Height          =   240
      Index           =   15
      Left            =   3480
      TabIndex        =   62
      Top             =   3000
      Width           =   1680
   End
   Begin VB.Label Label 
      Alignment       =   1  '�E����
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   21
      Left            =   5040
      TabIndex        =   61
      Top             =   3360
      Width           =   120
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "���Y�v��p"
      Height          =   240
      Index           =   23
      Left            =   3960
      TabIndex        =   60
      Top             =   3720
      Width           =   1200
   End
   Begin VB.Label Label 
      Alignment       =   1  '�E����
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   25
      Left            =   5040
      TabIndex        =   59
      Top             =   4080
      Width           =   120
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���Y��"
      Height          =   255
      Index           =   32
      Left            =   12600
      TabIndex        =   58
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���i����(*)"
      Height          =   255
      Index           =   31
      Left            =   7320
      TabIndex        =   56
      Top             =   1200
      Width           =   780
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���v"
      Height          =   255
      Index           =   30
      Left            =   13200
      TabIndex        =   55
      Top             =   8880
      Width           =   495
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�݌ɐ�"
      Height          =   255
      Index           =   29
      Left            =   11400
      TabIndex        =   54
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���ד�"
      Height          =   255
      Index           =   28
      Left            =   9960
      TabIndex        =   53
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "  �I��"
      Height          =   255
      Index           =   27
      Left            =   8040
      TabIndex        =   52
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label 
      Alignment       =   1  '�E����
      BackColor       =   &H00FFFFFF&
      Caption         =   "�@��(1)"
      Height          =   255
      Index           =   26
      Left            =   3360
      TabIndex        =   51
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   1  '�E����
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   24
      Left            =   1485
      TabIndex        =   50
      Top             =   4080
      Width           =   120
   End
   Begin VB.Label Label 
      Alignment       =   1  '�E����
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   16
      Left            =   1485
      TabIndex        =   46
      Top             =   3360
      Width           =   120
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0:��Ώ�/1:�Ώ�/2:�Ő؈ē���/3:�Ő�"
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   41
      Top             =   5760
      Width           =   4695
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0:��Ώ�/1:�Ώ�/2:�Ő؈ē���/3:�Ő�"
      Height          =   255
      Index           =   9
      Left            =   2760
      TabIndex        =   40
      Top             =   5280
      Width           =   4575
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0:�P�i/1�ƯĐe/2:�ƯĎq/3:�P�i�Ư�"
      Height          =   255
      Index           =   8
      Left            =   2760
      TabIndex        =   39
      Top             =   4800
      Width           =   4455
   End
   Begin VB.Label Label 
      Alignment       =   1  '�E����
      BackColor       =   &H00FFFFFF&
      Caption         =   "�W���P��"
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
      Left            =   1440
      TabIndex        =   37
      Top             =   6960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label 
      Alignment       =   1  '�E����
      BackColor       =   &H00FFFFFF&
      Caption         =   "�C�O�������i�敪"
      Height          =   255
      Index           =   6
      Left            =   -120
      TabIndex        =   33
      Top             =   5760
      Width           =   2295
   End
   Begin VB.Label Label 
      Alignment       =   1  '�E����
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����������i�敪"
      Height          =   255
      Index           =   4
      Left            =   -120
      TabIndex        =   32
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Label Label 
      Alignment       =   1  '�E����
      BackColor       =   &H00FFFFFF&
      Caption         =   "�Ưĕ��i�敪"
      Height          =   255
      Index           =   3
      Left            =   -120
      TabIndex        =   31
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  '�E����
      BorderStyle     =   1  '����
      Height          =   375
      Index           =   0
      Left            =   8640
      TabIndex        =   27
      Top             =   9120
      Width           =   1380
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�{"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   10080
      TabIndex        =   26
      Top             =   9120
      Width           =   435
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  '�E����
      BorderStyle     =   1  '����
      Height          =   375
      Index           =   1
      Left            =   10440
      TabIndex        =   25
      Top             =   9120
      Width           =   1380
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   11880
      TabIndex        =   24
      Top             =   9120
      Width           =   435
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  '�E����
      BorderStyle     =   1  '����
      Height          =   375
      Index           =   2
      Left            =   12360
      TabIndex        =   23
      Top             =   9120
      Width           =   1380
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
      TabIndex        =   21
      Top             =   10200
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����O"
      Height          =   255
      Index           =   33
      Left            =   1440
      TabIndex        =   20
      Top             =   8400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�ԁi�����j"
      Height          =   255
      Index           =   14
      Left            =   4440
      TabIndex        =   19
      Top             =   720
      Width           =   1485
   End
   Begin VB.Label Label 
      Alignment       =   1  '�E����
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i��"
      Height          =   255
      Index           =   1
      Left            =   8400
      TabIndex        =   18
      Top             =   720
      Width           =   750
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   " �i�ԁi�O���j"
      Height          =   240
      Index           =   0
      Left            =   165
      TabIndex        =   17
      Top             =   720
      Width           =   1560
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���i����(*)"
      Height          =   255
      Index           =   17
      Left            =   8640
      TabIndex        =   28
      Top             =   8880
      Width           =   1380
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����i"
      Height          =   255
      Index           =   19
      Left            =   11040
      TabIndex        =   29
      Top             =   8880
      Width           =   750
   End
   Begin VB.Menu MainMenu 
      Caption         =   "���ƕ�"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1040101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pcmbNAIGAI% = 0           '�����O

Private Const ptxHin_Gai% = 0           '�i�ԁi�O���j
Private Const ptxHin_Nai% = 1           '�i�ԁi�����j
Private Const ptxHin_Name% = 2          '�i��
Private Const ptxSt_Tana% = 3           '�W���I��
Private Const ptxHS_ZAIQTY% = 4         '�z�X�g�݌�
    
Private Const ptxPPSC_ZAI_QTY% = 5      'PPSC�݌�       '2007.06.16
Private Const ptxBU_ZAI_QTY% = 6        'BU�݌�         '2007.06.16
    
Private Const ptxAVE_SYUKA% = 7         '�����Ϗo�א�   '2007.06.16
    
    
Private Const ptxUNIT_BUHIN% = 8        '�Ưĕ��i�敪
Private Const ptxNAI_BUHIN% = 9         '�����������i�敪
Private Const ptxGAI_BUHIN% = 10        '�C�O�������i�敪
Private Const ptxHYO_TANKA% = 11        '�W���P��
    
Private Const ptxAVE_SYUKA_CNT% = 12    '�����Ϗo�׌��� 2011.07.02
Private Const ptxS_AVE_SYUKA_QTY1% = 13 '���Y�v��p�����Ϗo�א�(1) 2011.07.02
Private Const ptxS_AVE_SYUKA_QTY2% = 14 '���Y�v��p�����Ϗo�א�(2) 2011.07.02
    
Private Const ptxL_KISHU1% = 15         '��\�@��@2012.12.22
    
    
    
    
Private Const pLstZaiko% = 0            '�݌�ؽ�
    
Private Const Text_Max% = 3

'Private Const LAST_UPDATE_DAY$ = " (F104010 2017.04.28 14:00)"
'Private Const LAST_UPDATE_DAY$ = " (F104010 2018.09.18 11:30)"
'Private Const LAST_UPDATE_DAY$ = " [F104010] 2019.07.31 15:00"
Private Const LAST_UPDATE_DAY$ = " [F104010] 2019.08.01 17:00 �Γ��i��20���Ή�"

Private Function List_Dsp() As Integer

Dim sts             As Integer
Dim com             As Integer
Dim i               As Integer
Dim NAIGAI          As String * 1
Dim RetBuf          As String
Dim Edit            As String
    
Dim GK_GOODS_ON     As Long
Dim GK_GOODS_OFF    As Long
    
    List_Dsp = True
    
    Call Input_Lock
    
    List1.Clear
    List2.Clear
    
    
    
    If Combo(pcmbNAIGAI).Text = NAIGAI1 Then
        NAIGAI = NAIGAI_NAI
    Else
        NAIGAI = NAIGAI_GAI
    End If
        
                                            '�݌Ƀf�[�^�Ǎ���
    Call UniCode_Conv(K1_ZAIKO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K1_ZAIKO.NAIGAI, NAIGAI)
    Call UniCode_Conv(K1_ZAIKO.HIN_GAI, RTrim(Text(ptxHin_Gai).Text))
    Call UniCode_Conv(K1_ZAIKO.GOODS_ON, "")
    Call UniCode_Conv(K1_ZAIKO.NYUKA_DT, "")
    Call UniCode_Conv(K1_ZAIKO.Soko_No, "")
    Call UniCode_Conv(K1_ZAIKO.Retu, "")
    Call UniCode_Conv(K1_ZAIKO.Ren, "")
    Call UniCode_Conv(K1_ZAIKO.Dan, "")
    
    com = BtOpGetGreaterEqual
    
    GK_GOODS_ON = 0
    GK_GOODS_OFF = 0
    
    Do
        sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
        Select Case sts
            Case BtNoErr
                
                If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> NAIGAI Or _
                    RTrim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) <> RTrim(Text(ptxHin_Gai).Text) Then
                    Exit Do
                End If
                        
                If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = GOODS_OFF Then
                    Edit = "  "
                Else
                    Edit = "* "
                End If
                        
                        
Edit = Edit & Space(5)
                        
                Edit = Edit & StrConv(ZAIKOREC.Soko_No, vbUnicode) & "-" & StrConv(ZAIKOREC.Retu, vbUnicode) & "-" & StrConv(ZAIKOREC.Ren, vbUnicode) & "-" & StrConv(ZAIKOREC.Dan, vbUnicode) & " "
                
Edit = Edit & Space(2)
                
                Edit = Edit & Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 1, 4) & "/" & Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 5, 2) & "/" & Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 7, 2) & " "
                RetBuf = Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)), "#,##0")
                
                
                If Len(Trim(RetBuf)) < 7 Then
                    RetBuf = Space(7 - Len(Trim(RetBuf))) & Trim(RetBuf)
                End If
Edit = Edit & Space(1)
                Edit = Edit + RetBuf + " "
                
Edit = Edit & Space(2)
                
                '2010.07.14 ��
                Edit = Edit + Trim(StrConv(ZAIKOREC.GENSANKOKU, vbUnicode))
                '2010.07.14 ��
                
                If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = GOODS_ON Then
                    GK_GOODS_ON = GK_GOODS_ON + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                Else
                    GK_GOODS_OFF = GK_GOODS_OFF + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                End If
                
                List1.AddItem Edit
            
            
            
            
                List2.AddItem StrConv(ZAIKOREC.SHIIRE_CODE, vbUnicode) & " " & StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode)
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call Input_UnLock
                Call File_Error(sts, com, "�݌Ƀf�[�^")
                List_Dsp = True
                Exit Function
        End Select
        
        
        com = BtOpGetNext
    
    Loop
    
    lblTotal(0).Caption = Format(GK_GOODS_ON, "#,##0")
    lblTotal(1).Caption = Format(GK_GOODS_OFF, "#,##0")
    lblTotal(2).Caption = Format(GK_GOODS_ON + GK_GOODS_OFF, "#,##0")
    
    Call Input_UnLock
    
    List_Dsp = False

End Function
                                    '��ʏ�����Ԃ�ݒ肷��
Private Sub Clear_Field(Mode As Integer)

Dim i   As Integer


    For i = Mode To Text_Max
            Text(i).Text = ""
    Next i
    
    List1.Clear
End Sub

                                    '�i�ڃ}�X�^���e���ڂ�\������
Private Function Item_Read_Proc() As Integer

Dim sts     As Integer
Dim NAIGAI  As String * 1

    Item_Read_Proc = True
                                                '�����O�̔���
    If Combo(pcmbNAIGAI).Text = NAIGAI1 Then
        NAIGAI = NAIGAI_NAI
    Else
        NAIGAI = NAIGAI_GAI
    End If
                                                
                                                '�܂��O���i�Ԃœǂݍ���
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text(ptxHin_Gai).Text)
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            
            Text(ptxHin_Nai).Text = StrConv(ITEMREC.HIN_NAI, vbUnicode)
            Text(ptxHin_Name).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
            Text(ptxSt_Tana).Text = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
        
        
        
            Text(ptxUNIT_BUHIN).Text = StrConv(ITEMREC.UNIT_BUHIN, vbUnicode)
            Text(ptxNAI_BUHIN).Text = StrConv(ITEMREC.NAI_BUHIN, vbUnicode)
            Text(ptxGAI_BUHIN).Text = StrConv(ITEMREC.GAI_BUHIN, vbUnicode)
        
            If IsNumeric(StrConv(ITEMREC.HYO_TANKA, vbUnicode)) Then
                Text(ptxHYO_TANKA).Text = Format(CDbl(StrConv(ITEMREC.HYO_TANKA, vbUnicode)), "#0.00")
            Else
                Text(ptxHYO_TANKA).Text = ""
            End If
        
        
            Text(ptxL_KISHU1).Text = Trim(StrConv(ITEMREC.L_KISHU1, vbUnicode))     '2012.12.22
        
        Case BtErrKeyNotFound
                                                '�����i�Ԃōēx�ǂݍ���
            Call UniCode_Conv(K2_ITEM.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(K2_ITEM.NAIGAI, NAIGAI)
            Call UniCode_Conv(K2_ITEM.HIN_NAI, Text(ptxHin_Nai).Text)
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K3_ITEM, Len(K3_ITEM), 3)
            Select Case sts
                Case BtNoErr
                    Text(ptxHin_Gai).Text = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                    Text(ptxHin_Name).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                    Text(ptxSt_Tana).Text = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_DAN, vbUnicode)


                    Text(ptxUNIT_BUHIN).Text = StrConv(ITEMREC.UNIT_BUHIN, vbUnicode)
                    Text(ptxNAI_BUHIN).Text = StrConv(ITEMREC.NAI_BUHIN, vbUnicode)
                    Text(ptxGAI_BUHIN).Text = StrConv(ITEMREC.GAI_BUHIN, vbUnicode)
                
                    If IsNumeric(StrConv(ITEMREC.HYO_TANKA, vbUnicode)) Then
                        Text(ptxHYO_TANKA).Text = Format(CDbl(StrConv(ITEMREC.HYO_TANKA, vbUnicode)), "#0.00")
                    Else
                        Text(ptxHYO_TANKA).Text = ""
                    End If
                    
                    
                    Text(ptxL_KISHU1).Text = Trim(StrConv(ITEMREC.L_KISHU1, vbUnicode))     '2012.12.22
        
                Case BtErrKeyNotFound
        
                    Exit Function
        
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                    Item_Read_Proc = SYS_ERR
                    Exit Function
            End Select
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Item_Read_Proc = SYS_ERR
            Exit Function
    End Select
                                    '�݌ɏW�v�f�[�^���z�X�g���_�݌Ɋl��
    Call UniCode_Conv(K0_SUMZ.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_SUMZ.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_SUMZ.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
    
    Select Case sts
        Case BtNoErr
            Text(ptxHS_ZAIQTY).Text = Format(CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode)), "#,##0")
            '2007.06.16
            If IsNumeric(StrConv(SUMZREC.PPSC_ZAI_QTY, vbUnicode)) Then
                Text(ptxPPSC_ZAI_QTY).Text = Format(CLng(StrConv(SUMZREC.PPSC_ZAI_QTY, vbUnicode)), "#,##0")
            Else
                Text(ptxPPSC_ZAI_QTY).Text = 0
            End If
            
            '2007.06.16
            If IsNumeric(StrConv(SUMZREC.BU_ZAI_QTY, vbUnicode)) Then
                Text(ptxBU_ZAI_QTY).Text = Format(CLng(StrConv(SUMZREC.BU_ZAI_QTY, vbUnicode)), "#,##0")
            Else
                Text(ptxBU_ZAI_QTY).Text = 0
            End If
        Case BtErrKeyNotFound
            Text(ptxHS_ZAIQTY).Text = ""
            '2007.06.16
            Text(ptxPPSC_ZAI_QTY).Text = ""
            '2007.06.16
            Text(ptxBU_ZAI_QTY).Text = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�݌ɏW�v�f�[�^")
            Item_Read_Proc = SYS_ERR
            Exit Function
    End Select
            
                                    '�����Ϗo�א��W�v�f�[�^��茎���Ϗo�א��l�� 2007.06.16
    Call UniCode_Conv(K0_AVE_SYUKA.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_AVE_SYUKA.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_AVE_SYUKA.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
    
    Select Case sts
        Case BtNoErr
            Text(ptxAVE_SYUKA).Text = Format(Val(StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode)), "#,##0.0")
        
            '2011.07.02
            Text(ptxAVE_SYUKA_CNT).Text = Format(Val(StrConv(AVE_SYUKAREC.TOTAL_AVE_CNT, vbUnicode)), "#,##0.0")
            Text(ptxS_AVE_SYUKA_QTY1).Text = Format(Val(StrConv(AVE_SYUKAREC.S_AVE_SYUKA_QTY1, vbUnicode)), "#,##0.0")
            Text(ptxS_AVE_SYUKA_QTY2).Text = Format(Val(StrConv(AVE_SYUKAREC.S_AVE_SYUKA_QTY2, vbUnicode)), "#,##0.0")
            '2011.07.02
        
        
        
        
        Case BtErrKeyNotFound
            Text(ptxAVE_SYUKA).Text = ""
        
            '2011.07.02
            Text(ptxAVE_SYUKA_CNT).Text = ""
            Text(ptxS_AVE_SYUKA_QTY1).Text = ""
            Text(ptxS_AVE_SYUKA_QTY2).Text = ""
            '2011.07.02
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�����Ϗo�א��W�v�f�[�^")
            Item_Read_Proc = SYS_ERR
            Exit Function
    End Select
            
            
            
    lblNOW.Caption = Format(Now, "yyyy/mm/dd HH:MM")        '2018.09.18
            
            
            
            
    Item_Read_Proc = False

End Function

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F1040101.MousePointer = vbHourglass

    Call Ctrl_Lock(F1040101)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1040101)


    F1040101.MousePointer = vbDefault

End Sub


Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Select Case Index
        Case pcmbNAIGAI
            Call Clear_Field(0)
            Text(ptxHin_Gai).SetFocus
    End Select

End Sub



Private Sub Command_Click(Index As Integer)

Dim yn As Integer
Dim sts As Integer
    
    Select Case Index
        Case 7                              '�ŐV�\��
            
            
    
            Text(ptxHin_Gai).Text = StrConv(RTrim(Text(ptxHin_Gai).Text), vbUpperCase)
    
            
            
            If Len(Trim(Text(ptxHin_Gai).Text)) <> 0 Then
                sts = Item_Read_Proc()
                Select Case sts
                    Case False
                    Case True
                        Beep
                        MsgBox "���͂������ڂ̓G���[�ł�� (���o�^�G���[)"
                        Text(ptxHin_Gai).SetFocus
                        Exit Sub
                    Case SYS_ERR
                        Unload Me
                End Select
                        
                If List_Dsp() Then
                    Unload Me
                End If
                
                Text(ptxHin_Gai).SetFocus
                        
            Else
                If Len(Trim(Text(ptxHin_Nai).Text)) <> 0 Then
                    Text(ptxHin_Nai).Text = StrConv(RTrim(Text(ptxHin_Nai).Text), vbUpperCase)
                    sts = Item_Read_Proc()
                    Select Case sts
                        Case False
                        Case True
                            Beep
                            MsgBox "���͂������ڂ̓G���[�ł�� (���o�^�G���[)"
                            Text(ptxHin_Nai).SetFocus
                            Exit Sub
                        Case SYS_ERR
                            Unload Me
                    End Select
                        
                    If List_Dsp() Then
                        Unload Me
                    End If
            
                    Text(ptxHin_Nai).SetFocus
                
                End If
            End If
                        
        Case 11                             '�I��
            Unload Me
        Case Else
            Beep
    End Select
    
End Sub

Private Sub Command1_Click()
    
    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "�݌ɖ⍇�킹�i�i�ԕʁj ��ʈ�����J�n���܂��� ", Me.hwnd, 0)
    
    
    
    
    Call Form_HCopy_Win7_NEW(Picture1, vbPRPSA4, vbPRORLandscape)   '2017.04.27


    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "�݌ɖ⍇�킹�i�i�ԕʁj ��ʈ�����I�����܂��� ", Me.hwnd, 0)

End Sub

Private Sub Command2_Click() '2019/12/24 F120050 ZENKAI_YMD�ǉ�

    MsgBox "ZENKAI_YMD=" & Chr(13) & Chr(10)

End Sub

Private Sub Form_DblClick()
'    PrintForm                                                      '2016.12.16
'    Call Form_HCopy_Win7(Picture1, vbPRPSA4, vbPRORLandscape)       '2016.12.16
    
    
'    Me.Enabled = False
    
'    Call Form_HCopy_Win7_NEW(Picture1, vbPRPSA4, vbPRORLandscape)   '2017.02.14        2017.04.27

'    Me.Enabled = True

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
    
Dim TUKI1   As Integer  '2011.07.02
Dim TUKI2   As Integer  '2011.07.02
Dim TUKI3   As Integer  '2011.07.02
        
'    If App.PrevInstance Then
'        Beep
'        MsgBox "����v���O�������s���ł��B"
'        End
'    End If
    
    
    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "�݌ɖ⍇�킹�i�i�ԕʁj", Me.hwnd, 0)
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
                                

                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            Me.Caption = "�݌ɖ⍇�킹(�i�ԕ�)(" + RTrim(JGYOBU_T(i).NAME) + ")" & LAST_UPDATE_DAY
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)

                                
                                '�݌Ƀf�[�^OPEN
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌ɏW�v�f�[�^�n�o�d�m
    If SUMZ_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�����Ϗo�א��W�v�f�[�^�n�o�d�m 2007.06.15
    If AVE_SYUKA_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
'------------------------------------   2011.07.02  ���ϊ��Ԃ̊l��
'    If GetIni(App.EXEName, "TUKI1", "F120050", c) Then                 '2016.12.15 SYS.INI --> F120050 ��
    If GetIni("F120050", "TUKI1", "F120050", c) Then                    '2016.12.15
        TUKI1 = 3
    Else
        If Not IsNumeric(RTrim(c)) Then
            TUKI1 = 3
        Else
            TUKI1 = Val(RTrim(c))
        End If
    End If
    Label(16).Caption = "(" & Format(TUKI1, "#0") & "����)"
    Label(21).Caption = "(" & Format(TUKI1, "#0") & "����)"


'    If GetIni(App.EXEName, "TUKI2", "F120050", c) Then                 '2016.12.15 SYS.INI --> F120050 ��
    If GetIni("F120050", "TUKI2", "F120050", c) Then                    '2016.12.15
        TUKI2 = 3
    Else
        If Not IsNumeric(RTrim(c)) Then
            TUKI2 = 3
        Else
            TUKI2 = Val(RTrim(c))
        End If
    End If
    Label(24).Caption = "�o�א�(" & Format(TUKI2, "#0") & "����)"


'    If GetIni(App.EXEName, "TUKI3", "F120050", c) Then                 '2016.12.15 SYS.INI --> F120050 ��
    If GetIni("F120050", "TUKI3", "F120050", c) Then                    '2016.12.15
        TUKI3 = 12
    Else
        If Not IsNumeric(RTrim(c)) Then
            TUKI3 = 12
        Else
            TUKI3 = Val(RTrim(c))
        End If
    End If
    Label(25).Caption = "�o�א�(" & Format(TUKI3, "#0") & "����)"







'------------------------------------   2011.07.01
                                
                                
                                
                                
                                '��ʏ����ݒ�
    Call Clear_Field(0)
    
    
                                '�����O��荞��
    Combo(pcmbNAIGAI).AddItem NAIGAI1
    Combo(pcmbNAIGAI).AddItem NAIGAI2
    Combo(pcmbNAIGAI).Text = NAIGAI1
    
'    Combo(pcmbNAIGAI).SetFocus
    Text(ptxHin_Gai).SetFocus
    
    End Sub



Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        End If
    End If
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '�݌ɏW�v�f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌ɏW�v�f�[�^")
        End If
    End If
                                            '�����Ϗo�א��W�v�f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�����Ϗo�א��f�[�^")
        End If
    End If
    
    sts = BTRV(BtOpReset, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    Set F1040101 = Nothing

    End
End Sub



Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    '���j���[���I���v��
'    If JGYOBU_T(Index).CODE = " " Then
'        Unload Me
'    End If




    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '���ƕ��؂�ւ�
    Me.Caption = "�݌ɖ⍇�킹�i�i�ԕʁj�i" + RTrim(JGYOBU_T(Index).NAME) + ") " & LAST_UPDATE_DAY
    Last_JGYOBU = JGYOBU_T(Index).CODE
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

Dim i As Integer
Dim sts As Integer
    If KeyCode <> vbKeyReturn Then Exit Sub

    If Index = 0 Or Index = 1 Then
    
        Text(Index).Text = StrConv(RTrim(Text(Index).Text), vbUpperCase)
    
    End If



    Select Case Index
        Case ptxHin_Gai             '�O���i��
            
            If Len(Trim(Text(ptxHin_Gai).Text)) <> 0 Then
                sts = Item_Read_Proc()
                Select Case sts
                    Case False
                    Case True
                        Beep
                        MsgBox "���͂������ڂ̓G���[�ł�� (���o�^�G���[)"
                        Text(Index).SetFocus
                        Exit Sub
                    Case SYS_ERR
                        Unload Me
                End Select
                        
                If List_Dsp() Then
                    Unload Me
                End If
                        
            
            End If
        Case ptxHin_Nai             '�����i��
            If Len(Trim(Text(ptxHin_Gai).Text)) = 0 Then
                sts = Item_Read_Proc()
                Select Case sts
                    Case False
                    Case True
                        Beep
                        MsgBox "���͂������ڂ̓G���[�ł�� (���o�^�G���[)"
                        Text(Index).SetFocus
                        Exit Sub
                    Case SYS_ERR
                        Unload Me
                End Select
            
                If List_Dsp() Then
                    Unload Me
                End If
            
                Text(Index).SetFocus
            
            End If
    End Select
        
    For i = Index + 1 To Text_Max
        If Text(i).Enabled Then
            Text(i).SetFocus
            Exit For
        End If
    Next i

End Sub


Private Sub Text_LostFocus(Index As Integer)
    
    If Index = 0 Or Index = 1 Then
    
        Text(Index).Text = StrConv(RTrim(Text(Index).Text), vbUpperCase)
    
    End If

End Sub
