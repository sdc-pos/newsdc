VERSION 5.00
Begin VB.Form F1020121 
   BackColor       =   &H00FFFFFF&
   Caption         =   "���o�ח\��f�[�^�捞�݁u�܈�p�v"
   ClientHeight    =   6324
   ClientLeft      =   1908
   ClientTop       =   2388
   ClientWidth     =   11220
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
   ScaleHeight     =   6324
   ScaleWidth      =   11220
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   1
      Left            =   4800
      TabIndex        =   38
      Top             =   1560
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   0
      Left            =   3480
      TabIndex        =   37
      Top             =   1560
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ListBox LBox_Dup 
      ForeColor       =   &H00000080&
      Height          =   288
      Left            =   9840
      Sorted          =   -1  'True
      TabIndex        =   31
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton SelCmd 
      Caption         =   "�Ď捞"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1440
      MaskColor       =   &H0080C0FF&
      TabIndex        =   4
      Top             =   3000
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.CommandButton SelCmd 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   1440
      MaskColor       =   &H0080C0FF&
      TabIndex        =   3
      Top             =   2400
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.ListBox LBox_Hin 
      ForeColor       =   &H00000080&
      Height          =   288
      Left            =   9840
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox LBox_Etc 
      ForeColor       =   &H00000080&
      Height          =   288
      Left            =   9840
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton SelCmd 
      Caption         =   "�P��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1440
      MaskColor       =   &H0080C0FF&
      TabIndex        =   1
      Top             =   1200
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.CommandButton SelCmd 
      Caption         =   "�Q��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1440
      MaskColor       =   &H0080C0FF&
      TabIndex        =   2
      Top             =   1800
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.CommandButton SelCmd 
      Caption         =   "�R��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1440
      MaskColor       =   &H0080C0FF&
      TabIndex        =   0
      Top             =   600
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.CommandButton Command 
      Caption         =   "�I  ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   10320
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   9480
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   8640
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   7800
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6480
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5640
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4800
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3960
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2640
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�^"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Index           =   3
      Left            =   8400
      TabIndex        =   42
      Top             =   4080
      Width           =   492
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   2
      Left            =   9000
      TabIndex        =   41
      Top             =   4080
      Width           =   1452
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   6720
      TabIndex        =   40
      Top             =   4080
      Width           =   1452
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�捞�ݏ������@�捞�݌�����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   0
      Left            =   1680
      TabIndex        =   39
      Top             =   4080
      Width           =   4932
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   4920
      TabIndex        =   36
      Top             =   1320
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "SPIC"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   3720
      TabIndex        =   35
      Top             =   1320
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�o�׏d����������I"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   9
      Left            =   1920
      TabIndex        =   34
      Top             =   5040
      Visible         =   0   'False
      Width           =   4320
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�i�ԕύX�ۗ����E�o�׏d�����j"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   2640
      TabIndex        =   32
      Top             =   3240
      Visible         =   0   'False
      Width           =   3780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�o�׏d��"
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   9840
      TabIndex        =   30
      Top             =   2880
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�ۗ��f�[�^�ď������I"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   0
      Left            =   1920
      TabIndex        =   27
      Top             =   5040
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�ԕύX���X�g������I"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   7
      Left            =   2520
      TabIndex        =   26
      Top             =   5040
      Visible         =   0   'False
      Width           =   5280
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�ԓ`�A�����A�o�׊m�F�ꗗ������I"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   6
      Left            =   1440
      TabIndex        =   24
      Top             =   5040
      Visible         =   0   'False
      Width           =   7680
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "��M�G���[���X�g������I"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   5
      Left            =   2400
      TabIndex        =   22
      Top             =   5040
      Visible         =   0   'False
      Width           =   5760
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�ԕύX"
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   9840
      TabIndex        =   20
      Top             =   1920
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "��M�װؽ�"
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   9840
      TabIndex        =   18
      Top             =   840
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�R�֎捞�ݏ������I"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   3
      Left            =   3000
      TabIndex        =   12
      Top             =   5040
      Visible         =   0   'False
      Width           =   4320
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�Q�֎捞�ݏ������I"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   2
      Left            =   3240
      TabIndex        =   11
      Top             =   5040
      Visible         =   0   'False
      Width           =   4320
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�P�֎捞�ݏ������I"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   1
      Left            =   3000
      TabIndex        =   9
      Top             =   5040
      Visible         =   0   'False
      Width           =   4320
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�u�R�ցv�u�P�ցv�u�Q�ցv���]�\���͖{���捞�ݍρj"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   2640
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   6300
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�O�؂���ׁA���юc�`�F�b�N���I�I"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   8
      Left            =   1440
      TabIndex        =   29
      Top             =   5040
      Visible         =   0   'False
      Width           =   7680
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����捞�ݏ������I"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   4
      Left            =   1920
      TabIndex        =   33
      Top             =   5040
      Visible         =   0   'False
      Width           =   4320
   End
End
Attribute VB_Name = "F1020121"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'#################################################################################################
'�m�e�L�X�g�t�@�C�������̒��ӁI�I�n
'
'�@�@���̃v���O�����ł́A�e�L�X�g�t�@�C�������[�U�[��`�^�̍\���̂œǏ������Ă���B
'�@�@���̌`���ɂ��h�|�n�́A�uGET�v����сuPUT�v�ð���Ăɂ��s�����ɂȂ邪�A�����̽ð����
'�@�@���g�p����ׂɂ́A�uRANDOM�v�܂��́uBINARY�v���[�h�ŁuOPEN�v���Ȃ���΂Ȃ�Ȃ��B
'�@�@�A���A�ȉ��̃��W�b�N�ł́A�e�L�X�g�t�@�C���̑��݂��`�F�b�N����ړI�ŁuINPUT�v���[�h�ɂ��
'�@�@�uOPEN�v���s���Ă���A���݃`�F�b�N��A�����ɓǍ��݂��s���l�ȏꍇ�ɂ́A��U�uCLOSE�v���Ă���
'�@�@�uBINARY���[�h��OPEN�v���Ă��鎖�ɒ��ӂ��K�v�B
'
'�@�@���D�uINPUT�v���[�hOPEN�ł́A�uINPUT#�v�uOUTPUT#�v�̂ݎg�p�ł��邪�A�����͍\���̂̃����o
'�@�@�@�@�P�ʂ̂h�|�n�ɂȂ�B
'#################################################################################################
Private Type YUKO_SOKO_TBL1             '�L��νđq�Ɏ�荞�݃e�[�u���i����@���ƕ��j
    HS_SOKO As String * 2
    NAIGAI As String * 1
End Type
Dim SOKO_T1(ZERO To 9) As YUKO_SOKO_TBL1

Dim WS_NO As String * 2                 'ܰ��ð��ݔԍ�

Dim HS_NaiG As String                   '�����O�i������e�j�����@ν��ް����e�ɂ��ݒ�
Dim BEF_GAI As String * 13              '�ύX�O�i�ԁi�O���j
Dim BEF_NAI As String * 13              '�� �X�O�i�ԁi�����j

Dim PRT_CAN As Boolean                  '����r���L�����Z���v��
Dim NormalFont As New StdFont           '����t�H���g

Private Const LMAX% = 46                '�œ��ő�s��
Private Const MGN_L% = 1                '���׈���J�n���ʒu�i�P����j
Private Const MGN_U% = 1                '��]���i�s���F�P����j
Private Const MGN_L2% = 20              '�u�ߏ�O�ؕiؽāv���׈���J�n���ʒu�i�P����j
Private Const MGN_U2% = 1               '�u�ߏ�O�ؕiؽāv��]���i�s���F�P����j
Dim Pdate As String                     '����J�n���t�iͯ�ް�p�j
Dim Ptime As String                     '����J�n�����iͯ�ް�p�j

Dim Proc_F As Integer                   '�i�ԁ��݌ɗL���@����t���O
Dim Last_Proc_F As Integer              '���������ް��폜�����@���s�L���t���O

Dim Shori_Mode  As Integer
Private Const Text_Max% = 1
                                    
Private Const Er_Soko_NoT% = True       '�z�X�g�q�Ɉُ�
Private Const Er_Item_NoT% = True       '�`�[���t�^�i�ԁ^���o�ɋ敪�ُ�
Private Const Er_Dup_NoT% = True        '�`�[�d��
Private Const Er_Muke_NoT% = True       '������ُ�
                                    
                                    '��ʏ����\���i�����ρu�ցv�\���Ȃǁj
Private Sub Scr_Init()
Dim i As Integer
Dim sts As Integer
Dim Work As String

    For i = 1 To 3
        Work = Format(Date, "ddd") & "." & Format(i, "0") & "00"
        sts = XX_SIJ_Open(Work, ZERO)
        If sts = False Then
            Close #XX_SIJ_No
            If Format(FileDateTime(Work), "yyyy/mm/dd") = Format(Date, "yyyy/mm/dd") Then
                SelCmd(i).Enabled = False
            Else
                SelCmd(i).Enabled = True
            End If
        Else
            SelCmd(i).Enabled = True
        End If
    Next i


End Sub
                                            '�z�X�g�f�[�^�捞�ݏ���
Private Function Data_Inport() As Integer
Dim sts         As Integer
Dim ans         As Integer
Dim Command     As Integer
Dim FPass       As String
Dim Work        As String
Dim FP_XX_SIJ   As String
Dim FP_ER_SIJ   As String
Dim FP_CHGHIN   As String
Dim FP_SYUDUP   As String

Dim In_Cnt      As Integer

    On Error Resume Next    'FileCopy / Kill �ð���Ăł�̧�ٖ����͎��ï�߂����s

    Call Input_Lock                                 '��ʍ��ڃ��b�N

    MsgLab(Shori_Mode).Visible = True       '�X�V��ү���ޕ\��
    DoEvents

'�z�X�g�֕ʃ��[�N�@�폜

    sts = GetIni("FILE", XX_SIJ_ID, "SYS", FPass)
    If sts <> False Then
        MsgBox "SYS.INI [FILE][" & XX_SIJ_ID & "]�ǂݍ��݃G���[ "
        Call Log_Out(LOG_F, "SYS.INI [FILE][" & XX_SIJ_ID & "]�ǂݍ��݃G���[ ")
        Unload Me
    End If
    FP_XX_SIJ = RTrim(FPass) & Format(Date, "ddd") & "." & Format(Shori_Mode, "0") & "00"
    Kill FP_XX_SIJ
    FP_ER_SIJ = RTrim(FPass) & Format(Date, "ddd") & "." & Format(Shori_Mode, "0") & "0E"
    Kill FP_XX_SIJ

'̧�� OPEN
    Work = Format(Date, "ddd") & "." & Format(Shori_Mode, "0") & "00"
    If XX_SIJ_Open(Work, 1) Then            '�捞��ܰ�
        Unload Me
    End If

    Work = Format(Date, "ddd") & "." & Format(Shori_Mode, "0") & "0E"
    If ER_SIJ_Open(Work, 1) Then            '�捞��ܰ�[�ΏۊO�ް��p]
        Close #XX_SIJ_No
        Unload Me
    End If
    
    If CHGH_Open() Then                     '�i�ԕύX�ۗ��ް�
        Unload Me
    End If
    
    If SYUDUP_Open() Then                   '�o�ח\��d���ް�
        Unload Me
    End If


    Call Data_Load                          '�֕ʎ捞�݃��[�N���ް�۰��

        
    Close #CHGH_No                          '�i�ԕύX�ۗ��ް� CLOSE
    Close #SYUDUP_No                        '�o�ח\��d���ް� CLOSE
                                            
                                            
                                                '�i�ԕύX�ۗ��ް��폜
    sts = GetIni("FILE", CHGH_ID, "SYS", FP_CHGHIN)
    If sts <> False Then
        MsgBox "SYS.INI [FILE][" & CHGH_ID & "]�ǂݍ��݃G���[ "
        Call Log_Out(LOG_F, "SYS.INI [FILE][" & CHGH_ID & "]�ǂݍ��݃G���[ ")
        Unload Me
    End If
    FP_CHGHIN = RTrim(FP_CHGHIN)
    Kill FP_CHGHIN                          '�i�ԕύX�ۗ��ް� �N���A

                                                '�o�ח\��d���ް��폜
    sts = GetIni("FILE", SYUDUP_ID, "SYS", FP_SYUDUP)
    If sts <> False Then
        MsgBox "SYS.INI [FILE][" & SYUDUP_ID & "]�ǂݍ��݃G���[ "
        Call Log_Out(LOG_F, "SYS.INI [FILE][" & SYUDUP_ID & "]�ǂݍ��݃G���[ ")
        Unload Me
    End If
    FP_SYUDUP = RTrim(FP_SYUDUP)
    Kill FP_SYUDUP                          '�o�ח\��d���ް� �N���A

        
    If CHGH_Open() Then                     '�i�ԕύX�ۗ��ް�
        Unload Me
    End If
        
    If SYUDUP_Open() Then                   '�o�ח\��d���ް�
        Close #CHGH_No
        Unload Me
    End If
'�z�X�g�f�[�^�@�`�F�b�N���捞��
    LBox_Etc.Clear      '����ް��pؽ��ޯ���@�N���A
    LBox_Hin.Clear
    LBox_Dup.Clear

    In_Cnt = ZERO

    Do
        If XX_SIJ_Get Then          '�捞��ܰ� �Ǎ���
            Exit Do
        End If
        If Left(StrConv(XX_SIJREC.TEXT_NO, vbUnicode), 1) < " " Then    'EOF ?
            Exit Do
        End If

        sts = Data_Chk              '�ް��d���^���ړ��e �`�F�b�N
        
        If sts = False Then
                                    '�ԓ`�E�o�׊m�F�E�����ް��͈���̂� �o�׎��сE�Ǖi�ԕi���ǉ�
            If StrConv(XX_SIJREC.PM_KBN, vbUnicode) = "-" Or _
               StrConv(XX_SIJREC.IO_KBN, vbUnicode) = "2" Or _
               StrConv(XX_SIJREC.IO_KBN, vbUnicode) = "3" Or _
               StrConv(XX_SIJREC.IO_KBN, vbUnicode) = "4" Or _
               StrConv(XX_SIJREC.IO_KBN, vbUnicode) = "5" Then
                Call PDat_Etc_Add("1", " ")         '�ԓ`�E�o�׊m�F�E��������ް��ۑ�
            Else
                Call Proc_Sel(In_Cnt)               '����(0)�C����(1)�ް��X�V
            End If
        End If
    Loop

    Close #XX_SIJ_No            '�捞��ܰ� CLOSE
    Close #ER_SIJ_No            '�捞��ܰ�[�ΏۊO�ް��p] CLOSE
    Close #CHGH_No              '�i�ԕύX�ۗ��ް� CLOSE
    Close #SYUDUP_No            '�o�׏d���ۗ��ް� CLOSE

    MsgLab(Format(Shori_Mode, "0")).Visible = False      '�X�V��ү���� �ر

    Pdate = Date
    Ptime = Time
    Printer.Orientation = vbPRORLandscape       '�p���̒��ӂ���ɂ��Ĉ��
    
'��M�G���[���X�g���
    If LBox_Etc.ListCount > ZERO Then
        MsgLab(5).Visible = True                '�����ү���ޕ\��
        DoEvents
        Set Printer.Font = NormalFont           '����t�H���g�ݒ�
        Call P_Etc_Proc(ZERO)                   '��M�G���[���X�g���
        MsgLab(5).Visible = False               '�����ү���ރN���A
'        DoEvents
    End If

'�ԓ`�A�����A�o�׊m�F�ꗗ�\���
    If LBox_Etc.ListCount > ZERO Then
        MsgLab(6).Visible = True                '�����ү���ޕ\��
        DoEvents
        Set Printer.Font = NormalFont           '����t�H���g�ݒ�
        Call P_Etc_Proc(1)                      '�ԓ`�A�����A�o�׊m�F�ꗗ�\���
        MsgLab(6).Visible = False               '�����ү���ރN���A
'        DoEvents
    End If

'�i�ԕύX���X�g���
    If LBox_Hin.ListCount > ZERO Then
        MsgLab(7).Visible = True                '�����ү���ޕ\��
        DoEvents
        Set Printer.Font = NormalFont           '����t�H���g�ݒ�
        Call P_Hin_Proc                         '�i�ԕύX���X�g���
        MsgLab(7).Visible = False               '�����ү���ރN���A
        DoEvents
    End If

'�o�׏d�����X�g���
    If LBox_Dup.ListCount > ZERO Then
        MsgLab(9).Visible = True                '�����ү���ޕ\��
        DoEvents
        Set Printer.Font = NormalFont           '����t�H���g�ݒ�
        Call P_Dup_Proc                         '�o�׏d�����X�g���
        MsgLab(9).Visible = False               '�����ү���ރN���A
        DoEvents
    End If


    Call Input_UnLock                               '��ʍ��ڃ��b�N����

    Call Scr_Init                                   '��ʃN���A

End Function
                                            '�f�[�^���[�h(�e�ƕ���ν��ް����֕ʎ捞��ܰ��j
Private Sub Data_Load()
Dim sts As Integer
Dim Work As String

'�֕ʎ捞�݃��[�N�Ƀf�[�^���[�h�i�i�ԕۗ��f�[�^�j
    Do
        If CHGH_Get Then            '�ۗ��f�[�^ �Ǎ���
            Exit Do
        End If

        If Left(StrConv(CHGHREC.TEXT_NO, vbUnicode), 1) < " " Then    'EOF ?
            Exit Do
        End If

        If CHGH_Put(1) Then         '�捞��ܰ�������
            Close #CHGH_No
            Unload Me
        End If
    Loop

'�֕ʎ捞�݃��[�N�Ƀf�[�^���[�h�i�o�ח\��f�[�^�j
    Do
        If SYUDUP_Get Then          '�ۗ��f�[�^ �Ǎ���
            Exit Do
        End If

        If Left(StrConv(SYUDUPREC.TEXT_NO, vbUnicode), 1) < " " Then    'EOF ?
            Exit Do
        End If

        If SYUDUP_Put(1) Then       '�捞��ܰ�������
            Close #CHGH_No
            Close #SYUDUP_No
            Unload Me
        End If
    Loop
'�֕ʎ捞�݃��[�N�Ƀf�[�^���[�h�i�u1�`3�ցv�y�сu������v�w�莞�̂� �j
'�d�v�@���@������͂S�ւƂ݂Ȃ��@��
    If Shori_Mode > ZERO Then
                                        '����@���ƕ�  �z�X�g��M�f�[�^�捞��
        If HS_SIJ_Open1(ZERO, Format(Shori_Mode, "0")) = False Then
                                                        '̧�ٖ����Ȃ珈�����Ȃ�
            Close #HS_SIJ_No
                                                        '����@���ƕ�ν��ް� OPEN
            If HS_SIJ_Open1(1, Format(Shori_Mode, "0")) Then
                Unload Me
            End If

            Call Data_Load_Sub              '÷�ć�

            Close #HS_SIJ_No                'ν��ް� CLOSE
        End If
    End If

'�捞�݃��[�N �Ăn�o�d�m
    Close #XX_SIJ_No
    Work = Format(Date, "ddd") & "." & Format(Shori_Mode, "0") & "00"
    If XX_SIJ_Open(Work, 1) Then
        Close #CHGH_No
        Close #SYUDUP_No
        Unload Me
    End If

End Sub
                                            '���ƕ��ʃf�[�^���[�h(���ƕ���ν��ް����֕ʎ捞��ܰ��j
Private Sub Data_Load_Sub()
Dim sts As Integer
Dim Put_Sel As Integer
Dim Work As String
Dim i As Integer

Dim In_Cnt  As Integer


    In_Cnt = ZERO
    
    Do
        DoEvents
        If HS_SIJ_Get Then          'ν��ް� �Ǎ���
            Exit Do
        End If
        If Left(StrConv(HS_SIJREC.TEXT_NO, vbUnicode), 1) < " " Then
            Exit Do
        End If

        In_Cnt = In_Cnt + 1
                                
        Label3(2).Caption = StrConv(Format(In_Cnt, "#0"), vbWide)
                                
                                '÷�ć�
        Call UniCode_Conv(XX_SIJREC.TEXT_NO, StrConv(HS_SIJREC.TEXT_NO, vbUnicode))
                                '���ƕ��敪
        Call UniCode_Conv(XX_SIJREC.JGYOBU, StrConv(HS_SIJREC.JGYOBU, vbUnicode))
                                '�����敪
        Call UniCode_Conv(XX_SIJREC.CYOK_KBN, StrConv(HS_SIJREC.CYOK_KBN, vbUnicode))
                                '�`�[���t
        Call UniCode_Conv(XX_SIJREC.DEN_DT, StrConv(HS_SIJREC.DEN_DT, vbUnicode))
                                '���o�ɋ敪
        Call UniCode_Conv(XX_SIJREC.IO_KBN, StrConv(HS_SIJREC.IO_KBN, vbUnicode))
                                '�ԍ��敪
        Call UniCode_Conv(XX_SIJREC.PM_KBN, StrConv(HS_SIJREC.PM_KBN, vbUnicode))
                                '�`�[���
        Call UniCode_Conv(XX_SIJREC.DEN_SYU, StrConv(HS_SIJREC.DEN_SYU, vbUnicode))
                                '�`�[��
        Call UniCode_Conv(XX_SIJREC.DEN_NO, StrConv(HS_SIJREC.DEN_NO, vbUnicode))
                                '�����敪
        Call UniCode_Conv(XX_SIJREC.CYU_KBN, StrConv(HS_SIJREC.CYU_KBN, vbUnicode))
                                '�i�ԁi�O���j
        Call UniCode_Conv(XX_SIJREC.HIN_GAI, StrConv(HS_SIJREC.HIN_GAI, vbUnicode))
                                '�i�ԁi�����j
        Call UniCode_Conv(XX_SIJREC.HIN_NAI, StrConv(HS_SIJREC.HIN_NAI, vbUnicode))
                                '�i��
        Call UniCode_Conv(XX_SIJREC.HIN_NAME, StrConv(HS_SIJREC.HIN_NAME, vbUnicode))
                                '����
        Call UniCode_Conv(XX_SIJREC.YOTEI_QTY, StrConv(HS_SIJREC.YOTEI_QTY, vbUnicode))
                                '�\�Z�P�ʁi���j
        Call UniCode_Conv(XX_SIJREC.YOSAN_FROM, StrConv(HS_SIJREC.YOSAN_FROM, vbUnicode))
                                '�\�Z�P�ʁi��j
        Call UniCode_Conv(XX_SIJREC.YOSAN_TO, StrConv(HS_SIJREC.YOSAN_TO, vbUnicode))
                                '�q�ɋ敪�iνāj
        Call UniCode_Conv(XX_SIJREC.HOST_SOKO, StrConv(HS_SIJREC.HOST_SOKO, vbUnicode))
                                '�I�ԁiνāj
        Call UniCode_Conv(XX_SIJREC.HOST_TANA, StrConv(HS_SIJREC.HOST_TANA, vbUnicode))
                                '�x����^�o�א�
        Call UniCode_Conv(XX_SIJREC.SYUK_CODE, StrConv(HS_SIJREC.SYUK_CODE, vbUnicode))
                                '�x����^�o�א於
        Call UniCode_Conv(XX_SIJREC.SYUK_NAME, StrConv(HS_SIJREC.SYUK_NAME, vbUnicode))
                                'ں��ޏI�[ϰ�(@)
        Call UniCode_Conv(XX_SIJREC.REC_END, StrConv(HS_SIJREC.REC_END, vbUnicode))
                                'CR.LF
        Call UniCode_Conv(XX_SIJREC.CR_LF, StrConv(HS_SIJREC.CR_LF, vbUnicode))

        Put_Sel = True
                                                '���ƕ��敪�@�͈͊O�H
        For i = ZERO To UBound(JGYOBU_T) - 1
            If JGYOBU_T(i).Code = " " Then
                Put_Sel = False
                Exit For
            End If
            If JGYOBU_T(i).Code = StrConv(HS_SIJREC.JGYOBU, vbUnicode) Then
                Exit For
            End If
        Next i
                                                '�z�X�g�q�Ɂ@�͈͊O�H
        For i = ZERO To UBound(SOKO_T1) - 1
            If SOKO_T1(i).HS_SOKO = "  " Then
                Put_Sel = False
                Exit For
            End If
            If RTrim(StrConv(HS_SIJREC.HOST_SOKO, vbUnicode)) = RTrim(SOKO_T1(i).HS_SOKO) Then
                Exit For
            End If
        Next i
        
        If Put_Sel = True Then
            sts = XX_SIJ_Put                    '�捞��ܰ������݁i�Ώۑq�Ɂj
        Else
        '�ΏۊO�z�X�g�q�ɁI�I   �G���[���O
            If Er_Soko_NoT Then
                Call Err_Log_Out("�g�q�ɑΏۊO")
            End If
            sts = ER_SIJ_Put                    '�捞��ܰ������݁i�ΏۊO�q�Ɂj
        End If
        If sts Then
            Close #HS_SIJ_No
            Close #XX_SIJ_No
            Unload Me
        End If
    Loop

End Sub
                                            '�f�[�^�d���^���ړ��e �`�F�b�N
Private Function Data_Chk() As Integer

Dim sts     As Integer
Dim Work    As String
Dim i       As Integer

Dim MUKECHG As String * 10      '2001.07.04

    Data_Chk = False

'�f�[�^�d���`�F�b�N�i�ď����f�[�^�i�I�[���u�H�v�u���v�̓`�F�b�N�����j
    If StrConv(XX_SIJREC.REC_END, vbUnicode) <> "?" And StrConv(XX_SIJREC.REC_END, vbUnicode) <> "*" Then
        Call UniCode_Conv(K0_SEQCK.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
        If Shori_Mode = 4 Then
            Call UniCode_Conv(K0_SEQCK.SEQ_MODE, "2")
        Else
            Call UniCode_Conv(K0_SEQCK.SEQ_MODE, "1")
        End If
        sts = BTRV(BtOpGetEqual, SEQCK_POS, SEQCKREC, Len(SEQCKREC), K0_SEQCK, Len(K0_SEQCK), ZERO)
        If sts Then
            If sts <> BtErrKeyNotFound Then
                Call File_Error(sts, BtOpGetEqual, "�\��捞�݃`�F�b�N")
                Unload Me
            End If
            Call UniCode_Conv(SEQCKREC.LAST_TXTNO, "000000000")
        End If
                '
                '�u�O��÷�ć�������÷�ć��v�̓G���[
                '�i�A���A�O��÷�ć���12(��)���͖������n�j�j
                '
        If Shori_Mode = 4 Then
                '������f�[�^�̓e�L�X�g���̕ց��T�A�����敪���S�ȊO�G���[
            If Mid(StrConv(XX_SIJREC.TEXT_NO, vbUnicode), 5, 1) <> "5" Or StrConv(XX_SIJREC.CYU_KBN, vbUnicode) <> CYU_KBN_TOK Then
                Data_Chk = True
                Exit Function
            End If
                '����ȑO�̃e�L�X�g���̌����̓G���[
            If Left(StrConv(XX_SIJREC.TEXT_NO, vbUnicode), 4) < Right(Format(Date, "yyyymmdd"), 4) Then
                Data_Chk = True
                Exit Function
            End If
        Else
'-----------�e�L�X�g���`�F�b�N����
'            If Left(StrConv(SEQCKREC.LAST_TXTNO, vbUnicode), 2) <> "12" Then
'                If StrConv(SEQCKREC.LAST_TXTNO, vbUnicode) >= StrConv(XX_SIJREC.TEXT_NO, vbUnicode) Then
'                    Data_Chk = True
'                    Exit Function
'                End If
'            End If
'-----------�e�L�X�g���`�F�b�N����
        End If
    Else
        If Shori_Mode > ZERO Then
                                '�ď����v������Ȃ����
            Select Case StrConv(XX_SIJREC.REC_END, vbUnicode)
                Case "?"
                    If CHin_Put() Then            '�O���i�ԕύX�ۗ��ް��쐬
                        Unload Me
                    End If
                
                Case "*"
                    Call MAKE_SYUDUP_Put                    '�d���ۗ��ް��쐬
            End Select
            Data_Chk = True
            Exit Function
        End If
    End If
'�捞�݃f�[�^ ���ړ��e�`�F�b�N
'   [���� ����]
'       1) �\����@�@�@�F���t����
'       2) �i�ԁi�O���j�F����
'       4) ���o�ɋ敪�@�F�͈́�"0�`3"or"E" �C"=0"�̎��A�o�א恂��
'
'97.07.29  �����i�Ԃ́u�󔒁v���Ƃ���B�A���A�u�󔒁v�̎��̓}�X�^�̕i�Ԃƒu�����Ȃ��B
'
    Work = StrConv(XX_SIJREC.DEN_DT, vbUnicode)     '�`�[���t
    If IsDate(Left(Work, 4) & "/" & Mid(Work, 5, 2) & "/" & Right(Work, 2)) = False _
     Or StrConv(XX_SIJREC.HIN_GAI, vbUnicode) = Space(13) _
     Or StrConv(XX_SIJREC.IO_KBN, vbUnicode) < "0" _
     Or (StrConv(XX_SIJREC.IO_KBN, vbUnicode) > "3" _
       And StrConv(XX_SIJREC.IO_KBN, vbUnicode) <> "E") _
     Or (StrConv(XX_SIJREC.IO_KBN, vbUnicode) = "0" _
       And StrConv(XX_SIJREC.SYUK_CODE, vbUnicode) = Space(5)) Then
        
        '�`�[���t�^�i�ԁ^���o�ɋ敪�ُ�I�I   �G���[���O
        If Er_Item_NoT Then
            Call Err_Log_Out("�`�[���t�^�i�ԁ^���o�ɋ敪�ُ�")
        End If
        
        Call PDat_Etc_Add("0", "0")                 '�װؽĈ���ް��ۑ�
        Data_Chk = True
        Exit Function
    End If


'�����O�敪�̐ݒ�
    For i = ZERO To UBound(SOKO_T1)
        If SOKO_T1(i).HS_SOKO = "  " Then
            Exit For
        End If
        If RTrim(StrConv(XX_SIJREC.HOST_SOKO, vbUnicode)) = RTrim(SOKO_T1(i).HS_SOKO) Then
            HS_NaiG = SOKO_T1(i).NAIGAI
            Exit For
        End If
    Next i


    If StrConv(XX_SIJREC.IO_KBN, vbUnicode) = "0" Then
'����o�ד`�[���݃`�F�b�N(���ƕ��{�����敪�{�`�[��)
                                                    '���ƕ�
        Call UniCode_Conv(K0_Y_SYU.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
                                                    '�����敪
        If StrConv(XX_SIJREC.CYU_KBN, vbUnicode) = CYU_KBN_TOK Then
            Call UniCode_Conv(K0_Y_SYU.HS_CYU_KBN, CYU_KBN_SPO)
        Else
            Call UniCode_Conv(K0_Y_SYU.HS_CYU_KBN, StrConv(XX_SIJREC.CYU_KBN, vbUnicode))
        End If
                                                    '�`�[��
        Call UniCode_Conv(K0_Y_SYU.DEN_NO, StrConv(XX_SIJREC.DEN_NO, vbUnicode))
                                                    '�r�r�ǔԁi�󔒌Œ�j
        Call UniCode_Conv(K0_Y_SYU.SS_CODE, "")
        sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), ZERO)
        Select Case sts
            Case BtNoErr
                If StrConv(Y_SYUREC.TOK_KBN, vbUnicode) = "1" Then
                    If StrConv(XX_SIJREC.SYUK_CODE, vbUnicode) = StrConv(Y_SYUREC.SYUK_CODE, vbUnicode) And _
                        StrConv(XX_SIJREC.HIN_GAI, vbUnicode) = StrConv(Y_SYUREC.HIN_GAI, vbUnicode) And _
                        CLng(StrConv(XX_SIJREC.YOTEI_QTY, vbUnicode)) = CLng(StrConv(Y_SYUREC.YOTEI_QTY, vbUnicode)) Then
                                '�����莞�A������^�i�ԁ^���ʂ��������Ƃ��͎̂Ă�
                    Else
        '�`�[���t�^�i�ԁ^���o�ɋ敪�ُ�I�I   �G���[���O
                        If Er_Dup_NoT Then
                            Call Err_Log_Out("�d���`�[")
                        End If
                        
                        Call MAKE_SYUDUP_Put                    '�d���ۗ��ް��쐬
                        Call PDat_DUP_Add("0", "2")             '�װؽĈ���ް��ۑ�
                    End If
                Else
        '�`�[���t�^�i�ԁ^���o�ɋ敪�ُ�I�I   �G���[���O
                    If Er_Dup_NoT Then
                        Call Err_Log_Out("�d���`�[")
                    End If
                    
                    Call MAKE_SYUDUP_Put                    '�d���ۗ��ް��쐬
                    Call PDat_DUP_Add("0", "2")             '�װؽĈ���ް��ۑ�
                End If
                Data_Chk = True
                Exit Function
            Case BtErrKeyNotFound
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�o�ח\��")
                Unload Me
        End Select
    
'������R�[�h�Ǒ֏��� 2001.07.04
        Call UniCode_Conv(K0_MTSCHG.RYAKU, StrConv(XX_SIJREC.SYUK_CODE, vbUnicode))
        sts = BTRV(BtOpGetEqual, MTSCHG_POS, MTSCHGREC, Len(MTSCHGREC), K0_MTSCHG, Len(K0_MTSCHG), ZERO)
        Select Case sts
            Case BtNoErr
                MUKECHG = StrConv(MTSCHGREC.MUKE_CODE, vbUnicode)
            Case BtErrKeyNotFound
                MUKECHG = StrConv(XX_SIJREC.SYUK_CODE, vbUnicode)
            Case Else
                Call File_Error(sts, BtOpGetEqual, "������Ǒփ}�X�^")
                Unload Me
        End Select
'������R�[�h���݃`�F�b�N
        Call UniCode_Conv(K0_MTS.MUKE_CODE, MUKECHG)
        sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), ZERO)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Call UniCode_Conv(K0_MTS.MUKE_CODE, ETS_MTS & HS_NaiG)
                sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), ZERO)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                                            '���̑�������Ή����Ȃ�
                        If CHin_Put() Then              '�O���i�ԕύX�ۗ��ް��쐬
                            Unload Me
                        End If
                        
                        '������ُ�I�I   �G���[���O
                        If Er_Muke_NoT Then
                            Call Err_Log_Out("������ُ�")
                        End If
                        
                        
                        Call PDat_Etc_Add("0", "1")     '�װؽĈ���ް��ۑ�
                        Data_Chk = True
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "������Ǘ��}�X�^")
                        Unload Me
                End Select
            Case Else
                Call File_Error(sts, BtOpGetEqual, "������Ǘ��}�X�^")
                Unload Me
        End Select
    
    End If
End Function
                                            '�f�[�^�d���`�F�b�N�p�f�[�^�@�X�V
Private Function Seq_Update() As Boolean

Dim sts     As Integer
Dim Command As Integer
Dim ans     As Integer

    Seq_Update = True
'�d���`�F�b�N�p�f�[�^�X�V�i�ď����f�[�^�i�I�[���u�H�v�u���v�͍X�V���Ȃ��j
    If StrConv(XX_SIJREC.REC_END, vbUnicode) <> "?" And StrConv(XX_SIJREC.REC_END, vbUnicode) <> "*" Then
        Call UniCode_Conv(K0_SEQCK.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
        If Shori_Mode = 4 Then
            Call UniCode_Conv(K0_SEQCK.SEQ_MODE, "2")
        Else
            Call UniCode_Conv(K0_SEQCK.SEQ_MODE, "1")
        End If
        Do
            sts = BTRV(BtOpGetEqual + BtSNoWait, SEQCK_POS, SEQCKREC, Len(SEQCKREC), K0_SEQCK, Len(K0_SEQCK), ZERO)
                        
            Select Case sts
                Case BtNoErr
                    Command = BtOpUpdate
                    Exit Do
                Case BtErrEOF, BtErrKeyNotFound
                    Command = BtOpInsert
                    Call UniCode_Conv(SEQCKREC.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
                    If Shori_Mode = 4 Then
                        Call UniCode_Conv(SEQCKREC.SEQ_MODE, "2")
                    Else
                        Call UniCode_Conv(SEQCKREC.SEQ_MODE, "1")
                    End If
                    Call UniCode_Conv(SEQCKREC.LAST_TXTNO, "000000000")
                    Call UniCode_Conv(SEQCKREC.LAST_GET_DT, "00000000")
                    Call UniCode_Conv(SEQCKREC.LAST_GET_TM, "000000")
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SEQCK.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�\��捞�݃`�F�b�N")
                    Exit Function
            End Select
        Loop
        
        Call UniCode_Conv(SEQCKREC.LAST_TXTNO, StrConv(XX_SIJREC.TEXT_NO, vbUnicode))   '�ŏI�e�L�X�g��
        Call UniCode_Conv(SEQCKREC.LAST_GET_DT, Format(Date, "yyyymmdd"))               '�ŏI�捞�ݓ��t
        Call UniCode_Conv(SEQCKREC.LAST_GET_DT, Format(Time, "hhmm"))                   '�ŏI�捞�ݎ���
        
        Do
            sts = BTRV(Command, SEQCK_POS, SEQCKREC, Len(SEQCKREC), K0_SEQCK, Len(K0_SEQCK), ZERO)
            
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SEQCK.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, Command, "�\��捞�݃`�F�b�N")
                    Exit Function
            End Select
        Loop
    End If

    Seq_Update = False

End Function
                                            '����(0)�C����(1)�f�[�^�X�V
Private Sub Proc_Sel(In_Cnt As Integer)
Dim sts As Integer
Dim ans As Integer
Dim Command As Integer
Dim i As Integer



    Proc_F = ZERO

    DoEvents
'�i�ԃ}�X�^�f�[�^�i�O���j�L���`�F�b�N
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, HS_NaiG)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(XX_SIJREC.HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
    If sts = BtNoErr Then
        Proc_F = Proc_F + 1
        BEF_GAI = StrConv(ITEMREC.HIN_GAI, vbUnicode)
        BEF_NAI = StrConv(ITEMREC.HIN_NAI, vbUnicode)
    Else
        If sts <> BtErrEOF And sts <> BtErrKeyNotFound Then
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Unload Me
        End If
    End If

'�i�ԃ}�X�^�f�[�^�i�����j�L���`�F�b�N
    Call UniCode_Conv(K3_ITEM.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K3_ITEM.NAIGAI, HS_NaiG)
    Call UniCode_Conv(K3_ITEM.HIN_NAI, StrConv(XX_SIJREC.HIN_NAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K3_ITEM, Len(K3_ITEM), 3)
    If sts = BtNoErr Then
        Proc_F = Proc_F + 2
        BEF_GAI = StrConv(ITEMREC.HIN_GAI, vbUnicode)
        BEF_NAI = StrConv(ITEMREC.HIN_NAI, vbUnicode)
    Else
        If sts <> BtErrEOF And sts <> BtErrKeyNotFound Then
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Unload Me
        End If
    End If
                                        '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Unload Me
    End If

'�f�[�^��������
    Select Case Proc_F
        Case ZERO, 3                   '�u�O���������C�����������v�u�O�����L��C�������L��v
            If Upd_Item() Then                                      '�i�ڃ}�X�^�X�V
                GoTo Abort_Tran
            End If
            
            If StrConv(XX_SIJREC.IO_KBN, vbUnicode) = "1" Then
                If NyukaY_Put() Then        '���ח\��o�^
                    GoTo Abort_Tran
                End If
            
                In_Cnt = In_Cnt + 1
                                
                Label3(1).Caption = StrConv(Format(In_Cnt, "#0"), vbWide)
            
            Else
                If SyukaY_Put() Then        '�o�ח\��o�^
                    GoTo Abort_Tran
                End If
                
                In_Cnt = In_Cnt + 1
                                
                Label3(1).Caption = StrConv(Format(In_Cnt, "#0"), vbWide)
            
            End If

        Case 1                      '�u�O�����L��C�����������v�������i�ԕύX
            If Upd_Item() Then              '�i�ڃ}�X�^�X�V
                GoTo Abort_Tran
            End If
            
            If StrConv(XX_SIJREC.IO_KBN, vbUnicode) = "1" Then
                If NyukaY_Put() Then        '���ח\��o�^
                    GoTo Abort_Tran
                End If
                
                In_Cnt = In_Cnt + 1
                                
                Label3(1).Caption = StrConv(Format(In_Cnt, "#0"), vbWide)
            
            Else
                If SyukaY_Put() Then        '�o�ח\��o�^
                    GoTo Abort_Tran
                End If
                
                In_Cnt = In_Cnt + 1
                                
                Label3(1).Caption = StrConv(Format(In_Cnt, "#0"), vbWide)
            
            End If
    
            Call PDat_Hin_Add("0")          '�i�ԕύXؽ��ް��ۑ��i�����i�ԕύX�j

        Case Else                   '�u�O���������C�������L��v���O���i�ԕύX
            sts = Hin_Chg_Chk()            '�O���i�ԕύX�@������
            Select Case sts
                Case False
                    If StrConv(XX_SIJREC.IO_KBN, vbUnicode) = "1" Then      '���o�ח\��o�^
                        If NyukaY_Put() Then    '���ח\��o�^
                            GoTo Abort_Tran
                        End If
                        
                        In_Cnt = In_Cnt + 1
                                
                        Label3(1).Caption = StrConv(Format(In_Cnt, "#0"), vbWide)
                    
                    Else
                        If SyukaY_Put() Then    '�o�ח\��o�^
                            GoTo Abort_Tran
                        End If
                        
                        In_Cnt = In_Cnt + 1
                                
                        Label3(1).Caption = StrConv(Format(In_Cnt, "#0"), vbWide)
                    
                    End If
                
                    Call PDat_Hin_Add("1")      '�i�ԕύXؽ��ް��ۑ��i�O���i�ԕύX�j
                Case True
                                
                    If CHin_Put() Then            '�O���i�ԕύX�ۗ��ް��쐬
                        GoTo Abort_Tran
                    End If
                
                    Call PDat_Hin_Add("2")      '�i�ԕύXؽ��ް��ۑ��i�݌ɗL�I�i�ԕύX�s�j
                Case Else
                    GoTo Abort_Tran
            End Select
    End Select
            
    If Seq_Update() Then            '�ް��d�������p�ް��@�X�V
        GoTo Abort_Tran
    End If
                                        '�g�����U�N�V�����I��
    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If

    If StrConv(XX_SIJREC.PM_KBN, vbUnicode) <> "-" And _
        StrConv(XX_SIJREC.IO_KBN, vbUnicode) = "1" Then
        Call PDat_Etc_Add("1", " ") '�ԓ`�E�o�׊m�F�E��������ް��ۑ�
    End If
    
    Exit Sub

Abort_Tran:
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If

    Unload Me

End Sub
                                            '��װؽģ��m�Fؽģ����f�[�^�ۑ�(�� List Box)
                                            '�@�����@ؽċ敪�F�O���װؽā@�P���ԓ`�A�����A�o�׊m�F
Private Sub PDat_Etc_Add(List_Kbn As String, Err_Kbn As String)
Dim sts As Integer
Dim Work As String

    Work = List_Kbn                                             'ؽċ敪
    Work = Work & StrConv(XX_SIJREC.JGYOBU, vbUnicode)          '���ƕ��敪
    Work = Work & StrConv(XX_SIJREC.HIN_GAI, vbUnicode)         '�i�ԁi�O���j
    Work = Work & StrConv(XX_SIJREC.HIN_NAI, vbUnicode)         '�i�ԁi�����j
    Work = Work & StrConv(XX_SIJREC.DEN_DT, vbUnicode)          '�`�[���t
    Work = Work & StrConv(XX_SIJREC.IO_KBN, vbUnicode)          '���o�ɋ敪
    Work = Work & StrConv(XX_SIJREC.DEN_NO, vbUnicode)          '�`�[��
    Work = Work & StrConv(XX_SIJREC.YOTEI_QTY, vbUnicode)       '����
    Work = Work & StrConv(XX_SIJREC.PM_KBN, vbUnicode)          '�ԍ��敪
    Work = Work & StrConv(XX_SIJREC.HOST_SOKO, vbUnicode)       '�q�ɋ敪�iνāj
    Work = Work & StrConv(XX_SIJREC.CYU_KBN, vbUnicode)         '�����敪
    Work = Work & StrConv(XX_SIJREC.CYOK_KBN, vbUnicode)        '�����敪
    Work = Work & StrConv(XX_SIJREC.YOSAN_FROM, vbUnicode)      '�\�Z�P�ʁi���j
    Work = Work & StrConv(XX_SIJREC.YOSAN_TO, vbUnicode)        '�\�Z�P�ʁi��j
    Work = Work & StrConv(XX_SIJREC.HOST_TANA, vbUnicode)       '�I�ԁiνāj
'   Work = Work & StrConv(XX_SIJREC.SYUK_CODE, vbUnicode)       '�x����^�o�א�
'   Work = Work & StrConv(XX_SIJREC.SYUK_NAME, vbUnicode)       '�x����^�o�א於
    Work = Work & Err_Kbn                                       '�G���[�敪
    
'    Work = Work & StrConv(XX_SIJREC.TEXT_NO, vbUnicode)     '÷�ć�
'    Work = Work & StrConv(XX_SIJREC.CYOK_KBN, vbUnicode)    '�����敪
'    Work = Work & StrConv(XX_SIJREC.PM_KBN, vbUnicode)      '�ԍ��敪
'    Work = Work & StrConv(XX_SIJREC.DEN_SYU, vbUnicode)     '�`�[���
'    Work = Work & StrConv(XX_SIJREC.HIN_NAME, vbUnicode)    '�i��
'    Work = Work & StrConv(XX_SIJREC.YOSAN_FROM, vbUnicode)  '�\�Z�P�ʁi���j
'    Work = Work & StrConv(XX_SIJREC.YOSAN_TO, vbUnicode)    '�\�Z�P�ʁi��j
'    Work = Work & StrConv(XX_SIJREC.HOST_TANA, vbUnicode)   '�I�ԁiνāj

    LBox_Etc.AddItem Work

End Sub

                                            '��װؽģ��m�Fؽģ����f�[�^�ۑ�(�� List Box)
                                            '�@�����@ؽċ敪�F�O���װؽā@�P���ԓ`�A�����A�o�׊m�F
Private Sub PDat_DUP_Add(List_Kbn As String, Err_Kbn As String)
Dim sts As Integer
Dim Work As String

    Work = List_Kbn                                             'ؽċ敪
    Work = Work & StrConv(XX_SIJREC.JGYOBU, vbUnicode)          '���ƕ��敪
    Work = Work & StrConv(XX_SIJREC.HIN_GAI, vbUnicode)         '�i�ԁi�O���j
    Work = Work & StrConv(XX_SIJREC.HIN_NAI, vbUnicode)         '�i�ԁi�����j
    Work = Work & StrConv(XX_SIJREC.DEN_DT, vbUnicode)          '�`�[���t
    Work = Work & StrConv(XX_SIJREC.IO_KBN, vbUnicode)          '���o�ɋ敪
    Work = Work & StrConv(XX_SIJREC.DEN_NO, vbUnicode)          '�`�[��
    Work = Work & StrConv(XX_SIJREC.YOTEI_QTY, vbUnicode)       '����
    Work = Work & StrConv(XX_SIJREC.PM_KBN, vbUnicode)          '�ԍ��敪
    Work = Work & StrConv(XX_SIJREC.HOST_SOKO, vbUnicode)       '�q�ɋ敪�iνāj
    Work = Work & StrConv(XX_SIJREC.CYU_KBN, vbUnicode)         '�����敪
    Work = Work & StrConv(XX_SIJREC.CYOK_KBN, vbUnicode)        '�����敪
    Work = Work & StrConv(XX_SIJREC.YOSAN_FROM, vbUnicode)      '�\�Z�P�ʁi���j
    Work = Work & StrConv(XX_SIJREC.YOSAN_TO, vbUnicode)        '�\�Z�P�ʁi��j
    Work = Work & StrConv(XX_SIJREC.HOST_TANA, vbUnicode)       '�I�ԁiνāj
'   Work = Work & StrConv(XX_SIJREC.SYUK_CODE, vbUnicode)       '�x����^�o�א�
'   Work = Work & StrConv(XX_SIJREC.SYUK_NAME, vbUnicode)       '�x����^�o�א於
    Work = Work & Err_Kbn                                       '�G���[�敪
    
'    Work = Work & StrConv(XX_SIJREC.TEXT_NO, vbUnicode)     '÷�ć�
'    Work = Work & StrConv(XX_SIJREC.CYOK_KBN, vbUnicode)    '�����敪
'    Work = Work & StrConv(XX_SIJREC.PM_KBN, vbUnicode)      '�ԍ��敪
'    Work = Work & StrConv(XX_SIJREC.DEN_SYU, vbUnicode)     '�`�[���
'    Work = Work & StrConv(XX_SIJREC.HIN_NAME, vbUnicode)    '�i��
'    Work = Work & StrConv(XX_SIJREC.YOSAN_FROM, vbUnicode)  '�\�Z�P�ʁi���j
'    Work = Work & StrConv(XX_SIJREC.YOSAN_TO, vbUnicode)    '�\�Z�P�ʁi��j
'    Work = Work & StrConv(XX_SIJREC.HOST_TANA, vbUnicode)   '�I�ԁiνāj

    LBox_Dup.AddItem Work

End Sub

                                            '��i�ԕύXؽģ����f�[�^�ۑ�(�� List Box)
                                            '�@�����@�ύX�敪�F�O�������i�ԕύX�@�P���O���i�ԕύX
Private Sub PDat_Hin_Add(Chg_Kbn As String)
Dim sts As Integer
Dim Work As String

    Work = StrConv(XX_SIJREC.JGYOBU, vbUnicode)                 '���ƕ��敪
    Work = Work & StrConv(XX_SIJREC.TEXT_NO, vbUnicode)         '÷�ć�
    Work = Work & StrConv(XX_SIJREC.HIN_GAI, vbUnicode)         '�i�ԁi�O���j
    Work = Work & BEF_GAI
    Work = Work & StrConv(XX_SIJREC.HIN_NAI, vbUnicode)         '�i�ԁi�����j
    Work = Work & BEF_NAI
    Work = Work & StrConv(XX_SIJREC.DEN_DT, vbUnicode)          '�`�[���t
    Work = Work & StrConv(XX_SIJREC.IO_KBN, vbUnicode)          '���o�ɋ敪
    Work = Work & StrConv(XX_SIJREC.DEN_NO, vbUnicode)          '�`�[��
    Work = Work & StrConv(XX_SIJREC.YOTEI_QTY, vbUnicode)       '����
    Work = Work & StrConv(XX_SIJREC.HOST_SOKO, vbUnicode)       '�q�ɋ敪�iνāj
    Work = Work & StrConv(XX_SIJREC.CYU_KBN, vbUnicode)         '�����敪
    Work = Work & StrConv(XX_SIJREC.SYUK_CODE, vbUnicode)       '�x����^�o�א�
    Work = Work & Chg_Kbn                                       '�ύX�敪
    Work = Work & HS_NaiG                                       '�����O
    
'    Work = Work & StrConv(XX_SIJREC.CYOK_KBN, vbUnicode)    '�����敪
'    Work = Work & StrConv(XX_SIJREC.PM_KBN, vbUnicode)      '�ԍ��敪
'    Work = Work & StrConv(XX_SIJREC.DEN_SYU, vbUnicode)     '�`�[���
'    Work = Work & StrConv(XX_SIJREC.HIN_NAME, vbUnicode)    '�i��
'    Work = Work & StrConv(XX_SIJREC.YOSAN_FROM, vbUnicode)  '�\�Z�P�ʁi���j
'    Work = Work & StrConv(XX_SIJREC.YOSAN_TO, vbUnicode)    '�\�Z�P�ʁi��j
'    Work = Work & StrConv(XX_SIJREC.HOST_TANA, vbUnicode)   '�I�ԁiνāj
'    Work = Work & StrConv(XX_SIJREC.SYUK_NAME, vbUnicode)   '�x����^�o�א於

    LBox_Hin.AddItem Work

End Sub
                                            '�O���i�ԕύX�@�ۃ`�F�b�N
Private Function Hin_Chg_Chk() As Integer

Dim sts         As Integer
Dim Work        As String
Dim ZAIKO_QTY   As Long
Dim ans         As Integer

    Hin_Chg_Chk = False

    Call UniCode_Conv(K3_ITEM.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K3_ITEM.NAIGAI, HS_NaiG)
    Call UniCode_Conv(K3_ITEM.HIN_NAI, StrConv(XX_SIJREC.HIN_NAI, vbUnicode))
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K3_ITEM, Len(K3_ITEM), 3)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Hin_Chg_Chk = SYS_ERR
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
                Hin_Chg_Chk = SYS_ERR
                Exit Function
        End Select
    Loop

'�݌ɗL���`�F�b�N
    If Zaiko_Syukei_Proc(ZAIKO_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), StrConv(ITEMREC.NAIGAI, vbUnicode), StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
        Unload Me
    End If
    If ZAIKO_QTY <> ZERO Then
        Hin_Chg_Chk = True
        Exit Function
    End If

'�i�ڃ}�X�^�@�O���i�ԕύX
    Do
    
        sts = BTRV(BtOpDelete, ITEM_POS, ITEMREC, Len(ITEMREC), K3_ITEM, Len(K3_ITEM), 3)
        
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Hin_Chg_Chk = SYS_ERR
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpDelete, "�i�ڃ}�X�^")
                Hin_Chg_Chk = SYS_ERR
                Exit Function
        End Select
    
    Loop

    Call UniCode_Conv(ITEMREC.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(ITEMREC.HIN_GAI, StrConv(XX_SIJREC.HIN_GAI, vbUnicode))
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, HS_NaiG)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(XX_SIJREC.HIN_GAI, vbUnicode))
    
    Do
        sts = BTRV(BtOpInsert, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
        Select Case sts
        
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Hin_Chg_Chk = SYS_ERR
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpInsert, "�i�ڃ}�X�^")
                Hin_Chg_Chk = SYS_ERR
                Exit Function
        End Select
    
    Loop
End Function
                                            '�i�ڃ}�X�^�X�V
Private Function Upd_Item() As Boolean
Dim sts As Integer
Dim ans As Integer
Dim Command As Integer
Dim Work As String

    Upd_Item = True

    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, HS_NaiG)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(XX_SIJREC.HIN_GAI, vbUnicode))
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
        Select Case sts
            Case BtNoErr
                Command = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                Command = BtOpInsert
                Call UniCode_Conv(ITEMREC.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(ITEMREC.NAIGAI, HS_NaiG)
                Call UniCode_Conv(ITEMREC.HIN_GAI, StrConv(XX_SIJREC.HIN_GAI, vbUnicode))
                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                Call UniCode_Conv(ITEMREC.ST_SET_DT, "")
                Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                Call UniCode_Conv(ITEMREC.ST_RETU, "")
                Call UniCode_Conv(ITEMREC.ST_REN, "")
                Call UniCode_Conv(ITEMREC.ST_DAN, "")
                Call UniCode_Conv(ITEMREC.BEF_SOKO, "")
                Call UniCode_Conv(ITEMREC.BEF_RETU, "")
                Call UniCode_Conv(ITEMREC.BEF_REN, "")
                Call UniCode_Conv(ITEMREC.BEF_DAN, "")
                Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")
                Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")
                Call UniCode_Conv(ITEMREC.BIKOU_SOKO, "")
                Call UniCode_Conv(ITEMREC.BIKOU_TANA, "")
                Call UniCode_Conv(ITEMREC.SIZAI_CD, "")
                Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")
                Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")
                Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "0")
                Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")
                
                Call UniCode_Conv(ITEMREC.LOCK_F, "0")          '�r���t���O
                Call UniCode_Conv(ITEMREC.WEL_ID, "")           '�g�p���q�@�h�c
                Call UniCode_Conv(ITEMREC.PRG_ID, "")           '�g�p���v���O����
                
                Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")
                Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "0000000")
                
                Call UniCode_Conv(ITEMREC.FILLER, "")
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
                Exit Function
        End Select
    Loop
                                '�i�ԁi�����j(���󔒂̎��̂݃Z�b�g)
    If StrConv(XX_SIJREC.HIN_NAI, vbUnicode) <> Space(13) Then
        Call UniCode_Conv(ITEMREC.HIN_NAI, StrConv(XX_SIJREC.HIN_NAI, vbUnicode))
    End If
                                
Debug.Print StrConv(ITEMREC.HIN_NAI, vbUnicode)
                                '�i���i���󔒂̎��̂݃Z�b�g�j
    If StrConv(XX_SIJREC.HIN_NAME, vbUnicode) <> Space(25) Then
        Call UniCode_Conv(ITEMREC.HIN_NAME, StrConv(XX_SIJREC.HIN_NAME, vbUnicode))
    End If
                                '���l�Fνđq�ɋ敪
                                '�@�@�m�q�ɋ敪�̓Ǒւ��@�ݒ����]
                                '�@�@�@�@��M�f�[�^�̑q�ɋ敪����
                                '�@�@�@�A�@�@�V�@�@�@�q�ɋ敪���i��Ͻ��̑q�ɋ敪
    If StrConv(XX_SIJREC.HOST_SOKO, vbUnicode) <> Space(2) And _
       StrConv(XX_SIJREC.HOST_SOKO, vbUnicode) > StrConv(ITEMREC.BIKOU_SOKO, vbUnicode) Then
        Call UniCode_Conv(ITEMREC.BIKOU_SOKO, StrConv(XX_SIJREC.HOST_SOKO, vbUnicode))
    End If

                                '���l�FνĒI�ԁi���󔒂̎��j
    If StrConv(XX_SIJREC.HOST_TANA, vbUnicode) <> Space(8) Then
        Call UniCode_Conv(ITEMREC.BIKOU_TANA, StrConv(XX_SIJREC.HOST_TANA, vbUnicode))
    End If

    
    Do
        sts = BTRV(Command, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
        Select Case sts
            Case BtNoErr, BtErrEOF, BtErrKeyNotFound
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, Command, "�i�ڃ}�X�^")
                Exit Function
        End Select
    Loop

    Upd_Item = False

End Function
                                            '���ח\��쐬 �� ���׍X�V
Private Function NyukaY_Put() As Boolean
Dim sts As Integer
Dim Work As String
Dim com As Integer
Dim W_Qty As Long
Dim W_Y_Qty As Long         '���ח\�萔
Dim W_E_Qty As Long         '�O�؂���א�
Dim W_Date As String        '�������t

Dim ans     As Integer

    NyukaY_Put = True
    Call UniCode_Conv(K4_Y_NYU.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K4_Y_NYU.TEXT_NO, StrConv(XX_SIJREC.TEXT_NO, vbUnicode))
    sts = BTRV(BtOpGetEqual, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K4_Y_NYU, Len(K4_Y_NYU), 4)
    Select Case sts
        Case BtNoErr
            Work = "Y_NYUKA.DAT DUP ""���ƕ�=" & StrConv(XX_SIJREC.JGYOBU, vbUnicode) & "TEXTNo=" & StrConv(XX_SIJREC.TEXT_NO, vbUnicode)
            Work = "�`�[���t=" & StrConv(XX_SIJREC.DEN_DT, vbUnicode) & "�`�[��=" & StrConv(XX_SIJREC.DEN_NO, vbUnicode)
            Work = "�i��=" & StrConv(XX_SIJREC.HIN_GAI, vbUnicode)
            Call Log_Out(LOG_F, Work)
            NyukaY_Put = False
            Exit Function
        Case BtErrKeyNotFound
        Case Else
            Call File_Error(sts, BtOpGetEqual, "���ח\��")
            Exit Function
    End Select
'////////////////////////////////////////////////////////�@���׃f�[�^�r�����W�b�N
'
'
'                 �i�@���@�݁@�\�@��@�n�@�j
'1997.08.22
    If RTrim(StrConv(XX_SIJREC.YOSAN_FROM, vbUnicode)) = "555" And _
        RTrim(StrConv(XX_SIJREC.HOST_SOKO, vbUnicode)) = "0" Then
        Call UniCode_Conv(XX_SIJREC.CYOK_KBN, "*")          '�r���f�[�^�́u�����敪�v�Ɂu���v���Z�b�g
    End If

    If RTrim(StrConv(XX_SIJREC.YOSAN_FROM, vbUnicode)) = "111" And _
        RTrim(StrConv(XX_SIJREC.HOST_SOKO, vbUnicode)) = "5" Then
        Call UniCode_Conv(XX_SIJREC.CYOK_KBN, "*")          '�r���f�[�^�́u�����敪�v�Ɂu���v���Z�b�g
    End If

    If RTrim(StrConv(XX_SIJREC.YOSAN_FROM, vbUnicode)) = "555" And _
        RTrim(StrConv(XX_SIJREC.HOST_SOKO, vbUnicode)) = "1" Then
        Call UniCode_Conv(XX_SIJREC.CYOK_KBN, "*")          '�r���f�[�^�́u�����敪�v�Ɂu���v���Z�b�g
    End If

'
'
'        Call UniCode_Conv(XX_SIJREC.CYOK_KBN, "*")          '�r���f�[�^�́u�����敪�v�Ɂu���v���Z�b�g
'
'
'////////////////////////////////////////////////////////


    W_Date = Format(Date, "yyyymmdd")

'���ח\��쐬
                                '�����敪
    Call UniCode_Conv(Y_NYUREC.KAN_KBN, KAN_KBN_ON)
                                '�f�[�^���
    Call UniCode_Conv(Y_NYUREC.DT_SYU, "0")
                                '�\�萔��
    Call UniCode_Conv(Y_NYUREC.YOTEI_QTY, Format(CLng(StrConv(XX_SIJREC.YOTEI_QTY, vbUnicode)), "00000000"))
                                '�m�萔��
    Call UniCode_Conv(Y_NYUREC.FIX_QTY, "00000000")
                                '�����O
    Call UniCode_Conv(Y_NYUREC.NAIGAI, HS_NaiG)
                                '���ƕ��敪
    Call UniCode_Conv(Y_NYUREC.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
                                '÷�ć�
    If StrConv(XX_SIJREC.CYOK_KBN, vbUnicode) = "*" Then        '�r���f�[�^�H
                                '�����敪
        Call UniCode_Conv(Y_NYUREC.CYOK_KBN, "C")
    Else
                                '�����敪
        Call UniCode_Conv(Y_NYUREC.CYOK_KBN, StrConv(XX_SIJREC.CYOK_KBN, vbUnicode))
    End If
    
    Call UniCode_Conv(Y_NYUREC.TEXT_NO, StrConv(XX_SIJREC.TEXT_NO, vbUnicode))
                                '�`�[���t
    Call UniCode_Conv(Y_NYUREC.DEN_DT, StrConv(XX_SIJREC.DEN_DT, vbUnicode))
                                '���o�ɋ敪
    Call UniCode_Conv(Y_NYUREC.IO_KBN, StrConv(XX_SIJREC.IO_KBN, vbUnicode))
                                '�ԍ��敪
    Call UniCode_Conv(Y_NYUREC.PM_KBN, StrConv(XX_SIJREC.PM_KBN, vbUnicode))
                                '�`�[���
    
    Call UniCode_Conv(Y_NYUREC.DEN_SYU, StrConv(XX_SIJREC.DEN_SYU, vbUnicode))
                                '�`�[��
    Call UniCode_Conv(Y_NYUREC.DEN_NO, StrConv(XX_SIJREC.DEN_NO, vbUnicode))
                                '�����敪
    Call UniCode_Conv(Y_NYUREC.CYU_KBN, StrConv(XX_SIJREC.CYU_KBN, vbUnicode))
                                '�i�ԁi�O���j
    Call UniCode_Conv(Y_NYUREC.HIN_GAI, StrConv(XX_SIJREC.HIN_GAI, vbUnicode))
                                '�i�ԁi�����j
    Call UniCode_Conv(Y_NYUREC.HIN_NAI, StrConv(XX_SIJREC.HIN_NAI, vbUnicode))
                                '�i��
    Call UniCode_Conv(Y_NYUREC.HIN_NAME, StrConv(XX_SIJREC.HIN_NAME, vbUnicode))
                                '�\�Z�P�ʁi���j
    Call UniCode_Conv(Y_NYUREC.YOSAN_FROM, StrConv(XX_SIJREC.YOSAN_FROM, vbUnicode))
                                '�\�Z�P�ʁi��j
    Call UniCode_Conv(Y_NYUREC.YOSAN_TO, StrConv(XX_SIJREC.YOSAN_TO, vbUnicode))
                                '�q�ɋ敪�iνāj
    Call UniCode_Conv(Y_NYUREC.HOST_SOKO, StrConv(XX_SIJREC.HOST_SOKO, vbUnicode))
                                '�I�ԁiνāj���@�W�����ɒI�ԁi�i��Ͻ��j
    Work = StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
           StrConv(ITEMREC.ST_RETU, vbUnicode) & _
           StrConv(ITEMREC.ST_REN, vbUnicode) & _
           StrConv(ITEMREC.ST_DAN, vbUnicode)
    Call UniCode_Conv(Y_NYUREC.HOST_TANA, Work)
                                '�x����^�o�א�
    Call UniCode_Conv(Y_NYUREC.SYUK_CODE, StrConv(XX_SIJREC.SYUK_CODE, vbUnicode))
                                '�x����^�o�א於
    Call UniCode_Conv(Y_NYUREC.SYUK_NAME, StrConv(XX_SIJREC.SYUK_NAME, vbUnicode))
                                '��s���א�
    Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, "00000000")
                                '�������t
    Call UniCode_Conv(Y_NYUREC.KAN_DT, W_Date)
                                'FILLER
    Call UniCode_Conv(Y_NYUREC.FILLER, "")


'�r�����̓��׃f�[�^�́A���ח\��o�^�̂�
    If StrConv(XX_SIJREC.CYOK_KBN, vbUnicode) = "*" Then        '�r���Ώۃf�[�^�͓��׍X�V����
        Do
            sts = BTRV(BtOpInsert, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), ZERO)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "���ח\��")
                    Exit Function
            End Select
        Loop
'�����r��������I��
        NyukaY_Put = False
        Exit Function
    End If
    
    If Shori_Mode = ZERO Then       '�Ď�荞�ݎw�����́A�������Ȃ� 01.05.03 **
        NyukaY_Put = False
        Exit Function
    End If

    W_Y_Qty = CLng(StrConv(XX_SIJREC.YOTEI_QTY, vbUnicode))
    Last_Proc_F = True              '���������ް��폜�����@���s�L��

'���������ް��X�V
    Call UniCode_Conv(K0_J_NYU.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_J_NYU.NAIGAI, HS_NaiG)
    Call UniCode_Conv(K0_J_NYU.HIN_GAI, StrConv(XX_SIJREC.HIN_GAI, vbUnicode))

    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), ZERO)
        Select Case sts
            Case BtNoErr
                If CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode)) > W_Y_Qty Then
                    W_Qty = CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode)) - W_Y_Qty
                    Call UniCode_Conv(J_NYUREC.JITU_QTY, Format(W_Qty, "00000000"))
                    
                    Do
                        sts = BTRV(BtOpUpdate, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), ZERO)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Exit Function
                                End If
                            Case Else
                                Call File_Error(sts, BtOpUpdate, "���������ް�")
                                Exit Function
                        End Select
                    
                    Loop
                    W_E_Qty = CLng(StrConv(XX_SIJREC.YOTEI_QTY, vbUnicode))
                Else
                    Do
                        sts = BTRV(BtOpDelete, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), ZERO)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Exit Function
                                End If
                            Case Else
                                Call File_Error(sts, BtOpDelete, "���������ް�")
                                Exit Function
                        End Select
                    Loop
                    W_E_Qty = CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode))
                End If
                
                Exit Do
            Case BtErrKeyNotFound
                W_E_Qty = ZERO
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "���������ް�")
                Exit Function
        End Select
    Loop

'���ח\��f�[�^�ǉ��i���ו��j
                                '��s���א��i���׎��ѐ��j
    Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, Format(W_E_Qty, "00000000"))
    Do
        sts = BTRV(BtOpInsert, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), ZERO)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpInsert, "���ח\��")
                Exit Function
        End Select
    Loop

'���א��ō݌Ƀf�[�^�X�V�i�{�j
    If Nyuko_Update_Proc(StrConv(XX_SIJREC.JGYOBU, vbUnicode), _
                            HS_NaiG, _
                            StrConv(XX_SIJREC.HIN_GAI, vbUnicode), _
                            StrConv(XX_SIJREC.DEN_DT, vbUnicode), _
                            (KASO_NYUKA_Soko & "01" & "01" & "01"), _
                            YOIN_TU_NYUKA, _
                            CLng(StrConv(XX_SIJREC.YOTEI_QTY, vbUnicode)), _
                            WS_NO) Then
        Exit Function
    
    End If


'�O�؂萔�ō݌Ƀf�[�^�X�V�i�|�j
    If W_E_Qty <> ZERO Then
'�݌Ƀf�[�^LOCK
        If Zaiko_Lock_Proc((KASO_NYUKA_Soko & "01" & "01" & "01"), _
                            StrConv(XX_SIJREC.JGYOBU, vbUnicode), _
                            HS_NaiG, _
                            StrConv(XX_SIJREC.HIN_GAI, vbUnicode), _
                            WS_NO) Then
            Exit Function

        End If
        
        
        If Syuko_Update_Proc(StrConv(XX_SIJREC.JGYOBU, vbUnicode), _
                                HS_NaiG, _
                                StrConv(XX_SIJREC.HIN_GAI, vbUnicode), _
                                StrConv(XX_SIJREC.DEN_DT, vbUnicode), _
                                (KASO_NYUKA_Soko & "01" & "01" & "01"), _
                                YOIN_MAE_SOUSAI, _
                                W_E_Qty, _
                                WS_NO) Then
            Exit Function
    
        End If

'�݌Ƀf�[�^UNLOCK
        If Zaiko_UNLock_Proc((KASO_NYUKA_Soko & "01" & "01" & "01"), _
                                StrConv(XX_SIJREC.JGYOBU, vbUnicode), _
                                HS_NaiG, _
                                StrConv(XX_SIJREC.HIN_GAI, vbUnicode), _
                            "") Then
            Exit Function
        End If
    End If
    
    NyukaY_Put = False

End Function
                                            '�o�ח\��쐬
Private Function SyukaY_Put() As Boolean
Dim sts     As Integer
Dim Work    As String
Dim Command As Integer
                     
Dim ans     As Integer
    
    SyukaY_Put = True
                                '�����敪
    Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_SOFF_POFF_KOFF)
                                '�f�[�^���
    Call UniCode_Conv(Y_SYUREC.DT_SYU, "0")
                                '�\�萔��
    Call UniCode_Conv(Y_SYUREC.YOTEI_QTY, Format(CLng(StrConv(XX_SIJREC.YOTEI_QTY, vbUnicode)), "00000000"))
                                '�m�萔��
    Call UniCode_Conv(Y_SYUREC.FIX_QTY, "00000000")
                                '�����O
    Call UniCode_Conv(Y_SYUREC.NAIGAI, HS_NaiG)
                                '���ƕ��敪
    Call UniCode_Conv(Y_SYUREC.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
                                '÷�ć�
    Call UniCode_Conv(Y_SYUREC.TEXT_NO, StrConv(XX_SIJREC.TEXT_NO, vbUnicode))
                                '�����敪
    Call UniCode_Conv(Y_SYUREC.CYOK_KBN, StrConv(XX_SIJREC.CYOK_KBN, vbUnicode))
                                '�`�[���t
    Call UniCode_Conv(Y_SYUREC.DEN_DT, StrConv(XX_SIJREC.DEN_DT, vbUnicode))
                                '���o�ɋ敪
    Call UniCode_Conv(Y_SYUREC.IO_KBN, StrConv(XX_SIJREC.IO_KBN, vbUnicode))
                                '�ԍ��敪
    Call UniCode_Conv(Y_SYUREC.PM_KBN, StrConv(XX_SIJREC.PM_KBN, vbUnicode))
                                '�`�[���
    Call UniCode_Conv(Y_SYUREC.DEN_SYU, StrConv(XX_SIJREC.DEN_SYU, vbUnicode))
                                '�`�[��
    Call UniCode_Conv(Y_SYUREC.DEN_NO, StrConv(XX_SIJREC.DEN_NO, vbUnicode))
                                '�����敪�i��[�ƽ�߯Ắu��[�E��߯āv�敪�ɒu�����j
                                            '��������u��[�E��߯āv�敪�ɒu����
    If StrConv(XX_SIJREC.CYU_KBN, vbUnicode) = CYU_KBN_SPO Or _
       StrConv(XX_SIJREC.CYU_KBN, vbUnicode) = CYU_KBN_HJU Or _
       StrConv(XX_SIJREC.CYU_KBN, vbUnicode) = CYU_KBN_TOK Then
        Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN_HSP)
    Else
        Call UniCode_Conv(Y_SYUREC.CYU_KBN, StrConv(XX_SIJREC.CYU_KBN, vbUnicode))
    End If
                                '�i�ԁi�O���j
    Call UniCode_Conv(Y_SYUREC.HIN_GAI, StrConv(XX_SIJREC.HIN_GAI, vbUnicode))
                                '�i�ԁi�����j
    Call UniCode_Conv(Y_SYUREC.HIN_NAI, StrConv(XX_SIJREC.HIN_NAI, vbUnicode))
                                '�i��
    Call UniCode_Conv(Y_SYUREC.HIN_NAME, StrConv(XX_SIJREC.HIN_NAME, vbUnicode))
                                '�\�Z�P�ʁi���j
    Call UniCode_Conv(Y_SYUREC.YOSAN_FROM, StrConv(XX_SIJREC.YOSAN_FROM, vbUnicode))
                                '�\�Z�P�ʁi��j
    Call UniCode_Conv(Y_SYUREC.YOSAN_TO, StrConv(XX_SIJREC.YOSAN_TO, vbUnicode))
                                '�q�ɋ敪�iνāj
    Call UniCode_Conv(Y_SYUREC.HOST_SOKO, StrConv(XX_SIJREC.HOST_SOKO, vbUnicode))
                                '�I�ԁiνāj���@�W�����ɒI�ԁi�i��Ͻ��j
    Work = StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
           StrConv(ITEMREC.ST_RETU, vbUnicode) & _
           StrConv(ITEMREC.ST_REN, vbUnicode) & _
           StrConv(ITEMREC.ST_DAN, vbUnicode)
    Call UniCode_Conv(Y_SYUREC.HOST_TANA, Work)
                                '�x����^�o�א�
    Call UniCode_Conv(Y_SYUREC.SYUK_CODE, StrConv(XX_SIJREC.SYUK_CODE, vbUnicode))
                                '�x����^�o�א於
    Call UniCode_Conv(Y_SYUREC.SYUK_NAME, StrConv(XX_SIJREC.SYUK_NAME, vbUnicode))
                                '�������t
    Call UniCode_Conv(Y_SYUREC.KAN_DT, "")
                                '�����敪�iνāj
    If StrConv(XX_SIJREC.CYU_KBN, vbUnicode) = CYU_KBN_TOK Then
                                '������̓X�|�b�g�ɒu��������2001.06.21
        Call UniCode_Conv(Y_SYUREC.HS_CYU_KBN, CYU_KBN_SPO)
    Else
        Call UniCode_Conv(Y_SYUREC.HS_CYU_KBN, StrConv(XX_SIJREC.CYU_KBN, vbUnicode))
    End If
                                '�r�r�ǔ�
    Call UniCode_Conv(Y_SYUREC.SS_CODE, "")
                                '�g�p�q�@�h�c
    Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
                                '�g�p���v���O����
    Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
                                '���i���t
    Call UniCode_Conv(Y_SYUREC.KENPIN_DT, "")
                                
                                
                                '�����溰��
    Call UniCode_Conv(Y_SYUREC.MUKE_CODE, StrConv(MTSREC.MUKE_CODE, vbUnicode))
                                '������Ǒւ�����
    Call UniCode_Conv(Y_SYUREC.MUKE_CHG_CD, StrConv(MTSREC.MUKE_CHG_CD, vbUnicode))
                                '������}�[�N   2001.06.21
    If StrConv(XX_SIJREC.CYU_KBN, vbUnicode) = CYU_KBN_TOK Then
        Call UniCode_Conv(Y_SYUREC.TOK_KBN, "1")
    Else
        Call UniCode_Conv(Y_SYUREC.TOK_KBN, " ")
    End If
                                'FILLER
    Call UniCode_Conv(Y_SYUREC.FILLER, "")
    
    Do
        sts = BTRV(BtOpInsert, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), ZERO)
        
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpInsert, "�o�ח\��")
                Exit Function
        End Select
    Loop
    
    If SYUKA_LOG_ON Then
        Call SYUKA_LOG_OUT_PROC("INS", "AFT")
    End If
        
    SyukaY_Put = False

End Function
                                            '�O���i�ԕύX�ۗ��f�[�^�쐬
Private Function CHin_Put() As Integer
           
    CHin_Put = True
                                '÷�ć�
    Call UniCode_Conv(CHGHREC.TEXT_NO, StrConv(XX_SIJREC.TEXT_NO, vbUnicode))
                                '���ƕ��敪
    Call UniCode_Conv(CHGHREC.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
                                '�����敪
    Call UniCode_Conv(CHGHREC.CYOK_KBN, StrConv(XX_SIJREC.CYOK_KBN, vbUnicode))
                                '�`�[���t
    Call UniCode_Conv(CHGHREC.DEN_DT, StrConv(XX_SIJREC.DEN_DT, vbUnicode))
                                '���o�ɋ敪
    Call UniCode_Conv(CHGHREC.IO_KBN, StrConv(XX_SIJREC.IO_KBN, vbUnicode))
                                '�ԍ��敪
    Call UniCode_Conv(CHGHREC.PM_KBN, StrConv(XX_SIJREC.PM_KBN, vbUnicode))
                                '�`�[���
    Call UniCode_Conv(CHGHREC.DEN_SYU, StrConv(XX_SIJREC.DEN_SYU, vbUnicode))
                                '�`�[��
    Call UniCode_Conv(CHGHREC.DEN_NO, StrConv(XX_SIJREC.DEN_NO, vbUnicode))
                                '�����敪
    Call UniCode_Conv(CHGHREC.CYU_KBN, StrConv(XX_SIJREC.CYU_KBN, vbUnicode))
                                '�i�ԁi�O���j
    Call UniCode_Conv(CHGHREC.HIN_GAI, StrConv(XX_SIJREC.HIN_GAI, vbUnicode))
                                '�i�ԁi�����j
    Call UniCode_Conv(CHGHREC.HIN_NAI, StrConv(XX_SIJREC.HIN_NAI, vbUnicode))
                                '�i��
    Call UniCode_Conv(CHGHREC.HIN_NAME, StrConv(XX_SIJREC.HIN_NAME, vbUnicode))
                                '����
    Call UniCode_Conv(CHGHREC.YOTEI_QTY, StrConv(XX_SIJREC.YOTEI_QTY, vbUnicode))
                                '�\�Z�P�ʁi���j
    Call UniCode_Conv(CHGHREC.YOSAN_FROM, StrConv(XX_SIJREC.YOSAN_FROM, vbUnicode))
                                '�\�Z�P�ʁi��j
    Call UniCode_Conv(CHGHREC.YOSAN_TO, StrConv(XX_SIJREC.YOSAN_TO, vbUnicode))
                                '�q�ɋ敪�iνāj
    Call UniCode_Conv(CHGHREC.HOST_SOKO, StrConv(XX_SIJREC.HOST_SOKO, vbUnicode))
                                '�I�ԁiνāj
    Call UniCode_Conv(CHGHREC.HOST_TANA, StrConv(XX_SIJREC.HOST_TANA, vbUnicode))
                                '�x����^�o�א�
    Call UniCode_Conv(CHGHREC.SYUK_CODE, StrConv(XX_SIJREC.SYUK_CODE, vbUnicode))
                                '�x����^�o�א於
    Call UniCode_Conv(CHGHREC.SYUK_NAME, StrConv(XX_SIJREC.SYUK_NAME, vbUnicode))
                                'ں��ޏI�[ϰ�(@)
                                '�@�ď������̏d����������ׁ̈A�i�ԕύX�f�[�^�̏I�[�}�[�N�ɂ�
                                '�@�u�H�v���Z�b�g����i�d�������ł́u�I�[���H�v�̎��������Ȃ��j
    Call UniCode_Conv(CHGHREC.REC_END, "?")
                                'CR & LF
    Call UniCode_Conv(CHGHREC.CR, Chr(13))
    Call UniCode_Conv(CHGHREC.LF, Chr(10))

    If CHGH_Put(ZERO) Then         '�O���i�ԕύX�ۗ��f�[�^������
        Exit Function
    End If

    CHin_Put = False

End Function
                                            '�d���o�ח\��f�[�^�쐬
Private Sub MAKE_SYUDUP_Put()
                                '÷�ć�
    Call UniCode_Conv(SYUDUPREC.TEXT_NO, StrConv(XX_SIJREC.TEXT_NO, vbUnicode))
                                '���ƕ��敪
    Call UniCode_Conv(SYUDUPREC.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
                                '�����敪
    Call UniCode_Conv(SYUDUPREC.CYOK_KBN, StrConv(XX_SIJREC.CYOK_KBN, vbUnicode))
                                '�`�[���t
    Call UniCode_Conv(SYUDUPREC.DEN_DT, StrConv(XX_SIJREC.DEN_DT, vbUnicode))
                                '���o�ɋ敪
    Call UniCode_Conv(SYUDUPREC.IO_KBN, StrConv(XX_SIJREC.IO_KBN, vbUnicode))
                                '�ԍ��敪
    Call UniCode_Conv(SYUDUPREC.PM_KBN, StrConv(XX_SIJREC.PM_KBN, vbUnicode))
                                '�`�[���
    Call UniCode_Conv(SYUDUPREC.DEN_SYU, StrConv(XX_SIJREC.DEN_SYU, vbUnicode))
                                '�`�[��
    Call UniCode_Conv(SYUDUPREC.DEN_NO, StrConv(XX_SIJREC.DEN_NO, vbUnicode))
                                '�����敪
    Call UniCode_Conv(SYUDUPREC.CYU_KBN, StrConv(XX_SIJREC.CYU_KBN, vbUnicode))
                                '�i�ԁi�O���j
    Call UniCode_Conv(SYUDUPREC.HIN_GAI, StrConv(XX_SIJREC.HIN_GAI, vbUnicode))
                                '�i�ԁi�����j
    Call UniCode_Conv(SYUDUPREC.HIN_NAI, StrConv(XX_SIJREC.HIN_NAI, vbUnicode))
                                '�i��
    Call UniCode_Conv(SYUDUPREC.HIN_NAME, StrConv(XX_SIJREC.HIN_NAME, vbUnicode))
                                '����
    Call UniCode_Conv(SYUDUPREC.YOTEI_QTY, StrConv(XX_SIJREC.YOTEI_QTY, vbUnicode))
                                '�\�Z�P�ʁi���j
    Call UniCode_Conv(SYUDUPREC.YOSAN_FROM, StrConv(XX_SIJREC.YOSAN_FROM, vbUnicode))
                                '�\�Z�P�ʁi��j
    Call UniCode_Conv(SYUDUPREC.YOSAN_TO, StrConv(XX_SIJREC.YOSAN_TO, vbUnicode))
                                '�q�ɋ敪�iνāj
    Call UniCode_Conv(SYUDUPREC.HOST_SOKO, StrConv(XX_SIJREC.HOST_SOKO, vbUnicode))
                                '�I�ԁiνāj
    Call UniCode_Conv(SYUDUPREC.HOST_TANA, StrConv(XX_SIJREC.HOST_TANA, vbUnicode))
                                '�x����^�o�א�
    Call UniCode_Conv(SYUDUPREC.SYUK_CODE, StrConv(XX_SIJREC.SYUK_CODE, vbUnicode))
                                '�x����^�o�א於
    Call UniCode_Conv(SYUDUPREC.SYUK_NAME, StrConv(XX_SIJREC.SYUK_NAME, vbUnicode))
                                'ں��ޏI�[ϰ�(@)
                                '�@�ď������̏d����������ׁ̈A�i�ԕύX�f�[�^�̏I�[�}�[�N�ɂ�
                                '�@�u*�v���Z�b�g����i�d�������ł́u�I�[��*�v�̎��������Ȃ��j
    Call UniCode_Conv(SYUDUPREC.REC_END, "*")
                                'CR & LF
    Call UniCode_Conv(SYUDUPREC.CR, Chr(13))
    Call UniCode_Conv(SYUDUPREC.LF, Chr(10))

    If SYUDUP_Put(ZERO) Then       '�d���o�ח\��f�[�^������
        Unload Me
    End If
End Sub

                                            '�w�b�_�[����i�u�װؽāv�u�ԓ`�A�����A�o�׊m�F�ꗗ�\�v�j
Private Sub P_Etc_Head(Lst_Kbn As Integer, Lcnt As Integer, JGYO_Kbn As String)
Dim i As Integer
Dim sts As Integer
Dim RetBuf As String

    If Lcnt <> 99 Then
        Printer.NewPage
    End If

    Printer.Print


    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print
                                        '�w�b�_�[�i�P�j
    Printer.Print Tab(3);
    For i = ZERO To UBound(JGYOBU_T)
        If JGYO_Kbn = JGYOBU_T(i).Code Then
            Printer.Print "�i" & RTrim(JGYOBU_T(i).NAME) & "�j";
            Exit For
        End If
    Next i
    Printer.Print Tab(37);
    If Lst_Kbn = ZERO Then
        Printer.Print "�������@��M�G���[���X�g�@������";
    Else
        Printer.Print "���� �ԓ`�A�o�ɁA�����f�[�^�m�F���X�g�@����";
    End If
    Printer.Print Tab(91);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")

    Printer.Print
                                        '���׃w�b�_���
    Printer.Print Tab(MGN_L);
    Printer.Print "�i�ԁi�O���j";
    Printer.Print Tab(MGN_L + 14);
    Printer.Print "�i�ԁi�����j";
    Printer.Print Tab(MGN_L + 28);
    Printer.Print "�`�[���t";
    Printer.Print Tab(MGN_L + 39);
    Printer.Print "���o�ɋ敪";
    Printer.Print Tab(MGN_L + 50);
    Printer.Print "�`�[��";
    Printer.Print Tab(MGN_L + 58);
    Printer.Print "���ɐ�";
    Printer.Print Tab(MGN_L + 68);
    Printer.Print "�o�ɐ�";
    Printer.Print Tab(MGN_L + 77);
    Printer.Print "�q��";
    Printer.Print Tab(MGN_L + 82);
    Printer.Print "�敪";
    Printer.Print Tab(MGN_L + 92);
    Printer.Print "����";
    Printer.Print Tab(MGN_L + 98);
    Printer.Print "�\�Z�P�ʁ@�@ �W���I��"
    Printer.Print

    Lcnt = 6 + MGN_U

End Sub
                                            '���׈���i�u�װؽāv�u�ԓ`�A�����A�o�׊m�F�ꗗ�\�v�j
Private Sub P_Etc_Proc(Lst_Kbn As Integer)

Dim Lcnt As Integer
Dim Ldata As String
Dim wk_IO As String
Dim Work As String
Dim i As Integer
Dim sts As Integer
Dim B_Jgyobu As String

    Lcnt = 99
    B_Jgyobu = Space(1)

    For i = ZERO To LBox_Etc.ListCount - 1
        If Left(LBox_Etc.List(i), 1) = Lst_Kbn Then

            Ldata = LBox_Etc.List(i)

                                        '�w�b�_�[�R���g���[��
            If Lcnt > LMAX Or _
               Mid(Ldata, 2, 1) <> B_Jgyobu Then
                Call P_Etc_Head(Left(Ldata, 1), Lcnt, Mid(Ldata, 2, 1))
                B_Jgyobu = Mid(Ldata, 2, 1)
            End If
                                        '���׈��
            Ldata = Mid(Ldata, 3, Len(Ldata) - 2)

            Printer.Print Tab(MGN_L);
            Printer.Print ChrCut(Ldata, 13);            '�i�ԁi�O���j

            Printer.Print Tab(MGN_L + 14);
            Printer.Print ChrCut(Ldata, 13);            '�i�ԁi�����j

            Printer.Print Tab(MGN_L + 28);              '�`�[���t
            Printer.Print ChrCut(Ldata, 4) & "/" & ChrCut(Ldata, 2) & "/" & ChrCut(Ldata, 2);

            Printer.Print Tab(MGN_L + 40);              '���o�ɋ敪
            wk_IO = ChrCut(Ldata, 1)
            Select Case wk_IO
                Case IO_KBN_URI
                    Printer.Print wk_IO & " " & (IO_KBN_0);
                Case IO_KBN_NYU
                    Printer.Print wk_IO & " " & (IO_KBN_1);
                Case IO_KBN_SYU
                    Printer.Print wk_IO & " " & (IO_KBN_2);
                Case IO_KBN_ZAT
                    Printer.Print wk_IO & " " & (IO_KBN_3);
                Case IO_KBN_SYU_JITU
                    Printer.Print wk_IO & " " & (IO_KBN_4);
                Case IO_KBN_HENPIN
                    Printer.Print wk_IO & " " & (IO_KBN_5);
                Case Else
                    Printer.Print wk_IO;
            End Select

            Printer.Print Tab(MGN_L + 50);
            Printer.Print ChrCut(Ldata, 6);             '�`�[��

                                                        '����
            If wk_IO = IO_KBN_NYU Or wk_IO = IO_KBN_ZAT Or wk_IO = IO_KBN_HENPIN Then
                Printer.Print Tab(MGN_L + 57);          '���ɐ�
            Else
                Printer.Print Tab(MGN_L + 67);          '�o�ɐ�
            End If
            sts = Numeric_Check(EDIT_ONLY, 7, ZERO, NEGA_DIS, ZSUP_ENA, COMA_ENA, ChrCut(Ldata, 6), Work)
            
            Printer.Print Work;
            
            Printer.Print ChrCut(Ldata, 1);             ' �ԍ��敪
            
            Printer.Print Tab(MGN_L + 78);
            Printer.Print ChrCut(Ldata, 2);             '�q�ɋ敪�iνāj

            Printer.Print Tab(MGN_L + 83);              '�����敪
            Select Case Left(Ldata, 1)
                Case CYU_KBN_TUK
                    Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_1);
                Case CYU_KBN_SPO
                    Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_2);
                Case CYU_KBN_HJU
                    Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_3);
                Case CYU_KBN_TOK
                    Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_4);
                Case CYU_KBN_BOU
                    Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_E);
                Case Else
                    Printer.Print ChrCut(Ldata, 1);
            End Select

            Printer.Print Tab(MGN_L + 92);              ' �����敪
            Select Case Left(Ldata, 1)
                Case "*"
                    Printer.Print ChrCut(Ldata, 1) & " ��";
                                    
                Case Else
                    Printer.Print ChrCut(Ldata, 1);
            End Select
            Printer.Print Tab(MGN_L + 98);
            Printer.Print ChrCut(Ldata, 5);             '�����(���j
            Printer.Print " ";
            Printer.Print ChrCut(Ldata, 5);             '�����(��j

            Printer.Print Tab(MGN_L + 111);
            Printer.Print ChrCut(Ldata, 8);             '�W���I��

            'Select Case Ldata
            '    Case "0"
            '        Printer.Print Tab(MGN_L + 115);
            '        Printer.Print "�f�[�^�G���[";
            '    Case "1"
            '        Printer.Print Tab(MGN_L + 115);
            '        Printer.Print "�o�א斢�o�^";
            '    Case Else
            'End Select
                                                    '1997.10.30
'            Select Case Ldata
'                Case "2"
'                    Printer.Print "  �`�[�d��";
'            End Select
                                                    '1997.10.30
            
            Printer.Print
            Printer.Print

            Lcnt = Lcnt + 2
        End If
    Next i

    If Lcnt <> 99 Then
        Printer.EndDoc
        Printer.Orientation = vbPRORLandscape   '�p���̒��ӂ���ɂ��Ĉ��
    End If

End Sub
                                            '�w�b�_�[����i�u�i�ԕύX���X�g�v�j
Private Sub P_Hin_Head(Lcnt As Integer, JGYO_Kbn As String)
Dim i As Integer
Dim sts As Integer
Dim RetBuf As String

    If Lcnt <> 99 Then
        Printer.NewPage
    End If

    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print
    
    Printer.Print Tab(3);
    For i = ZERO To UBound(JGYOBU_T)
        If JGYO_Kbn = JGYOBU_T(i).Code Then
            Printer.Print "�i" & RTrim(JGYOBU_T(i).NAME) & "�j";
            Exit For
        End If
    Next i
    Printer.Print Tab(37);
    Printer.Print "�������@�i�ԕύX���X�g�@������";
    Printer.Print Tab(91);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")
    
    Printer.Print
                                        '���׃w�b�_���
    Printer.Print "------- �i�ԁi�O���j-------";
    Printer.Print Tab(30);
    Printer.Print "------- �i�ԁi�����j-------";
    Printer.Print
    
    Printer.Print Tab(MGN_L);
    Printer.Print "��M�f�[�^";
    Printer.Print Tab(MGN_L + 15);
    Printer.Print "�}�X�^";
    Printer.Print Tab(MGN_L + 30);
    Printer.Print "��M�f�[�^";
    Printer.Print Tab(MGN_L + 44);
    Printer.Print "�}�X�^";
    Printer.Print Tab(MGN_L + 58);
    Printer.Print "�`�[���t";
    Printer.Print Tab(MGN_L + 69);
    Printer.Print "���o�ɋ�";
    Printer.Print Tab(MGN_L + 78);
    Printer.Print "�`�[��";
    Printer.Print Tab(MGN_L + 86);
    Printer.Print "���o��";
    Printer.Print Tab(MGN_L + 93);
    Printer.Print "�q";
    Printer.Print Tab(MGN_L + 96);
    Printer.Print "������";
    Printer.Print Tab(MGN_L + 103);
    Printer.Print "�o�א�"
    Printer.Print

    Lcnt = 7 + MGN_U

End Sub
                                            '���׈���i�u�i�ԕύX���X�g�v�j
Private Sub P_Hin_Proc()

Dim Lcnt As Integer
Dim Ldata As String
Dim wk_IO As String
Dim Work As String
Dim Emsg As String
Dim Wqty As Long
Dim i As Integer
Dim sts As Integer
Dim B_Jgyobu As String

    Lcnt = 99

    For i = ZERO To LBox_Hin.ListCount - 1
        
        Ldata = LBox_Hin.List(i)

                                        '�w�b�_�[�R���g���[��
        If Lcnt > LMAX Or _
           B_Jgyobu <> Left(Ldata, 1) Then
            Call P_Hin_Head(Lcnt, Left(Ldata, 1))
            B_Jgyobu = Left(Ldata, 1)
        End If

                                        '���׈��
        Ldata = Mid(Ldata, 11, Len(Ldata) - 11)                     '���ƕ��C÷�ć��C�����O�@���O

        Printer.Print Tab(MGN_L);
        Printer.Print ChrCut(Ldata, 13);                            '��M�ް��i�ԁi�O���j
        Work = ChrCut(Ldata, 13)
        If Right(Ldata, 1) = "1" Or Right(Ldata, 1) = "2" Then      '�O���i�ԕύX�H
            Printer.Print Tab(MGN_L + 15);
            Printer.Print Work;                                     '�}�X�^�i�ԁi�O���j
        End If

        Printer.Print Tab(MGN_L + 30);
        Printer.Print ChrCut(Ldata, 13);                            '��M�ް��i�ԁi�����j
        Work = ChrCut(Ldata, 13)
        If Right(Ldata, 1) = "0" Then                               '�����i�ԕύX�H
            Printer.Print Tab(MGN_L + 44);
            Printer.Print Work;                                     '�}�X�^�i�ԁi�����j
        End If

        Printer.Print Tab(MGN_L + 58);                              '�`�[���t
        Printer.Print ChrCut(Ldata, 4) & "/" & ChrCut(Ldata, 2) & "/" & ChrCut(Ldata, 2);

        Printer.Print Tab(MGN_L + 69);                              '���o�ɋ敪
        wk_IO = ChrCut(Ldata, 1)
        Select Case wk_IO
            Case IO_KBN_URI
                Printer.Print wk_IO & " " & (IO_KBN_0);
            Case IO_KBN_NYU
                Printer.Print wk_IO & " " & (IO_KBN_1);
            Case IO_KBN_SYU
                Printer.Print wk_IO & " " & (IO_KBN_2);
            Case IO_KBN_ZAT
                Printer.Print wk_IO & " " & (IO_KBN_3);
            Case Else
                Printer.Print wk_IO;
        End Select

        Printer.Print Tab(MGN_L + 78);
        Printer.Print ChrCut(Ldata, 6);                             '�`�[��

        Printer.Print Tab(MGN_L + 85);                              '���o�ɐ�
        Wqty = CLng(ChrCut(Ldata, 6))
        
        
        sts = Numeric_Check(EDIT_ONLY, 7, ZERO, NEGA_DIS, ZSUP_ENA, COMA_ENA, Format(Wqty, "00000000"), Work)
        
        Printer.Print Work;

        Printer.Print Tab(MGN_L + 93);
        Printer.Print ChrCut(Ldata, 2);                             '�q�ɋ敪�iνāj

        Printer.Print Tab(MGN_L + 96);                              '�����敪
        Select Case Left(Ldata, 1)
            Case CYU_KBN_TUK
                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_1);
            Case CYU_KBN_SPO
                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_2);
            Case CYU_KBN_HJU
                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_3);
            Case CYU_KBN_BOU
                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_E);
            Case Else
                Printer.Print ChrCut(Ldata, 1);
        End Select

        Printer.Print Tab(MGN_L + 103);
        Printer.Print ChrCut(Ldata, 5);                             '�x����^�o�א�

        Printer.Print Tab(MGN_L + 110);                             '�ύX���b�Z�[�W
        Select Case Left(Ldata, 1)
            Case "0"
                Printer.Print "�����ύX Ͻ��i�ԓ���";
            Case "1"
                Printer.Print "�O���ύX Ͻ��i�ԓ���";
            Case "2"
                Printer.Print "�݌ɗL�I�O���ύX�s��";
        End Select
        Printer.Print

        Printer.Print

        Lcnt = Lcnt + 2
    Next i

    If Lcnt <> 99 Then
        Printer.EndDoc
        Printer.Orientation = vbPRORLandscape   '�p���̒��ӂ���ɂ��Ĉ��
    End If

End Sub
                                            '���׈���i�o�ח\��d�����X�g�j
Private Sub P_Dup_Proc()

Dim Lcnt As Integer
Dim Ldata As String
Dim wk_IO As String
Dim Work As String
Dim i As Integer
Dim sts As Integer
Dim B_Jgyobu As String

    Lcnt = 99
    B_Jgyobu = Space(1)

    For i = ZERO To LBox_Dup.ListCount - 1

        Ldata = LBox_Dup.List(i)

                                        '�w�b�_�[�R���g���[��
        If Lcnt > LMAX Or _
            Mid(Ldata, 2, 1) <> B_Jgyobu Then
            Call P_Dup_Head(Lcnt, Mid(Ldata, 2, 1))
            B_Jgyobu = Mid(Ldata, 2, 1)
        End If
                                        '���׈��
        Ldata = Mid(Ldata, 3, Len(Ldata) - 2)

        Printer.Print Tab(MGN_L);
        Printer.Print ChrCut(Ldata, 13);            '�i�ԁi�O���j

        Printer.Print Tab(MGN_L + 14);
        Printer.Print ChrCut(Ldata, 13);            '�i�ԁi�����j

        Printer.Print Tab(MGN_L + 28);              '�`�[���t
        Printer.Print ChrCut(Ldata, 4) & "/" & ChrCut(Ldata, 2) & "/" & ChrCut(Ldata, 2);

        Printer.Print Tab(MGN_L + 40);              '���o�ɋ敪
        wk_IO = ChrCut(Ldata, 1)
        Select Case wk_IO
            Case IO_KBN_URI
                Printer.Print wk_IO & " " & (IO_KBN_0);
            Case IO_KBN_NYU
                Printer.Print wk_IO & " " & (IO_KBN_1);
            Case IO_KBN_SYU
                Printer.Print wk_IO & " " & (IO_KBN_2);
            Case IO_KBN_ZAT
                Printer.Print wk_IO & " " & (IO_KBN_3);
            Case IO_KBN_SYU_JITU
                Printer.Print wk_IO & " " & (IO_KBN_4);
            Case IO_KBN_HENPIN
                Printer.Print wk_IO & " " & (IO_KBN_5);
            Case Else
                Printer.Print wk_IO;
        End Select

        Printer.Print Tab(MGN_L + 50);
        Printer.Print ChrCut(Ldata, 6);             '�`�[��

        Printer.Print Tab(MGN_L + 67);              '�o�ɐ�
            
        
        sts = Numeric_Check(EDIT_ONLY, 7, ZERO, NEGA_DIS, ZSUP_ENA, COMA_ENA, ChrCut(Ldata, 6), Work)
        
        Printer.Print Work;
            
        Printer.Print ChrCut(Ldata, 1);             ' �ԍ��敪
            
        Printer.Print Tab(MGN_L + 78);
        Printer.Print ChrCut(Ldata, 2);             '�q�ɋ敪�iνāj

        Printer.Print Tab(MGN_L + 83);              '�����敪
        Select Case Left(Ldata, 1)
            Case CYU_KBN_TUK
                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_1);
            Case CYU_KBN_SPO
                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_2);
            Case CYU_KBN_HJU
                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_3);
            Case CYU_KBN_TOK
                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_4);
            Case CYU_KBN_BOU
                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_E);
            Case Else
                Printer.Print ChrCut(Ldata, 1);
        End Select

        Printer.Print Tab(MGN_L + 92);              ' �����敪
        Select Case Left(Ldata, 1)
            Case "*"
                Printer.Print ChrCut(Ldata, 1) & " ��";
                                    
            Case Else
                Printer.Print ChrCut(Ldata, 1);
        End Select
        Printer.Print Tab(MGN_L + 98);
        Printer.Print ChrCut(Ldata, 5);             '�����(���j
        Printer.Print " ";
        Printer.Print ChrCut(Ldata, 5);             '�����(��j

        Printer.Print Tab(MGN_L + 111);
        Printer.Print ChrCut(Ldata, 8);             '�W���I��

            
        Printer.Print
        Printer.Print

        Lcnt = Lcnt + 2
    Next i

    If Lcnt <> 99 Then
        Printer.EndDoc
        Printer.Orientation = vbPRORLandscape   '�p���̒��ӂ���ɂ��Ĉ��
    End If

End Sub
                                            '�w�b�_�[����i�o�ח\��d�����X�g�j
Private Sub P_Dup_Head(Lcnt As Integer, JGYO_Kbn As String)
Dim i As Integer
Dim sts As Integer
Dim RetBuf As String

    If Lcnt <> 99 Then
        Printer.NewPage
    End If

    Printer.Print


    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print
                                        '�w�b�_�[�i�P�j
    Printer.Print Tab(3);
    For i = ZERO To UBound(JGYOBU_T)
        If JGYO_Kbn = JGYOBU_T(i).Code Then
            Printer.Print "�i" & RTrim(JGYOBU_T(i).NAME) & "�j";
            Exit For
        End If
    Next i
    Printer.Print Tab(37);
    Printer.Print "���� �o�ח\��d�����X�g�@����";
    Printer.Print Tab(91);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")

    Printer.Print
                                        '���׃w�b�_���
    Printer.Print Tab(MGN_L);
    Printer.Print "�i�ԁi�O���j";
    Printer.Print Tab(MGN_L + 14);
    Printer.Print "�i�ԁi�����j";
    Printer.Print Tab(MGN_L + 28);
    Printer.Print "�`�[���t";
    Printer.Print Tab(MGN_L + 39);
    Printer.Print "���o�ɋ敪";
    Printer.Print Tab(MGN_L + 50);
    Printer.Print "�`�[��";
    Printer.Print Tab(MGN_L + 58);
    Printer.Print "���ɐ�";
    Printer.Print Tab(MGN_L + 68);
    Printer.Print "�o�ɐ�";
    Printer.Print Tab(MGN_L + 77);
    Printer.Print "�q��";
    Printer.Print Tab(MGN_L + 82);
    Printer.Print "�敪";
    Printer.Print Tab(MGN_L + 92);
    Printer.Print "����";
    Printer.Print Tab(MGN_L + 98);
    Printer.Print "�\�Z�P�ʁ@�@ �W���I��"
    Printer.Print

    Lcnt = 6 + MGN_U

End Sub
                                            '������̐؏o��
Private Function ChrCut(Moto As String, Leng As Long) As String
    ChrCut = Left(Moto, Leng)

    If Len(Moto) <= Leng Then
        Moto = ""
        Exit Function
    End If

    Moto = Mid(Moto, Leng + 1, Len(Moto) - Leng)
End Function
                                            '�w�b�_�[����i�u�ߏ�O�ؕiؽāv�j
Private Sub P_Last_Head(Lcnt As Integer, JGYO_Kbn As String)
Dim i As Integer
Dim sts As Integer
Dim RetBuf As String

    If Lcnt <> 99 Then
        Printer.NewPage
    End If

    Printer.Print


    For i = 1 To MGN_U2
        Printer.Print
    Next i

    Printer.Print
                                        '�w�b�_�[�i�P�j
    Printer.Print Tab(3);
    For i = ZERO To UBound(JGYOBU_T)
        If JGYO_Kbn = JGYOBU_T(i).Code Then
            Printer.Print "�i" & RTrim(JGYOBU_T(i).NAME) & "�j";
            Exit For
        End If
    Next i
    Printer.Print Tab(39);
    Printer.Print "������ �ߏ�O�ؕi���X�g�@������";
    Printer.Print Tab(91);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")

    Printer.Print
                                        '���׃w�b�_���
    Printer.Print Tab(MGN_L2 + 1);
    Printer.Print "�����O";
    Printer.Print Tab(MGN_L2 + 11);
    Printer.Print "�i�ԁi�O���j";
    Printer.Print Tab(MGN_L2 + 29);
    Printer.Print "�i�@��";
    Printer.Print Tab(MGN_L2 + 62);
    Printer.Print "�ߏ萔"
    Printer.Print

    Lcnt = 6 + MGN_U2

End Sub
                                            '�O�؂���ׁA���юc�`�F�b�N
Private Sub Last_Proc()
Dim sts         As Integer
Dim Lcnt        As Integer
Dim Command     As Integer
Dim RetBuf      As String
Dim B_Jgyobu    As String

Dim ans         As Integer

    Call Input_Lock           '��ʍ��ڃ��b�N

    MsgLab(7).Visible = True       '�X�V��ү���ޕ\��
    DoEvents

'    Printer.Orientation = vbPRORLandscape   '�p���̒��ӂ���ɂ��Ĉ��
    Set Printer.Font = NormalFont           '����t�H���g�ݒ�
    Lcnt = 99
    B_Jgyobu = Space(1)

'���������ް��X�V
    Command = BtOpGetFirst
    Do
        Do
            sts = BTRV(Command + BtSNoWait, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), ZERO)
            Select Case sts
                Case BtNoErr, BtErrKeyNotFound, BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Sub
                    End If
                Case Else
                    Call File_Error(sts, Command, "���������ް�")
                    Exit Sub
            End Select
        Loop

        If sts = BtErrKeyNotFound Or sts = BtErrEOF Then
            Exit Do
        End If

                                        '�w�b�_�[�R���g���[��
        If Lcnt > LMAX Or _
           StrConv(J_NYUREC.JGYOBU, vbUnicode) <> B_Jgyobu Then
            Call P_Last_Head(Lcnt, StrConv(J_NYUREC.JGYOBU, vbUnicode))
            B_Jgyobu = StrConv(J_NYUREC.JGYOBU, vbUnicode)
        End If
                                        '���׈��
        Printer.Print Tab(MGN_L2 + 2);
        If StrConv(J_NYUREC.NAIGAI, vbUnicode) = "1" Then
            Printer.Print NAIGAI1;         '����
        Else
            Printer.Print NAIGAI2;         '�C�O
        End If

        Printer.Print Tab(MGN_L2 + 11);
        Printer.Print StrConv(J_NYUREC.HIN_GAI, vbUnicode);

        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(J_NYUREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(J_NYUREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(J_NYUREC.HIN_GAI, vbUnicode))
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
        Select Case sts
            Case BtNoErr
                Printer.Print Tab(MGN_L2 + 29);
                Printer.Print StrConv(ITEMREC.HIN_NAME, vbUnicode);
            Case BtErrKeyNotFound
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                Exit Sub
        End Select

        sts = Numeric_Check(EDIT_ONLY, 10, ZERO, NEGA_DIS, ZSUP_ENA, COMA_ENA, StrConv(J_NYUREC.JITU_QTY, vbUnicode), RetBuf)
        Printer.Print Tab(MGN_L2 + 58);
        Printer.Print RetBuf;
        Printer.Print

        Printer.Print

        Lcnt = Lcnt + 2
''�܈�O�؂�f�[�^�͎c��
''        Do
''            sts = BTRV(BtOpDelete, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), ZERO)
''            Select Case sts
''                Case BtNoErr
''                    Exit Do
''                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
''                    Beep
''                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
''                    If ans = vbCancel Then
''                        Exit Sub
''                    End If
''                Case Else
''                    Call File_Error(sts, BtOpDelete, "���������ް�")
''                    Exit Sub
''            End Select
''
''        Loop
        
        Command = BtOpGetNext
    Loop

    If Lcnt <> 99 Then
        Printer.EndDoc
    End If

End Sub
Private Sub RERUN_PROC()

Dim i   As Integer
Dim sts As Integer

'�����Ď��s����
    SelCmd(3).Enabled = True
    SelCmd(1).Enabled = True
    SelCmd(2).Enabled = True
    
    Label2(ZERO).Visible = True
    Text1(ZERO).Visible = True
                                                '����@ SPIC
    Call UniCode_Conv(K0_SEQCK.JGYOBU, "1")
    Call UniCode_Conv(K0_SEQCK.SEQ_MODE, "1")
        
    sts = BTRV(BtOpGetEqual, SEQCK_POS, SEQCKREC, Len(SEQCKREC), K0_SEQCK, Len(K0_SEQCK), ZERO)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(SEQCKREC.LAST_TXTNO, "000000000")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�\��捞�݃`�F�b�N")
            Unload Me
    End Select

    Text1(ZERO).Text = StrConv(SEQCKREC.LAST_TXTNO, vbUnicode)

                                                '����@ ������
    Call UniCode_Conv(K0_SEQCK.JGYOBU, "1")
    Call UniCode_Conv(K0_SEQCK.SEQ_MODE, "2")
        
    sts = BTRV(BtOpGetEqual, SEQCK_POS, SEQCKREC, Len(SEQCKREC), K0_SEQCK, Len(K0_SEQCK), ZERO)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(SEQCKREC.LAST_TXTNO, "000000000")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�\��捞�݃`�F�b�N")
            Unload Me
    End Select

    Text1(1).Text = StrConv(SEQCKREC.LAST_TXTNO, vbUnicode)


    Text1(ZERO).SetFocus
End Sub
Private Sub Command_Click(Index As Integer)

    Select Case Index
        Case 11
            Unload Me
        Case Else
            Beep
    End Select

End Sub
Private Sub Command_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyF12
            Command(11).Value = True
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
        Case vbKeyZ
            If Shift = 1 Then
                Call RERUN_PROC
            End If
    End Select


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = ZERO
    End If
End Sub
Private Sub Form_Load()
Dim i As Integer
Dim c As String * 128
Dim sts As Integer


Dim sBuffer As String * 255
Dim com     As String
    
    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If

    Show
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
                                '�o�׃��O�t�@�C������荞��
    If SYUKA_LOGF_GET_PROC() Then
        Beep
        MsgBox "�o�׃��O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
                                
                                '�V�X�e���\��ϗv����荞��
    If SYSTEM_YOIN_Set() Then
        Beep
        MsgBox "�V�X�e���\��ϗv���̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If

                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
                                '�L��νđq�Ɏ�荞�݁i�|���@�j
    For i = ZERO To UBound(SOKO_T1) - 1
        SOKO_T1(i).HS_SOKO = "  "
        SOKO_T1(i).NAIGAI = " "
    Next i

    i = ZERO
    Do
        If GetIni("NYUSYU_OK_SOKO", "SOKO1" & RTrim(Format(i + 1, "#0")), "SYS", c) Then
            Beep
            MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
            Call Log_Out(LOG_F, "[SYS.INI] [NYUSYU_OK_SOKO] [SOKO] READ ERROR")
            End
        End If
        If RTrim(c) = "**" Then
            Exit Do
        End If
        SOKO_T1(i).HS_SOKO = RTrim(c)
        If GetIni("NYUSYU_OK_SOKO", "NAIG1" & RTrim(Format(i + 1, "#0")), "SYS", c) Then
            Beep
            MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
            Call Log_Out(LOG_F, "[SYS.INI] [NYUSYU_OK_SOKO] [NAIG] READ ERROR")
            End
        End If
        SOKO_T1(i).NAIGAI = RTrim(c)
        i = i + 1
    Loop

    If Kaso_Soko_No_Set() Then
        Beep
        MsgBox "���z�q�ɂ̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
                                'ܰ��ð��ݔԍ���荞��
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> ZERO Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "??"
    End If
    WS_NO = RTrim(com)

                                '�q�Ƀ}�X�^�n�o�d�m
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^�i�X�V�p���[�N�j�n�o�d�m
    If wITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '������Ǘ��}�X�^�n�o�d�m
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '������Ǒփ}�X�^�n�o�d�m�@2001.07.04
    If MTSCHG_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�v���X�^�n�o�d�m
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ƀf�[�^�n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ɉړ����n�o�d�m
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���ח\��n�o�d�m
    If Y_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�o�ח\��n�o�d�m
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���������ް��n�o�d�m
    If J_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�\��捞�������n�o�d�m
    If SEQCK_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '����t�H���g�ݒ�
    With NormalFont
        .NAME = F1020121.FontName
        .Size = F1020121.FontSize
    End With
    Set Printer.Font = NormalFont
                                '��ʏ����ݒ�
    Call Scr_Init
    
    Last_Proc_F = False         '���������ް��폜�����@���s�L���t���O�N���A

    Shori_Mode = 3
    Call Data_Inport            '�z�X�g�f�[�^�捞�ݏ���
    
    Unload Me



End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer

    If Last_Proc_F = True Then              '���������ް��폜�����@���s�L��H
        Call Last_Proc
    End If

                                            '�q�Ƀ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�q�Ƀ}�X�^")
        End If
    End If
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '�i�ڃ}�X�^�i�X�V�p���[�N�j�b�k�n�r�d
    sts = BTRV(BtOpClose, wITEM_POS, wITEMREC, Len(wITEMREC), K0_wITEM, Len(K0_wITEM), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '������Ǘ��}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "������Ǘ��}�X�^")
        End If
    End If
                                            '������Ǒփ}�X�^�b�k�n�r�d�@2001.07.04
    sts = BTRV(BtOpClose, MTSCHG_POS, MTSCHGREC, Len(MTSCHGREC), K0_MTSCHG, Len(K0_MTSCHG), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "������Ǒփ}�X�^")
        End If
    End If
                                            '�v���}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�v���}�X�^")
        End If
    End If
                                            '�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        End If
    End If
                                            '�݌Ɉړ����b�k�n�r�d
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ɉړ���")
        End If
    End If
                                            '���ח\��b�k�n�r�d
    sts = BTRV(BtOpClose, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ח\��")
        End If
    End If
                                            '�o�ח\��b�k�n�r�d
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�o�ח\��")
        End If
    End If
                                            '���������ް��b�k�n�r�d
    sts = BTRV(BtOpClose, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���������ް�")
        End If
    End If
                                            '�\��捞�������b�k�n�r�d
    sts = BTRV(BtOpClose, SEQCK_POS, SEQCKREC, Len(SEQCKREC), K0_SEQCK, Len(K0_SEQCK), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�\��捞������")
        End If
    End If
                                            '�a���������������Z�b�g
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), ZERO)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F1020121 = Nothing

    End
End Sub


Private Sub SelCmd_Click(Index As Integer)

Dim ans As Integer
        
    Beep
    ans = MsgBox("�u" & SelCmd(Index).Caption & "�v" & "�@�捞�ݏ����@���s���܂����H", vbYesNo + vbQuestion, "�m�F����")
    If ans = vbYes Then
        If Text1(ZERO).Visible Then
            Call SEQCHEK_PUT
        End If
        Shori_Mode = Index
        Call Data_Inport            '�z�X�g�f�[�^�捞�ݏ���
    End If

End Sub
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F1020121.MousePointer = vbHourglass

    Call Ctrl_Lock(F1020121)

End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1020121)

    F1020121.MousePointer = vbDefault

End Sub

Private Sub Text1_GotFocus(Index As Integer)
    
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = ZERO
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If


End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
Dim i   As Integer
    
    If KeyCode <> vbKeyReturn Then Exit Sub
            

    For i = Index + 1 To Text_Max
        If Text1(i).Enabled And Text1(i).Visible And Text1(i).TabStop Then
            Text1(i).SetFocus
            Exit For
        End If
    Next i

End Sub

Private Sub SEQCHEK_PUT()
    
Dim sts As Integer
Dim com As Integer
Dim ans As Integer
                                    '����@ SPIC
    Call UniCode_Conv(K0_SEQCK.JGYOBU, "1")
    Call UniCode_Conv(K0_SEQCK.SEQ_MODE, "1")
        
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, SEQCK_POS, SEQCKREC, Len(SEQCKREC), K0_SEQCK, Len(K0_SEQCK), ZERO)
    
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SEQCHK.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Unload Me
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�\��捞�݃`�F�b�N")
                Unload Me
        End Select

    Loop

    If com = BtOpInsert Then
        Call UniCode_Conv(SEQCKREC.JGYOBU, "1")
        Call UniCode_Conv(SEQCKREC.SEQ_MODE, "1")
    End If

    Call UniCode_Conv(SEQCKREC.LAST_TXTNO, Text1(ZERO).Text)
    Call UniCode_Conv(SEQCKREC.LAST_GET_DT, Format(Date, "yyyymmdd"))
    Call UniCode_Conv(SEQCKREC.LAST_GET_TM, Format(Time, "HHmmss"))

    Do
        sts = BTRV(com, SEQCK_POS, SEQCKREC, Len(SEQCKREC), K0_SEQCK, Len(K0_SEQCK), ZERO)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SEQCHK.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Unload Me
                End If
            Case Else
                Call File_Error(sts, com, "�\��捞�݃`�F�b�N")
                Unload Me
        End Select
    
    Loop

                                    '����@ ������
    Call UniCode_Conv(K0_SEQCK.JGYOBU, "1")
    Call UniCode_Conv(K0_SEQCK.SEQ_MODE, "2")
        
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, SEQCK_POS, SEQCKREC, Len(SEQCKREC), K0_SEQCK, Len(K0_SEQCK), ZERO)
    
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SEQCHK.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Unload Me
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�\��捞�݃`�F�b�N")
                Unload Me
        End Select

    Loop

    If com = BtOpInsert Then
        Call UniCode_Conv(SEQCKREC.JGYOBU, "1")
        Call UniCode_Conv(SEQCKREC.SEQ_MODE, "2")
    End If

    Call UniCode_Conv(SEQCKREC.LAST_TXTNO, Text1(1).Text)
    Call UniCode_Conv(SEQCKREC.LAST_GET_DT, Format(Date, "yyyymmdd"))
    Call UniCode_Conv(SEQCKREC.LAST_GET_TM, Format(Time, "HHmmss"))

    Do
        sts = BTRV(com, SEQCK_POS, SEQCKREC, Len(SEQCKREC), K0_SEQCK, Len(K0_SEQCK), ZERO)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SEQCHK.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Unload Me
                End If
            Case Else
                Call File_Error(sts, com, "�\��捞�݃`�F�b�N")
                Unload Me
        End Select
    
    Loop

End Sub


Private Sub Err_Log_Out(Mesg As String)
Dim Work_Rec    As String

                                '÷�ć�
        Work_Rec = StrConv(XX_SIJREC.TEXT_NO, vbUnicode)
                                '���ƕ��敪
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.JGYOBU, vbUnicode)
                                '�����敪
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.CYOK_KBN, vbUnicode)
                                '�`�[���t
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.DEN_DT, vbUnicode)
                                '���o�ɋ敪
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.IO_KBN, vbUnicode)
                                '�ԍ��敪
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.PM_KBN, vbUnicode)
                                '�`�[���
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.DEN_SYU, vbUnicode)
                                '�`�[��
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.DEN_NO, vbUnicode)
                                '�����敪
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.CYU_KBN, vbUnicode)
                                '�i�ԁi�O���j
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.HIN_GAI, vbUnicode)
                                '�i�ԁi�����j
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.HIN_NAI, vbUnicode)
                                '�i��
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.HIN_NAME, vbUnicode)
                                '����
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.YOTEI_QTY, vbUnicode)
                                '�\�Z�P�ʁi���j
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.YOSAN_FROM, vbUnicode)
                                '�\�Z�P�ʁi��j
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.YOSAN_TO, vbUnicode)
                                '�q�ɋ敪�iνāj
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.HOST_SOKO, vbUnicode)
                                '�I�ԁiνāj
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.HOST_TANA, vbUnicode)
                                '�x����^�o�א�
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.SYUK_CODE, vbUnicode)
                                '�x����^�o�א於
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.SYUK_NAME, vbUnicode)
                                'ں��ޏI�[ϰ�(@)
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.REC_END, vbUnicode)

        Call Log_Out(LOG_F, Mesg & " " & Work_Rec)

End Sub
