VERSION 5.00
Begin VB.Form F1030901 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�o�ח\�芮������"
   ClientHeight    =   8070
   ClientLeft      =   2445
   ClientTop       =   3315
   ClientWidth     =   12195
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
   ScaleHeight     =   8070
   ScaleWidth      =   12195
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   2760
      MaxLength       =   20
      TabIndex        =   9
      Top             =   3600
      Width           =   2535
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   7
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   6
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   2520
      MaxLength       =   2
      TabIndex        =   5
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   4
      Top             =   2280
      Width           =   615
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   2
      Left            =   1680
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   8
      Top             =   3600
      Width           =   972
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   1
      Left            =   1560
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   1320
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   1
      Top             =   720
      Width           =   972
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   1320
      MaxLength       =   12
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   4  '�S�p�Ђ炪��
      Index           =   8
      Left            =   7200
      MaxLength       =   10
      TabIndex        =   12
      Top             =   6240
      Width           =   2532
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   7
      Left            =   5280
      MaxLength       =   6
      TabIndex        =   11
      Top             =   6240
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   4140
      Index           =   0
      Left            =   5400
      TabIndex        =   10
      Top             =   1560
      Width           =   6255
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
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6960
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
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6960
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
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6960
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
      Left            =   7800
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6960
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
      Left            =   6480
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6960
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6960
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
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6960
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
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6960
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
      Left            =   2640
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6960
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6960
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6960
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�X  �V"
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6960
      Width           =   855
   End
   Begin VB.Label lblGOODS_F 
      Caption         =   "Label2"
      Height          =   255
      Left            =   3480
      TabIndex        =   62
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "������"
      Height          =   255
      Index           =   23
      Left            =   840
      TabIndex        =   61
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Height          =   375
      Index           =   11
      Left            =   3480
      TabIndex        =   60
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Height          =   375
      Index           =   10
      Left            =   1680
      TabIndex        =   59
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Height          =   375
      Index           =   9
      Left            =   3600
      TabIndex        =   58
      Top             =   5760
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Height          =   375
      Index           =   8
      Left            =   2760
      TabIndex        =   57
      Top             =   5760
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Height          =   375
      Index           =   7
      Left            =   1680
      TabIndex        =   56
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Height          =   375
      Index           =   6
      Left            =   4200
      TabIndex        =   55
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Height          =   375
      Index           =   5
      Left            =   3360
      TabIndex        =   54
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Height          =   375
      Index           =   4
      Left            =   2520
      TabIndex        =   53
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Height          =   375
      Index           =   3
      Left            =   1680
      TabIndex        =   52
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Height          =   375
      Index           =   2
      Left            =   4200
      TabIndex        =   51
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   50
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   49
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�h�c��"
      Height          =   255
      Index           =   22
      Left            =   480
      TabIndex        =   48
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�S��"
      Height          =   255
      Index           =   21
      Left            =   720
      TabIndex        =   47
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblTanto_Name 
      BackColor       =   &H80000005&
      Height          =   255
      Left            =   2160
      TabIndex        =   46
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���@��"
      Height          =   255
      Index           =   17
      Left            =   6360
      TabIndex        =   45
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����敪"
      Height          =   240
      Index           =   12
      Left            =   240
      TabIndex        =   44
      Top             =   840
      Width           =   960
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�j"
      Height          =   255
      Index           =   20
      Left            =   5160
      TabIndex        =   43
      Top             =   4800
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "(�o�ח\�萔"
      Height          =   255
      Index           =   19
      Left            =   2760
      TabIndex        =   42
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�o�׎c��"
      Height          =   255
      Index           =   18
      Left            =   600
      TabIndex        =   41
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�^"
      Height          =   255
      Index           =   16
      Left            =   2400
      TabIndex        =   40
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�^"
      Height          =   255
      Index           =   15
      Left            =   3240
      TabIndex        =   39
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`�[��"
      Height          =   255
      Index           =   14
      Left            =   840
      TabIndex        =   38
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`�[���t"
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   37
      Top             =   2400
      Width           =   975
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
      TabIndex        =   36
      Top             =   7560
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�o�ɐ�"
      Height          =   255
      Index           =   11
      Left            =   4440
      TabIndex        =   35
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�ԁi�����j"
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   34
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�^"
      Height          =   255
      Index           =   9
      Left            =   2880
      TabIndex        =   33
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�^"
      Height          =   255
      Index           =   8
      Left            =   2280
      TabIndex        =   32
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���ד�"
      Height          =   255
      Index           =   7
      Left            =   840
      TabIndex        =   31
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i  ��"
      Height          =   255
      Index           =   6
      Left            =   840
      TabIndex        =   30
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i  ��"
      Height          =   255
      Index           =   5
      Left            =   840
      TabIndex        =   29
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�|"
      Height          =   255
      Index           =   3
      Left            =   3840
      TabIndex        =   28
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�|"
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   27
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�|"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   26
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�I��"
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   25
      Top             =   5400
      Width           =   495
   End
   Begin VB.Menu MainMenu 
      Caption         =   "���ƕ�"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1030901"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WS_NO As String * 3

Dim MENU_NO As String * 2   '2007.07.11


Private Const ptxTANTO_CODE% = 0        '�S���R�[�h
Private Const ptxID_No% = 1             'ID��
Private Const ptxSYUKA_YY% = 2          '�o�ד��@�N
Private Const ptxSYUKA_MM% = 3          '�o�ד��@��
Private Const ptxSYUKA_DD% = 4          '�o�ד��@��
Private Const ptxDEN_NO% = 5            '�`�[��
Private Const ptxHIN_NO% = 6            '�i�ԁi�O���j
Private Const ptxJitu_QTY% = 7          '�o�ɐ�
Private Const ptxMEMO% = 8              '����
Private Const Text_Max% = 8            '

Private Const plblHIN_NAME% = 0         '�i��
Private Const plblSURYO_ZAN% = 1        '�o�׎c��
Private Const plblSURYO% = 2            '�o�א�
Private Const plblSoko_No% = 3          '�q�ɇ�
Private Const plblRetu% = 4             '��
Private Const plblRen% = 5              '�A
Private Const plblDan% = 6              '�i
Private Const plblNYUKA_YY% = 7         '���ד��@�N
Private Const plblNYUKA_MM% = 8         '���ד��@��
Private Const plblNYUKA_DD% = 9         '���ד��@��
Private Const plblHIN_NAI% = 10         '�i�ԁi�����j
Private Const plblGOODS_F% = 11         '���i�^�����i
Private Const Label_Max% = 11           '

Private Const pcmbCYU_KBN% = 0          '�����敪
Private Const pcmbMUKE_CODE% = 1        '������R�[�h
Private Const pcmbNAIGAI% = 2           '���O�敪

Private Const plstZaiko% = 0            '�݌ɏ��

Private Function Update_Proc() As Integer

Dim sts             As Integer

Dim HS_CYU_KBN      As String * 1
Dim YOIN            As String * 2
Dim SUMI_JITU_QTY   As Long
Dim MI_JITU_QTY     As Long
Dim SYUKA_QTY       As Long
    
    
    Update_Proc = True

    Call Input_Lock
                                        '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, ZAIKO_POS, ZAIKOREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If
                                    
    If Right(Combo(pcmbCYU_KBN).Text, 1) = CYU_KBN_KIN Then
        YOIN = ACT_SYUKA_GAI & Right(Combo(pcmbCYU_KBN).Text, 1)
    Else
        YOIN = ACT_SYUKA_KEI & Right(Combo(pcmbCYU_KBN).Text, 1)
    End If
                                    
    If lblGOODS_F.Caption = GOODS_ON Then
        SUMI_JITU_QTY = CLng(Text(ptxJitu_QTY).Text)
    Else
        MI_JITU_QTY = CLng(Text(ptxJitu_QTY).Text)
    End If
                                    
    SYUKA_QTY = 0                   '���i�^�����i���ʖ��̐��ʁi�u�O�v�Œ�j
                                    '�o�ɏ���
    sts = Syuko_Update_Proc(Last_JGYOBU, _
                            Right(Combo(pcmbNAIGAI).Text, 1), _
                            Text(ptxHIN_NO).Text, _
                            (Label1(plblNYUKA_YY).Caption & Label1(plblNYUKA_MM).Caption & Label1(plblNYUKA_DD).Caption), _
                            (Label1(plblSoko_No).Caption & Label1(plblRetu).Caption & Label1(plblRen).Caption & Label1(plblDan).Caption), _
                            YOIN, _
                            SUMI_JITU_QTY, _
                            MI_JITU_QTY, _
                            SYUKA_QTY, _
                            WS_NO, _
                            Text(ptxTANTO_CODE).Text, , _
                            Text(ptxMEMO).Text, _
                            Right(Combo(pcmbCYU_KBN).Text, 1), _
                            Right(Combo(pcmbMUKE_CODE).Text, 16), _
                            Text(ptxSYUKA_YY).Text & Text(ptxSYUKA_MM).Text & Text(ptxSYUKA_DD).Text, _
                            Text(ptxDEN_NO).Text, _
                            Text(ptxID_No).Text, MENU_NO)   ''2007.07.11 MENU_NO�ǉ�
    Select Case sts
        Case False
        Case Else
            Update_Proc = sts
            GoTo Abort_Tran
    End Select


End_Tran:
                                        '�g�����U�N�V�����I��
    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
    Call Input_UnLock
    
    Update_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If

    Call Input_UnLock


End Function

Private Function Err_Chk() As Integer

Dim sts     As Integer

Dim CYU_KBN As String * 1

    Err_Chk = True
                
    Call Input_Lock
                                        '�S����
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text(ptxTANTO_CODE).Text)
    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
        
            lblTanto_Name.Caption = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
        
        Case BtErrKeyNotFound
            Beep
            MsgBox "���͂������ڂ̓G���[�ł��B�i�S���ҁj"
            Call Input_UnLock
            Text(ptxTANTO_CODE).SetFocus
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
            Err_Chk = SYS_ERR
            Exit Function
    End Select

    Select Case Right(Combo(pcmbCYU_KBN).Text, 1)
        Case CYU_KBN_KIN
            '-------------------------  '�ً}�̃`�F�b�N
            If Len(Trim(Text(ptxID_No).Text)) = 0 Then      'IDNo
            Else
                If IsNumeric(Text(ptxID_No).Text) Then
                    Beep
                    MsgBox "���͂��č��ڂ̓G���[�ł��B"
                    Call Input_UnLock
                    Text(ptxID_No).SetFocus
                    Exit Function
                End If
                                        '�Y���f�[�^���L������G���[
                Call UniCode_Conv(K0_Y_SYU.JGYOBU, Last_JGYOBU)
'                Call UniCode_Conv(K0_Y_SYU.KEY_CYU_KBN, CYU_KBN_KIN)2004.04.08
                Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, Text(ptxID_No).Text)
                sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                Select Case sts
                    Case BtNoErr
                        Beep
                        MsgBox "���͂������ڂ̓G���[�ł��B�i�o�ח\��o�^�ς݁j"
                        Call Input_UnLock
                        Text(ptxID_No).SetFocus
                        Exit Function
                    Case BtErrKeyNotFound
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�o�ח\��")
                        Err_Chk = SYS_ERR
                        Exit Function
                End Select
            
            End If
                                        '�`�[���t�`�`�[���̃`�F�b�N���ҏW
            If IsNumeric(Text(ptxSYUKA_MM).Text) Then
                Text(ptxSYUKA_MM).Text = Format(CInt(Text(ptxSYUKA_MM).Text), "00")
            Else
                Beep
                MsgBox "���͂��č��ڂ̓G���[�ł��B"
                Call Input_UnLock
                Text(ptxSYUKA_MM).SetFocus
                Exit Function
            End If
            
            If IsNumeric(Text(ptxSYUKA_DD).Text) Then
                Text(ptxSYUKA_DD).Text = Format(CInt(Text(ptxSYUKA_DD).Text), "00")
            Else
                Beep
                MsgBox "���͂��č��ڂ̓G���[�ł��B"
                Call Input_UnLock
                Text(ptxSYUKA_DD).SetFocus
                Exit Function
            End If

            If Not IsDate(Text(ptxSYUKA_YY).Text & "/" & Text(ptxSYUKA_MM).Text & "/" & Text(ptxSYUKA_DD).Text) Then
                Beep
                MsgBox "���͂��č��ڂ̓G���[�ł��B"
                Call Input_UnLock
                Text(ptxSYUKA_YY).SetFocus
                Exit Function
            End If
            
            sts = Item_Read_Proc
            Select Case sts
                Case False
                    Label1(plblHIN_NAME).Caption = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                Case True
                    Label1(plblHIN_NAME).Caption = ""
                    Beep
                    MsgBox "���͂������ڂ̓G���[�ł��B�i�i�ԁj"
                    Call Input_UnLock
                    Text(ptxHIN_NO).SetFocus
                    Exit Function
                Case Else
                    Call Input_UnLock
                    Err_Chk = SYS_ERR
                    Exit Function
            End Select
        Case Else
                                        '���̏o�ח\��͊m�ۍς݁H

            Call UniCode_Conv(K0_Y_SYU.JGYOBU, Last_JGYOBU)     '���ƕ�
                                                                '�����敪
'            Call UniCode_Conv(K0_Y_SYU.KEY_CYU_KBN, Right(Combo(pcmbCYU_KBN), 1))2004.04.08
                                                                'ID��
            Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, Text(ptxID_No).Text)
            sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
            Select Case sts
                Case BtNoErr
                    If Len(Trim(StrConv(Y_SYUREC.WEL_ID, vbUnicode))) = 0 Then
                        Beep
                        MsgBox "�X�V�Ώۂ̏o�ח\�肪�m�肵�Ă��܂���B"
                        Call Input_UnLock
                        Text(ptxID_No).SetFocus
                        Exit Function
                    End If

                    If StrConv(Y_SYUREC.WEL_ID, vbUnicode) <> WS_NO Or _
                        Trim(StrConv(Y_SYUREC.PRG_ID, vbUnicode)) <> StrConv(App.EXEName, vbUpperCase) Then
                        Beep
                        MsgBox "���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>"
                        Call Input_UnLock
                        Text(ptxID_No).SetFocus
                        Exit Function
                    End If
                
                
                Case BtErrKeyNotFound
                    Beep
                    MsgBox "���͂����R�[�h�͓o�^����Ă��܂���B�i�o�ח\��j"
                    Call Clear_Field(ptxHIN_NO)
                    Call Input_UnLock
                    Text(ptxID_No).SetFocus
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�o�ח\��")
                    Err_Chk = SYS_ERR
                    Exit Function
            End Select
    End Select

    If List1(plstZaiko).ListCount = 0 Then
        Beep
        MsgBox "�o�ɉ\�ȍ݌ɂ����݂��܂���B"
        Call Input_UnLock
        If Combo(pcmbNAIGAI).TabStop Then
            Combo(pcmbNAIGAI).SetFocus
        Else
            Text(ptxID_No).SetFocus
        End If
        Exit Function
    End If

    If Len(Trim(Label1(plblSoko_No).Caption)) = 0 Then
        Beep
        MsgBox "�o�Ɍ��݌ɂ��I������Ă��܂���B"
        Call Input_UnLock
        List1(plstZaiko).ListIndex = 0
        List1(plstZaiko).SetFocus
        Exit Function
    End If

    If IsNumeric(Text(ptxJitu_QTY).Text) Then
        If CLng(Text(ptxJitu_QTY).Text) = 0 Then
            Beep
            MsgBox "���͂��č��ڂ̓G���[�ł��B"
            Call Input_UnLock
            Text(ptxJitu_QTY).SetFocus
            Exit Function
        Else
            If Label1(plblSURYO_ZAN).Visible Then
                If CLng(Label1(plblSURYO_ZAN).Caption) < CLng(Text(ptxJitu_QTY).Text) Then
                    Beep
                    MsgBox "���͂��č��ڂ̓G���[�ł��B"
                    Call Input_UnLock
                    Text(ptxJitu_QTY).SetFocus
                    Exit Function
                End If
            End If
        End If
    Else
        Beep
        MsgBox "���͂��č��ڂ̓G���[�ł��B"
        Call Input_UnLock
        Text(ptxJitu_QTY).SetFocus
        Exit Function
    End If
                                            '�݌ɐ��ʂ̃`�F�b�N
    Call UniCode_Conv(K1_ZAIKO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K1_ZAIKO.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))
    Call UniCode_Conv(K1_ZAIKO.HIN_GAI, Text(ptxHIN_NO).Text)
    Call UniCode_Conv(K1_ZAIKO.GOODS_ON, lblGOODS_F)
    Call UniCode_Conv(K1_ZAIKO.NYUKA_DT, Label1(plblNYUKA_YY).Caption & Label1(plblNYUKA_MM).Caption & Label1(plblNYUKA_DD).Caption)
    Call UniCode_Conv(K1_ZAIKO.Soko_No, Label1(plblSoko_No).Caption)
    Call UniCode_Conv(K1_ZAIKO.Retu, Label1(plblRetu).Caption)
    Call UniCode_Conv(K1_ZAIKO.Ren, Label1(plblRen).Caption)
    Call UniCode_Conv(K1_ZAIKO.Dan, Label1(plblDan).Caption)
    sts = BTRV(BtOpGetEqual, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
    Select Case sts
        Case BtNoErr
            If StrConv(ZAIKOREC.LOCK_F, vbUnicode) = LOCK_OFF Then
                Beep
                MsgBox "�o�Ɍ��݌ɂ��m�肵�Ă��܂���B"
                Call Input_UnLock
                List1(plstZaiko).ListIndex = 0
                List1(plstZaiko).SetFocus
                Exit Function
            End If

            If StrConv(ZAIKOREC.WEL_ID, vbUnicode) <> WS_NO Or _
                Trim(StrConv(ZAIKOREC.PRG_ID, vbUnicode)) <> StrConv(App.EXEName, vbUpperCase) Then
                Beep
                MsgBox "���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>"
                Call Input_UnLock
                List1(plstZaiko).ListIndex = 0
                List1(plstZaiko).SetFocus
                Exit Function
            End If
        Case BtErrKeyNotFound
            Beep
            MsgBox "�݌Ƀf�[�^�����ŕύX����Ă��܂��B"
            Call Input_UnLock
            List1(plstZaiko).ListIndex = 0
            List1(plstZaiko).SetFocus
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�݌Ƀf�[�^")
            Err_Chk = SYS_ERR
            Exit Function
    End Select

    If CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) < CLng(Text(ptxJitu_QTY).Text) Then
        Beep
        MsgBox "���͂��č��ڂ̓G���[�ł��B"
        Call Input_UnLock
        Text(ptxJitu_QTY).SetFocus
        Exit Function
    End If
    
    Call Input_UnLock

    Err_Chk = False

End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F1030901.MousePointer = vbHourglass

    Call Ctrl_Lock(F1030901)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1030901)


    F1030901.MousePointer = vbDefault

End Sub

Private Sub Combo_Click(Index As Integer)
Dim sts As Integer
    
    Select Case Index
    
        Case pcmbCYU_KBN
            Call Input_Lock
            
                                            '�o�ח\��̊J��
            If Y_Syuka_UnLock() Then
                Unload Me
            End If
            

            If Zaiko_UNLock_Proc("", "", "", "", WS_NO) Then
                Unload Me
            End If
    
            Call Clear_Field(1)
    
            Call Input_UnLock
            Call Input_Change_Proc
    
            Text(ptxID_No).SetFocus
    End Select



End Sub

Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim sts As Integer
'----------------------------------------------------------------------------
'                   �R���{�{�b�N�X���́i�j�����c�������j����
'----------------------------------------------------------------------------
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    Select Case Index
        Case pcmbCYU_KBN    '�����敪
            
                                        
                                            '�o�ח\��̊J��
            If Y_Syuka_UnLock() Then
                Unload Me
            End If
            
            If Zaiko_UNLock_Proc("", "", "", "", WS_NO) Then
                Unload Me
            End If
                
                
                
            Call Clear_Field(1)
            
            
            Input_Change_Proc          '���͍��ڂ�؂�ւ���
            
            Text(ptxID_No).SetFocus
        
        Case pcmbMUKE_CODE  '������
            Text(ptxSYUKA_YY).SetFocus
    
        Case pcmbNAIGAI  '�����O
            Text(ptxHIN_NO).SetFocus
    End Select

End Sub

Private Sub Command_Click(Index As Integer)
Dim yn As Integer
Dim sts As Integer

    Select Case Index
        Case 0
                                            '�G���[�`�F�b�N
            sts = Err_Chk()
            Select Case sts
                Case False
                Case True
                    Exit Sub
                Case SYS_ERR
                    Unload Me
            End Select
            
            Beep
            yn = MsgBox("�X�V���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                sts = Update_Proc()
                Select Case sts
                    Case False
                    Case True, SYS_CANCEL
                        Call Clear_Field(1)
                        Text(ptxID_No).SetFocus
                        Exit Sub
                    Case SYS_ERR
                        Unload Me
                End Select
                If Label1(plblSURYO_ZAN).Visible Then
                    If CLng(Label1(plblSURYO_ZAN).Caption) > CLng(Text(ptxJitu_QTY).Text) Then
                        Call Clear_Field(8)
                        sts = Y_Syuka_Disp_Proc()       '�o�ח\����e�\��
                
                        Select Case sts
                            Case False
                            Case True, SYS_CANCEL
                                Call Clear_Field(1)
                                Text(ptxID_No).SetFocus
                                Exit Sub
                            Case SYS_ERR
                                Unload Me
                        End Select
                        Text(ptxJitu_QTY).Text = ""
                
                
                        sts = Zaiko_Disp_Proc()         '�o�׉\�݌ɕ\��
                        Select Case sts
                            Case False
                                List1(plstZaiko).ListIndex = 0
                                List1(plstZaiko).SetFocus
                                                
                            Case True, SYS_CANCEL
                                Call Clear_Field(1)
                                Text(ptxID_No).SetFocus
                            Case SYS_ERR
                                Unload Me
                        End Select
                    Else
                        Call Clear_Field(1)
                        Text(ptxID_No).SetFocus
                    End If
                Else
                    Call Clear_Field(1)
                    Text(ptxID_No).SetFocus
                End If
            Else
                Call Clear_Field(1)
                Text(ptxID_No).SetFocus
            End If
            Exit Sub
            
        Case 7
            If Len(Trim(Text(ptxID_No).Text)) = 0 Then
                Beep
                Text(ptxID_No).SetFocus
                Exit Sub
            End If
            
            sts = Zaiko_Disp_Proc()         '�o�׉\�݌ɕ\��
            Select Case sts
                Case False
                    List1(plstZaiko).ListIndex = 0
                    List1(plstZaiko).SetFocus
                    Exit Sub
                Case True, SYS_CANCEL
                    Text(ptxID_No).SetFocus
                    Exit Sub
                Case SYS_ERR
                    Unload Me
            End Select
        Case 11
            
                                            '�݌ɂ̊J��
            If Zaiko_UNLock_Proc("", "", "", "", WS_NO) Then
                Unload Me
            End If
                                            '�o�ח\��̊J��
            If Y_Syuka_UnLock() Then
                Unload Me
            End If
            
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
                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1030901.Caption = "�o�ח\�芮�������i" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i

    Unload SubMenu(i)
                                
                                
                                        '�ƭ����l�� 2007.07.11
    If GetIni(App.EXEName, "MENU_NO", "SYS", c) Then
        MENU_NO = ""
    Else
        MENU_NO = Trim(c)
    End If
                                
                                
                                'ܰ��ð��ݔԍ���荞��
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "??"
    End If
    WS_NO = RTrim(com)
                                '�q�Ƀ}�X�^�n�o�d�m
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�I�}�X�^�n�o�d�m
    If TANA_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�S���}�X�^�n�o�d�m
    If TANTO_Open(BtOpenNomal) Then
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
                                '�v���}�X�^�n�o�d�m
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ƀf�[�^�t�@�C���n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�o�ח\��f�[�^�t�@�C���n�o�d�m
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ɉړ����f�[�^�t�@�C���n�o�d�m
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���ԃ}�X�^�n�o�d�m
    If HATUBAN_Open(BtOpenNomal) Then
        Unload Me
    End If
'---------------------------------------------- '��Ǝ���۸ނn�o�d�m
    If P_SAGYO_LOG_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
    Combo(pcmbCYU_KBN).Clear
    Combo(pcmbCYU_KBN).AddItem CYU_KBN_1$ & Space(5) & CYU_KBN_TUK
    Combo(pcmbCYU_KBN).AddItem CYU_KBN_2$ & Space(5) & CYU_KBN_SPO
    Combo(pcmbCYU_KBN).AddItem CYU_KBN_3$ & Space(5) & CYU_KBN_HJU
    Combo(pcmbCYU_KBN).AddItem CYU_KBN_E$ & Space(5) & CYU_KBN_BOU
'    Combo(pcmbCYU_KBN).AddItem CYU_KBN_4$ & Space(5) & CYU_KBN_TOK
    Combo(pcmbCYU_KBN).AddItem CYU_KBN_T$ & Space(5) & CYU_KBN_KIN
    Combo(pcmbCYU_KBN).ListIndex = 0
                    
    Combo(pcmbNAIGAI).AddItem NAIGAI1 & Space(5) & NAIGAI_NAI
    Combo(pcmbNAIGAI).AddItem NAIGAI2 & Space(5) & NAIGAI_GAI
    Combo(pcmbNAIGAI).ListIndex = 0
                        '������ݒ�
    If MTS_Set_Proc() Then
        Unload Me
    End If
                                
                                
    Text(ptxTANTO_CODE).SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
                                            
                                    '�o�ח\��̊J��
    If Y_Syuka_UnLock() Then
    End If
                                            
                                    '�݌ɂ̊J��
    If Zaiko_UNLock_Proc("", "", "", "", WS_NO) Then
    End If
                                            
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
                                            '�S���҃}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�S���҃}�X�^")
        End If
    End If
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '�i�ڃ}�X�^�i�X�V�p���[�N�j�b�k�n�r�d
    sts = BTRV(BtOpClose, wITEM_POS, wITEMREC, Len(wITEMREC), K0_wITEM, Len(K0_wITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '������Ǘ��}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "������Ǘ��}�X�^")
        End If
    End If
                                            '�v���}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�v���}�X�^")
        End If
    End If
                                            '�݌Ƀf�[�^�t�@�C���b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        End If
    End If
                                            '�o�ח\��f�[�^CLOSE
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�o�ח\��f�[�^")
        End If
    End If
                                            '�݌Ɉړ���CLOSE
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ɉړ���")
        End If
    End If
                                            '���ԃ}�X�^CLOSE
    sts = BTRV(BtOpClose, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), K0_HATUBAN, Len(K0_HATUBAN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ɉړ���")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F1030901 = Nothing

    End
End Sub

Private Sub List1_DblClick(Index As Integer)
    
Dim sts As Integer
    
Dim LOCATION    As String * 8
Dim End_Flg     As Boolean
    
    Call Input_Lock
                                        '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Unload Me
    End If

    If Zaiko_UNLock_Proc("", "", "", "", WS_NO) Then
        Unload Me
    End If

                                                
    LOCATION = Mid(List1(Index).List(List1(Index).ListIndex), 14, 2) & _
                Mid(List1(Index).List(List1(Index).ListIndex), 17, 2) & _
                Mid(List1(Index).List(List1(Index).ListIndex), 20, 2) & _
                Mid(List1(Index).List(List1(Index).ListIndex), 23, 2)

    End_Flg = False
    sts = Zaiko_Lock_Proc(LOCATION, Last_JGYOBU, Right(Combo(pcmbNAIGAI).Text, 1), Text(ptxHIN_NO).Text, WS_NO)
    Select Case sts
        Case False
        Case True, SYS_CANCEL
            GoTo Abort_Tran
        Case SYS_ERR
            End_Flg = True
            GoTo Abort_Tran
    End Select
                                                '�݌Ƀf�[�^�t�@�C���ǂݍ���
    Call UniCode_Conv(K1_ZAIKO.JGYOBU, Last_JGYOBU)                             '���ƕ�
    Call UniCode_Conv(K1_ZAIKO.NAIGAI, Right(Combo(pcmbNAIGAI), 1))             '�����O
    Call UniCode_Conv(K1_ZAIKO.HIN_GAI, Text(ptxHIN_NO).Text)                   '�i�ԁi�O���j
                                                                                '���i�^�����i
    If Left(List1(plstZaiko).List(List1(plstZaiko).ListIndex), 1) = "*" Then
        Call UniCode_Conv(K1_ZAIKO.GOODS_ON, GOODS_ON)
        lblGOODS_F.Caption = GOODS_ON
    Else
        Call UniCode_Conv(K1_ZAIKO.GOODS_ON, GOODS_OFF)
        lblGOODS_F.Caption = GOODS_OFF
    End If
                                                                                
                                                                                '���ד�
    Call UniCode_Conv(K1_ZAIKO.NYUKA_DT, Mid(List1(plstZaiko).List(List1(plstZaiko).ListIndex), 3, 4) & _
                                            Mid(List1(plstZaiko).List(List1(plstZaiko).ListIndex), 8, 2) & _
                                            Mid(List1(plstZaiko).List(List1(plstZaiko).ListIndex), 11, 2))
    Call UniCode_Conv(K1_ZAIKO.Soko_No, Left(LOCATION, 2))                      '�q�ɇ�
    Call UniCode_Conv(K1_ZAIKO.Retu, Mid(LOCATION, 3, 2))                       '��
    Call UniCode_Conv(K1_ZAIKO.Ren, Mid(LOCATION, 5, 2))                        '�A
    Call UniCode_Conv(K1_ZAIKO.Dan, Right(LOCATION, 2))                         '�i
        
    sts = BTRV(BtOpGetEqual, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
    Select Case sts
        Case BtNoErr
            Call Zaiko_Detail_Proc
        Case BtErrKeyNotFound
            Beep
            MsgBox "�f�[�^���e���ύX����Ă��܂��B�u�ŐV�v�\����I�����Ă��������B"
            GoTo Abort_Tran
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�݌Ƀf�[�^")
            End_Flg = True
            GoTo Abort_Tran
    End Select
                                        '�g�����U�N�V�����I��
End_Tran:
    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        End_Flg = True
        GoTo Abort_Tran
    End If
    
    Call Input_UnLock

    Text(ptxJitu_QTY).SetFocus
    
    Exit Sub

Abort_Tran:
    Call Input_UnLock
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
        Unload Me
    End If

    If End_Flg Then
        Unload Me
    End If
    
    List1(plstZaiko).SetFocus

End Sub

Private Sub List1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
Dim sts As Integer
    
Dim LOCATION    As String * 8
Dim End_Flg     As Boolean
    
    If List1(plstZaiko).ListCount = 0 Then Exit Sub
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Call Input_Lock
                                        '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Unload Me
    End If

    If Zaiko_UNLock_Proc("", "", "", "", WS_NO) Then
        Unload Me
    End If

                                                
    LOCATION = Mid(List1(Index).List(List1(Index).ListIndex), 14, 2) & _
                Mid(List1(Index).List(List1(Index).ListIndex), 17, 2) & _
                Mid(List1(Index).List(List1(Index).ListIndex), 20, 2) & _
                Mid(List1(Index).List(List1(Index).ListIndex), 23, 2)

    End_Flg = False
    sts = Zaiko_Lock_Proc(LOCATION, Last_JGYOBU, Right(Combo(pcmbNAIGAI).Text, 1), Text(ptxHIN_NO).Text, WS_NO)
    Select Case sts
        Case False
        Case True, SYS_CANCEL
            GoTo Abort_Tran
        Case SYS_ERR
            End_Flg = True
            GoTo Abort_Tran
    End Select
                                                '�݌Ƀf�[�^�t�@�C���ǂݍ���
    Call UniCode_Conv(K1_ZAIKO.JGYOBU, Last_JGYOBU)                             '���ƕ�
    Call UniCode_Conv(K1_ZAIKO.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))        '�����O
    Call UniCode_Conv(K1_ZAIKO.HIN_GAI, Text(ptxHIN_NO).Text)                   '�i�ԁi�O���j
                                                                                
                                                                                '���i�^�����i
    If Left(List1(plstZaiko).List(List1(plstZaiko).ListIndex), 1) = "*" Then
        Call UniCode_Conv(K1_ZAIKO.GOODS_ON, GOODS_ON)
        lblGOODS_F.Caption = GOODS_ON
    Else
        Call UniCode_Conv(K1_ZAIKO.GOODS_ON, GOODS_OFF)
        lblGOODS_F.Caption = GOODS_OFF
    End If
                                                                                '���ד�
    Call UniCode_Conv(K1_ZAIKO.NYUKA_DT, Mid(List1(plstZaiko).List(List1(plstZaiko).ListIndex), 3, 4) & _
                                            Mid(List1(plstZaiko).List(List1(plstZaiko).ListIndex), 8, 2) & _
                                            Mid(List1(plstZaiko).List(List1(plstZaiko).ListIndex), 11, 2))
        
    Call UniCode_Conv(K1_ZAIKO.Soko_No, Left(LOCATION, 2))                      '�q�ɇ�
    Call UniCode_Conv(K1_ZAIKO.Retu, Mid(LOCATION, 3, 2))                       '��
    Call UniCode_Conv(K1_ZAIKO.Ren, Mid(LOCATION, 5, 2))                        '�A
    Call UniCode_Conv(K1_ZAIKO.Dan, Right(LOCATION, 2))                         '�i
        
    sts = BTRV(BtOpGetEqual, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
    Select Case sts
        Case BtNoErr
            Call Zaiko_Detail_Proc
        Case BtErrKeyNotFound
            Beep
            MsgBox "�f�[�^���e���ύX����Ă��܂��B�u�ŐV�v�\����I�����Ă��������B"
            GoTo Abort_Tran
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�݌Ƀf�[�^")
            End_Flg = True
            GoTo Abort_Tran
    End Select
                                        '�g�����U�N�V�����I��
End_Tran:
    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        End_Flg = True
        GoTo Abort_Tran
    End If
    
    Call Input_UnLock

    Text(ptxJitu_QTY).SetFocus
    
    Exit Sub

Abort_Tran:
    Call Input_UnLock
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
        Unload Me
    End If

    If End_Flg Then
        Unload Me
    End If

    List1(plstZaiko).SetFocus

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
    F1030901.Caption = "�o�ח\�芮�������i" + RTrim(JGYOBU_T(Index).NAME) + "�j"
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

Dim sts     As Integer
Dim i       As Integer
    
    If KeyCode <> vbKeyReturn Then Exit Sub
            
    Select Case Index
                
        Case ptxTANTO_CODE
                                            '�S���҂̃`�F�b�N
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text(ptxTANTO_CODE).Text)
            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                    lblTanto_Name.Caption = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
'                    Combo(pcmbCYU_KBN).SetFocus
'                    Exit Sub
                Case BtErrKeyNotFound
                    lblTanto_Name.Caption = ""
                    MsgBox "���͂������ڂ̓G���[�ł��B�i�S���ҁj"
                    Exit Sub
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
                    Unload Me
            End Select
        
        
            Combo(pcmbCYU_KBN).SetFocus
            Exit Sub
        
        
        Case ptxID_No
            If Right(Combo(pcmbCYU_KBN).Text, 1) = CYU_KBN_KIN Then
                                            '�ً}�͉������Ȃ�
            Else
                sts = Y_Syuka_Disp_Proc()   '�o�ח\����e�\��
                Select Case sts
                    Case False
                    Case True, SYS_CANCEL
                        Text(ptxID_No).SetFocus
                        Exit Sub
                    Case SYS_ERR
                        Unload Me
                End Select
                
                sts = Zaiko_Disp_Proc()         '�o�׉\�݌ɕ\��
                Select Case sts
                    Case False
                        List1(plstZaiko).ListIndex = 0
                        List1(plstZaiko).SetFocus
                        Exit Sub
                    Case True, SYS_CANCEL
                        Text(ptxID_No).SetFocus
                        Exit Sub
                    Case SYS_ERR
                        Unload Me
                End Select
            End If
        
        Case ptxHIN_NO          '�i�ԁi�O���j�ˋً}���̂�
            sts = Item_Read_Proc
            Select Case sts
                Case False
                    Label1(plblHIN_NAME).Caption = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                Case True
                    Label1(plblHIN_NAME).Caption = ""
                    MsgBox "���͂������ڂ̓G���[�ł��B�i�i�ԁj"
                    Exit Sub
                Case Else
                    Unload Me
            End Select
                
            sts = Zaiko_Disp_Proc()         '�o�׉\�݌ɕ\��
            Select Case sts
                Case False
                    List1(plstZaiko).ListIndex = 0
                    List1(plstZaiko).SetFocus
                    Exit Sub
                Case True, SYS_CANCEL
                    Text(ptxID_No).SetFocus
                    Exit Sub
                Case SYS_ERR
                    Unload Me
            End Select
    
    End Select
            
    For i = Index + 1 To Text_Max
        If Text(i).Visible And Text(i).Enabled And Text(i).TabStop Then
            Text(i).SetFocus
            Exit For
        End If
    Next i
End Sub


Private Sub Input_Change_Proc()

    Select Case Right(Combo(pcmbCYU_KBN).Text, 1)
        Case CYU_KBN_TUK, CYU_KBN_SPO, CYU_KBN_HJU, CYU_KBN_TOK, CYU_KBN_BOU
            
            Text(ptxDEN_NO).Locked = True          '�`�[��
            Text(ptxDEN_NO).TabStop = False
    
            Combo(pcmbNAIGAI).Locked = True         '�����O
            Combo(pcmbNAIGAI).TabStop = False
            
            Text(ptxHIN_NO).Locked = True           '�i�ԁi�O���j
            Text(ptxHIN_NO).TabStop = False
        
            Label1(plblSURYO_ZAN).Visible = True     '�o�׎c��
            Label1(plblSURYO).Visible = True         '�o�ח\�萔
        
        
        
                 
        
        Case CYU_KBN_KIN
            
            Text(ptxDEN_NO).Locked = False          '�`�[��
            Text(ptxDEN_NO).TabStop = True
        
        
            Combo(pcmbNAIGAI).Locked = False        '�����O
            Combo(pcmbNAIGAI).TabStop = True
            
            Text(ptxHIN_NO).Locked = False         '�i�ԁi�O���j
            Text(ptxHIN_NO).TabStop = True
        
            Label1(plblSURYO_ZAN).Visible = False    '�o�׎c��
            Label1(plblSURYO).Visible = False        '�o�ח\�萔
        
    
    End Select
    

End Sub

Private Function Y_Syuka_Disp_Proc() As Integer

Dim sts     As Integer
Dim CYU_KBN As String * 1

Dim ans     As Integer

Dim i       As Integer

    Y_Syuka_Disp_Proc = True
                                    
    Call Input_Lock
    
    sts = Y_Syuka_UnLock()          '�o�ח\��̊J��
    If sts Then
        Call Input_UnLock
        Y_Syuka_Disp_Proc = sts
        Exit Function
    End If
    
            
    sts = Zaiko_UNLock_Proc("", "", "", "", WS_NO)
    If sts Then
        Call Input_UnLock
        Y_Syuka_Disp_Proc = sts
        Exit Function
    End If
    
    
    sts = Y_Syuka_Lock()            '�o�ח\��̊m��
    If sts Then
        Call Input_UnLock
        Y_Syuka_Disp_Proc = sts
        Exit Function
    End If
                                        
                                        
                                        
                                        '�f�[�^�L���`�F�b�N
    If CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) = CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) Then
        Beep
        MsgBox "���͂����R�[�h�͏o�׏����ςł��B"
        sts = Y_Syuka_UnLock()          '�o�ח\��̊J��
        If sts Then
            Y_Syuka_Disp_Proc = sts
        End If
        Call Input_UnLock
        Exit Function
    End If
                                                
    If StrConv(Y_SYUREC.CYU_KBN, vbUnicode) <> Right(Combo(pcmbCYU_KBN).Text, 1) Then
        Beep
        MsgBox "�����敪�Ⴂ�ł��B"
        sts = Y_Syuka_UnLock()          '�o�ח\��̊J��
        If sts Then
            Y_Syuka_Disp_Proc = sts
        End If
        Call Input_UnLock
        Exit Function
    End If
                                                '�o�ד��e�̕\��
                                                                                                                                        
'    For i = 0 To Combo(pcmbCYU_KBN).ListCount - 1
'        If Right(Combo(pcmbCYU_KBN).List(i), 1) = StrConv(Y_SYUREC.CYU_KBN, vbUnicode) Then
'            Combo(pcmbCYU_KBN).ListIndex = i
'            Exit For
'        End If
'
'    Next i
    
    
    
    
    For i = 0 To Combo(pcmbMUKE_CODE).ListCount - 1 '������
    
        If Right(Combo(pcmbMUKE_CODE).List(i), 16) = StrConv(Y_SYUREC.MUKE_CODE, vbUnicode) & StrConv(Y_SYUREC.SS_CODE, vbUnicode) Then
            Combo(pcmbMUKE_CODE).ListIndex = i
            Exit For
        End If
    
    
    Next
                                                    '�`�[���t
    Text(ptxSYUKA_YY).Text = Left(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 4)
    Text(ptxSYUKA_MM).Text = Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 5, 2)
    Text(ptxSYUKA_DD).Text = Right(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 2)
                                                    '�`�[��
    Text(ptxDEN_NO).Text = Trim(StrConv(Y_SYUREC.DEN_NO, vbUnicode))
                                        
    For i = 0 To Combo(pcmbNAIGAI).ListCount - 1    '�����O
        If StrConv(Y_SYUREC.NAIGAI, vbUnicode) = Right(Combo(pcmbNAIGAI).List(i), 1) Then
            Combo(pcmbNAIGAI).ListIndex = i
            Exit For
        End If
    Next i
                                                        '�i�ԁi�O���j
    Text(ptxHIN_NO).Text = RTrim(StrConv(Y_SYUREC.HIN_NO, vbUnicode))

    sts = Item_Read_Proc
    Select Case sts
        Case False
            Label1(plblHIN_NAME).Caption = StrConv(ITEMREC.HIN_NAME, vbUnicode)
        Case True
            Label1(plblHIN_NAME).Caption = ""
        Case Else
            Call Input_UnLock
            Y_Syuka_Disp_Proc = sts
            Exit Function
    End Select
                                                        '�o�׎c��
    Label1(plblSURYO_ZAN).Caption = Format((CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) - _
                                    CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode))), "#0")
                                                        '�o�ח\�萔
    Label1(plblSURYO).Caption = Format(CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)), "#0")
    
    Call Input_UnLock

    Y_Syuka_Disp_Proc = False
End Function

Private Function Y_Syuka_UnLock() As Integer

Dim sts     As Integer
Dim ans     As Integer

    Y_Syuka_UnLock = True

    Call UniCode_Conv(K4_Y_SYU.WEL_ID, WS_NO)
    Call UniCode_Conv(K4_Y_SYU.PRG_ID, StrConv(App.EXEName, vbUpperCase))

    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K4_Y_SYU, Len(K4_Y_SYU), 4)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
                Y_Syuka_UnLock = False
                Exit Function
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Y_Syuka_UnLock = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��")
                Y_Syuka_UnLock = SYS_ERR
                Exit Function
        End Select
    Loop
                                        '�g�p�\�����
    
    Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
    Call UniCode_Conv(Y_SYUREC.PRG_ID, "")

    Do
        sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K4_Y_SYU, Len(K4_Y_SYU), BtNCC)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    
                    
                    '2004.07.07��
                    sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                    
                    If sts Then
                    
                        Call File_Error(sts, BtOpUnlock, "�o�ח\��")
                        Y_Syuka_UnLock = SYS_ERR
                        Exit Function
                    
                    End If
                    '2004.07.07��
                    
                    
                    Y_Syuka_UnLock = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpUpdate, "�o�ח\��")
                Y_Syuka_UnLock = SYS_ERR
                Exit Function
        End Select
    Loop

    Y_Syuka_UnLock = False

End Function

Private Function Y_Syuka_Lock() As Integer
    
Dim CYU_KBN As String * 1

Dim sts     As Integer
Dim ans     As Integer
    
    Y_Syuka_Lock = True
    
    
    Call UniCode_Conv(K0_Y_SYU.JGYOBU, Last_JGYOBU)     '���ƕ�
                                                        '�����敪
'    Call UniCode_Conv(K0_Y_SYU.KEY_CYU_KBN, Right(Combo(pcmbCYU_KBN).Text, 1))2004.04.08
                                                        '�h�c��
    Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, Text(ptxID_No).Text)
                    
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        Select Case sts
            Case BtNoErr
                                
                If Len(Trim(StrConv(Y_SYUREC.WEL_ID, vbUnicode))) <> 0 Then
                    If StrConv(Y_SYUREC.WEL_ID, vbUnicode) <> WS_NO Or _
                        Trim(StrConv(Y_SYUREC.PRG_ID, vbUnicode)) <> StrConv(App.EXEName, vbUpperCase) Then
                        '2004.07.07��
                        sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)

                        If sts Then

                            Call File_Error(sts, BtOpUnlock, "�o�ח\��")
                            Y_Syuka_Lock = SYS_ERR
                            Exit Function

                        End If
                        '2004.07.07��
                        
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Y_Syuka_Lock = SYS_CANCEL
                            Exit Function
                        End If
                    Else
                        
                        
                        '2004.07.07��
                        sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)

                        If sts Then

                            Call File_Error(sts, BtOpUnlock, "�o�ח\��")
                            Y_Syuka_Lock = SYS_ERR
                            Exit Function

                        End If
                        '2004.07.07��
                        
                        Y_Syuka_Lock = False
                        Exit Function
                    End If
                Else
                    Exit Do
                End If
            Case BtErrKeyNotFound
                Beep
                MsgBox "���͂����R�[�h�͓o�^����Ă��܂���B�i�o�ח\��j"
                Call Clear_Field(ptxHIN_NO)
                Exit Function
            
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Y_Syuka_Lock = SYS_CANCEL
                    Exit Function
                End If
           Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��")
                Y_Syuka_Lock = SYS_ERR
                Exit Function
        End Select
    Loop
                                        '�g�p�\��
    Call UniCode_Conv(Y_SYUREC.WEL_ID, WS_NO)
    Call UniCode_Conv(Y_SYUREC.PRG_ID, StrConv(App.EXEName, vbUpperCase))
    Do
        sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K5_Y_SYU, Len(K5_Y_SYU), 5)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    
                    
                    
                    '2004.07.07��
                    sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)

                    If sts Then

                        Call File_Error(sts, BtOpUnlock, "�o�ח\��")
                        Y_Syuka_Lock = SYS_ERR
                        Exit Function

                    End If
                    '2004.07.07��
                    
                    
                    Y_Syuka_Lock = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpUpdate, "�o�ח\��")
                Y_Syuka_Lock = SYS_ERR
                Exit Function
        End Select
    Loop

    Y_Syuka_Lock = False

End Function

Private Function Zaiko_Disp_Proc() As Integer

Dim sts     As Integer
Dim com     As Integer

Dim Edit    As String
Dim RetBuf  As String

    Zaiko_Disp_Proc = True

    Call Input_Lock

    Call UniCode_Conv(K1_ZAIKO.JGYOBU, Last_JGYOBU) '���ƕ�
                                                    '�����O
    Call UniCode_Conv(K1_ZAIKO.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))
                                                    '�i�ځi�O���j
    Call UniCode_Conv(K1_ZAIKO.HIN_GAI, Text(ptxHIN_NO).Text)
    Call UniCode_Conv(K1_ZAIKO.GOODS_ON, "")        '���i�^�����i
    Call UniCode_Conv(K1_ZAIKO.NYUKA_DT, "")        '���ד�
    Call UniCode_Conv(K1_ZAIKO.Soko_No, "")         '�q�ɇ�
    Call UniCode_Conv(K1_ZAIKO.Retu, "")            '��
    Call UniCode_Conv(K1_ZAIKO.Ren, "")             '�A
    Call UniCode_Conv(K1_ZAIKO.Dan, "")             '�i

    List1(plstZaiko).Clear

    com = BtOpGetGreaterEqual

    Do
        sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
        Select Case sts
            Case BtNoErr
                
                If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNAIGAI).Text, 1) Or _
                    RTrim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) <> Text(ptxHIN_NO).Text Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call Input_UnLock
                Call File_Error(sts, com, "�݌Ƀf�[�^")
                Zaiko_Disp_Proc = SYS_ERR
                Exit Function
        End Select
                                            '�I�̎g�p�ۃ`�F�b�N�i�I�}�X�^���j
        Call UniCode_Conv(K0_TANA.Soko_No, StrConv(ZAIKOREC.Soko_No, vbUnicode))
        Call UniCode_Conv(K0_TANA.Retu, StrConv(ZAIKOREC.Retu, vbUnicode))
        Call UniCode_Conv(K0_TANA.Ren, StrConv(ZAIKOREC.Ren, vbUnicode))
        Call UniCode_Conv(K0_TANA.Dan, StrConv(ZAIKOREC.Dan, vbUnicode))
        
        sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_OK Then
                
                    If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = GOODS_OFF Then
                        Edit = "  "
                    Else
                        Edit = "* "
                    End If
                                            
                    Edit = Edit & Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 1, 4) & "/"
                    Edit = Edit & Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 5, 2) & "/"
                    Edit = Edit & Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 7, 2) & " "
                    Edit = Edit & StrConv(ZAIKOREC.Soko_No, vbUnicode) & "-"
                    Edit = Edit & StrConv(ZAIKOREC.Retu, vbUnicode) & "-"
                    Edit = Edit & StrConv(ZAIKOREC.Ren, vbUnicode) & "-"
                    Edit = Edit & StrConv(ZAIKOREC.Dan, vbUnicode) & " "
                    Edit = Edit & Left(StrConv(ZAIKOREC.HIN_NAI, vbUnicode), 13) & " "
                    
                    RetBuf = Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)), "#0")
                    RetBuf = Space(10 - Len(RetBuf)) & RetBuf
                    Edit = Edit & RetBuf
                    
                    List1(plstZaiko).AddItem Edit
                                    
                End If
            Case BtErrKeyNotFound           '���o�^�͎g�p�s�Ƃ݂Ȃ�
            Case Else
                Call Input_UnLock
                Call File_Error(sts, BtOpGetEqual, "�I�}�X�^")
                Zaiko_Disp_Proc = SYS_ERR
                Exit Function
        End Select
        
        com = BtOpGetNext
    Loop
    
    
    '��ɓ��͂��ꂽ���e���N���A�[
    Label1(plblSoko_No).Caption = ""
    Label1(plblRetu).Caption = ""
    Label1(plblRen).Caption = ""
    Label1(plblDan).Caption = ""
    
    Label1(plblNYUKA_YY).Caption = ""
    Label1(plblNYUKA_MM).Caption = ""
    Label1(plblNYUKA_DD).Caption = ""
    Label1(plblHIN_NAI).Caption = ""
    
    Label1(plblGOODS_F).Caption = ""
    
    Text(ptxJitu_QTY).Text = ""
    
    
    
    Call Input_UnLock
    
    If List1(plstZaiko).ListCount = 0 Then
        Beep
        MsgBox "�o�ɉ\�ȍ݌ɂ����݂��܂���B"
    Else
        Zaiko_Disp_Proc = False
    End If

End Function

Private Sub Zaiko_Detail_Proc()
        
        
    Label1(plblSoko_No).Caption = StrConv(ZAIKOREC.Soko_No, vbUnicode)
    Label1(plblRetu).Caption = StrConv(ZAIKOREC.Retu, vbUnicode)
    Label1(plblRen).Caption = StrConv(ZAIKOREC.Ren, vbUnicode)
    Label1(plblDan).Caption = StrConv(ZAIKOREC.Dan, vbUnicode)
    
    Label1(plblNYUKA_YY).Caption = Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 1, 4)
    Label1(plblNYUKA_MM).Caption = Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 5, 2)
    Label1(plblNYUKA_DD).Caption = Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 7, 2)
    Label1(plblHIN_NAI).Caption = StrConv(ZAIKOREC.HIN_NAI, vbUnicode)
    
    If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = "0" Then
        Label1(plblGOODS_F).Caption = "���@�i"
    Else
        Label1(plblGOODS_F).Caption = "�����i"
    End If
    
    If Label1(plblSURYO_ZAN).Visible Then
        If CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) > CLng(Label1(plblSURYO_ZAN).Caption) Then
            Text(ptxJitu_QTY).Text = Label1(plblSURYO_ZAN).Caption
        Else
            Text(ptxJitu_QTY).Text = Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)), "#0")
        End If
    Else
        Text(ptxJitu_QTY).Text = Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)), "#0")
    End If
End Sub

Private Function Item_Read_Proc() As Integer
            
Dim sts As Integer
            
    Item_Read_Proc = True
            
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)      '�i���i�i�ڃ}�X�^���j
    Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text(ptxHIN_NO).Text)
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
            Exit Function
    End Select

    Item_Read_Proc = False

End Function

Private Sub Clear_Field(Start_Pos As Integer)

Dim i   As Integer

    For i = Start_Pos To Text_Max
        Text(i).Text = ""
    Next i

    For i = 0 To Label_Max
        Label1(i).Caption = ""
    Next i

    List1(plstZaiko).Clear


End Sub
Private Function MTS_Set_Proc() As Integer

Dim sts         As Integer
Dim com         As Integer
Dim Edit        As String


    MTS_Set_Proc = True
    
    com = BtOpGetFirst
    
    Combo(pcmbMUKE_CODE).Clear
    
    Do
        sts = BTRV(com, MTS_POS, MTSREC, Len(MTSREC), K1_MTS, Len(K1_MTS), 1)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "������}�X�^")
                MTS_Set_Proc = SYS_ERR
                Exit Function
        End Select
    
    
        Edit = StrConv(MTSREC.MUKE_DNAME, vbUnicode) & "   "
        Edit = Edit & StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode)
        Combo(pcmbMUKE_CODE).AddItem Edit
    
    
        com = BtOpGetNext
    Loop




    If Combo(pcmbMUKE_CODE).ListCount = 0 Then
    Else
        Combo(pcmbMUKE_CODE).ListIndex = 0
    End If


    MTS_Set_Proc = False
End Function

