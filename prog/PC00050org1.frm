VERSION 5.00
Begin VB.Form PC00050org1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "�i�ڃ}�X�^�R���o�[�g����"
   ClientHeight    =   7230
   ClientLeft      =   2325
   ClientTop       =   2625
   ClientWidth     =   9120
   ControlBox      =   0   'False
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
   ScaleHeight     =   7230
   ScaleWidth      =   9120
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CommandButton Command1 
      Caption         =   "�I��"
      Height          =   375
      Index           =   10
      Left            =   6960
      TabIndex        =   30
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�S��"
      Height          =   375
      Index           =   9
      Left            =   1920
      TabIndex        =   29
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   8
      Left            =   1920
      TabIndex        =   28
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   7
      Left            =   1920
      TabIndex        =   27
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   6
      Left            =   1920
      TabIndex        =   26
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   5
      Left            =   1920
      TabIndex        =   25
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   4
      Left            =   1920
      TabIndex        =   24
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   3
      Left            =   1920
      TabIndex        =   23
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   22
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   21
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   20
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '�E����
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   5160
      TabIndex        =   19
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�@���i�h�~�x�����O��"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   2760
      TabIndex        =   18
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '�E����
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   5160
      TabIndex        =   17
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�@�@�@�I�����f�[�^��"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   2760
      TabIndex        =   16
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '�E����
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   5160
      TabIndex        =   15
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�@�@�݌ɏW�v�f�[�^��"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   2760
      TabIndex        =   14
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '�E����
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   5160
      TabIndex        =   13
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "���׃`�F�b�N�f�[�^��"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   2760
      TabIndex        =   12
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '�E����
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   5160
      TabIndex        =   11
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�@�@�@�@�݌Ɉړ�����"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   2760
      TabIndex        =   10
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '�E����
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   5160
      TabIndex        =   9
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�@�@�ߓ����o�ח\�聁"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   8
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '�E����
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   5160
      TabIndex        =   7
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�@�@�@�@�@�o�ח\�聁"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   6
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '�E����
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�@�@�@�@�݌Ƀf�[�^��"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   4
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '�E����
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�@�@�@�@�i�ڃ}�X�^��"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   2
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Width           =   240
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "�f�[�^�R���o�[�g����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   4800
   End
End
Attribute VB_Name = "PC00050org1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Function Update_Proc(Mode As Integer) As Integer

Dim sts             As Integer
Dim Upd_com         As Integer
Dim com             As Integer
Dim ans             As Integer
Dim Count           As Long

Dim i               As Integer

Dim DISP_INTERVAL   As Long




    Update_Proc = True


    Select Case Mode
        Case 1              '�i�ڂ�
            GoTo ITEM_CONV
        Case 2              '�݌ɂ�
            GoTo ZAIKO_CONV
        Case 3              '�o�ח\���
            GoTo Y_SYU_CONV
        Case 4              '�ߓ����o�ח\���
            GoTo DEL_SYU_CONV
        Case 5              '�݌Ɉړ�����
            GoTo IDO_CONV
        Case 6              '����������
            GoTo J_NYU_CONV
        Case 7              '�݌ɏW�v��
            GoTo SUMZ_CONV
        Case 8              '�I������
            GoTo STOCK_CONV
        Case 9              '���i�h�~۸ނ�
            GoTo KEPPINLOG_CONV
    End Select


    
    '---------------------------------------------------------  �i�ڃ}�X�^����
    
    
ITEM_CONV:
    
    MsgLab(1) = "�i�ڃ}�X�^�R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(0).Caption = Format(Count, "#0")
                                        
                                '(��)�i�ڃ}�X�^�n�o�d�m
    If OLD_ITEM_Open(BtOpenNomal) Then
        If sts = BtErrFileNotFound Then
            If Mode = 0 Then
                GoTo ZAIKO_CONV
            Else
                MsgBox "�Ώۃf�[�^�Ȃ�"
                Unload Me
            End If
        End If
        Unload Me
    End If
                                        
                                        
                                        
    com = BtOpGetFirst
                                        
                                        
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_ITEM_POS, OLD_ITEMREC, Len(OLD_ITEMREC), K0_OLD_ITEM, Len(K0_OLD_ITEM), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "(��)�i�ڃ}�X�^")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(0).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
                                                
        Call UniCode_Conv(ITEMREC.JGYOBU, StrConv(OLD_ITEMREC.JGYOBU, vbUnicode))               '���ƕ�
        Call UniCode_Conv(ITEMREC.NAIGAI, StrConv(OLD_ITEMREC.NAIGAI, vbUnicode))               '�����O
        Call UniCode_Conv(ITEMREC.HIN_GAI, StrConv(OLD_ITEMREC.HIN_GAI, vbUnicode))             '�i��(�O)
        Call UniCode_Conv(ITEMREC.HIN_NAME, StrConv(OLD_ITEMREC.HIN_NAME, vbUnicode))           '�i��
        Call UniCode_Conv(ITEMREC.ST_SET_DT, StrConv(OLD_ITEMREC.ST_SET_DT, vbUnicode))         '�W���q�ɐݒ���t
        Call UniCode_Conv(ITEMREC.ST_SOKO, StrConv(OLD_ITEMREC.ST_SOKO, vbUnicode))             '�W���I�ԁ@�q��
        Call UniCode_Conv(ITEMREC.ST_RETU, StrConv(OLD_ITEMREC.ST_RETU, vbUnicode))             '�W���I�ԁ@��
        Call UniCode_Conv(ITEMREC.ST_REN, StrConv(OLD_ITEMREC.ST_REN, vbUnicode))               '�W���I�ԁ@�A
        Call UniCode_Conv(ITEMREC.ST_DAN, StrConv(OLD_ITEMREC.ST_DAN, vbUnicode))               '�W���I�ԁ@�i
        Call UniCode_Conv(ITEMREC.BEF_SOKO, StrConv(OLD_ITEMREC.BEF_SOKO, vbUnicode))           '�O��I�ԁ@�q��
        Call UniCode_Conv(ITEMREC.BEF_RETU, StrConv(OLD_ITEMREC.BEF_RETU, vbUnicode))           '�O��I�ԁ@��
        Call UniCode_Conv(ITEMREC.BEF_REN, StrConv(OLD_ITEMREC.BEF_REN, vbUnicode))             '�O��I�ԁ@�A
        Call UniCode_Conv(ITEMREC.BEF_DAN, StrConv(OLD_ITEMREC.BEF_DAN, vbUnicode))             '�O��I�ԁ@�i
                                                                                                '�ŏI���ɓ�
        Call UniCode_Conv(ITEMREC.LAST_NYU_DT, StrConv(OLD_ITEMREC.LAST_NYU_DT, vbUnicode))
                                                                                                '�ŏI�o�ɓ�
        Call UniCode_Conv(ITEMREC.LAST_SYU_DT, StrConv(OLD_ITEMREC.LAST_SYU_DT, vbUnicode))
        Call UniCode_Conv(ITEMREC.HIN_NAI, StrConv(OLD_ITEMREC.HIN_NAI, vbUnicode))             '�i��(��)
        Call UniCode_Conv(ITEMREC.BIKOU_SOKO, StrConv(OLD_ITEMREC.BIKOU_SOKO, vbUnicode))       '���l�@νđq��
        Call UniCode_Conv(ITEMREC.BIKOU_TANA, StrConv(OLD_ITEMREC.BIKOU_TANA, vbUnicode))       '���l�@νĒI��
        Call UniCode_Conv(ITEMREC.HOJYU_P, StrConv(OLD_ITEMREC.HOJYU_P, vbUnicode))             '��[�_
        Call UniCode_Conv(ITEMREC.AVE_SYUKA, StrConv(OLD_ITEMREC.AVE_SYUKA, vbUnicode))         '�����Ϗo�א�
                
        Call UniCode_Conv(ITEMREC.SAMPLE_QTY, StrConv(OLD_ITEMREC.SAMPLE_QTY, vbUnicode))       '����ِ�
        Call UniCode_Conv(ITEMREC.LAST_INP_DT, StrConv(OLD_ITEMREC.LAST_INP_DT, vbUnicode))     '�ŏI���ד��t
        
        Call UniCode_Conv(ITEMREC.LAST_CHK_DT, StrConv(OLD_ITEMREC.LAST_CHK_DT, vbUnicode))     '�ŏI�ƍ����t
        Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, StrConv(OLD_ITEMREC.LAST_CHK_QTY, vbUnicode))   '�ŏI�ƍ����݌ɐ�
        
        Call UniCode_Conv(ITEMREC.BIKOU, StrConv(OLD_ITEMREC.BIKOU, vbUnicode))                 '������l
        Call UniCode_Conv(ITEMREC.IRI_QTY, StrConv(OLD_ITEMREC.IRI_QTY, vbUnicode))             '������萔
        Call UniCode_Conv(ITEMREC.JAN_CODE, StrConv(OLD_ITEMREC.JAN_CODE, vbUnicode))           'JAN�R�[�h
        Call UniCode_Conv(ITEMREC.HIN_CHANGE, StrConv(OLD_ITEMREC.HIN_CHANGE, vbUnicode))       '�i�ԓǂݑւ�
        Call UniCode_Conv(ITEMREC.GOODS_KBN, StrConv(OLD_ITEMREC.GOODS_KBN, vbUnicode))         '���i���L��
        Call UniCode_Conv(ITEMREC.PACKING_NO, StrConv(OLD_ITEMREC.PACKING_NO, vbUnicode))       '������
        Call UniCode_Conv(ITEMREC.RANK, StrConv(OLD_ITEMREC.RANK, vbUnicode))                   '�����ݸ
        Call UniCode_Conv(ITEMREC.NEW_RANK, StrConv(OLD_ITEMREC.NEW_RANK, vbUnicode))           '�V�ݸ
        Call UniCode_Conv(ITEMREC.GLICS1_TANA, StrConv(OLD_ITEMREC.GLICS1_TANA, vbUnicode))     '��د���I��1
        Call UniCode_Conv(ITEMREC.GLICS2_TANA, StrConv(OLD_ITEMREC.GLICS2_TANA, vbUnicode))     '��د���I��2
        Call UniCode_Conv(ITEMREC.GLICS3_TANA, StrConv(OLD_ITEMREC.GLICS3_TANA, vbUnicode))     '��د���I��3
        
        
        
        Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, "")                                 '�Ɩ��Ǘ��@ �d���敪
        Call UniCode_Conv(ITEMREC.G_HANBAI_KBN, "")                                 '           �̔��敪
        Call UniCode_Conv(ITEMREC.G_SYUSHI, "")                                     '           ���x�P��
        Call UniCode_Conv(ITEMREC.G_KUMITATE, "")                                   '           �g�����i
        Call UniCode_Conv(ITEMREC.G_ST_URITAN, "")                                  '           �W���e�������P���@9(8)V99
        Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, "")                               '           �W���e�������ݒ��
        Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "")                                  '           �W���e�������P��  9(8)V99
        Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, "")                               '           �W���e�������ݒ��
        
        For i = 0 To 2                                                              '�d������
            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).CODE, "")                     '           �d����R�[�h
            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).TANKA, "")                    '           �P��
            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).TANKA_DT, "")                 '           �P���ݒ��
            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LOT, "")                      '           �P���ݒ��
            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LEAD_TIME, "")                '           ���[�h�^�C��
            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, "")            '           �ŏI������
            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_QTY, "")           '           �ŏI������
        
        Next i
    
        Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, "")                              '           �O���݌ɋ��z
        Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, "")                                 '           ���ދ敪
        Call UniCode_Conv(ITEMREC.G_LABEL_NON, "")                                  '           ���ٓ\��t��
                
        
        Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, "")                                 '�i��E
        Call UniCode_Conv(ITEMREC.L_BIKOU, "")                                      '���l
        Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, "")                                '��Ж�
        Call UniCode_Conv(ITEMREC.L_KISHU1, "")                                     '�@��(1)
        Call UniCode_Conv(ITEMREC.L_KISHU2, "")                                     '�@��(2)
        Call UniCode_Conv(ITEMREC.L_KISHU3, "")                                     '�@��(3)
        Call UniCode_Conv(ITEMREC.L_PAPER, "")                                      '��
        Call UniCode_Conv(ITEMREC.L_PLASTIC, "")                                    '��׽���
        Call UniCode_Conv(ITEMREC.L_URIKIN1, "")                                    '���i(1)
        Call UniCode_Conv(ITEMREC.L_URIKIN2, "")                                    '���i(2)
        Call UniCode_Conv(ITEMREC.L_URIKIN3, "")                                    '���i(3)
        Call UniCode_Conv(ITEMREC.L_LABEL, "")                                      '�K�p�@������
        Call UniCode_Conv(ITEMREC.L_MAISU, "")                                      '���ٖ���
        Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, "")                                '�K�p�@����l
        Call UniCode_Conv(ITEMREC.L_SAGYO_SHIJI, "")                                '��Ǝw��
        Call UniCode_Conv(ITEMREC.L_BIKOU3, "")                                     '���l(3)
        Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, "")                                '���ƕ���
        Call UniCode_Conv(ITEMREC.L_IRI_QTY, "")                                    '���萔
        Call UniCode_Conv(ITEMREC.L_TANA1, "")                                      '�I��(1)
        Call UniCode_Conv(ITEMREC.L_TANA2, "")                                      '�I��(2)
        
        
        
        Call UniCode_Conv(ITEMREC.S_TANTO, "")                                      '���P�^�S����
        Call UniCode_Conv(ITEMREC.ZAIKO_F, P_ZAIKO_F_ON)                            '�݌ɊǗ��Ώ�
        
        Do
            sts = BTRV(BtOpInsert, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "�i�ڃ}�X�^")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

'---------------------------------------------  �I��
    Cnt(0).Caption = Format(Count, "#0")

    If Mode <> 0 Then
        Me.MousePointer = vbDefault
        
        MsgBox "�R���o�[�g�I��"
        Update_Proc = False
        Exit Function
    End If
    '---------------------------------------------------------  �݌Ƀf�[�^�̏���

ZAIKO_CONV:

    MsgLab(1) = "�݌Ƀf�[�^�R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(1).Caption = Format(Count, "#0")
                                        
                                '(��)�݌Ƀf�[�^�n�o�d�m
    If OLD_ZAIKO_Open(BtOpenNomal) Then
        If sts = BtErrFileNotFound Then
            If Mode = 0 Then
                GoTo Y_SYU_CONV
            Else
                MsgBox "�Ώۃf�[�^�Ȃ�"
                Unload Me
            End If
        End If
        Unload Me
    End If
                                        
                                        
    com = BtOpGetFirst
                                        
                                        
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_ZAIKO_POS, OLD_ZAIKOREC, Len(OLD_ZAIKOREC), K0_OLD_ZAIKO, Len(K0_OLD_ZAIKO), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "(��)�݌Ƀf�[�^")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(1).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        Call UniCode_Conv(ZAIKOREC.Soko_No, StrConv(OLD_ZAIKOREC.Soko_No, vbUnicode))       '�q�ɇ�
        Call UniCode_Conv(ZAIKOREC.Retu, StrConv(OLD_ZAIKOREC.Retu, vbUnicode))             '��
        Call UniCode_Conv(ZAIKOREC.Ren, StrConv(OLD_ZAIKOREC.Ren, vbUnicode))               '�A
        Call UniCode_Conv(ZAIKOREC.Dan, StrConv(OLD_ZAIKOREC.Dan, vbUnicode))               '�i
                                                
        Call UniCode_Conv(ZAIKOREC.JGYOBU, StrConv(OLD_ZAIKOREC.JGYOBU, vbUnicode))         '���ƕ�����
        Call UniCode_Conv(ZAIKOREC.NAIGAI, StrConv(OLD_ZAIKOREC.NAIGAI, vbUnicode))         '�����O
        Call UniCode_Conv(ZAIKOREC.HIN_GAI, StrConv(OLD_ZAIKOREC.HIN_GAI, vbUnicode))       '�i�ԁi�O���j
        
        Call UniCode_Conv(ZAIKOREC.GOODS_ON, StrConv(OLD_ZAIKOREC.GOODS_ON, vbUnicode))     '���i�^�����i
        Call UniCode_Conv(ZAIKOREC.NYUKA_DT, StrConv(OLD_ZAIKOREC.NYUKA_DT, vbUnicode))     '���ד�
        Call UniCode_Conv(ZAIKOREC.NYUKO_DT, StrConv(OLD_ZAIKOREC.NYUKO_DT, vbUnicode))     '���ɓ�
        Call UniCode_Conv(ZAIKOREC.HIN_NAI, StrConv(OLD_ZAIKOREC.HIN_NAI, vbUnicode))       '�i��(����)
        
        Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, StrConv(OLD_ZAIKOREC.YUKO_Z_QTY, vbUnicode)) '�L���݌�
        
        Call UniCode_Conv(ZAIKOREC.LOCK_F, "")                                              '�r���t���O
        Call UniCode_Conv(ZAIKOREC.WEL_ID, "")                                              '�g�p���[��
        Call UniCode_Conv(ZAIKOREC.PRG_ID, "")                                              '�g�p���v���O����
        
        Call UniCode_Conv(ZAIKOREC.GOODS_YMD, StrConv(OLD_ZAIKOREC.GOODS_YMD, vbUnicode))   '���i�����t
        
        Call UniCode_Conv(ZAIKOREC.SHIIRE_CODE, "")                                         '�d���溰��
        Call UniCode_Conv(ZAIKOREC.SHIIRE_TANKA, "")                                        '�d���P��
        Call UniCode_Conv(ZAIKOREC.KEIJYO_YM, "")                                           '�v��N��
                
        Call UniCode_Conv(ZAIKOREC.FILLER, "")
        
        Do
            sts = BTRV(BtOpInsert, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "�݌Ƀf�[�^")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

'---------------------------------------------  �I��
    Cnt(1).Caption = Format(Count, "#0")
    
    If Mode <> 0 Then
        Me.MousePointer = vbDefault
        MsgBox "�R���o�[�g�I��"
        Update_Proc = False
        Exit Function
    End If

    '---------------------------------------------------------  �o�ח\��f�[�^�̏���
Y_SYU_CONV:


    MsgLab(1) = "�o�ח\��f�[�^�R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(2).Caption = Format(Count, "#0")
                                        
                                        
                                '(��)�o�ח\��f�[�^�n�o�d�m
    If OLD_Y_SYU_Open(BtOpenNomal) Then
        If sts = BtErrFileNotFound Then
            If Mode = 0 Then
                GoTo DEL_SYU_CONV
            Else
                MsgBox "�Ώۃf�[�^�Ȃ�"
                Unload Me
            End If
        End If
        Unload Me
    End If
                                        
                                        
    com = BtOpGetFirst
                                        
                                        
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_Y_SYU_POS, OLD_Y_SYUREC, Len(OLD_Y_SYUREC), K0_OLD_Y_SYU, Len(K0_OLD_Y_SYU), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "(��)�o�ח\��f�[�^")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(2).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        
        
        
        
        
        Call UniCode_Conv(Y_SYUREC.WEL_ID, "")                                                      '�g�p�[��ID
        Call UniCode_Conv(Y_SYUREC.PRG_ID, "")                                                      '�g�p����۸���ID
        
        Call UniCode_Conv(Y_SYUREC.KAN_KBN, StrConv(OLD_Y_SYUREC.KAN_KBN, vbUnicode))               '�����敪
        Call UniCode_Conv(Y_SYUREC.DT_SYU, StrConv(OLD_Y_SYUREC.DT_SYU, vbUnicode))                 '�ް����
        Call UniCode_Conv(Y_SYUREC.JGYOBU, StrConv(OLD_Y_SYUREC.JGYOBU, vbUnicode))                 '���ƕ�
        
        Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, StrConv(OLD_Y_SYUREC.KEY_CYU_KBN, vbUnicode))       '�����敪
        Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, StrConv(OLD_Y_SYUREC.KEY_ID_NO, vbUnicode))           'ID-NO
        
        Call UniCode_Conv(Y_SYUREC.NAIGAI, StrConv(OLD_Y_SYUREC.NAIGAI, vbUnicode))                 '�����O
        Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, StrConv(OLD_Y_SYUREC.KEY_HIN_NO, vbUnicode))         '�i�ڔԍ�
        
        Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, StrConv(OLD_Y_SYUREC.KEY_MUKE_CODE, vbUnicode))   '���Ӑ溰��
        Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, StrConv(OLD_Y_SYUREC.KEY_SS_CODE, vbUnicode))       '�����溰��
        
        Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, StrConv(OLD_Y_SYUREC.KEY_SYUKA_YMD, vbUnicode))   '�o�ד��t
        
        Call UniCode_Conv(Y_SYUREC.JGYOBA, StrConv(OLD_Y_SYUREC.JGYOBA, vbUnicode))                 '���Ə�
        Call UniCode_Conv(Y_SYUREC.DATA_KBN, StrConv(OLD_Y_SYUREC.DATA_KBN, vbUnicode))             '�ް��敪
        Call UniCode_Conv(Y_SYUREC.TORI_KBN, StrConv(OLD_Y_SYUREC.TORI_KBN, vbUnicode))             '����敪
        Call UniCode_Conv(Y_SYUREC.ID_NO, StrConv(OLD_Y_SYUREC.ID_NO, vbUnicode))                   'ID-NO
        Call UniCode_Conv(Y_SYUREC.HIN_NO, StrConv(OLD_Y_SYUREC.HIN_NO, vbUnicode))                 '�i�ڔԍ�
        
        Call UniCode_Conv(Y_SYUREC.DEN_NO, StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode))                 '�`�[�ԍ�
        Call UniCode_Conv(Y_SYUREC.SURYO, StrConv(OLD_Y_SYUREC.SURYO, vbUnicode))                   '����
        Call UniCode_Conv(Y_SYUREC.MUKE_CODE, StrConv(OLD_Y_SYUREC.MUKE_CODE, vbUnicode))           '���Ӑ溰��
        Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, StrConv(OLD_Y_SYUREC.SYUKO_SYUSI, vbUnicode))       '�o�Ɏ��x
        Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, StrConv(OLD_Y_SYUREC.SYUKA_YMD, vbUnicode))           '�o�ד��t
        Call UniCode_Conv(Y_SYUREC.ODER_NO, StrConv(OLD_Y_SYUREC.ODER_NO, vbUnicode))               '���ް�ԍ�
        Call UniCode_Conv(Y_SYUREC.ITEM_NO, StrConv(OLD_Y_SYUREC.ITEM_NO, vbUnicode))               '���єԍ�
        Call UniCode_Conv(Y_SYUREC.MUKE_NAME, StrConv(OLD_Y_SYUREC.MUKE_NAME, vbUnicode))           '���Ӑ於��
        
        Call UniCode_Conv(Y_SYUREC.CYU_KBN, StrConv(OLD_Y_SYUREC.CYU_KBN, vbUnicode))               '�����敪
        Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, StrConv(OLD_Y_SYUREC.CYU_KBN_NAME, vbUnicode))     '�����敪����
        Call UniCode_Conv(Y_SYUREC.EXPORT_KBN, StrConv(OLD_Y_SYUREC.EXPORT_KBN, vbUnicode))         '�A�o�o�׌����敪
        Call UniCode_Conv(Y_SYUREC.LABEL_ISSUE_KBN, StrConv(OLD_Y_SYUREC.LABEL_ISSUE_KBN, vbUnicode))   '�����x�����s�敪
        Call UniCode_Conv(Y_SYUREC.LABEL_ISSUE_UNIT, StrConv(OLD_Y_SYUREC.LABEL_ISSUE_UNIT, vbUnicode)) '�����x�����s�P�ʐ�
        Call UniCode_Conv(Y_SYUREC.LABEL_TANKA_KBN, StrConv(OLD_Y_SYUREC.LABEL_TANKA_KBN, vbUnicode))   '�����x���P���\���敪
        Call UniCode_Conv(Y_SYUREC.TANKA, StrConv(OLD_Y_SYUREC.TANKA, vbUnicode))                   '�P��
        Call UniCode_Conv(Y_SYUREC.KINGAKU, StrConv(OLD_Y_SYUREC.KINGAKU, vbUnicode))               '���z
        
        Call UniCode_Conv(Y_SYUREC.BIKOU2, StrConv(OLD_Y_SYUREC.BIKOU2, vbUnicode))                 '���l�Q
        Call UniCode_Conv(Y_SYUREC.REBATE_KBN, StrConv(OLD_Y_SYUREC.REBATE_KBN, vbUnicode))         '��ްċ敪
        Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, StrConv(OLD_Y_SYUREC.CHOHA_KBN, vbUnicode))           '���[�敪
        Call UniCode_Conv(Y_SYUREC.ATAISA_KBN, StrConv(OLD_Y_SYUREC.ATAISA_KBN, vbUnicode))         '�l���敪
        Call UniCode_Conv(Y_SYUREC.REP_KISHU, StrConv(OLD_Y_SYUREC.REP_KISHU, vbUnicode))           '��\�@��
        Call UniCode_Conv(Y_SYUREC.NS_KANRI_NO, StrConv(OLD_Y_SYUREC.NS_KANRI_NO, vbUnicode))       'NS�Ǘ��敪
        Call UniCode_Conv(Y_SYUREC.MTS_HIN_CODE, StrConv(OLD_Y_SYUREC.MTS_HIN_CODE, vbUnicode))     'MTS���i����
        Call UniCode_Conv(Y_SYUREC.BIKOU1, StrConv(OLD_Y_SYUREC.BIKOU1, vbUnicode))                 '���l1
        Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, StrConv(OLD_Y_SYUREC.CHOKU_KBN, vbUnicode))           '�����敪
        Call UniCode_Conv(Y_SYUREC.REBATE_RATE, StrConv(OLD_Y_SYUREC.REBATE_RATE, vbUnicode))       '��ްė�
        Call UniCode_Conv(Y_SYUREC.HIN_NAME, StrConv(OLD_Y_SYUREC.HIN_NAME, vbUnicode))             '�i��
        Call UniCode_Conv(Y_SYUREC.JGYOBA_GAI, StrConv(OLD_Y_SYUREC.JGYOBA_GAI, vbUnicode))         '�ΊO���Ə�
        Call UniCode_Conv(Y_SYUREC.KISHU_CODE, StrConv(OLD_Y_SYUREC.KISHU_CODE, vbUnicode))         '�@����
        Call UniCode_Conv(Y_SYUREC.SS_CODE, StrConv(OLD_Y_SYUREC.SS_CODE, vbUnicode))               '�����溰��
        Call UniCode_Conv(Y_SYUREC.HIN_NAI, StrConv(OLD_Y_SYUREC.HIN_NAI, vbUnicode))               '�i��(����)
        Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(OLD_Y_SYUREC.HTANABAN, vbUnicode))             'νĒI��
        
        Call UniCode_Conv(Y_SYUREC.PRINT_YMD, StrConv(OLD_Y_SYUREC.PRINT_YMD, vbUnicode))           '�o�ɕ\������t
        Call UniCode_Conv(Y_SYUREC.KAN_YMD, StrConv(OLD_Y_SYUREC.KAN_YMD, vbUnicode))               '�������t
        Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, StrConv(OLD_Y_SYUREC.KENPIN_YMD, vbUnicode))         '���i���t
        Call UniCode_Conv(Y_SYUREC.TOK_KBN, StrConv(OLD_Y_SYUREC.TOK_KBN, vbUnicode))               '������敪
        Call UniCode_Conv(Y_SYUREC.JITU_SURYO, StrConv(OLD_Y_SYUREC.JITU_SURYO, vbUnicode))         '�o�Ɏ��ѐ���
        Call UniCode_Conv(Y_SYUREC.INS_NOW, StrConv(OLD_Y_SYUREC.INS_NOW, vbUnicode))               '��荞�ݓ���
        
        
        
        Call UniCode_Conv(Y_SYUREC.FILLER, "")
        
        
        Do
            sts = BTRV(BtOpInsert, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYU.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "�o�ח\��f�[�^")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

'---------------------------------------------  �I��
    Cnt(2).Caption = Format(Count, "#0")
    If Mode <> 0 Then
        Me.MousePointer = vbDefault
        MsgBox "�R���o�[�g�I��"
        Update_Proc = False
        Exit Function
    End If


    '---------------------------------------------------------  �ߓ����o�ח\��f�[�^�̏���
DEL_SYU_CONV:


    MsgLab(1) = "�ߓ����o�ח\��f�[�^�R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(3).Caption = Format(Count, "#0")
                                        
                                '(��)�ߓ����o�ח\��f�[�^�n�o�d�m
    If OLD_DEL_SYU_Open(BtOpenNomal) Then
        If sts = BtErrFileNotFound Then
            If Mode = 0 Then
                GoTo IDO_CONV
            Else
                MsgBox "�Ώۃf�[�^�Ȃ�"
                Unload Me
            End If
        End If
        Unload Me
    End If
                                        
                                        
    com = BtOpGetFirst
                                        
                                        
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_DEL_SYU_POS, OLD_DEL_SYUREC, Len(OLD_DEL_SYUREC), K0_OLD_DEL_SYU, Len(K0_OLD_DEL_SYU), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "(��)�ߓ����o�ח\��f�[�^")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(3).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        
        
        
        
        
        Call UniCode_Conv(DEL_SYUREC.WEL_ID, "")                                                      '�g�p�[��ID
        Call UniCode_Conv(DEL_SYUREC.PRG_ID, "")                                                      '�g�p����۸���ID
        
        Call UniCode_Conv(DEL_SYUREC.KAN_KBN, StrConv(OLD_DEL_SYUREC.KAN_KBN, vbUnicode))               '�����敪
        Call UniCode_Conv(DEL_SYUREC.DT_SYU, StrConv(OLD_DEL_SYUREC.DT_SYU, vbUnicode))                 '�ް����
        Call UniCode_Conv(DEL_SYUREC.JGYOBU, StrConv(OLD_DEL_SYUREC.JGYOBU, vbUnicode))                 '���ƕ�
        
        Call UniCode_Conv(DEL_SYUREC.KEY_CYU_KBN, StrConv(OLD_DEL_SYUREC.KEY_CYU_KBN, vbUnicode))       '�����敪
        Call UniCode_Conv(DEL_SYUREC.KEY_ID_NO, StrConv(OLD_DEL_SYUREC.KEY_ID_NO, vbUnicode))           'ID-NO
        
        Call UniCode_Conv(DEL_SYUREC.NAIGAI, StrConv(OLD_DEL_SYUREC.NAIGAI, vbUnicode))                 '�����O
        Call UniCode_Conv(DEL_SYUREC.KEY_HIN_NO, StrConv(OLD_DEL_SYUREC.KEY_HIN_NO, vbUnicode))         '�i�ڔԍ�
        
        Call UniCode_Conv(DEL_SYUREC.KEY_MUKE_CODE, StrConv(OLD_DEL_SYUREC.KEY_MUKE_CODE, vbUnicode))   '���Ӑ溰��
        Call UniCode_Conv(DEL_SYUREC.KEY_SS_CODE, StrConv(OLD_DEL_SYUREC.KEY_SS_CODE, vbUnicode))       '�����溰��
        
        Call UniCode_Conv(DEL_SYUREC.KEY_SYUKA_YMD, StrConv(OLD_DEL_SYUREC.KEY_SYUKA_YMD, vbUnicode))   '�o�ד��t
        
        Call UniCode_Conv(DEL_SYUREC.JGYOBA, StrConv(OLD_DEL_SYUREC.JGYOBA, vbUnicode))                 '���Ə�
        Call UniCode_Conv(DEL_SYUREC.DATA_KBN, StrConv(OLD_DEL_SYUREC.DATA_KBN, vbUnicode))             '�ް��敪
        Call UniCode_Conv(DEL_SYUREC.TORI_KBN, StrConv(OLD_DEL_SYUREC.TORI_KBN, vbUnicode))             '����敪
        Call UniCode_Conv(DEL_SYUREC.ID_NO, StrConv(OLD_DEL_SYUREC.ID_NO, vbUnicode))                   'ID-NO
        Call UniCode_Conv(DEL_SYUREC.HIN_NO, StrConv(OLD_DEL_SYUREC.HIN_NO, vbUnicode))                 '�i�ڔԍ�
        
        Call UniCode_Conv(DEL_SYUREC.DEN_NO, StrConv(OLD_DEL_SYUREC.DEN_NO, vbUnicode))                 '�`�[�ԍ�
        Call UniCode_Conv(DEL_SYUREC.SURYO, StrConv(OLD_DEL_SYUREC.SURYO, vbUnicode))                   '����
        Call UniCode_Conv(DEL_SYUREC.MUKE_CODE, StrConv(OLD_DEL_SYUREC.MUKE_CODE, vbUnicode))           '���Ӑ溰��
        Call UniCode_Conv(DEL_SYUREC.SYUKO_SYUSI, StrConv(OLD_DEL_SYUREC.SYUKO_SYUSI, vbUnicode))       '�o�Ɏ��x
        Call UniCode_Conv(DEL_SYUREC.SYUKA_YMD, StrConv(OLD_DEL_SYUREC.SYUKA_YMD, vbUnicode))           '�o�ד��t
        Call UniCode_Conv(DEL_SYUREC.ODER_NO, StrConv(OLD_DEL_SYUREC.ODER_NO, vbUnicode))               '���ް�ԍ�
        Call UniCode_Conv(DEL_SYUREC.ITEM_NO, StrConv(OLD_DEL_SYUREC.ITEM_NO, vbUnicode))               '���єԍ�
        Call UniCode_Conv(DEL_SYUREC.MUKE_NAME, StrConv(OLD_DEL_SYUREC.MUKE_NAME, vbUnicode))           '���Ӑ於��
        
        Call UniCode_Conv(DEL_SYUREC.CYU_KBN, StrConv(OLD_DEL_SYUREC.CYU_KBN, vbUnicode))               '�����敪
        Call UniCode_Conv(DEL_SYUREC.CYU_KBN_NAME, StrConv(OLD_DEL_SYUREC.CYU_KBN_NAME, vbUnicode))     '�����敪����
        Call UniCode_Conv(DEL_SYUREC.EXPORT_KBN, StrConv(OLD_DEL_SYUREC.EXPORT_KBN, vbUnicode))         '�A�o�o�׌����敪
        Call UniCode_Conv(DEL_SYUREC.LABEL_ISSUE_KBN, StrConv(OLD_DEL_SYUREC.LABEL_ISSUE_KBN, vbUnicode))   '�����x�����s�敪
        Call UniCode_Conv(DEL_SYUREC.LABEL_ISSUE_UNIT, StrConv(OLD_DEL_SYUREC.LABEL_ISSUE_UNIT, vbUnicode)) '�����x�����s�P�ʐ�
        Call UniCode_Conv(DEL_SYUREC.LABEL_TANKA_KBN, StrConv(OLD_DEL_SYUREC.LABEL_TANKA_KBN, vbUnicode))   '�����x���P���\���敪
        Call UniCode_Conv(DEL_SYUREC.TANKA, StrConv(OLD_DEL_SYUREC.TANKA, vbUnicode))                   '�P��
        Call UniCode_Conv(DEL_SYUREC.KINGAKU, StrConv(OLD_DEL_SYUREC.KINGAKU, vbUnicode))               '���z
        
        Call UniCode_Conv(DEL_SYUREC.BIKOU2, StrConv(OLD_DEL_SYUREC.BIKOU2, vbUnicode))                 '���l�Q
        Call UniCode_Conv(DEL_SYUREC.REBATE_KBN, StrConv(OLD_DEL_SYUREC.REBATE_KBN, vbUnicode))         '��ްċ敪
        Call UniCode_Conv(DEL_SYUREC.CHOHA_KBN, StrConv(OLD_DEL_SYUREC.CHOHA_KBN, vbUnicode))           '���[�敪
        Call UniCode_Conv(DEL_SYUREC.ATAISA_KBN, StrConv(OLD_DEL_SYUREC.ATAISA_KBN, vbUnicode))         '�l���敪
        Call UniCode_Conv(DEL_SYUREC.REP_KISHU, StrConv(OLD_DEL_SYUREC.REP_KISHU, vbUnicode))           '��\�@��
        Call UniCode_Conv(DEL_SYUREC.NS_KANRI_NO, StrConv(OLD_DEL_SYUREC.NS_KANRI_NO, vbUnicode))       'NS�Ǘ��敪
        Call UniCode_Conv(DEL_SYUREC.MTS_HIN_CODE, StrConv(OLD_DEL_SYUREC.MTS_HIN_CODE, vbUnicode))     'MTS���i����
        Call UniCode_Conv(DEL_SYUREC.BIKOU1, StrConv(OLD_DEL_SYUREC.BIKOU1, vbUnicode))                 '���l1
        Call UniCode_Conv(DEL_SYUREC.CHOKU_KBN, StrConv(OLD_DEL_SYUREC.CHOKU_KBN, vbUnicode))           '�����敪
        Call UniCode_Conv(DEL_SYUREC.REBATE_RATE, StrConv(OLD_DEL_SYUREC.REBATE_RATE, vbUnicode))       '��ްė�
        Call UniCode_Conv(DEL_SYUREC.HIN_NAME, StrConv(OLD_DEL_SYUREC.HIN_NAME, vbUnicode))             '�i��
        Call UniCode_Conv(DEL_SYUREC.JGYOBA_GAI, StrConv(OLD_DEL_SYUREC.JGYOBA_GAI, vbUnicode))         '�ΊO���Ə�
        Call UniCode_Conv(DEL_SYUREC.KISHU_CODE, StrConv(OLD_DEL_SYUREC.KISHU_CODE, vbUnicode))         '�@����
        Call UniCode_Conv(DEL_SYUREC.SS_CODE, StrConv(OLD_DEL_SYUREC.SS_CODE, vbUnicode))               '�����溰��
        Call UniCode_Conv(DEL_SYUREC.HIN_NAI, StrConv(OLD_DEL_SYUREC.HIN_NAI, vbUnicode))               '�i��(����)
        Call UniCode_Conv(DEL_SYUREC.HTANABAN, StrConv(OLD_DEL_SYUREC.HTANABAN, vbUnicode))             'νĒI��
        
        Call UniCode_Conv(DEL_SYUREC.PRINT_YMD, StrConv(OLD_DEL_SYUREC.PRINT_YMD, vbUnicode))           '�o�ɕ\������t
        Call UniCode_Conv(DEL_SYUREC.KAN_YMD, StrConv(OLD_DEL_SYUREC.KAN_YMD, vbUnicode))               '�������t
        Call UniCode_Conv(DEL_SYUREC.KENPIN_YMD, StrConv(OLD_DEL_SYUREC.KENPIN_YMD, vbUnicode))         '���i���t
        Call UniCode_Conv(DEL_SYUREC.TOK_KBN, StrConv(OLD_DEL_SYUREC.TOK_KBN, vbUnicode))               '������敪
        Call UniCode_Conv(DEL_SYUREC.JITU_SURYO, StrConv(OLD_DEL_SYUREC.JITU_SURYO, vbUnicode))         '�o�Ɏ��ѐ���
        Call UniCode_Conv(DEL_SYUREC.INS_NOW, StrConv(OLD_DEL_SYUREC.INS_NOW, vbUnicode))               '��荞�ݓ���
        
        
        
        Call UniCode_Conv(DEL_SYUREC.FILLER, "")
        
        
        Do
            sts = BTRV(BtOpInsert, DEL_SYU_POS, DEL_SYUREC, Len(DEL_SYUREC), K0_DEL_SYU, Len(K0_DEL_SYU), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<DEL_SYU.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "�ߓ����o�ח\��f�[�^")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

'---------------------------------------------  �I��
    Cnt(3).Caption = Format(Count, "#0")
    If Mode <> 0 Then
        Me.MousePointer = vbDefault
        MsgBox "�R���o�[�g�I��"
        Update_Proc = False
        Exit Function
    End If



    '---------------------------------------------------------  �݌Ɉړ����f�[�^�̏���
IDO_CONV:


    MsgLab(1) = "�݌Ɉړ����f�[�^�R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(4).Caption = Format(Count, "#0")
                                '(��)�݌Ɉړ����f�[�^�n�o�d�m
    If OLD_IDO_Open(BtOpenNomal) Then
        If sts = BtErrFileNotFound Then
            If Mode = 0 Then
                GoTo J_NYU_CONV
            Else
                MsgBox "�Ώۃf�[�^�Ȃ�"
                Unload Me
            End If
        End If
        Unload Me
    End If
                                        
    com = BtOpGetFirst
                                        
                                        
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_IDO_POS, OLD_IDOREC, Len(OLD_IDOREC), K0_OLD_IDO, Len(K0_OLD_IDO), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "(��)�݌Ɉړ����f�[�^")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(4).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        
        
        
        
        
        
        Call UniCode_Conv(IDOREC.JITU_DT, StrConv(OLD_IDOREC.JITU_DT, vbUnicode))           '���ѓ��t
        Call UniCode_Conv(IDOREC.JITU_TM, StrConv(OLD_IDOREC.JITU_TM, vbUnicode))           '���ю���
        Call UniCode_Conv(IDOREC.JGYOBU, StrConv(OLD_IDOREC.JGYOBU, vbUnicode))             '���ƕ��敪
        Call UniCode_Conv(IDOREC.NAIGAI, StrConv(OLD_IDOREC.NAIGAI, vbUnicode))             '�����O
        Call UniCode_Conv(IDOREC.HIN_GAI, StrConv(OLD_IDOREC.HIN_GAI, vbUnicode))           '�i��(�O��)
        Call UniCode_Conv(IDOREC.RIRK_ID, StrConv(OLD_IDOREC.RIRK_ID, vbUnicode))           '�������
        Call UniCode_Conv(IDOREC.SUMI_JITU_QTY, StrConv(OLD_IDOREC.SUMI_JITU_QTY, vbUnicode))   '���ѐ���(���i���ς�)
        Call UniCode_Conv(IDOREC.MI_JITU_QTY, StrConv(OLD_IDOREC.MI_JITU_QTY, vbUnicode))   '���ѐ���(�����i)
        Call UniCode_Conv(IDOREC.FROM_SOKO, StrConv(OLD_IDOREC.FROM_SOKO, vbUnicode))       'From �q�ɇ�
        Call UniCode_Conv(IDOREC.FROM_RETU, StrConv(OLD_IDOREC.FROM_RETU, vbUnicode))       'From ��
        Call UniCode_Conv(IDOREC.FROM_REN, StrConv(OLD_IDOREC.FROM_REN, vbUnicode))         'From �A
        Call UniCode_Conv(IDOREC.FROM_DAN, StrConv(OLD_IDOREC.FROM_DAN, vbUnicode))         'From �i
        
        Call UniCode_Conv(IDOREC.TO_SOKO, StrConv(OLD_IDOREC.TO_SOKO, vbUnicode))           'To �q�ɇ�
        Call UniCode_Conv(IDOREC.TO_RETU, StrConv(OLD_IDOREC.TO_RETU, vbUnicode))           'To ��
        Call UniCode_Conv(IDOREC.TO_REN, StrConv(OLD_IDOREC.TO_REN, vbUnicode))             'To �A
        Call UniCode_Conv(IDOREC.TO_DAN, StrConv(OLD_IDOREC.TO_DAN, vbUnicode))             'To �i
        
        Call UniCode_Conv(IDOREC.DEN_DT, StrConv(OLD_IDOREC.DEN_DT, vbUnicode))             '�`�[���t
        Call UniCode_Conv(IDOREC.DEN_NO, StrConv(OLD_IDOREC.DEN_NO, vbUnicode))             '�`�[��
        Call UniCode_Conv(IDOREC.PRG_ID, StrConv(OLD_IDOREC.PRG_ID, vbUnicode))             '�o�͌���۸���
        
        Call UniCode_Conv(IDOREC.HIN_NAI, StrConv(OLD_IDOREC.HIN_NAI, vbUnicode))           '�i��(����)
        
        Call UniCode_Conv(IDOREC.NYUKA_DT, StrConv(OLD_IDOREC.NYUKA_DT, vbUnicode))         '���ד��t
        Call UniCode_Conv(IDOREC.NYUKO_DT, StrConv(OLD_IDOREC.NYUKO_DT, vbUnicode))         '���ɓ��t
        
        Call UniCode_Conv(IDOREC.WEL_ID, StrConv(OLD_IDOREC.WEL_ID, vbUnicode))             '�Ώے[����
        
        Call UniCode_Conv(IDOREC.RIRK_NAME, StrConv(OLD_IDOREC.RIRK_NAME, vbUnicode))       '������ʖ���
        
        Call UniCode_Conv(IDOREC.HIN_NAME, StrConv(OLD_IDOREC.HIN_NAME, vbUnicode))         '�i��
        
        Call UniCode_Conv(IDOREC.SUMI_HIN_Zaiko_Qty, _
                                StrConv(OLD_IDOREC.SUMI_HIN_Zaiko_Qty, vbUnicode))          '�i�ڕʍ݌ɐ��i���i���ς݁j
        Call UniCode_Conv(IDOREC.MI_HIN_Zaiko_Qty, _
                                StrConv(OLD_IDOREC.MI_HIN_Zaiko_Qty, vbUnicode))            '�i�ڕʍ݌ɐ��i�����i�j
        
        Call UniCode_Conv(IDOREC.SUMI_FROM_TANA_Zaiko_Qty, _
                                StrConv(OLD_IDOREC.SUMI_FROM_TANA_Zaiko_Qty, vbUnicode))    'FROM�I�ʕi�ڕʍ݌ɐ�
        Call UniCode_Conv(IDOREC.SUMI_TO_TANA_Zaiko_Qty, _
                                StrConv(OLD_IDOREC.SUMI_TO_TANA_Zaiko_Qty, vbUnicode))      'TO�I�ʕi�ڕʍ݌ɐ�
        Call UniCode_Conv(IDOREC.MI_FROM_TANA_Zaiko_Qty, _
                                StrConv(OLD_IDOREC.MI_FROM_TANA_Zaiko_Qty, vbUnicode))      'FROM�I�ʕi�ڕʍ݌ɐ�
        Call UniCode_Conv(IDOREC.MI_TO_TANA_Zaiko_Qty, _
                                StrConv(OLD_IDOREC.MI_TO_TANA_Zaiko_Qty, vbUnicode))        'TO�I�ʕi�ڕʍ݌ɐ�
        
        
        Call UniCode_Conv(IDOREC.TOKU_MARK, StrConv(OLD_IDOREC.TOKU_MARK, vbUnicode))       '������ϰ�
        Call UniCode_Conv(IDOREC.MEMO, StrConv(OLD_IDOREC.MEMO, vbUnicode))                 '����
        Call UniCode_Conv(IDOREC.TANTO_CODE, StrConv(OLD_IDOREC.TANTO_CODE, vbUnicode))     '�S���Һ���
        Call UniCode_Conv(IDOREC.TANTO_NAME, StrConv(OLD_IDOREC.TANTO_NAME, vbUnicode))     '�S���Җ���
        Call UniCode_Conv(IDOREC.MUKE_CODE, StrConv(OLD_IDOREC.MUKE_CODE, vbUnicode))       '���Ӑ溰��
        Call UniCode_Conv(IDOREC.MUKE_NAME, StrConv(OLD_IDOREC.MUKE_NAME, vbUnicode))       '���Ӑ於��
        Call UniCode_Conv(IDOREC.SS_CODE, StrConv(OLD_IDOREC.SS_CODE, vbUnicode))           '�����溰��
        Call UniCode_Conv(IDOREC.SS_NAME, StrConv(OLD_IDOREC.SS_NAME, vbUnicode))           '�����於��
        Call UniCode_Conv(IDOREC.MUKE_DNAME, StrConv(OLD_IDOREC.MUKE_DNAME, vbUnicode))     '���Ӑ旪��
        Call UniCode_Conv(IDOREC.MUKE_CHG_CD, StrConv(OLD_IDOREC.MUKE_CHG_CD, vbUnicode))   '������Ǒւ�����
        Call UniCode_Conv(IDOREC.SUM_KBN, StrConv(OLD_IDOREC.SUM_KBN, vbUnicode))           '�W�v�敪
        Call UniCode_Conv(IDOREC.ID_NO, StrConv(OLD_IDOREC.ID_NO, vbUnicode))               'ID-NO
        Call UniCode_Conv(IDOREC.Ins_DateTime, StrConv(OLD_IDOREC.Ins_DateTime, vbUnicode)) '�}������
        
        Call UniCode_Conv(IDOREC.SHIIRE_CODE, "")                                           '�d���溰��
        Call UniCode_Conv(IDOREC.SHIIRE_TANKA, "")                                          '�d���P��
        Call UniCode_Conv(IDOREC.KEIJYO_YM, "")                                             '�v��N��
        
        
        Call UniCode_Conv(IDOREC.FILLER, "")
        
        
        
        
        
        
        Do
            sts = BTRV(BtOpInsert, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<IDO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "�݌Ɉړ����f�[�^")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

'---------------------------------------------  �I��
    Cnt(4).Caption = Format(Count, "#0")
    If Mode <> 0 Then
        Me.MousePointer = vbDefault
        MsgBox "�R���o�[�g�I��"
        Update_Proc = False
        Exit Function
    End If

    '---------------------------------------------------------  ���׃`�F�b�N�f�[�^�̏���
J_NYU_CONV:


    MsgLab(1) = "���׃`�F�b�N�f�[�^�R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(5).Caption = Format(Count, "#0")
                                        
                                '(��)���׃`�F�b�N�f�[�^�n�o�d�m
    If OLD_J_NYU_Open(BtOpenNomal) Then
        If sts = BtErrFileNotFound Then
            If Mode = 0 Then
                GoTo SUMZ_CONV
            Else
                MsgBox "�Ώۃf�[�^�Ȃ�"
                Unload Me
            End If
        End If
        Unload Me
    End If
                                        
                                        
    com = BtOpGetFirst
                                        
                                        
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_J_NYU_POS, OLD_J_NYUREC, Len(OLD_J_NYUREC), K0_OLD_J_NYU, Len(K0_OLD_J_NYU), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "(��)���׃`�F�b�N�f�[�^")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(5).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        Call UniCode_Conv(J_NYUREC.JGYOBU, StrConv(OLD_J_NYUREC.JGYOBU, vbUnicode))         '���ƕ�
        Call UniCode_Conv(J_NYUREC.NAIGAI, StrConv(OLD_J_NYUREC.NAIGAI, vbUnicode))         '�����O
        Call UniCode_Conv(J_NYUREC.HIN_GAI, StrConv(OLD_J_NYUREC.HIN_GAI, vbUnicode))       '�i��(�O��)
        Call UniCode_Conv(J_NYUREC.JITU_QTY, StrConv(OLD_J_NYUREC.JITU_QTY, vbUnicode))     '���ѐ���
        Call UniCode_Conv(J_NYUREC.INS_DATE, Format(Now, "YYYYMMDD"))                       '�o�^��
        Call UniCode_Conv(J_NYUREC.FILLER, "")
        
        Do
            sts = BTRV(BtOpInsert, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<J_NYU.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "���������ް�")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

'---------------------------------------------  �I��
    Cnt(5).Caption = Format(Count, "#0")

    If Mode <> 0 Then
        Me.MousePointer = vbDefault
        MsgBox "�R���o�[�g�I��"
        Update_Proc = False
        Exit Function
    End If

    '---------------------------------------------------------  �݌ɏW�v�f�[�^�̏���
SUMZ_CONV:


    MsgLab(1) = "�݌ɏW�v�f�[�^�R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(6).Caption = Format(Count, "#0")
                                        
    
                                '(��)�݌ɏW�v�f�[�^�n�o�d�m
    If OLD_SUMZ_Open(BtOpenNomal) Then
        If sts = BtErrFileNotFound Then
            If Mode = 0 Then
                GoTo STOCK_CONV
            Else
                MsgBox "�Ώۃf�[�^�Ȃ�"
                Unload Me
            End If
        End If
        Unload Me
    End If
    
    
    com = BtOpGetFirst
                                        
                                        
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_SUMZ_POS, OLD_SUMZREC, Len(OLD_SUMZREC), K0_OLD_SUMZ, Len(K0_OLD_SUMZ), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "(��)�݌ɏW�v�f�[�^")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(6).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        Call UniCode_Conv(SUMZREC.JGYOBU, StrConv(OLD_SUMZREC.JGYOBU, vbUnicode))       '���ƕ�
        Call UniCode_Conv(SUMZREC.NAIGAI, StrConv(OLD_SUMZREC.NAIGAI, vbUnicode))       '�����O
        Call UniCode_Conv(SUMZREC.HIN_GAI, StrConv(OLD_SUMZREC.HIN_GAI, vbUnicode))     '�i��(�O��)
        
        Call UniCode_Conv(SUMZREC.ST_SOKO, StrConv(OLD_SUMZREC.ST_SOKO, vbUnicode))     '�W���I�� �q�ɇ�
        Call UniCode_Conv(SUMZREC.ST_RETU, StrConv(OLD_SUMZREC.ST_RETU, vbUnicode))     '�W���I�� ��
        Call UniCode_Conv(SUMZREC.ST_REN, StrConv(OLD_SUMZREC.ST_REN, vbUnicode))       '�W���I�� �A
        Call UniCode_Conv(SUMZREC.ST_DAN, StrConv(OLD_SUMZREC.ST_DAN, vbUnicode))       '�W���I�� �i
        
        Call UniCode_Conv(SUMZREC.T_Zai_Qty, StrConv(OLD_SUMZREC.T_Zai_Qty, vbUnicode))     '�݌ɑ���(����)
        Call UniCode_Conv(SUMZREC.ZEN_Zai_Qty, StrConv(OLD_SUMZREC.ZEN_Zai_Qty, vbUnicode)) '�݌ɑ���(�O��)
                        
        Call UniCode_Conv(SUMZREC.SYK_E_QTY, StrConv(OLD_SUMZREC.SYK_E_QTY, vbUnicode))     '�o�ɍϐ�
        Call UniCode_Conv(SUMZREC.NYUKA_YQTY, StrConv(OLD_SUMZREC.NYUKA_YQTY, vbUnicode))   '���ח\�萔
                        
        Call UniCode_Conv(SUMZREC.HS_ZAIQTY, StrConv(OLD_SUMZREC.HS_ZAIQTY, vbUnicode))         'νč݌ɐ�(����)
        Call UniCode_Conv(SUMZREC.ZEN_HS_ZAIQTY, StrConv(OLD_SUMZREC.ZEN_HS_ZAIQTY, vbUnicode)) 'νč݌ɐ�(�O��)
                        
        
        Call UniCode_Conv(SUMZREC.SAI_QTY, StrConv(OLD_SUMZREC.SAI_QTY, vbUnicode))     '���ِ�
        Call UniCode_Conv(SUMZREC.SUM_DT, StrConv(OLD_SUMZREC.SUM_DT, vbUnicode))       '�W�v���t
        
        
        
        
        Call UniCode_Conv(SUMZREC.FILLER, "")
        
        Do
            sts = BTRV(BtOpInsert, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "�݌ɏW�v�ް�")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

'---------------------------------------------  �I��
    Cnt(6).Caption = Format(Count, "#0")

    If Mode <> 0 Then
        Me.MousePointer = vbDefault
        MsgBox "�R���o�[�g�I��"
        Update_Proc = False
        Exit Function
    End If

    '---------------------------------------------------------  �I�����f�[�^�̏���

STOCK_CONV:

    MsgLab(1) = "�I�����f�[�^�R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(7).Caption = Format(Count, "#0")
                                        
    
                                '(��)�I�����f�[�^�n�o�d�m
    If OLD_STOCK_Open(BtOpenNomal) Then
        If sts = BtErrFileNotFound Then
            If Mode = 0 Then
                GoTo KEPPINLOG_CONV
            Else
                MsgBox "�Ώۃf�[�^�Ȃ�"
                Unload Me
            End If
        End If
        Unload Me
    End If
    
    
    com = BtOpGetFirst
                                        
                                        
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_STOCK_POS, OLD_STOCKREC, Len(OLD_STOCKREC), K0_OLD_STOCK, Len(K0_OLD_STOCK), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "(��)�I�����f�[�^")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(7).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        Call UniCode_Conv(STOCKREC.JGYOBU, StrConv(OLD_STOCKREC.JGYOBU, vbUnicode))     '���ƕ�
        Call UniCode_Conv(STOCKREC.NAIGAI, StrConv(OLD_STOCKREC.NAIGAI, vbUnicode))     '�����O
        Call UniCode_Conv(STOCKREC.HIN_GAI, StrConv(OLD_STOCKREC.HIN_GAI, vbUnicode))   '�i��(�O��)
        
        Call UniCode_Conv(STOCKREC.ST_LOCATION, StrConv(OLD_STOCKREC.ST_LOCATION, vbUnicode))   '�W�����ɑq��
        Call UniCode_Conv(STOCKREC.HOST_ZAIKO, StrConv(OLD_STOCKREC.HOST_ZAIKO, vbUnicode))     'νė��_�݌�
        Call UniCode_Conv(STOCKREC.POS_ZAIKO, StrConv(OLD_STOCKREC.POS_ZAIKO, vbUnicode))       'POS�݌�
        
        Call UniCode_Conv(STOCKREC.ST_ZAIKO, StrConv(OLD_STOCKREC.ST_ZAIKO, vbUnicode))         '�W���I�ԍ݌�
        
        Call UniCode_Conv(STOCKREC.EE1_LOCATION, StrConv(OLD_STOCKREC.EE1_LOCATION, vbUnicode)) '�ʒu���I��1
        Call UniCode_Conv(STOCKREC.EE1_ZAIKO, StrConv(OLD_STOCKREC.EE1_ZAIKO, vbUnicode))       '�ʒu���I��1 �݌�
        Call UniCode_Conv(STOCKREC.EE2_LOCATION, StrConv(OLD_STOCKREC.EE2_LOCATION, vbUnicode)) '�ʒu���I��2
        Call UniCode_Conv(STOCKREC.EE2_ZAIKO, StrConv(OLD_STOCKREC.EE2_ZAIKO, vbUnicode))       '�ʒu���I��2 �݌�
        Call UniCode_Conv(STOCKREC.EE3_LOCATION, StrConv(OLD_STOCKREC.EE3_LOCATION, vbUnicode)) '�ʒu���I��3
        Call UniCode_Conv(STOCKREC.EE3_ZAIKO, StrConv(OLD_STOCKREC.EE3_ZAIKO, vbUnicode))       '�ʒu���I��3 �݌�
        
        Call UniCode_Conv(STOCKREC.ETC_ZAIKO, StrConv(OLD_STOCKREC.ETC_ZAIKO, vbUnicode))       '���̑��݌�
        
        Call UniCode_Conv(STOCKREC.CHECK_MARK, StrConv(OLD_STOCKREC.CHECK_MARK, vbUnicode))     '�ƍ�ϰ�
        
        Call UniCode_Conv(STOCKREC.PRINT_YMD, StrConv(OLD_STOCKREC.PRINT_YMD, vbUnicode))       '������t
        Call UniCode_Conv(STOCKREC.INPUT_YMD, StrConv(OLD_STOCKREC.INPUT_YMD, vbUnicode))       '���͓��t
        
        Call UniCode_Conv(STOCKREC.SAI_QTY, StrConv(OLD_STOCKREC.SAI_QTY, vbUnicode))           '���ِ�
        
        
        
        
        
        
        
        Call UniCode_Conv(STOCKREC.FILLER, "")
        
        Do
            sts = BTRV(BtOpInsert, STOCK_POS, STOCKREC, Len(STOCKREC), K0_STOCK, Len(K0_STOCK), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<STOCK.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "�I�����ް�")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

'---------------------------------------------  �I��
    Cnt(7).Caption = Format(Count, "#0")
    If Mode <> 0 Then
        Me.MousePointer = vbDefault
        MsgBox "�R���o�[�g�I��"
        Update_Proc = False
        Exit Function
    End If

    '---------------------------------------------------------  ���i�h�~�x�����O�f�[�^�̏���

KEPPINLOG_CONV:
    
    MsgLab(1) = "���i�h�~�x�����O�f�[�^�R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(8).Caption = Format(Count, "#0")
                                        
                                '(��)���i�h�~�x�����O�f�[�^�n�o�d�m
    If OLD_KEPPINLOG_Open(BtOpenNomal) Then
        If sts = BtErrFileNotFound Then
            If Mode = 0 Then
                MsgBox "�R���o�[�g�I��"
                Update_Proc = False
            Else
                MsgBox "�Ώۃf�[�^�Ȃ�"
                Unload Me
            End If
        End If
        Unload Me
    End If
                                        
                                        
    com = BtOpGetFirst
                                        
                                        
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_KEPPINLOG_POS, OLD_KEPPINLOGREC, Len(OLD_KEPPINLOGREC), K0_OLD_KEPPINLOG, Len(K0_OLD_KEPPINLOG), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "(��)���i�h�~�x�����O�f�[�^")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(8).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        Call UniCode_Conv(KEPPINLOGREC.JGYOBU, StrConv(OLD_KEPPINLOGREC.JGYOBU, vbUnicode))         '���ƕ�
        Call UniCode_Conv(KEPPINLOGREC.NAIGAI, StrConv(OLD_KEPPINLOGREC.NAIGAI, vbUnicode))         '�����O
        Call UniCode_Conv(KEPPINLOGREC.HIN_GAI, StrConv(OLD_KEPPINLOGREC.HIN_GAI, vbUnicode))       '�i��(�O��)
        
        Call UniCode_Conv(KEPPINLOGREC.CREATE_DT, StrConv(OLD_KEPPINLOGREC.CREATE_DT, vbUnicode))   '�쐬���t
        
        
        Call UniCode_Conv(KEPPINLOGREC.FILLER, "")
        
        Do
            sts = BTRV(BtOpInsert, KEPPINLOG_POS, KEPPINLOGREC, Len(KEPPINLOGREC), K0_KEPPINLOG, Len(K0_KEPPINLOG), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<KEPPINLOG.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "���i�h�~۸��ް�")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

'---------------------------------------------  �I��
    Cnt(8).Caption = Format(Count, "#0")

    If Mode <> 0 Then
        Me.MousePointer = vbDefault
        MsgBox "�R���o�[�g�I��"
        Update_Proc = False
        Exit Function
    End If

    Me.MousePointer = vbDefault
    MsgBox "�R���o�[�g�I��"
    Update_Proc = False


End Function


Private Sub Command1_Click(Index As Integer)


Dim ans As Integer
                                
    If Index = 10 Then
        Unload Me
    End If
                                '�����I��
    Beep
    ans = MsgBox("���s���܂����H", vbYesNo + vbQuestion, "�m�F����")
    If ans = vbYes Then
        
        
        If Index = 9 Then
        
        
            If Update_Proc(0) Then
                Unload Me
            End If
        Else
        
        
            If Update_Proc(Index + 1) Then
                Unload Me
            End If
        End If
    End If
'    Unload Me



End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_Load()
Dim i As Integer
Dim c As String * 128
Dim sts As Integer
    
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
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '�݌Ƀf�[�^�n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '�o�ח\��f�[�^�n�o�d�m
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '�ߓ����o�׃f�[�^�n�o�d�m
    If DEL_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
    
                                '�݌Ɉړ����f�[�^�n�o�d�m
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
                                '���׃`�F�b�N�f�[�^�n�o�d�m
    If J_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
                                '�݌ɏW�v�f�[�^�n�o�d�m
    If SUMZ_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
                                '�I�����f�[�^�n�o�d�m
    If STOCK_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
                                '���i�h�~�x�����O�f�[�^�n�o�d�m
    If KEPPINLOG_Open(BtOpenNomal) Then
        Unload Me
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
                                            '(��)�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, OLD_ITEM_POS, OLD_ITEMREC, Len(OLD_ITEMREC), K0_OLD_ITEM, Len(K0_OLD_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(��)�i�ڃ}�X�^")
        End If
    End If
    
                                            '�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        End If
    End If
                                            '(��)�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, OLD_ZAIKO_POS, OLD_ZAIKOREC, Len(OLD_ZAIKOREC), K0_OLD_ZAIKO, Len(K0_OLD_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(��)�݌Ƀf�[�^")
        End If
    End If
                                            '�o�ח\��f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�o�ח\��f�[�^")
        End If
    End If
                                            '(��)�o�ח\��f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, OLD_ZAIKO_POS, OLD_ZAIKOREC, Len(OLD_ZAIKOREC), K0_OLD_ZAIKO, Len(K0_OLD_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(��)�݌Ƀf�[�^")
        End If
    End If
                                            '�ߓ����o�׃f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, DEL_SYU_POS, DEL_SYUREC, Len(DEL_SYUREC), K0_DEL_SYU, Len(K0_DEL_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�ߓ����o�׃f�[�^")
        End If
    End If
                                            '(��)�ߓ����o�׃f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, OLD_DEL_SYU_POS, OLD_DEL_SYUREC, Len(OLD_DEL_SYUREC), K0_DEL_SYU, Len(K0_DEL_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(��)�ߓ����o�׃f�[�^")
        End If
    End If
    
    
                                            '�݌Ɉړ����f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ɉړ����f�[�^")
        End If
    End If
                                            '(��)�݌Ɉړ����f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, OLD_IDO_POS, OLD_IDOREC, Len(OLD_IDOREC), K0_OLD_IDO, Len(K0_OLD_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(��)�݌Ɉړ����f�[�^")
        End If
    End If
    
    
                                            '���׃`�F�b�N�f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���׃`�F�b�N�f�[�^")
        End If
    End If
                                            '(��)���׃`�F�b�N�f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, OLD_J_NYU_POS, OLD_J_NYUREC, Len(OLD_J_NYUREC), K0_OLD_J_NYU, Len(K0_OLD_J_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(��)���׃`�F�b�N�f�[�^")
        End If
    End If
    
                                            '�݌ɏW�v�f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌ɏW�v�f�[�^")
        End If
    End If
                                            '(��)�݌ɏW�v�f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, OLD_SUMZ_POS, OLD_SUMZREC, Len(OLD_SUMZREC), K0_OLD_SUMZ, Len(K0_OLD_SUMZ), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(��)�݌ɏW�v�f�[�^")
        End If
    End If
    
    
                                            '�I�����f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, STOCK_POS, STOCKREC, Len(STOCKREC), K0_STOCK, Len(K0_STOCK), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�I�����f�[�^")
        End If
    End If
                                            '(��)�I�����f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, OLD_STOCK_POS, OLD_STOCKREC, Len(OLD_STOCKREC), K0_OLD_STOCK, Len(K0_OLD_STOCK), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(��)�I�����f�[�^")
        End If
    End If
    
    
                                            '���i�h�~�x�����O�f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, KEPPINLOG_POS, KEPPINLOGREC, Len(KEPPINLOGREC), K0_KEPPINLOG, Len(K0_KEPPINLOG), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���i�h�~�x�����O�f�[�^")
        End If
    End If
                                            '(��)���i�h�~�x�����O�f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, OLD_KEPPINLOG_POS, OLD_KEPPINLOGREC, Len(OLD_KEPPINLOGREC), K0_OLD_KEPPINLOG, Len(K0_OLD_KEPPINLOG), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(��)���i�h�~�x�����O�f�[�^")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set PC000501 = Nothing

    End
End Sub

