VERSION 5.00
Begin VB.Form CONV20041 
   BackColor       =   &H00C0C0C0&
   Caption         =   "�f�[�^�R���o�[�g����"
   ClientHeight    =   7230
   ClientLeft      =   2325
   ClientTop       =   2625
   ClientWidth     =   10095
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
   ScaleWidth      =   10095
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   3000
      TabIndex        =   9
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�@�o�ח\�聁"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   8
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�݌Ɉړ�����"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   7
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   6
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   3000
      TabIndex        =   4
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�݌Ƀf�[�^��"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   1440
      TabIndex        =   5
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�i�ڃ}�X�^��"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   3000
      Width           =   1455
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
      Top             =   2160
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
      Top             =   960
      Width           =   4800
   End
End
Attribute VB_Name = "CONV20041"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function Update_Proc() As Integer

Dim sts             As Integer
Dim Upd_com         As Integer
Dim com             As Integer
Dim ans             As Integer
Dim Count           As Long

Dim DISP_INTERVAL   As Long


Dim MTS_CODE        As String * 8
Dim SS_CODE         As String * 8

Dim c               As String * 128

    Update_Proc = True

    GoTo ido_upd
'---------------------------------------------  �i�ڃ}�X�^�̃R���o�[�g
    MsgLab(1) = "�i�ڃ}�X�^�R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(0).Caption = Format(Count, "#0")
                                        
                                        
    com = BtOpGetFirst
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_ITEM_POS, OLD_ITEMREC, Len(OLD_ITEMREC), K0_OLD_ITEM, Len(K0_OLD_ITEM), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�i���j�i�ڃ}�X�^")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(0).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        
        Call UniCode_Conv(ITEMREC.JGYOBU, StrConv(OLD_ITEMREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(ITEMREC.NAIGAI, StrConv(OLD_ITEMREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(ITEMREC.HIN_GAI, StrConv(OLD_ITEMREC.HIN_GAI, vbUnicode))
        Call UniCode_Conv(ITEMREC.HIN_NAME, StrConv(OLD_ITEMREC.HIN_NAME, vbUnicode))
        Call UniCode_Conv(ITEMREC.ST_SET_DT, StrConv(OLD_ITEMREC.ST_SET_DT, vbUnicode))
        Call UniCode_Conv(ITEMREC.ST_SOKO, StrConv(OLD_ITEMREC.ST_SOKO, vbUnicode))
        Call UniCode_Conv(ITEMREC.ST_RETU, StrConv(OLD_ITEMREC.ST_RETU, vbUnicode))
        Call UniCode_Conv(ITEMREC.ST_REN, StrConv(OLD_ITEMREC.ST_REN, vbUnicode))
        Call UniCode_Conv(ITEMREC.ST_DAN, StrConv(OLD_ITEMREC.ST_DAN, vbUnicode))
        Call UniCode_Conv(ITEMREC.BEF_SOKO, StrConv(OLD_ITEMREC.BEF_SOKO, vbUnicode))
        Call UniCode_Conv(ITEMREC.BEF_RETU, StrConv(OLD_ITEMREC.BEF_RETU, vbUnicode))
        Call UniCode_Conv(ITEMREC.BEF_REN, StrConv(OLD_ITEMREC.BEF_REN, vbUnicode))
        Call UniCode_Conv(ITEMREC.BEF_DAN, StrConv(OLD_ITEMREC.BEF_DAN, vbUnicode))
        Call UniCode_Conv(ITEMREC.LAST_NYU_DT, StrConv(OLD_ITEMREC.LAST_NYU_DT, vbUnicode))
        Call UniCode_Conv(ITEMREC.LAST_SYU_DT, StrConv(OLD_ITEMREC.LAST_SYU_DT, vbUnicode))
        Call UniCode_Conv(ITEMREC.HIN_NAI, StrConv(OLD_ITEMREC.HIN_NAI, vbUnicode))
        Call UniCode_Conv(ITEMREC.BIKOU_SOKO, StrConv(OLD_ITEMREC.BIKOU_SOKO, vbUnicode))
        Call UniCode_Conv(ITEMREC.BIKOU_TANA, StrConv(OLD_ITEMREC.BIKOU_TANA, vbUnicode))
        Call UniCode_Conv(ITEMREC.SIZAI_CD, StrConv(OLD_ITEMREC.SIZAI_CD, vbUnicode))
        Call UniCode_Conv(ITEMREC.HOJYU_P, StrConv(OLD_ITEMREC.HOJYU_P, vbUnicode))
        Call UniCode_Conv(ITEMREC.AVE_SYUKA, StrConv(OLD_ITEMREC.AVE_SYUKA, vbUnicode))
        Call UniCode_Conv(ITEMREC.SAMPLE_QTY, StrConv(OLD_ITEMREC.SAMPLE_QTY, vbUnicode))
        Call UniCode_Conv(ITEMREC.LAST_INP_DT, StrConv(OLD_ITEMREC.LAST_INP_DT, vbUnicode))
        Call UniCode_Conv(ITEMREC.LOCK_F, StrConv(OLD_ITEMREC.LOCK_F, vbUnicode))
        Call UniCode_Conv(ITEMREC.WEL_ID, StrConv(OLD_ITEMREC.WEL_ID, vbUnicode))
        Call UniCode_Conv(ITEMREC.PRG_ID, StrConv(OLD_ITEMREC.PRG_ID, vbUnicode))
        Call UniCode_Conv(ITEMREC.LAST_CHK_DT, StrConv(OLD_ITEMREC.LAST_CHK_DT, vbUnicode))
        Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, StrConv(OLD_ITEMREC.LAST_CHK_QTY, vbUnicode))
        Call UniCode_Conv(ITEMREC.MOTO_JIGYOBU, StrConv(OLD_ITEMREC.MOTO_JIGYOBU, vbUnicode))
        Call UniCode_Conv(ITEMREC.BIKOU, StrConv(OLD_ITEMREC.BIKOU, vbUnicode))
        Call UniCode_Conv(ITEMREC.IRI_QTY, StrConv(OLD_ITEMREC.IRI_QTY, vbUnicode))
        Call UniCode_Conv(ITEMREC.JAN_CODE, "")
        Call UniCode_Conv(ITEMREC.HIN_CHANGE, "")
        Call UniCode_Conv(ITEMREC.GOODS_KBN, "0")
        Call UniCode_Conv(ITEMREC.PACKING_NO, "")
        Call UniCode_Conv(ITEMREC.FILLER, "")
        
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
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

    Cnt(0).Caption = Format(Count, "#0")
'---------------------------------------------  �݌Ƀf�[�^�̃R���o�[�g
zaiko_upd:
    MsgLab(1) = "�݌Ƀf�[�^�R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(1).Caption = Format(Count, "#0")
                                        
                                        
    com = BtOpGetFirst
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_ZAIKO_POS, OLD_ZAIKOREC, Len(OLD_ZAIKOREC), K0_OLD_ZAIKO, Len(K0_OLD_ZAIKO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�i���j�݌Ƀf�[�^")
                Exit Function
        End Select
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(1).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        Call UniCode_Conv(ZAIKOREC.Soko_No, StrConv(OLD_ZAIKOREC.Soko_No, vbUnicode))       '�q�ɇ�
        Call UniCode_Conv(ZAIKOREC.Retu, StrConv(OLD_ZAIKOREC.Retu, vbUnicode))             '�I�ԁ@��
        Call UniCode_Conv(ZAIKOREC.Ren, StrConv(OLD_ZAIKOREC.Ren, vbUnicode))               '�I�ԁ@�A
        Call UniCode_Conv(ZAIKOREC.Dan, StrConv(OLD_ZAIKOREC.Dan, vbUnicode))               '�I�ԁ@�i
        Call UniCode_Conv(ZAIKOREC.JGYOBU, StrConv(OLD_ZAIKOREC.JGYOBU, vbUnicode))         '���ƕ�
        Call UniCode_Conv(ZAIKOREC.NAIGAI, StrConv(OLD_ZAIKOREC.NAIGAI, vbUnicode))         '�����O
        Call UniCode_Conv(ZAIKOREC.HIN_GAI, StrConv(OLD_ZAIKOREC.HIN_GAI, vbUnicode))       '�i�ځi�O���j
        
        
        If StrConv(OLD_ZAIKOREC.Soko_No, vbUnicode) = "92" Or _
            StrConv(OLD_ZAIKOREC.Soko_No, vbUnicode) = "93" Or _
            StrConv(OLD_ZAIKOREC.Soko_No, vbUnicode) = "81" Then
            Call UniCode_Conv(ZAIKOREC.GOODS_ON, "0")                                       '���i���^�����i��
        Else
            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(OLD_ZAIKOREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(OLD_ZAIKOREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(OLD_ZAIKOREC.HIN_GAI, vbUnicode))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts = BtNoErr Then
                If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = "1" Then
                    Call UniCode_Conv(ZAIKOREC.GOODS_ON, "0")                                   '���i���^�����i��
                Else
                    Call UniCode_Conv(ZAIKOREC.GOODS_ON, "1")                                   '���i���^�����i��
                End If
            Else
                Call UniCode_Conv(ZAIKOREC.GOODS_ON, "1")                                   '���i���^�����i��
            End If
        End If
        
        Call UniCode_Conv(ZAIKOREC.NYUKA_DT, StrConv(OLD_ZAIKOREC.NYUKA_DT, vbUnicode))     '���ד�
        Call UniCode_Conv(ZAIKOREC.NYUKO_DT, StrConv(OLD_ZAIKOREC.NYUKO_DT, vbUnicode))     '���ɓ�
        Call UniCode_Conv(ZAIKOREC.HIN_NAI, StrConv(OLD_ZAIKOREC.HIN_NAI, vbUnicode))       '�i�ԁi�����j
        Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, StrConv(OLD_ZAIKOREC.YUKO_Z_QTY, vbUnicode)) '�L���݌ɐ�
        Call UniCode_Conv(ZAIKOREC.LOCK_F, "0")                                             '�r���t���O
        Call UniCode_Conv(ZAIKOREC.WEL_ID, "")                                              '�g�p�q�@ID
        Call UniCode_Conv(ZAIKOREC.PRG_ID, "")                                              '�g�p�q�@ID
        Call UniCode_Conv(ZAIKOREC.GOODS_YMD, "")                                           '���i�����t
        Call UniCode_Conv(ZAIKOREC.FILLER, Format(Now, "YYYYMMDD"))                                             ''FILLER
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
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌Ƀf�[�^")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

    
    Cnt(1).Caption = Format(Count, "#0")

    GoTo Update_End

'---------------------------------------------  �݌Ɉړ����̃R���o�[�g
ido_upd:
    MsgLab(1) = "�݌Ɉړ����R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(2).Caption = Format(Count, "#0")


    com = BtOpGetFirst
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_IDO_POS, OLD_IDOREC, Len(OLD_IDOREC), K0_OLD_IDO, Len(K0_OLD_IDO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�i���j�݌Ɉړ���")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(2).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        
        Call UniCode_Conv(IDOREC.JITU_DT, StrConv(OLD_IDOREC.JITU_DT, vbUnicode))           '���ѓ��t
        Call UniCode_Conv(IDOREC.JITU_TM, StrConv(OLD_IDOREC.JITU_TM, vbUnicode))           '���ю���
        Call UniCode_Conv(IDOREC.JGYOBU, StrConv(OLD_IDOREC.JGYOBU, vbUnicode))             '���ƕ��敪
        Call UniCode_Conv(IDOREC.NAIGAI, StrConv(OLD_IDOREC.NAIGAI, vbUnicode))             '�����O
        Call UniCode_Conv(IDOREC.HIN_GAI, StrConv(OLD_IDOREC.HIN_GAI, vbUnicode))           '�i�ځi�O���j
        Call UniCode_Conv(IDOREC.RIRK_ID, StrConv(OLD_IDOREC.RIRK_ID, vbUnicode))           '�������
        Call UniCode_Conv(IDOREC.SUMI_JITU_QTY, StrConv(OLD_IDOREC.JITU_QTY, vbUnicode))    '���ѐ���(���i���ς�)
        Call UniCode_Conv(IDOREC.MI_JITU_QTY, "00000000")                                   '���ѐ���(���ѐ���(�����i))
        Call UniCode_Conv(IDOREC.FROM_SOKO, StrConv(OLD_IDOREC.FROM_SOKO, vbUnicode))       'From �q�ɇ�
        Call UniCode_Conv(IDOREC.FROM_RETU, StrConv(OLD_IDOREC.FROM_RETU, vbUnicode))       'From ��
        Call UniCode_Conv(IDOREC.FROM_REN, StrConv(OLD_IDOREC.FROM_REN, vbUnicode))         'From �A
        Call UniCode_Conv(IDOREC.FROM_DAN, StrConv(OLD_IDOREC.FROM_DAN, vbUnicode))         'From �i
        Call UniCode_Conv(IDOREC.TO_SOKO, StrConv(OLD_IDOREC.TO_SOKO, vbUnicode))         'TO �q�ɇ�
        Call UniCode_Conv(IDOREC.TO_RETU, StrConv(OLD_IDOREC.TO_RETU, vbUnicode))         'TO ��
        Call UniCode_Conv(IDOREC.TO_REN, StrConv(OLD_IDOREC.TO_REN, vbUnicode))           'TO �A
        Call UniCode_Conv(IDOREC.TO_DAN, StrConv(OLD_IDOREC.TO_DAN, vbUnicode))             'TO �i
        Call UniCode_Conv(IDOREC.DEN_DT, StrConv(OLD_IDOREC.DEN_DT, vbUnicode))             '�`�[���t
        Call UniCode_Conv(IDOREC.DEN_NO, StrConv(OLD_IDOREC.DEN_NO, vbUnicode))             '�`�[��
        Call UniCode_Conv(IDOREC.PRG_ID, StrConv(OLD_IDOREC.PRG_ID, vbUnicode))             '�o�͌��v���O����
        Call UniCode_Conv(IDOREC.HIN_NAI, StrConv(OLD_IDOREC.HIN_NAI, vbUnicode))           '�i�ԁi�����j
        Call UniCode_Conv(IDOREC.NYUKA_DT, StrConv(OLD_IDOREC.NYUKA_DT, vbUnicode))         '���ד��t
        Call UniCode_Conv(IDOREC.NYUKO_DT, StrConv(OLD_IDOREC.NYUKO_DT, vbUnicode))         '���ɓ��t
        Call UniCode_Conv(IDOREC.WEL_ID, StrConv(OLD_IDOREC.WEL_ID, vbUnicode))             '�Ώے[����
        Call UniCode_Conv(IDOREC.RIRK_NAME, StrConv(OLD_IDOREC.RIRK_NAME, vbUnicode))       '������ʖ���
        Call UniCode_Conv(IDOREC.HIN_NAME, StrConv(OLD_IDOREC.HIN_NAME, vbUnicode))         '�i��
        Call UniCode_Conv(IDOREC.SUMI_HIN_Zaiko_Qty, _
                                    StrConv(OLD_IDOREC.HIN_Zaiko_Qty, vbUnicode))           '�i�ڕʍ݌ɐ��i���i���ς݁j
        Call UniCode_Conv(IDOREC.MI_HIN_Zaiko_Qty, "00000000")                              '�i�ڕʍ݌ɐ��i�����i�j
        Call UniCode_Conv(IDOREC.SUMI_FROM_TANA_Zaiko_Qty, _
                                    StrConv(OLD_IDOREC.FROM_TANA_Zaiko_Qty, vbUnicode))     'FROM�I�ʕi�ڕʍ݌ɐ��i���i���ς݁j
        Call UniCode_Conv(IDOREC.SUMI_TO_TANA_Zaiko_Qty, _
                                    StrConv(OLD_IDOREC.TO_TANA_Zaiko_Qty, vbUnicode))       'TO�I�ʕi�ڕʍ݌ɐ��i���i���ς݁j
        Call UniCode_Conv(IDOREC.MI_FROM_TANA_Zaiko_Qty, "00000000")                        'FROM�I�ʕi�ڕʍ݌ɐ��i�����i�j
        Call UniCode_Conv(IDOREC.MI_TO_TANA_Zaiko_Qty, "00000000")                          'TO�I�ʕi�ڕʍ݌ɐ��i�����i�j
        Call UniCode_Conv(IDOREC.TOKU_MARK, StrConv(OLD_IDOREC.TOKU_MARK, vbUnicode))       '������}�[�N
        Call UniCode_Conv(IDOREC.MEMO, StrConv(OLD_IDOREC.MEMO, vbUnicode))                 '����
        Call UniCode_Conv(IDOREC.TANTO_CODE, "")                                            '�S���҃R�[�h
        Call UniCode_Conv(IDOREC.TANTO_NAME, "")                                            '�S���Җ���
        Call UniCode_Conv(IDOREC.MUKE_CODE, StrConv(OLD_IDOREC.MUKE_CODE, vbUnicode))       '���Ӑ�R�[�h
        Call UniCode_Conv(IDOREC.MUKE_NAME, StrConv(OLD_IDOREC.MUKE_DNAME, vbUnicode))       '���Ӑ於��
        Call UniCode_Conv(IDOREC.SS_CODE, "")                                               '������R�[�h
        Call UniCode_Conv(IDOREC.SS_NAME, "")                                               '�����於��
        Call UniCode_Conv(IDOREC.MUKE_DNAME, StrConv(OLD_IDOREC.MUKE_DNAME, vbUnicode))     '���Ӑ旪��
        Call UniCode_Conv(IDOREC.MUKE_CHG_CD, StrConv(OLD_IDOREC.MUKE_CHG_CD, vbUnicode))   '������Ǒւ��R�[�h
        Call UniCode_Conv(IDOREC.SUM_KBN, StrConv(OLD_IDOREC.SUM_KBN, vbUnicode))           '�W�v�敪
        Call UniCode_Conv(IDOREC.FILLER, "")
        
        Do
            sts = BTRV(BtOpInsert, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<IDOREKI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌Ɉړ���")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

    Cnt(2).Caption = Format(Count, "#0")
    GoTo Update_End
'---------------------------------------------  �o�ח\��̃R���o�[�g
syuka_upd:
    
    MsgLab(1) = "�o�ח\��f�[�^�R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(3).Caption = Format(Count, "#0")
                                        
                                        
    com = BtOpGetFirst
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_Y_SYU_POS, OLD_Y_SYUREC, Len(OLD_Y_SYUREC), K0_OLD_Y_SYU, Len(K0_OLD_Y_SYU), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�i���j�o�ח\��f�[�^")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(3).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
            
        Call UniCode_Conv(Y_SYUREC.WEL_ID, "")                  '�g�p�[���h�c
        Call UniCode_Conv(Y_SYUREC.PRG_ID, "")                  '�g�p���v�g�O�����h�c
        If CLng(StrConv(OLD_Y_SYUREC.FIX_QTY, vbUnicode)) >= CLng(StrConv(OLD_Y_SYUREC.YOTEI_QTY, vbUnicode)) Then
            Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_FIN)    '�����敪������
        Else
            Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_UN)     '�����敪������
        End If
                                                                '�f�[�^���
        Call UniCode_Conv(Y_SYUREC.DT_SYU, StrConv(OLD_Y_SYUREC.DT_SYU, vbUnicode))
                                                                '���ƕ��R�[�h
        Call UniCode_Conv(Y_SYUREC.JGYOBU, StrConv(OLD_Y_SYUREC.JGYOBU, vbUnicode))
                                                                '�����敪�i�j�d�x�j
        Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, StrConv(OLD_Y_SYUREC.HS_CYU_KBN, vbUnicode))
                                                                '�`�[�h�c�i�j�d�x�j�i���`�[�ԍ��j
        Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, Format(CLng(StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode)), "00000000"))
                                                                '�����O
        Call UniCode_Conv(Y_SYUREC.NAIGAI, StrConv(OLD_Y_SYUREC.NAIGAI, vbUnicode))
                                                                '�i�ڔԍ��i�j�d�x�j
        Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, StrConv(OLD_Y_SYUREC.HIN_GAI, vbUnicode))
                                                                
        sts = GetIni(App.EXEName, Trim(StrConv(OLD_Y_SYUREC.MUKE_CODE, vbUnicode)), "SETUP", c)
        
        If sts Then
            MTS_CODE = ETS_MTS & StrConv(OLD_Y_SYUREC.NAIGAI, vbUnicode)
        Else
            MTS_CODE = Trim(c)
        End If
                                                                
        SS_CODE = ""
        
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "095299" Then
            MTS_CODE = "20513770"
        End If
                                                                
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "095412" Then
            MTS_CODE = "75T"
            SS_CODE = "20099826"
        End If
                                                                
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "095414" Then
            MTS_CODE = "75T"
            SS_CODE = "20099826"
        End If
                                                                
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "095413" Then
            MTS_CODE = "75T"
            SS_CODE = "20099826"
        End If
                                                                
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "095056" Then
            MTS_CODE = "20006876"
            SS_CODE = ""
        End If
                                                                
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "095298" Then
            MTS_CODE = "20006876"
            SS_CODE = ""
        End If
                                                                
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "094626" Then
            MTS_CODE = "20513770"
            SS_CODE = ""
        End If
                                                                
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "095060" Then
            MTS_CODE = "29246"
            SS_CODE = ""
        End If
                                                                
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "006460" Then
            MTS_CODE = "20064371"
            SS_CODE = ""
        End If
                                                                
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "095059" Then
            MTS_CODE = "29246"
            SS_CODE = ""
        End If
                                                                
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "095058" Then
            MTS_CODE = "29246"
            SS_CODE = ""
        End If
        
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "095057" Then
            MTS_CODE = "29246"
            SS_CODE = ""
        End If
                                                                
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "094627" Then
            MTS_CODE = "20513770"
            SS_CODE = ""
        End If
        
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "095061" Then
            MTS_CODE = "29246"
            SS_CODE = ""
        End If
                                                                
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "095062" Then
            MTS_CODE = "29246"
            SS_CODE = ""
        End If
                                                                
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "095486" Then
            MTS_CODE = "20054433"
            SS_CODE = ""
        End If
                                                                
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "095487" Then
            MTS_CODE = "20054433"
            SS_CODE = ""
        End If
                                                                
                                                                
                                                                '���Ӑ�R�[�h�i�j�d�x�j
        Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, MTS_CODE)
        Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, SS_CODE)        '������R�[�h�i�j�d�x�j
                                                                '�o�ד��t�i�j�d�x�j
        Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, StrConv(OLD_Y_SYUREC.DEN_DT, vbUnicode))
        Select Case StrConv(OLD_Y_SYUREC.JGYOBU, vbUnicode)     '���Ə�
            Case SOJIKI         '�|���@
                Call UniCode_Conv(Y_SYUREC.JGYOBA, "00023210")
            Case DENKA          '�d������
                Call UniCode_Conv(Y_SYUREC.JGYOBA, "00023510")
            Case SUIHAN         '���ъ�
                Call UniCode_Conv(Y_SYUREC.JGYOBA, "00023410")
            Case SENTAKU        '����@�i�A�C�����j
                Call UniCode_Conv(Y_SYUREC.JGYOBA, "00023100")
        End Select
        Call UniCode_Conv(Y_SYUREC.DATA_KBN, "1")               '�f�[�^�敪�i�P�F����j
        Call UniCode_Conv(Y_SYUREC.TORI_KBN, "25")              '����敪
                                                                '�`�[�h�c�i���`�[�ԍ��j
        Call UniCode_Conv(Y_SYUREC.ID_NO, Format(CLng(StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode)), "00000000"))
                                                                '�i�ڔԍ�
        Call UniCode_Conv(Y_SYUREC.HIN_NO, StrConv(OLD_Y_SYUREC.HIN_GAI, vbUnicode))
                                                                '�`�[�ԍ�
        Call UniCode_Conv(Y_SYUREC.DEN_NO, StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode))
                                                                '�o�א���
        Call UniCode_Conv(Y_SYUREC.SURYO, Format(CLng(StrConv(OLD_Y_SYUREC.YOTEI_QTY, vbUnicode)), "0000000"))
                                                                '���Ӑ�R�[�h
        Call UniCode_Conv(Y_SYUREC.MUKE_CODE, MTS_CODE)
        Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, "")             '�o�Ɏ��x
                                                                '�o�ד��t
        Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, StrConv(OLD_Y_SYUREC.DEN_DT, vbUnicode))
        Call UniCode_Conv(Y_SYUREC.ODER_NO, "")                 '�I�[�_�[�ԍ�
        Call UniCode_Conv(Y_SYUREC.ITEM_NO, "")                 '�A�C�e���ԍ�
                                                                '���Ӑ於��
        Call UniCode_Conv(Y_SYUREC.MUKE_NAME, StrConv(OLD_Y_SYUREC.SYUK_NAME, vbUnicode))
                                                                '�����敪
        Call UniCode_Conv(Y_SYUREC.CYU_KBN, StrConv(OLD_Y_SYUREC.HS_CYU_KBN, vbUnicode))
        Select Case StrConv(OLD_Y_SYUREC.HS_CYU_KBN, vbUnicode) '�����敪����
            Case CYU_KBN_TUK        '����
                Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, CYU_KBN_1)
            Case CYU_KBN_SPO        '�X�|�b�g
                Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, CYU_KBN_2)
            Case CYU_KBN_HJU        '��[
                Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, CYU_KBN_3)
            Case CYU_KBN_TOK        '����
                Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, CYU_KBN_4)
            Case CYU_KBN_KIN        '�ً}
                Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, CYU_KBN_T)
            Case CYU_KBN_BOU        '�f��
                Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, CYU_KBN_E)
        End Select
        Call UniCode_Conv(Y_SYUREC.EXPORT_KBN, "")              '�A�o�o�׌����敪
        Call UniCode_Conv(Y_SYUREC.LABEL_ISSUE_KBN, "")         '�����x�����s�敪
        Call UniCode_Conv(Y_SYUREC.LABEL_ISSUE_UNIT, "")        '�����x�����s�P�ʐ�
        Call UniCode_Conv(Y_SYUREC.LABEL_TANKA_KBN, "")         '�����x���P���\���敪
        Call UniCode_Conv(Y_SYUREC.TANKA, "0000000.00")         '�P��
        Call UniCode_Conv(Y_SYUREC.TANKA, "0000000000")         '���z
        Call UniCode_Conv(Y_SYUREC.BIKOU2, "")                  '���l�Q
        Call UniCode_Conv(Y_SYUREC.REBATE_KBN, "")              '���x�[�g�敪
        Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, "")               '���[�敪
        Call UniCode_Conv(Y_SYUREC.ATAISA_KBN, "")              '�l���敪
        Call UniCode_Conv(Y_SYUREC.REP_KISHU, "")               '��\�@��
        Call UniCode_Conv(Y_SYUREC.NS_KANRI_NO, "")             '�m�r�Ǘ��ԍ�
        Call UniCode_Conv(Y_SYUREC.MTS_HIN_CODE, "")            '�l�s�r���i�R�[�h
        Call UniCode_Conv(Y_SYUREC.BIKOU1, "�R���o�[�g�f�[�^")   '���l�P
        Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, "")               '�����敪
        Call UniCode_Conv(Y_SYUREC.REBATE_RATE, "00.00")        '���x�[�g��
                                                                '�i��
        Call UniCode_Conv(Y_SYUREC.HIN_NAME, StrConv(OLD_Y_SYUREC.HIN_NAME, vbUnicode))
        Call UniCode_Conv(Y_SYUREC.JGYOBA_GAI, "")              '�ΊO���Ə�
        Call UniCode_Conv(Y_SYUREC.KISHU_CODE, "")              '�@��R�[�h
        Call UniCode_Conv(Y_SYUREC.SS_CODE, SS_CODE)            '������R�[�h
                                                                '�i�ԁi�����j
        Call UniCode_Conv(Y_SYUREC.HIN_NAI, StrConv(OLD_Y_SYUREC.HIN_NAI, vbUnicode))
                                                                '�z�X�g�I��
        Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(OLD_Y_SYUREC.HOST_TANA, vbUnicode))
                                                                '�o�ɕ\������t
        If StrConv(Y_SYUREC.KAN_KBN, vbUnicode) = "2" Or _
            StrConv(Y_SYUREC.KAN_KBN, vbUnicode) = "3" Or _
            StrConv(Y_SYUREC.KAN_KBN, vbUnicode) = "5" Or _
            StrConv(Y_SYUREC.KAN_KBN, vbUnicode) = "C" Or _
            StrConv(Y_SYUREC.KAN_KBN, vbUnicode) = "D" Then
            Call UniCode_Conv(Y_SYUREC.PRINT_YMD, Format(Now, "YYYYMMDD"))
        End If
                                                                '�������t
        Call UniCode_Conv(Y_SYUREC.KAN_YMD, StrConv(OLD_Y_SYUREC.KAN_DT, vbUnicode))
                                                                '���i���t
        Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, StrConv(OLD_Y_SYUREC.KENPIN_DT, vbUnicode))
                                                                '������敪
        Call UniCode_Conv(Y_SYUREC.TOK_KBN, StrConv(OLD_Y_SYUREC.TOK_KBN, vbUnicode))
                                                                '���яo�ɐ�
        Call UniCode_Conv(Y_SYUREC.JITU_SURYO, Format(CLng(StrConv(OLD_Y_SYUREC.FIX_QTY, vbUnicode)), "0000000"))
                                                                        
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
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

    Cnt(3).Caption = Format(Count, "#0")


'---------------------------------------------  �I��
Update_End:
    
    Update_Proc = False

End Function

Private Sub Form_Activate()

Dim ans As Integer
                                
                                '�����I��
    Beep
    ans = MsgBox("���s���܂����H", vbYesNo + vbQuestion, "�m�F����")
    If ans = vbYes Then
        If Update_Proc() Then
            Unload Me
        End If
    End If
    MsgBox "�I�����܂����B"
    Unload Me

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
'    If ITEM_Open(BtOpenNomal) Then
'        Unload Me
'    End If
                                '�i���j�i�ڃ}�X�^�n�o�d�m
'    If OLD_ITEM_Open(BtOpenNomal) Then
'        Unload Me
'    End If
    
                                '�݌Ƀf�[�^�n�o�d�m
'    If ZAIKO_Open(BtOpenNomal) Then
'        Unload Me
'    End If
                                '�i���j�݌Ƀf�[�^�n�o�d�m
    
'    If OLD_ZAIKO_Open(BtOpenNomal) Then
'        Unload Me
'    End If
                                '�݌Ɉړ����f�[�^�n�o�d�m
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i���j�݌Ɉړ����f�[�^�n�o�d�m
    If OLD_IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�o�ח\��f�[�^�n�o�d�m
'    If Y_SYU_Open(BtOpenNomal) Then
'        Unload Me
'    End If
                                '�i���j�o�ח\��f�[�^�n�o�d�m
'    If OLD_Y_SYU_Open(BtOpenNomal) Then
'        Unload Me
'    End If
    
                    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '�i�ڃ}�X�^�b�k�n�r�d
 '   sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
 '   If sts Then
 '       If sts <> BtErrNoOpen Then
'            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
 '       End If
 '   End If
                                            '(��)�i�ڃ}�X�^CLOSE
  '  sts = BTRV(BtOpClose, OLD_ITEM_POS, OLD_ITEMREC, Len(OLD_ITEMREC), K0_OLD_ITEM, Len(K0_OLD_ITEM), 0)
  '  If sts Then
  '      If sts <> BtErrNoOpen Then
  '          Call File_Error(sts, BtOpClose, "�i���j�i�ڃ}�X�^")
  '      End If
  '  End If
                                            '�݌Ƀf�[�^�b�k�n�r�d
   ' sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
   ' If sts Then
   '     If sts <> BtErrNoOpen Then
   '         Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
   '     End If
   ' End If
                                            '(��)�݌Ƀf�[�^�b�k�n�r�d
   ' sts = BTRV(BtOpClose, OLD_ZAIKO_POS, OLD_ZAIKOREC, Len(OLD_ZAIKOREC), K0_OLD_ZAIKO, Len(K0_OLD_ZAIKO), 0)
   ' If sts Then
   '     If sts <> BtErrNoOpen Then
   '         Call File_Error(sts, BtOpClose, "(��)�݌Ƀf�[�^")
   '     End If
   ' End If
    
                                            '�݌Ɉړ����b�k�n�r�d
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ɉړ���")
        End If
    End If
                                            '(��)�݌Ɉړ����b�k�n�r�d
    sts = BTRV(BtOpClose, OLD_IDO_POS, OLD_IDOREC, Len(OLD_IDOREC), K0_OLD_IDO, Len(K0_OLD_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(��)�݌Ɉړ���")
        End If
    End If
                                            '�o�ח\��f�[�^�b�k�n�r�d
    'sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    'If sts Then
    '    If sts <> BtErrNoOpen Then
    '        Call File_Error(sts, BtOpClose, "�o�ח\��f�[�^")
    '    End If
    'End If
                                            '(��)�o�ח\��f�[�^�b�k�n�r�d
    'sts = BTRV(BtOpClose, OLD_Y_SYU_POS, OLD_Y_SYUREC, Len(OLD_Y_SYUREC), K0_OLD_Y_SYU, Len(K0_OLD_Y_SYU), 0)
    'If sts Then
    '    If sts <> BtErrNoOpen Then
    '        Call File_Error(sts, BtOpClose, "(��)�o�ח\��f�[�^")
    '    End If
    'End If
    
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set CONV20041 = Nothing

    End
End Sub

