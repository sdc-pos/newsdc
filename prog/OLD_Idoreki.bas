Attribute VB_Name = "OLD_IDO"
Option Explicit
'********************************************************************
'*
'*              (��)�݌Ɉړ����@�t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const OLD_IDO_ID$ = "OLD_IDO"

'�y�[�W�T�C�Y
Public Const OLD_IDO_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public OLD_IDO_POS As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type OLD_IDOREC_Tag
    JITU_DT(0 To 7)                     As Byte     '���ѓ��t
    JITU_TM(0 To 5)                     As Byte     '���ю���
    JGYOBU(0 To 0)                      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)                      As Byte     '�����O
    HIN_GAI(0 To 19)                    As Byte     '�i�ԁi�O���j
    RIRK_ID(0 To 1)                     As Byte     '�������
    SUMI_JITU_QTY(0 To 7)               As Byte     '���ѐ���(���i���ς�)
    MI_JITU_QTY(0 To 7)                 As Byte     '���ѐ���(�����i)
    FROM_SOKO(0 To 1)                   As Byte     'From �q�ɇ�
    FROM_RETU(0 To 1)                   As Byte     '   �@��
    FROM_REN(0 To 1)                    As Byte     '   �@�A
    FROM_DAN(0 To 1)                    As Byte     '   �@�i
    TO_SOKO(0 To 1)                     As Byte     '�s�n �q�ɇ�
    TO_RETU(0 To 1)                     As Byte     '   �@��
    TO_REN(0 To 1)                      As Byte     '   �@�A
    TO_DAN(0 To 1)                      As Byte     '   �@�i
    DEN_DT(0 To 7)                      As Byte     '�`�[���t
    DEN_NO(0 To 9)                      As Byte     '�`�[��
    PRG_ID(0 To 7)                      As Byte     '�o�͌��v���O����
    HIN_NAI(0 To 19)                    As Byte     '�i�ԁi�����j
    NYUKA_DT(0 To 7)                    As Byte     '���ד��t
    NYUKO_DT(0 To 7)                    As Byte     '���ɓ��t
    WEL_ID(0 To 2)                      As Byte     '�Ώے[����
    RIRK_NAME(0 To 9)                   As Byte     '������ʖ���
    HIN_NAME(0 To 24)                   As Byte     '�i��
    SUMI_HIN_Zaiko_Qty(0 To 7)          As Byte     '�i�ڕʍ݌ɐ��i���i���ς݁j
    MI_HIN_Zaiko_Qty(0 To 7)            As Byte     '�i�ڕʍ݌ɐ��i�����i�j
    SUMI_FROM_TANA_Zaiko_Qty(0 To 7)    As Byte     'FROM�I�ʕi�ڕʍ݌ɐ�
    SUMI_TO_TANA_Zaiko_Qty(0 To 7)      As Byte     'TO�I�ʕi�ڕʍ݌ɐ�
    MI_FROM_TANA_Zaiko_Qty(0 To 7)      As Byte     'FROM�I�ʕi�ڕʍ݌ɐ�
    MI_TO_TANA_Zaiko_Qty(0 To 7)        As Byte     'TO�I�ʕi�ڕʍ݌ɐ�
    TOKU_MARK(0 To 0)                   As Byte     '������}�[�N
    MEMO(0 To 59)                       As Byte     '����
    TANTO_CODE(0 To 4)                  As Byte     '�S���҃R�[�h
    TANTO_NAME(0 To 19)                 As Byte     '�S���Җ���
    MUKE_CODE(0 To 7)                   As Byte     '���Ӑ�R�[�h
    MUKE_NAME(0 To 39)                  As Byte     '���Ӑ於��
    SS_CODE(0 To 7)                     As Byte     '������R�[�h
    SS_NAME(0 To 39)                    As Byte     '�����於��
    MUKE_DNAME(0 To 9)                  As Byte     '���Ӑ旪��
    MUKE_CHG_CD(0 To 1)                 As Byte     '������Ǒւ��R�[�h
    SUM_KBN(0 To 0)                     As Byte     '�W�v�敪
    ID_NO(0 To 7)                       As Byte     'ID-NO
    
    Ins_DateTime(0 To 13)               As Byte     '�}������2004.12.09
    FILLER(0 To 76)                     As Byte
    
End Type

'�f�[�^�E�o�b�t�@
Public OLD_IDOREC                       As OLD_IDOREC_Tag

'�L�[��`
Type KEY0_OLD_IDO            '�j�d�x�O
    JGYOBU(0 To 0)                      As Byte     '���ƕ��敪
    JITU_DT(0 To 7)                     As Byte     '���ѓ��t
    JITU_TM(0 To 5)                     As Byte     '���ю���
End Type
'�L�[�E�f�[�^
Public K0_OLD_IDO                       As KEY0_OLD_IDO

Public Function OLD_IDO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �i���j�݌Ɉړ����@�n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    OLD_IDO_Open = True
                                            '�݌Ɉړ����t���p�X�捞��
    sts = GetIni("FILE", OLD_IDO_ID, "CONV2006", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "CONV2006.INI [OLD_IDO]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, OLD_IDO_POS, OLD_IDOREC, Len(OLD_IDOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
            
                OLD_IDO_Open = sts
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "�i���j�݌Ɉړ���")
                Exit Function
        End Select
    Loop
    OLD_IDO_Open = False
End Function


