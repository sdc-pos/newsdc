Attribute VB_Name = "OLD_IDO2"
Option Explicit
'********************************************************************
'*
'*              �݌Ɉړ����@�t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const OLD_IDO2_ID$ = "OLD_IDO2"

'�y�[�W�T�C�Y
Public Const OLD_IDO2_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public OLD_IDO2_POS As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type OLD_IDO2REC_Tag
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
    FILLER(0 To 90)                     As Byte
    
End Type

'�f�[�^�E�o�b�t�@
Public OLD_IDO2REC   As OLD_IDO2REC_Tag

'�L�[��`
Type KEY0_OLD_IDO2            '�j�d�x�O
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    JITU_DT(0 To 7)             As Byte     '���ѓ��t
    JITU_TM(0 To 5)             As Byte     '���ю���
End Type

'�L�[�E�f�[�^
Public K0_OLD_IDO2                   As KEY0_OLD_IDO2

Public Function OLD_IDO2_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �݌Ɉړ����@�n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    OLD_IDO2_Open = True
                                            '�݌Ɉړ����t���p�X�捞��
    sts = GetIni("FILE", OLD_IDO2_ID, "SETUP", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SETUP.INI [OLD_IDO2]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, OLD_IDO2_POS, OLD_IDO2REC, Len(OLD_IDO2REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "�݌Ɉړ���")
                Exit Function
        End Select
    Loop
    OLD_IDO2_Open = False
End Function


