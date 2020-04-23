Attribute VB_Name = "OLD_ITEM2"
Option Explicit
'********************************************************************
'*
'*              �i�ڃ}�X�^  �t�@�C����`
'*
'*          CREATE 2004.02.19
'********************************************************************
'�t�@�C���h�c
Public Const OLD_ITEM2_ID$ = "OLD_ITEM2"

'�y�[�W�T�C�Y
Public Const OLD_ITEM2_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public OLD_ITEM2_POS         As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type OLD_ITEM2REC_Tag
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    HIN_GAI(0 To 12)    As Byte     '�i�ԁi�O���j
    HIN_NAME(0 To 24)   As Byte     '�i��
    ST_SET_DT(0 To 7)   As Byte     '�W���q�ɐݒ���t
    ST_SOKO(0 To 1)     As Byte     '�W�����ɑq�� �q��
    ST_RETU(0 To 1)     As Byte     '             ��
    ST_REN(0 To 1)      As Byte     '             �A
    ST_DAN(0 To 1)      As Byte     '             �i
    BEF_SOKO(0 To 1)    As Byte     '�O����ɑq�� �q��
    BEF_RETU(0 To 1)    As Byte     '             ��
    BEF_REN(0 To 1)     As Byte     '             �A
    BEF_DAN(0 To 1)     As Byte     '             �i
    LAST_NYU_DT(0 To 7) As Byte     '�ŏI���ɓ��t
    LAST_SYU_DT(0 To 7) As Byte     '�ŏI�o�ɓ��t
    HIN_NAI(0 To 12)    As Byte     '�i�ԁi�����j
    BIKOU_SOKO(0 To 1)  As Byte     '���l �z�X�g�q��
    BIKOU_TANA(0 To 7)  As Byte     '���l �z�X�g�I��
    SIZAI_CD(0 To 4)    As Byte     '���ރR�[�h
    HOJYU_P(0 To 7)     As Byte     '��[�_
    AVE_SYUKA(0 To 7)   As Byte     '�����Ϗo�א�
    SAMPLE_QTY(0 To 0)  As Byte     '�T���v����
    LAST_INP_DT(0 To 7) As Byte     '�ŏI���ד��t
'*------------------------------------------ 2001.02.15 �ǉ� ��
    LOCK_F(0 To 0)      As Byte     '�r���t���O
    WEL_ID(0 To 2)      As Byte     '�g�p�q�@ID
    PRG_ID(0 To 7)      As Byte     '�g�p���v���O����
'*------------------------------------------ 2001.02.15 �ǉ� ��
    LAST_CHK_DT(0 To 7) As Byte     '�ŏI�ƍ����t2001.06.12
    LAST_CHK_QTY(0 To 7) As Byte    '�ŏI�ƍ����݌ɐ�2001.06.12
    MOTO_JIGYOBU(0 To 0) As Byte    '�������ƕ�     '���g�p2004.02
    BIKOU(0 To 14)      As Byte     '������l
    IRI_QTY(0 To 7)     As Byte     '������萔
    
    JAN_CODE(0 To 12)   As Byte     'Jan�R�[�h      2004.02
    HIN_CHANGE(0 To 12) As Byte     '�i�ԓǂݑւ�   2004.02
    GOODS_KBN(0 To 0)   As Byte     '���i���L��     2004.02
    PACKING_NO(0 To 3)  As Byte     '������       2004.02
    
    FILLER(0 To 167)    As Byte     'FILLER         2004.02
End Type
'�f�[�^�E�o�b�t�@
Public OLD_ITEM2REC As OLD_ITEM2REC_Tag

'�L�[��`

Type KEY0_OLD_ITEM2            '�j�d�x�O
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    HIN_GAI(0 To 12)    As Byte     '�i�ԁi�O���j
End Type



'�L�[�E�f�[�^
Public K0_OLD_ITEM2 As KEY0_OLD_ITEM2

Public Function OLD_ITEM2_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �i�ڃ}�X�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    OLD_ITEM2_Open = True
                                            '�i�ڃ}�X�^�t���p�X�捞��
    sts = GetIni("FILE", OLD_ITEM2_ID, "SETUP", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SETUP.INI [OLD_ITEM2]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, OLD_ITEM2_POS, OLD_ITEM2REC, Len(OLD_ITEM2REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "�i�ڃ}�X�^")
                Exit Function
        End Select
    Loop

    OLD_ITEM2_Open = False

End Function


