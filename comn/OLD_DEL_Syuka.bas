Attribute VB_Name = "OLD_DEL_SYU"
Option Explicit
'********************************************************************
'*
'*              �i���j�폜�ςݏo�ח\��f�[�^  �t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const OLD_DEL_SYU_ID$ = "OLD_DEL_SYU"

'�y�[�W�T�C�Y
Public Const OLD_DEL_SYU_PG_SIZ% = 2048

'�|�W�V�����E�u���b�N
Public OLD_DEL_SYU_POS    As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type OLD_DEL_SYUREC_Tag
    WEL_ID(0 To 2)              As Byte     '�g�p�q�@ID
    PRG_ID(0 To 7)              As Byte     '�g�p���v���O����
    KAN_KBN(0 To 0)             As Byte     '�����敪
    DT_SYU(0 To 0)              As Byte     '�f�[�^���
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    KEY_CYU_KBN(0 To 0)         As Byte     '�����敪
    KEY_ID_NO(0 To 7)           As Byte     'ID-NO
    NAIGAI(0 To 0)              As Byte     '�����O
    KEY_HIN_NO(0 To 19)         As Byte     '�i�ڔԍ�
    KEY_MUKE_CODE(0 To 7)       As Byte     '���Ӑ�R�[�h
    KEY_SS_CODE(0 To 7)         As Byte     '������R�[�h
    KEY_SYUKA_YMD(0 To 7)       As Byte     '�o�ד��t
    JGYOBA(0 To 7)              As Byte     '���Ə�
    DATA_KBN(0 To 0)            As Byte     '�f�[�^�敪
    TORI_KBN(0 To 1)            As Byte     '����敪
    ID_NO(0 To 7)               As Byte     'ID-NO
    HIN_NO(0 To 19)             As Byte     '�i�ڔԍ�
    DEN_NO(0 To 9)              As Byte     '�`�[�ԍ�
    SURYO(0 To 6)               As Byte     '�o�ɐ���
    MUKE_CODE(0 To 7)           As Byte     '���Ӑ�R�[�h
    SYUKO_SYUSI(0 To 1)         As Byte     '�o�Ɏ��x
    SYUKA_YMD(0 To 7)           As Byte     '�o�ד��t
    ODER_NO(0 To 11)            As Byte     '�I�[�_�[�ԍ�
    ITEM_NO(0 To 4)             As Byte     '�A�C�e���ԍ�
    MUKE_NAME(0 To 23)          As Byte     '���Ӑ於��
    CYU_KBN(0 To 0)             As Byte     '�����敪
    CYU_KBN_NAME(0 To 9)        As Byte     '�����敪����
    EXPORT_KBN(0 To 0)          As Byte     '�A�o�o�׌����敪
    LABEL_ISSUE_KBN(0 To 0)     As Byte     '�����x�����s�敪
    LABEL_ISSUE_UNIT(0 To 4)    As Byte     '�����x�����s�P�ʐ�
    LABEL_TANKA_KBN(0 To 0)     As Byte     '�����x���P���\���敪
    TANKA(0 To 9)               As Byte     '�P��
    KINGAKU(0 To 9)             As Byte     '���z
    BIKOU2(0 To 19)             As Byte     '���l�Q
    REBATE_KBN(0 To 0)          As Byte     '���x�[�g�敪
    CHOHA_KBN(0 To 0)           As Byte     '���[�敪
    ATAISA_KBN(0 To 0)          As Byte     '�l���敪
    REP_KISHU(0 To 19)          As Byte     '��\�@��
    NS_KANRI_NO(0 To 8)         As Byte     '�m�r�Ǘ��ԍ�
    MTS_HIN_CODE(0 To 10)       As Byte     '�l�s�r���i�R�[�h
    BIKOU1(0 To 39)             As Byte     '���l�P
    CHOKU_KBN(0 To 0)           As Byte     '�����敪
    REBATE_RATE(0 To 4)         As Byte     '���x�[�g��
    HIN_NAME(0 To 19)           As Byte     '�i��
    JGYOBA_GAI(0 To 7)          As Byte     '�ΊO���Ə�
    KISHU_CODE(0 To 2)          As Byte     '�@��R�[�h
    SS_CODE(0 To 7)             As Byte     '������R�[�h
    HIN_NAI(0 To 12)            As Byte     '�i�ԁi�����j
    HTANABAN(0 To 7)            As Byte     '�z�X�g�I��
    PRINT_YMD(0 To 7)           As Byte     '�o�ɕ\������t
    KAN_YMD(0 To 7)             As Byte     '�������t
    KENPIN_YMD(0 To 7)          As Byte     '���i���t
    TOK_KBN(0 To 0)             As Byte     '������敪
    JITU_SURYO(0 To 6)          As Byte     '�o�Ɏ��ѐ���
    INS_NOW(0 To 13)            As Byte     '�捞�ݓ���
    FILLER(0 To 74)             As Byte     'FILLER
End Type

'�f�[�^�E�o�b�t�@
Public OLD_DEL_SYUREC           As OLD_DEL_SYUREC_Tag

'�L�[��`
Type KEY0_OLD_DEL_SYU            '�j�d�x�O
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    KEY_CYU_KBN(0 To 0)         As Byte     '�����敪
    KEY_MUKE_CODE(0 To 7)       As Byte     '���Ӑ�R�[�h
    KEY_SS_CODE(0 To 7)         As Byte     '������R�[�h
    KEY_HIN_NO(0 To 19)         As Byte     '�i�ڔԍ�
    KEY_SYUKA_YMD(0 To 7)       As Byte     '�o�ד��t
End Type



'�L�[�E�f�[�^
Public K0_OLD_DEL_SYU               As KEY0_OLD_DEL_SYU

Function OLD_DEL_SYU_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*             �i���j�폜�ςݏo�ח\��f�[�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    OLD_DEL_SYU_Open = True
                                            '�폜�ςݏo�ח\��f�[�^�t���p�X�捞��
    sts = GetIni("FILE", OLD_DEL_SYU_ID, "CONV2006", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "CONV2006.INI [OLD_DEL_SYU]�ǂݍ��݃G���[ ")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, OLD_DEL_SYU_POS, OLD_DEL_SYUREC, Len(OLD_DEL_SYUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            
            Case BtErrFileNotFound
            
                OLD_DEL_SYU_Open = sts
            
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "�i���j�폜�ςݏo�ח\��f�[�^")
                Exit Function
        End Select
    Loop
    
    OLD_DEL_SYU_Open = False

End Function
