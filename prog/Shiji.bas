Attribute VB_Name = "HS_SIJ"
Option Explicit
'********************************************************************
'*
'*              �z�X�g��M�f�[�^ �t�@�C����`
'*
'*          CREATE 2004.03.04
'********************************************************************
'�t�@�C���h�c
Public Const HS_IN_SIJ_ID$ = "HS_IN_SIJ"
Public Const HS_OUT_SIJ_ID$ = "HS_OUT_SIJ"
'�t�@�C����
Public HS_SIJ_No As Integer
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`(����)
Type HS_IN_SIJREC_Tag
    
    
    
    TEXT_NO(0 To 8) As Byte         '÷�ć�
    JGYOBU(0 To 0) As Byte          '���ƕ��敪
    CYOK_KBN(0 To 0) As Byte        '�����敪
    DEN_DT(0 To 7) As Byte          '�`�[���t
    IO_KBN(0 To 0) As Byte          '���o�ɋ敪
    PM_KBN(0 To 0) As Byte          '�ԍ��敪
    DEN_SYU(0 To 0) As Byte         '�`�[���
    DEN_NO(0 To 5) As Byte          '�`�[��
    CYU_KBN(0 To 0) As Byte         '�����敪
    HIN_GAI(0 To 12) As Byte        '�i�ԁi�O���j
    HIN_NAI(0 To 12) As Byte        '�i�ԁi�����j
    HIN_NAME(0 To 24) As Byte       '�i��
    YOTEI_QTY(0 To 5) As Byte       '����
    YOSAN_FROM(0 To 4) As Byte      '�\�Z�P�ʁi���j
    YOSAN_TO(0 To 4) As Byte        '�\�Z�P�ʁi��j
    HOST_SOKO(0 To 1) As Byte       '�q�ɋ敪�iνāj
    HOST_TANA(0 To 7) As Byte       '�I�ԁiνāj
    SYUK_CODE(0 To 4) As Byte       '�x����^�o�א�
    SYUK_NAME(0 To 19) As Byte      '�x����^�o�א於
    REC_END(0 To 0) As Byte         'ں��ޏI�[ϰ�(@)
    CR_LF(0 To 1) As Byte           'CR.LF
    
    
    
End Type




'�f�[�^�E�o�b�t�@
Public HS_IN_SIJREC As HS_IN_SIJREC_Tag
'-------------------------------------------'
'���R�[�h��`(�o��)
Type HS_OUT_SIJREC_Tag
'    JGYOBA(0 To 7)              As Byte     '���Ə�
'    DATA_KBN(0 To 0)            As Byte     '�f�[�^�敪
'    TORI_KBN(0 To 1)            As Byte     '����敪
'    ID_NO(0 To 7)               As Byte     'ID-NO
'    HIN_NO(0 To 19)             As Byte     '�i�ڔԍ�
'    DEN_NO(0 To 9)              As Byte     '�`�[�ԍ�
'    SURYO(0 To 6)               As Byte     '�o�ɐ���
'    MUKE_CODE(0 To 7)           As Byte     '���Ӑ�R�[�h
'    SYUKO_SYUSI(0 To 1)         As Byte     '�o�Ɏ��x
'    SYUKA_YMD(0 To 7)           As Byte     '�o�ד��t
'    ODER_NO(0 To 11)            As Byte     '�I�[�_�[�ԍ�
'    ITEM_NO(0 To 4)             As Byte     '�A�C�e���ԍ�
'    MUKE_NAME(0 To 23)          As Byte     '���Ӑ於��
'    CHU_KBN(0 To 0)             As Byte     '�����敪
'    CHU_KBN_NAME(0 To 9)        As Byte     '�����敪����
'    EXPORT_KBN(0 To 0)          As Byte     '�A�o�o�׌����敪
'    LABEL_ISSUE_KBN(0 To 0)     As Byte     '�����x�����s�敪
'    LABEL_ISSUE_UNIT(0 To 4)    As Byte     '�����x�����s�P�ʐ�
'    LABEL_TANKA_KBN(0 To 0)     As Byte     '�����x���P���\���敪
'    TANKA(0 To 9)               As Byte     '�P��
'    KINGAKU(0 To 9)             As Byte     '���z
'    BIKOU2(0 To 19)             As Byte     '���l�Q
'    REBATE_KBN(0 To 0)          As Byte     '���x�[�g�敪
'    CHOHA_KBN(0 To 0)           As Byte     '���[�敪
'    ATAISA_KBN(0 To 0)          As Byte     '�l���敪
'    REP_KISHU(0 To 19)          As Byte     '��\�@��
'    NS_KANRI_NO(0 To 8)         As Byte     '�m�r�Ǘ��ԍ�
'    MTS_HIN_CODE(0 To 10)       As Byte     '�l�s�r���i�R�[�h
'    BIKOU1(0 To 39)             As Byte     '���l�P
'    CHOKU_KBN(0 To 0)           As Byte     '�����敪
'    REBATE_RATE(0 To 4)         As Byte     '���x�[�g��
'    HIN_NAME(0 To 19)           As Byte     '�i��
'    JGYOBU_GAI(0 To 7)          As Byte     '�ΊO���Ə�
'    SS_CODE(0 To 7)             As Byte     '������R�[�h
'    KISHU_HIN_NO(0 To 2)        As Byte     '�@��i�ڃR�[�h
'    HIN_NAI(0 To 19)            As Byte     '�i�ԁi�����j
'    CRLF(0 To 1)                As Byte     'CRLF

    JGYOBA(0 To 7)              As Byte     '���Ə�
    DATA_KBN(0 To 0)            As Byte     '�f�[�^�敪
    TORI_KBN(0 To 1)            As Byte     '����敪
    ID_NO(0 To 7)               As Byte     'ID-NO
    HIN_NO(0 To 19)             As Byte     '�i�ڔԍ�
    DEN_NO(0 To 9)              As Byte     '�`�[�ԍ�
    SURYO(0 To 6)               As Byte     '�o�ɐ���
    MUKE_CODE(0 To 7)           As Byte     '�o�ɐ�
    SYUKO_SYUSI(0 To 1)         As Byte     '�o�Ɏ��x
    SYUKO_YMD(0 To 7)           As Byte     '�o�ɓ��t
    TANKA(0 To 9)               As Byte     '�P��
    ODER_NO(0 To 11)            As Byte     '�I�[�_�[�ԍ�
    ITEM_NO(0 To 4)             As Byte     '�A�C�e���ԍ�
    ODER_R_NO(0 To 4)           As Byte     '�I�[�_�[����
    KOSO_KEITAI(0 To 9)         As Byte     '���`��
    SYUKA_YMD(0 To 7)           As Byte     '�o�ד�
    TANABAN1(0 To 9)            As Byte     '�I�ԂP
    TANABAN2(0 To 9)            As Byte     '�I�ԂQ
    TANABAN3(0 To 9)            As Byte     '�I�ԂR
    MUKE_NAME(0 To 23)          As Byte     '�o�ɐ於��
    CHU_KBN(0 To 0)             As Byte     '�����敪
    CHU_KBN_NAME(0 To 9)        As Byte     '�����敪����
    ORIGIN1(0 To 9)             As Byte     '���Y���P
    ORIGIN2(0 To 9)             As Byte     '���Y���Q
    BIKOU2(0 To 39)             As Byte     '���l�Q
    HAN_KBN(0 To 0)             As Byte     '�̔��敪
    CHOKU_KBN(0 To 0)           As Byte     '�����敪
    UNIT_ID_NO(0 To 7)          As Byte     '�ƯďC��ID-NO
    ZAIKO_HIKIATE(0 To 2)       As Byte     '�݌Ɉ�������
    GOKON_KANRI_NO(0 To 8)      As Byte     '�����Ǘ��ԍ�
    JUCHU_ZAN(0 To 6)           As Byte     '�󒍎c����
    KYOKYU_KBN(0 To 0)          As Byte     '�����敪
    SHOHIN_SYUSI(0 To 1)        As Byte     '���i���[������x
    BIKOU1(0 To 39)             As Byte     '���l�P
    CHOHA_KBN(0 To 0)           As Byte     '���[�敪
    JYU_HIN_NO(0 To 19)         As Byte     '�󒍕i�ڔԍ�
    HIN_NAME(0 To 19)           As Byte     '�i��
    HIN_CHANGE_KBN(0 To 0)      As Byte     '�i�ԕύX�敪
    MODULE_EXCHANGE(0 To 0)     As Byte     '���W���[�������敪
    ZAIKO_SYUSI(0 To 1)         As Byte     '�c�݌ɂ܂Ƃߍ݌Ɏ��x�R�[�h
    NOUKI_YMD(0 To 7)           As Byte     '�w��[��
    SERVICE_KANRI_NO(0 To 8)    As Byte     '�T�[�r�X��ЊǗ��ԍ�
    KISHU_CODE(0 To 2)          As Byte     '�@��i�ڃR�[�h
    ENVIRONMENT_KBN(0 To 0)     As Byte     '���K�i���i�敪
    SS_CODE(0 To 7)             As Byte     '������R�[�h
    FILLER(0 To 4)              As Byte
    CRLF(0 To 1)                As Byte     'CRLF









End Type

'�f�[�^�E�o�b�t�@
Public HS_OUT_SIJREC As HS_OUT_SIJREC_Tag
Public Function HS_SIJ_Open(Mode As Integer, Data_Type As Integer) As Integer
'********************************************************************
'*
'*      �z�X�g��M�f�[�^  �n�o�d�m
'*
'*      �����@:OPEN���[�h�i0:�Q�Ɓ@1:�X�V�j
'*             �ް�����   (1:���Ɂ@2:�o��)
'*
'*      �߂�l:false ����
'*             true  �ُ�
'*
'*          CREATE 2004.03.05
'********************************************************************

Dim ans         As Integer
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    On Error GoTo HS_SIJ_Op_Err     '�װ�ׯ��ON

    HS_SIJ_Open = True
                                    
    Select Case Data_Type
        Case 1          '����
            If GetIni("FILE", HS_IN_SIJ_ID, "SYS", c) Then
                Call LOG_OUT(LOG_F, "SYS.INI [HS_IN_SIJ]�ǂݍ��݃G���[")
                Exit Function
            End If
        Case 2          '�o��
            If GetIni("FILE", HS_OUT_SIJ_ID, "SYS", c) Then
                Call LOG_OUT(LOG_F, "SYS.INI [HS_OUT_SIJ]�ǂݍ��݃G���[")
                Exit Function
            End If
    End Select
                                    
    FullPath = RTrim(c)
    
    HS_SIJ_No = FreeFile

    If Mode = 0 Then
        Open FullPath For Input As #HS_SIJ_No
    Else
        Open FullPath For Binary As #HS_SIJ_No
    End If
    
    HS_SIJ_Open = False

    Exit Function

HS_SIJ_Op_Err:     '�װ����ٰ��
Const ErrDiskNotReady = 71, ErrDeviceUnavailable = 68, ErrNotFound = 53
    Select Case Err.Number
        Case ErrDiskNotReady
            If Mode = 1 Then
                Beep
                ans = MsgBox("�h���C�u���m�F���ĉ�����", vbYesNo + vbExclamation + vbDefaultButton1, "�m�F����")
                If ans = vbYes Then
                    Resume
                End If
            End If
        Case ErrDeviceUnavailable
            If Mode = 1 Then
                Beep
                ans = MsgBox("�h���C�u�܂��̓p�X��������܂���" & FullPath, vbExclamation)
            End If
        Case ErrNotFound
            If Mode = 1 Then
                Beep
                ans = MsgBox("�t�@�C����������܂���" & FullPath, vbExclamation)
            End If
        Case Else
            If Mode = 1 Then
                Beep
                ans = MsgBox("�G���[ [HS_SIJ Open : " & Str(Err.Number) & "] " & Err.Description, vbCritical)
            End If
    End Select
End Function
