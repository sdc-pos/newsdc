Attribute VB_Name = "OLD_ITEM"
Option Explicit
'********************************************************************
'*
'*              �i���j�i�ڃ}�X�^  �t�@�C����`
'*
'*          CREATE 2005.12.02
'********************************************************************
'�t�@�C���h�c
Public Const OLD_ITEM_ID$ = "OLD_ITEM"

'�y�[�W�T�C�Y
Public Const OLD_ITEM_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public OLD_ITEM_POS     As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type OLD_ITEMREC_Tag
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
    RANK(0 To 2)        As Byte     '���݃����N     2004.06
    NEW_RANK(0 To 2)    As Byte     '���݃����N     2004.06
    GLICS1_TANA(0 To 9) As Byte     '�O���b�N�X�I�ԂP   2005.05
    GLICS2_TANA(0 To 9) As Byte     '�O���b�N�X�I�ԂQ   2005.05
    GLICS3_TANA(0 To 9) As Byte     '�O���b�N�X�I�ԂR   2005.05
    
    
    
    FILLER(0 To 131)    As Byte     'FILLER         2005.05
End Type
'�f�[�^�E�o�b�t�@
Public OLD_ITEMREC      As OLD_ITEMREC_Tag

'�L�[��`

Type KEY0_OLD_ITEM                  '�j�d�x�O
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    HIN_GAI(0 To 12)    As Byte     '�i�ԁi�O���j
End Type




'�L�[�E�f�[�^
Public K0_OLD_ITEM      As KEY0_OLD_ITEM

Public Function OLD_ITEM_Open(Mode As Integer) As Integer
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
    
    OLD_ITEM_Open = True
                                            '�i�ڃ}�X�^�t���p�X�捞��
    sts = GetIni("FILE", OLD_ITEM_ID, "CONV2006", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "CONV2006.INI [OLD_ITEM]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, OLD_ITEM_POS, OLD_ITEMREC, Len(OLD_ITEMREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
            
                OLD_ITEM_Open = sts
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "(��)�i�ڃ}�X�^")
                Exit Function
        End Select
    Loop

    OLD_ITEM_Open = False

End Function


