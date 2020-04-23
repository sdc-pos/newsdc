Attribute VB_Name = "RECOV_ITEM"
Option Explicit
'********************************************************************
'*
'*              �i���j�i�ڃ}�X�^  �t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Global Const RECOV_ITEM_ID = "RECOV_ITEM"

'�y�[�W�T�C�Y
Global Const RECOV_ITEM_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Global RECOV_ITEM_POS As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type RECOV_ITEMREC_Tag
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
    WEL_ID(0 To 1)      As Byte     '�g�p�q�@ID
    PRG_ID(0 To 7)      As Byte     '�g�p���v���O����
'*------------------------------------------ 2001.02.15 �ǉ� ��
    LAST_CHK_DT(0 To 7) As Byte     '�ŏI�ƍ����t2001.06.12
    LAST_CHK_QTY(0 To 7) As Byte    '�ŏI�ƍ����݌ɐ�2001.06.12
    MOTO_JIGYOBU(0 To 0) As Byte    '�������ƕ�
    BIKOU(0 To 14)      As Byte     '������l
    IRI_QTY(0 To 7)     As Byte     '������萔
    FILLER(0 To 7)     As Byte      'FILLER
End Type
'�f�[�^�E�o�b�t�@
Global RECOV_ITEMREC As RECOV_ITEMREC_Tag

'�L�[��`

Type KEY0_RECOV_ITEM            '�j�d�x�O
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    HIN_GAI(0 To 12)    As Byte     '�i�ԁi�O���j
End Type


'�L�[�E�f�[�^
Global K0_RECOV_ITEM As KEY0_RECOV_ITEM

Type RECOV_ITEM_FSpeck
    fs As BtFileSpeck                 ' ̧�� ��߯��\����
    ks0 As BtKeySpeck                 ' �� ��߯��\����
End Type

Global RECOV_ITEM_Speck As RECOV_ITEM_FSpeck
 

Function RECOV_ITEM_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              �i���j�i�ڃ}�X�^  �n�o�d�m                          *
'*                                                                  *
'*      ��  ��:Open Mode(Btrieve�Q��)                               *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String
    
    RECOV_ITEM_Open = False
                                            '�i�ڃ}�X�^�t���p�X�捞��
    sts = GetIni("FILE", RECOV_ITEM_ID, "SETUP", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SETUP.INI �ǂݍ��݃G���[")
        RECOV_ITEM_Open = True
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, RECOV_ITEM_POS, RECOV_ITEMREC, Len(RECOV_ITEMREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Function
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "�i�ڃ}�X�^")
                RECOV_ITEM_Open = True
                Exit Function
        End Select
    Loop
End Function


