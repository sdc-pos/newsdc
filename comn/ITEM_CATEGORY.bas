Attribute VB_Name = "ITEM_CATEGORY"
Option Explicit
'********************************************************************
'*
'*              �i���J�e�S���}�X�^  �t�@�C����`
'*
'*          CREATE 2011.12.07
'********************************************************************
'�t�@�C���h�c
Public Const ITEM_CATEGORY_ID$ = "ITEM_CATEGORY"

'�y�[�W�T�C�Y
Public Const ITEM_CATEGORY_PG_SIZ% = 4096

'�|�W�V�����E�u���b�N
Public ITEM_CATEGORY_POS            As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************



'���R�[�h��`
Type ITEM_CATEGORYREC_Tag
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    CATEGORY_CODE(0 To 7)       As Byte     '�i����ú�غ���
    CATEGORY_NAME(0 To 79)      As Byte     '�i����ú�ؖ���
    SEI_LOT(0 To 9)             As Byte     '���Y���b�g                 �����_��
    KOUSU_LOT(0 To 9)           As Byte     '�O��H��(�b/ۯ�)�@         �����_��
    KOUSU_QTY(0 To 9)           As Byte     '�O��H��(�b/��)�@          �����_��
    TOKU_TANKA_QTY(0 To 9)      As Byte     '���ʒP��(��ƍH���@�b/��)�@�����_��
    TOKU_TANKA_KOURYO(0 To 12)  As Byte     '���ʒP��(�H����)�@         9(10).99
    TOKU_TANKA_HAKO(0 To 12)    As Byte     '���ʒP��(���し)�@         9(10).99
    MEMO(0 To 79)               As Byte     '���l/����
    FILLER(0 To 228)            As Byte
    INS_TANTO(0 To 9)           As Byte     '�ǉ��@�S����
    Ins_DateTime(0 To 13)       As Byte     '�ǉ��@����
    UPD_TANTO(0 To 9)           As Byte     '�X�V�@�S����
    UPD_DATETIME(0 To 13)       As Byte     '�X�V�@����

End Type
'�f�[�^�E�o�b�t�@
Public ITEM_CATEGORYREC As ITEM_CATEGORYREC_Tag

'�L�[��`

Type KEY0_ITEM_CATEGORY                     '�j�d�x�O
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    CATEGORY_CODE(0 To 7)       As Byte     '�i����ú�غ���
End Type




'�L�[�E�f�[�^
Public K0_ITEM_CATEGORY         As KEY0_ITEM_CATEGORY

Type ITEM_CATEGORY_FSpeck
    fs      As BtFileSpeck                  ' ̧�� ��߯��\����
    ks0     As BtKeySpeck                   ' �� ��߯��\����
    ks1     As BtKeySpeck
End Type

Private ITEM_CATEGORY_Speck  As ITEM_CATEGORY_FSpeck

Private Function ITEM_CATEGORY_Create() As Integer
'********************************************************************
'*
'*              �i���J�e�S���}�X�^  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************

Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    ITEM_CATEGORY_Create = True
                                            '�i���J�e�S���}�X�^ �t���p�X�捞��
    sts = GetIni("FILE", ITEM_CATEGORY_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ITEM_CATEGORY]�ǂݍ��݃G���[ ")
        Exit Function
    End If

    FullPath = RTrim(c)

    ITEM_CATEGORY_Speck.fs.recoleng = Len(ITEM_CATEGORYREC)     ' ���R�[�h��
    ITEM_CATEGORY_Speck.fs.PageSize = ITEM_CATEGORY_PG_SIZ      ' �y�[�W�T�C�Y
    ITEM_CATEGORY_Speck.fs.idexnumb = 1                         ' �C���f�b�N�X��
    ITEM_CATEGORY_Speck.fs.fileflag = 0                         ' �t�@�C���t���O
    ITEM_CATEGORY_Speck.fs.reserve = &H0                        ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�O
    ITEM_CATEGORY_Speck.ks0.keypos = 1                          ' �L�[�|�W�V����
    ITEM_CATEGORY_Speck.ks0.keyleng = 1                         ' �L�[��
    ITEM_CATEGORY_Speck.ks0.keyflag = BtKfExt + BtKfSeg         ' �L�[�t���O
    ITEM_CATEGORY_Speck.ks0.keytype = Chr(BtKtString)           ' �L�[�^�C�v
    ITEM_CATEGORY_Speck.ks0.reserve = &H0                       ' �\��ς�

    ITEM_CATEGORY_Speck.ks1.keypos = 2                          ' �L�[�|�W�V����
    ITEM_CATEGORY_Speck.ks1.keyleng = 8                         ' �L�[��
    ITEM_CATEGORY_Speck.ks1.keyflag = BtKfExt                   ' �L�[�t���O
    ITEM_CATEGORY_Speck.ks1.keytype = Chr(BtKtString)           ' �L�[�^�C�v
    ITEM_CATEGORY_Speck.ks1.reserve = &H0                       ' �\��ς�
'-----------------------------------------------
    sts = BTRV(BtOpCreate, ITEM_CATEGORY_POS, ITEM_CATEGORY_Speck, Len(ITEM_CATEGORY_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�i���J�e�S���}�X�^")
        Exit Function
    End If

    ITEM_CATEGORY_Create = False

End Function

Public Function ITEM_CATEGORY_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �i���J�e�S���}�X�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    ITEM_CATEGORY_Open = True
                                            '�i���J�e�S���}�X�^ �t���p�X�捞��
    sts = GetIni("FILE", ITEM_CATEGORY_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ITEM_CATEGORY]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, ITEM_CATEGORY_POS, ITEM_CATEGORYREC, Len(ITEM_CATEGORYREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = ITEM_CATEGORY_Create()    '�i���J�e�S���}�X�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ITEM_CATEGORY_POS, ITEM_CATEGORYREC, Len(ITEM_CATEGORYREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�i���J�e�S���}�X�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�i���J�e�S���}�X�^")
                Exit Function
        End Select
    Loop

    ITEM_CATEGORY_Open = False

End Function

