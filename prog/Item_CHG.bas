Attribute VB_Name = "ITEM_CHG"
Option Explicit
'********************************************************************
'*
'*              �i�ړǂݑւ�  �t�@�C����`
'*
'*          CREATE 2018.02.03
'********************************************************************
'�t�@�C���h�c
Public Const ITEM_CHG_ID$ = "ITEM_CHG"

'�y�[�W�T�C�Y
Public Const ITEM_CHG_PG_SIZ% = 4096

'�|�W�V�����E�u���b�N
Public ITEM_CHG_POS             As POSBLK
'=
'=
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************

'���R�[�h��`
Type ITEM_CHG_REC_Tag
    N_JGYOBU(0 To 0)            As Byte     '�V�@���ƕ��敪
    N_NAIGAI(0 To 0)            As Byte     '�V�@�����O
    N_HIN_GAI(0 To 19)          As Byte     '�V�@�i�ԁi�O���j
    HIN_NAME(0 To 39)           As Byte     '�i��
    O_HIN_GAI(0 To 39)          As Byte     '���@�i�ԁi�O���j�i���l�j

End Type
'�f�[�^�E�o�b�t�@
Public ITEM_CHG_REC As ITEM_CHG_REC_Tag

'�L�[��`

Type KEY0_ITEM_CHG            '�j�d�x�O
    N_JGYOBU(0 To 0)            As Byte     '�V�@���ƕ��敪
    N_NAIGAI(0 To 0)            As Byte     '�V�@�����O
    N_HIN_GAI(0 To 19)          As Byte     '�V�@�i�ԁi�O���j
End Type




'�L�[�E�f�[�^
Public K0_ITEM_CHG  As KEY0_ITEM_CHG

Type ITEM_CHG_FSpeck
    fs      As BtFileSpeck                 ' ̧�� ��߯��\����
    ks0     As BtKeySpeck                 ' �� ��߯��\����
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
End Type

Private ITEM_CHG_Speck  As ITEM_CHG_FSpeck

Private Function ITEM_CHG_Create() As Integer
'********************************************************************
'*
'*              �i�ړǂݑւ�  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    ITEM_CHG_Create = True
                                            '�i�ڃ}�X�^�t���p�X�捞��
    sts = GetIni("FILE", ITEM_CHG_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ITEM_CHG]�ǂݍ��݃G���[ ")
        Exit Function
    End If

    FullPath = RTrim(c)

    ITEM_CHG_Speck.fs.recoleng = Len(ITEM_CHG_REC)  ' ���R�[�h��
    ITEM_CHG_Speck.fs.PageSize = ITEM_PG_SIZ       ' �y�[�W�T�C�Y
    ITEM_CHG_Speck.fs.idexnumb = 1                 ' �C���f�b�N�X��
    ITEM_CHG_Speck.fs.fileflag = 0                 ' �t�@�C���t���O
    ITEM_CHG_Speck.fs.reserve = &H0                ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�O
    ITEM_CHG_Speck.ks0.keypos = 1                              ' �L�[�|�W�V����
    ITEM_CHG_Speck.ks0.keyleng = 1                             ' �L�[��
    ITEM_CHG_Speck.ks0.keyflag = BtKfExt + BtKfChg + BtKfSeg   ' �L�[�t���O
    ITEM_CHG_Speck.ks0.keytype = Chr(BtKtString)               ' �L�[�^�C�v
    ITEM_CHG_Speck.ks0.reserve = &H0                           ' �\��ς�

    ITEM_CHG_Speck.ks1.keypos = 2                              ' �L�[�|�W�V����
    ITEM_CHG_Speck.ks1.keyleng = 1                             ' �L�[��
    ITEM_CHG_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfSeg   ' �L�[�t���O
    ITEM_CHG_Speck.ks1.keytype = Chr(BtKtString)               ' �L�[�^�C�v
    ITEM_CHG_Speck.ks1.reserve = &H0                           ' �\��ς�

    ITEM_CHG_Speck.ks2.keypos = 3                              ' �L�[�|�W�V����
    ITEM_CHG_Speck.ks2.keyleng = 20                            ' �L�[��
    ITEM_CHG_Speck.ks2.keyflag = BtKfExt + BtKfChg             ' �L�[�t���O
    ITEM_CHG_Speck.ks2.keytype = Chr(BtKtString)               ' �L�[�^�C�v
    ITEM_CHG_Speck.ks2.reserve = &H0                           ' �\��ς�
'-----------------------------------------------





    sts = BTRV(BtOpCreate, ITEM_CHG_POS, ITEM_CHG_Speck, Len(ITEM_CHG_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�i�ړǂݑւ�")
        Exit Function
    End If

    ITEM_CHG_Create = False

End Function

Public Function ITEM_CHG_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �i�ړǂݑւ�  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    ITEM_CHG_Open = True
                                            '�i�ڃ}�X�^�t���p�X�捞��
    sts = GetIni("FILE", ITEM_CHG_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ITEM_CHG]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, ITEM_CHG_POS, ITEM_CHG_REC, Len(ITEM_CHG_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = ITEM_CHG_Create()        '�i�ڃ}�X�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ITEM_CHG_POS, ITEM_CHG_REC, Len(ITEM_CHG_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�i�ړǂݑւ�")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�i�ړǂݑւ�")
                Exit Function
        End Select
    Loop

    ITEM_CHG_Open = False

End Function

