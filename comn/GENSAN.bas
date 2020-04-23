Attribute VB_Name = "GENSAN"
Option Explicit
'********************************************************************
'*
'*              ���Y���}�X�^  �t�@�C����`
'*
'*          CREATE 2010.07.08
'********************************************************************
'�t�@�C���h�c
Public Const GENSAN_ID$ = "GENSAN"

'�y�[�W�T�C�Y
Public Const GENSAN_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public GENSAN_POS       As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type GENSANREC_Tag
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    NAIGAI(0 To 0)              As Byte     '�����O
    HIN_GAI(0 To 19)            As Byte     '�i�ԁi�O���j
    GENSANKOKU(0 To 19)         As Byte     '���Y��
    FILLER(0 To 175)            As Byte     'FILLER
    INS_TANTO(0 To 4)           As Byte     '�ǉ��@�S����
    Ins_DateTime(0 To 13)       As Byte     '�ǉ��@����
    UPD_TANTO(0 To 4)           As Byte     '�X�V�@�S����
    UPD_DATETIME(0 To 13)       As Byte     '�X�V�@����

End Type
'�f�[�^�E�o�b�t�@
Public GENSANREC                As GENSANREC_Tag

'�L�[��`

Type KEY0_GENSAN                '�j�d�x�O
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    NAIGAI(0 To 0)              As Byte     '�����O
    HIN_GAI(0 To 19)            As Byte     '�i�ԁi�O���j
    GENSANKOKU(0 To 19)         As Byte     '���Y��
End Type




'�L�[�E�f�[�^
Public K0_GENSAN                As KEY0_GENSAN

Type GENSAN_FSpeck
    fs      As BtFileSpeck                 ' ̧�� ��߯��\����
    ks0     As BtKeySpeck                 ' �� ��߯��\����
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck
End Type

Private GENSAN_Speck  As GENSAN_FSpeck

Private Function GENSAN_Create() As Integer
'********************************************************************
'*
'*              ���Y�}�X�^  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    GENSAN_Create = True
                                            '���Y�}�X�^�t���p�X�捞��
    sts = GetIni("FILE", GENSAN_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [GENSAN]�ǂݍ��݃G���[ ")
        Exit Function
    End If

    FullPath = RTrim(c)

    GENSAN_Speck.fs.recoleng = Len(GENSANREC)       ' ���R�[�h��
    GENSAN_Speck.fs.PageSize = GENSAN_PG_SIZ        ' �y�[�W�T�C�Y
    GENSAN_Speck.fs.idexnumb = 1                    ' �C���f�b�N�X��
    GENSAN_Speck.fs.fileflag = 0                    ' �t�@�C���t���O
    GENSAN_Speck.fs.reserve = &H0                   ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�O
    GENSAN_Speck.ks0.keypos = 1                     ' �L�[�|�W�V����
    GENSAN_Speck.ks0.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    GENSAN_Speck.ks0.keyflag = BtKfExt + BtKfChg + BtKfSeg
    GENSAN_Speck.ks0.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    GENSAN_Speck.ks0.reserve = &H0                  ' �\��ς�

    GENSAN_Speck.ks1.keypos = 2                     ' �L�[�|�W�V����
    GENSAN_Speck.ks1.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    GENSAN_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfSeg
    GENSAN_Speck.ks1.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    GENSAN_Speck.ks1.reserve = &H0                  ' �\��ς�

    GENSAN_Speck.ks2.keypos = 3                     ' �L�[�|�W�V����
    GENSAN_Speck.ks2.keyleng = 20                   ' �L�[��
                                                    ' �L�[�t���O
    GENSAN_Speck.ks2.keyflag = BtKfExt + BtKfChg + BtKfSeg
    GENSAN_Speck.ks2.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    GENSAN_Speck.ks2.reserve = &H0                  ' �\��ς�

    GENSAN_Speck.ks3.keypos = 23                    ' �L�[�|�W�V����
    GENSAN_Speck.ks3.keyleng = 20                   ' �L�[��
    GENSAN_Speck.ks3.keyflag = BtKfExt + BtKfChg    ' �L�[�t���O
    GENSAN_Speck.ks3.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    GENSAN_Speck.ks3.reserve = &H0                  ' �\��ς�
'-----------------------------------------------

    sts = BTRV(BtOpCreate, GENSAN_POS, GENSAN_Speck, Len(GENSAN_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���Y�}�X�^")
        Exit Function
    End If

    GENSAN_Create = False

End Function

Public Function GENSAN_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���Y�}�X�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    GENSAN_Open = True
                                            '���Y�}�X�^�t���p�X�捞��
    sts = GetIni("FILE", GENSAN_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [GENSAN]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, GENSAN_POS, GENSANREC, Len(GENSANREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = GENSAN_Create()        '���Y�}�X�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, GENSAN_POS, GENSANREC, Len(GENSANREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���Y�}�X�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���Y�}�X�^")
                Exit Function
        End Select
    Loop

    GENSAN_Open = False

End Function

