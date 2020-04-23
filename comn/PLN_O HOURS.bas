Attribute VB_Name = "PLN_O_HOURS"
Option Explicit
'********************************************************************
'*
'*              �S���ҕʋΖ����ԃf�[�^  �t�@�C����`
'*
'*          CREATE 2011.09.13
'********************************************************************
'�t�@�C���h�c
Public Const PLN_O_HOURS_ID$ = "PLN_O_HOURS"

'�y�[�W�T�C�Y
Public Const PLN_O_HOURS_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public PLN_O_HOURS_POS            As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type PLN_O_HOURS_REC_Tag
    TANTO_CODE(0 To 4)              As Byte         '�f�[�^�쐬��
    O_DATE(0 To 7)                  As Byte         '�N����
    O_Time(0 To 3)                  As Byte         '�Ζ����� 99.9
    FILLER(0 To 62)                 As Byte
    INS_TANTO(0 To 9)               As Byte         '�ǉ��@�S����
    Ins_DateTime(0 To 13)           As Byte         '�ǉ��@����
    UPD_TANTO(0 To 9)               As Byte         '�X�V�@�S����
    UPD_DATETIME(0 To 13)           As Byte         '�X�V�@����



End Type
'�f�[�^�E�o�b�t�@
Public PLN_O_HOURS_REC              As PLN_O_HOURS_REC_Tag

'�L�[��`

Type KEY0_PLN_O_HOURS               '�j�d�x�O
    
    TANTO_CODE(0 To 4)              As Byte         '�S���Һ���
    O_DATE(0 To 7)                  As Byte         '�N����

End Type

Type KEY1_PLN_O_HOURS               '�j�d�x�P
    
    O_DATE(0 To 7)                  As Byte         '�N����
    TANTO_CODE(0 To 4)              As Byte         '�S���Һ���

End Type






'�L�[�E�f�[�^
Public K0_PLN_O_HOURS               As KEY0_PLN_O_HOURS
Public K1_PLN_O_HOURS               As KEY1_PLN_O_HOURS



Private Type PLN_O_HOURS_FSpeck
    fs      As BtFileSpeck              ' ̧�� ��߯��\����
    ks0     As BtKeySpeck               ' �� ��߯��\����
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck

End Type

Private PLN_O_HOURS_Speck           As PLN_O_HOURS_FSpeck

Private Function PLN_O_HOURS_Create() As Integer
'********************************************************************
'*
'*              �S���ҕʋΖ����ԃf�[�^  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    PLN_O_HOURS_Create = True
                                            '�S���ҕʋΖ����ԃf�[�^ �t���p�X�捞��
    sts = GetIni("FILE", PLN_O_HOURS_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [PLN_O_HOURS]�ǂݍ��݃G���[ ")
        Exit Function
    End If

    FullPath = RTrim(c)

    PLN_O_HOURS_Speck.fs.recoleng = Len(PLN_O_HOURS_REC)    ' ���R�[�h��
    PLN_O_HOURS_Speck.fs.PageSize = PLN_O_HOURS_PG_SIZ      ' �y�[�W�T�C�Y
    PLN_O_HOURS_Speck.fs.idexnumb = 2                       ' �C���f�b�N�X��
    PLN_O_HOURS_Speck.fs.fileflag = 0                       ' �t�@�C���t���O
    PLN_O_HOURS_Speck.fs.reserve = &H0                      ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�O
    PLN_O_HOURS_Speck.ks0.keypos = 1                        ' �L�[�|�W�V����
    PLN_O_HOURS_Speck.ks0.keyleng = 5                       ' �L�[��
    PLN_O_HOURS_Speck.ks0.keyflag = BtKfExt + BtKfSeg       ' �L�[�t���O
    PLN_O_HOURS_Speck.ks0.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    PLN_O_HOURS_Speck.ks0.reserve = &H0                     ' �\��ς�

    PLN_O_HOURS_Speck.ks1.keypos = 6                        ' �L�[�|�W�V����
    PLN_O_HOURS_Speck.ks1.keyleng = 8                       ' �L�[��
    PLN_O_HOURS_Speck.ks1.keyflag = BtKfExt                 ' �L�[�t���O
    PLN_O_HOURS_Speck.ks1.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    PLN_O_HOURS_Speck.ks1.reserve = &H0                     ' �\��ς�

'-----------------------------------------------
                                                ' �L�[�P
    PLN_O_HOURS_Speck.ks2.keypos = 6                        ' �L�[�|�W�V����
    PLN_O_HOURS_Speck.ks2.keyleng = 8                       ' �L�[��
    PLN_O_HOURS_Speck.ks2.keyflag = BtKfExt + BtKfSeg       ' �L�[�t���O
    PLN_O_HOURS_Speck.ks2.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    PLN_O_HOURS_Speck.ks2.reserve = &H0                     ' �\��ς�

    PLN_O_HOURS_Speck.ks3.keypos = 1                        ' �L�[�|�W�V����
    PLN_O_HOURS_Speck.ks3.keyleng = 5                       ' �L�[��
    PLN_O_HOURS_Speck.ks3.keyflag = BtKfExt                 ' �L�[�t���O
    PLN_O_HOURS_Speck.ks3.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    PLN_O_HOURS_Speck.ks3.reserve = &H0                     ' �\��ς�

'-----------------------------------------------

    sts = BTRV(BtOpCreate, PLN_O_HOURS_POS, PLN_O_HOURS_Speck, Len(PLN_O_HOURS_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�S���ҕʋΖ����ԃf�[�^")
        Exit Function
    End If

    PLN_O_HOURS_Create = False

End Function

Public Function PLN_O_HOURS_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �S���ҕʋΖ����ԃf�[�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    PLN_O_HOURS_Open = True
                                            '�S���ҕʋΖ����ԃf�[�^ �t���p�X�捞��
    sts = GetIni("FILE", PLN_O_HOURS_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [PLN_O_HOURS]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, PLN_O_HOURS_POS, PLN_O_HOURS_REC, Len(PLN_O_HOURS_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = PLN_O_HOURS_Create()  '�S���ҕʋΖ����ԃf�[�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, PLN_O_HOURS_POS, PLN_O_HOURS_REC, Len(PLN_O_HOURS_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�S���ҕʋΖ����ԃf�[�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�S���ҕʋΖ����ԃf�[�^")
                Exit Function
        End Select
    Loop

    PLN_O_HOURS_Open = False

End Function

