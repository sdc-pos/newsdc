Attribute VB_Name = "TANAPACKING"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �I�ʌ����}�X�^  �t�@�C����`                      *
'*                                                                  *
'*          CREATE 2004.02.16                                       *
'********************************************************************
'�t�@�C���h�c
Public Const TPACKING_ID$ = "TANAPACKING"

'�y�[�W�T�C�Y
Public Const TPACKING_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public TPACKING_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type TPACKINGREC_Tag
    Soko_No(0 To 1)     As Byte     '�q�ɇ�
    Retu(0 To 1)        As Byte     '�I�ԁ@��
    Ren(0 To 1)         As Byte     '�I�ԁ@�A
    PACKING_NO(0 To 3)  As Byte     '������
    RANK(0 To 2)        As Byte     '�����N
    FILLER(0 To 10)     As Byte     'FILLER
End Type
'�f�[�^�E�o�b�t�@
Public TPACKINGREC      As TPACKINGREC_Tag

'�L�[��`
Type KEY0_TPACKING                  '�j�d�x�O
    Soko_No(0 To 1)     As Byte     '�q�ɇ�
    Retu(0 To 1)        As Byte     '�I�ԁ@��
    Ren(0 To 1)         As Byte     '�I�ԁ@�A
    PACKING_NO(0 To 3)  As Byte     '������
    RANK(0 To 2)        As Byte     '�����N
End Type

Type KEY1_TPACKING                  '�j�d�x�P
    PACKING_NO(0 To 3)  As Byte     '������
    RANK(0 To 2)        As Byte     '�����N
    Soko_No(0 To 1)     As Byte     '�q�ɇ�
    Retu(0 To 1)        As Byte     '�I�ԁ@��
    Ren(0 To 1)         As Byte     '�I�ԁ@�A
End Type
    
'�L�[�E�f�[�^
Public K0_TPACKING      As KEY0_TPACKING
Public K1_TPACKING      As KEY1_TPACKING

Type TPACKING_FSpeck
    fs  As BtFileSpeck              ' ̧�� ��߯��\����
    ks0 As BtKeySpeck               ' �� ��߯��\����
    ks1 As BtKeySpeck               ' �� ��߯��\����
    ks2 As BtKeySpeck               ' �� ��߯��\����
    ks3 As BtKeySpeck               ' �� ��߯��\����
    ks4 As BtKeySpeck               ' �� ��߯��\����
    ks5 As BtKeySpeck               ' �� ��߯��\����
    ks6 As BtKeySpeck               ' �� ��߯��\����
    ks7 As BtKeySpeck               ' �� ��߯��\����
    ks8 As BtKeySpeck               ' �� ��߯��\����
    ks9 As BtKeySpeck               ' �� ��߯��\����
End Type

Public TPACKING_Speck   As TPACKING_FSpeck
Private Function TPACKING_Create() As Integer
'********************************************************************
'*                                                                  *
'*              �I�ʌ����}�X�^  �b�q�d�`�s�d                      *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 2004.02.13                                       *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    TPACKING_Create = True
                                            '�I�ʔ��}�X�^�t���p�X�捞��
    sts = GetIni("FILE", TPACKING_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        Exit Function
    End If

    FullPath = RTrim$(c)

    TPACKING_Speck.fs.recoleng = Len(TPACKINGREC)       ' ���R�[�h��
    TPACKING_Speck.fs.PageSize = TPACKING_PG_SIZ        ' �y�[�W�T�C�Y
    TPACKING_Speck.fs.idexnumb = 2                      ' �C���f�b�N�X��
    TPACKING_Speck.fs.fileflag = 0                      ' �t�@�C���t���O
    TPACKING_Speck.fs.reserve = &H0                     ' �\��ς�
'--------------------------------------------------------
                                                        ' �L�[�O
    TPACKING_Speck.ks0.keypos = 1                       ' �L�[�|�W�V����
    TPACKING_Speck.ks0.keyleng = 2                      ' �L�[��
                                                        ' �L�[�t���O
    TPACKING_Speck.ks0.keyflag = BtKfExt + BtKfSeg + BtKfChg
    TPACKING_Speck.ks0.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    TPACKING_Speck.ks0.reserve = &H0                    ' �\��ς�
                                                        ' �L�[�O
    TPACKING_Speck.ks1.keypos = 3                       ' �L�[�|�W�V����
    TPACKING_Speck.ks1.keyleng = 2                      ' �L�[��
                                                        ' �L�[�t���O
    TPACKING_Speck.ks1.keyflag = BtKfExt + BtKfSeg + BtKfChg
    TPACKING_Speck.ks1.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    TPACKING_Speck.ks1.reserve = &H0                    ' �\��ς�
                                                        ' �L�[�O
    TPACKING_Speck.ks2.keypos = 5                       ' �L�[�|�W�V����
    TPACKING_Speck.ks2.keyleng = 2                      ' �L�[��
                                                        ' �L�[�t���O
    TPACKING_Speck.ks2.keyflag = BtKfExt + BtKfSeg + BtKfChg
    TPACKING_Speck.ks2.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    TPACKING_Speck.ks2.reserve = &H0                    ' �\��ς�
                                                        ' �L�[�O
    TPACKING_Speck.ks3.keypos = 7                       ' �L�[�|�W�V����
    TPACKING_Speck.ks3.keyleng = 4                      ' �L�[��
                                                        ' �L�[�t���O
    TPACKING_Speck.ks3.keyflag = BtKfExt + BtKfSeg + BtKfChg
    TPACKING_Speck.ks3.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    TPACKING_Speck.ks3.reserve = &H0                    ' �\��ς�
                                                        ' �L�[�O
    TPACKING_Speck.ks4.keypos = 11                      ' �L�[�|�W�V����
    TPACKING_Speck.ks4.keyleng = 3                      ' �L�[��
                                                        ' �L�[�t���O
    TPACKING_Speck.ks4.keyflag = BtKfExt + BtKfChg
    TPACKING_Speck.ks4.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    TPACKING_Speck.ks4.reserve = &H0                    ' �\��ς�
'--------------------------------------------------------
                                                        ' �L�[�P
    TPACKING_Speck.ks5.keypos = 7                       ' �L�[�|�W�V����
    TPACKING_Speck.ks5.keyleng = 4                      ' �L�[��
                                                        ' �L�[�t���O
    TPACKING_Speck.ks5.keyflag = BtKfExt + BtKfSeg + BtKfChg
    TPACKING_Speck.ks5.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    TPACKING_Speck.ks5.reserve = &H0                    ' �\��ς�
                                                        ' �L�[�P
    TPACKING_Speck.ks6.keypos = 11                      ' �L�[�|�W�V����
    TPACKING_Speck.ks6.keyleng = 3                      ' �L�[��
                                                        ' �L�[�t���O
    TPACKING_Speck.ks6.keyflag = BtKfExt + BtKfSeg + BtKfChg
    TPACKING_Speck.ks6.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    TPACKING_Speck.ks6.reserve = &H0                    ' �\��ς�
                                                        ' �L�[�P
    TPACKING_Speck.ks7.keypos = 1                       ' �L�[�|�W�V����
    TPACKING_Speck.ks7.keyleng = 2                      ' �L�[��
                                                        ' �L�[�t���O
    TPACKING_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfChg
    TPACKING_Speck.ks7.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    TPACKING_Speck.ks7.reserve = &H0                    ' �\��ς�
                                                        ' �L�[�P
    TPACKING_Speck.ks8.keypos = 3                       ' �L�[�|�W�V����
    TPACKING_Speck.ks8.keyleng = 2                      ' �L�[��
                                                        ' �L�[�t���O
    TPACKING_Speck.ks8.keyflag = BtKfExt + BtKfSeg + BtKfChg
    TPACKING_Speck.ks8.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    TPACKING_Speck.ks8.reserve = &H0                    ' �\��ς�
                                                        ' �L�[�P
    TPACKING_Speck.ks9.keypos = 5                       ' �L�[�|�W�V����
    TPACKING_Speck.ks9.keyleng = 2                      ' �L�[��
                                                        ' �L�[�t���O
    TPACKING_Speck.ks9.keyflag = BtKfExt + BtKfChg
    TPACKING_Speck.ks9.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    TPACKING_Speck.ks9.reserve = &H0                    ' �\��ς�
'--------------------------------------------------------

    sts = BTRV(BtOpCreate, TPACKING_POS, TPACKING_Speck, Len(TPACKING_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�I�ʌ����}�X�^")
        Exit Function
    End If

    TPACKING_Create = False

End Function

Function TPACKING_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              �I�ʌ����}�X�^  �n�o�d�m                          *
'*                                                                  *
'*      ��  ��:Open Mode(Btrieve�Q��)                               *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 2004.02.13                                       *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    TPACKING_Open = True
                                            '�I�ʌ����}�X�^�t���p�X�捞��
    sts = GetIni("FILE", TPACKING_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim$(c)

    Do
        sts = BTRV(BtOpOpen, TPACKING_POS, TPACKINGREC, Len(TPACKINGREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = TPACKING_Create()         '�I�ʌ����}�X�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, TPACKING_POS, TPACKINGREC, Len(TPACKINGREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�I�ʌ����}�X�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�I�ʌ����}�X�^")
                Exit Function
        End Select
    Loop
    TPACKING_Open = False
End Function

Function TPACKING_ReCreate() As Integer
'********************************************************************
'*
'*              �I�ʌ����}�X�^  �t�@�C���č쐬
'*
'*      �߂�l:false ����
'*             true  �ُ�
'*
'*          CREATE 2004.06.16
'********************************************************************
Dim sts         As Integer

    TPACKING_ReCreate = True

    sts = TPACKING_Create()         '�I�ʌ����}�X�^�쐬
    If sts <> False Then
        Exit Function
    End If

    TPACKING_ReCreate = False

End Function

