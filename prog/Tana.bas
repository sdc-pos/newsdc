Attribute VB_Name = "TANA"
Option Explicit
'********************************************************************
'*
'*              �I�}�X�^  �t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const TANA_ID$ = "TANA"
'�y�[�W�T�C�Y
Public Const TANA_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public TANA_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type TANAREC_Tag
    SOKO_NO(0 To 1)         As Byte     '�q�ɇ�
    Retu(0 To 1)            As Byte     '�I�ԁ@��
    Ren(0 To 1)             As Byte     '�I�ԁ@�A
    Dan(0 To 1)             As Byte     '�I�ԁ@�i
    KAHI_KBN(0 To 0)        As Byte     '�g�p��
    TANA_COND(0 To 0)       As Byte     '�I���
    
    ZAIKO_SHOGO_FLG(0 To 0) As Byte     '�݌ɏƍ��t���O 2004.02
    
    Tana_Use(0 To 2)        As Byte     '�I�̎g�p��   2010.12.13
    
    
    FILLER(0 To 6)          As Byte     'FILLER         2010.12.13
End Type
'�f�[�^�E�o�b�t�@
Public TANAREC As TANAREC_Tag


'�L�[��`
Type KEY0_TANA                 '�j�d�x�O
    SOKO_NO(0 To 1)         As Byte     '�q�ɇ�
    Retu(0 To 1)            As Byte     '�I�ԁ@��
    Ren(0 To 1)             As Byte     '�I�ԁ@�A
    Dan(0 To 1)             As Byte     '�I�ԁ@�i
End Type

Type KEY1_TANA                 '�j�d�x�P
    KAHI_KBN(0 To 0)        As Byte     '�g�p��
    SOKO_NO(0 To 1)         As Byte     '�q�ɇ�
    Retu(0 To 1)            As Byte     '�I�ԁ@��
    Ren(0 To 1)             As Byte     '�I�ԁ@�A
    Dan(0 To 1)             As Byte     '�I�ԁ@�i
End Type

    
'�L�[�E�f�[�^
Public K0_TANA              As KEY0_TANA
Public K1_TANA              As KEY1_TANA

Type TANA_FSpeck
    fs              As BtFileSpeck      ' ̧�� ��߯��\����
    ks0             As BtKeySpeck       ' �� ��߯��\����
    ks1             As BtKeySpeck       ' �� ��߯��\����
    ks2             As BtKeySpeck       ' �� ��߯��\����
    ks3             As BtKeySpeck       ' �� ��߯��\����
    ks4             As BtKeySpeck       ' �� ��߯��\����
    ks5             As BtKeySpeck       ' �� ��߯��\����
    ks6             As BtKeySpeck       ' �� ��߯��\����
    ks7             As BtKeySpeck       ' �� ��߯��\����
    ks8             As BtKeySpeck       ' �� ��߯��\����
End Type

Public TANA_Speck   As TANA_FSpeck
Private Function TANA_Create() As Integer
'********************************************************************
'*                                                                  *
'*              �I�}�X�^  �b�q�d�`�s�d                              *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    TANA_Create = False
                                            '�I�}�X�^�t���p�X�捞��
    sts = GetIni("FILE", TANA_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        TANA_Create = True
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    TANA_Speck.fs.recoleng = Len(TANAREC)       ' ���R�[�h��
    TANA_Speck.fs.PageSize = TANA_PG_SIZ        ' �y�[�W�T�C�Y
    TANA_Speck.fs.idexnumb = 2                  ' �C���f�b�N�X��
    TANA_Speck.fs.fileflag = 0                  ' �t�@�C���t���O
    TANA_Speck.fs.reserve = &H0                 ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�O
    TANA_Speck.ks0.keypos = 1                   ' �L�[�|�W�V����
    TANA_Speck.ks0.keyleng = 2                  ' �L�[��
    TANA_Speck.ks0.keyflag = BtKfExt + BtKfSeg  ' �L�[�t���O
    TANA_Speck.ks0.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    TANA_Speck.ks0.reserve = &H0                ' �\��ς�

    TANA_Speck.ks1.keypos = 3                   ' �L�[�|�W�V����
    TANA_Speck.ks1.keyleng = 2                  ' �L�[��
    TANA_Speck.ks1.keyflag = BtKfExt + BtKfSeg  ' �L�[�t���O
    TANA_Speck.ks1.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    TANA_Speck.ks1.reserve = &H0                ' �\��ς�

    TANA_Speck.ks2.keypos = 5                   ' �L�[�|�W�V����
    TANA_Speck.ks2.keyleng = 2                  ' �L�[��
    TANA_Speck.ks2.keyflag = BtKfExt + BtKfSeg  ' �L�[�t���O
    TANA_Speck.ks2.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    TANA_Speck.ks2.reserve = &H0                ' �\��ς�

    TANA_Speck.ks3.keypos = 7                   ' �L�[�|�W�V����
    TANA_Speck.ks3.keyleng = 2                  ' �L�[��
    TANA_Speck.ks3.keyflag = BtKfExt            ' �L�[�t���O
    TANA_Speck.ks3.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    TANA_Speck.ks3.reserve = &H0                ' �\��ς�

'-----------------------------------------------
                                                ' �L�[�P
    TANA_Speck.ks4.keypos = 9                   ' �L�[�|�W�V����
    TANA_Speck.ks4.keyleng = 1                  ' �L�[��
                                                ' �L�[�t���O
    TANA_Speck.ks4.keyflag = BtKfExt + BtKfSeg + BtKfChg
    TANA_Speck.ks4.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    TANA_Speck.ks4.reserve = &H0                ' �\��ς�
                                                ' �L�[�P
    TANA_Speck.ks5.keypos = 1                   ' �L�[�|�W�V����
    TANA_Speck.ks5.keyleng = 2                  ' �L�[��
                                                ' �L�[�t���O
    TANA_Speck.ks5.keyflag = BtKfExt + BtKfSeg + BtKfChg
    TANA_Speck.ks5.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    TANA_Speck.ks5.reserve = &H0                ' �\��ς�
                                                ' �L�[�P
    TANA_Speck.ks6.keypos = 3                   ' �L�[�|�W�V����
    TANA_Speck.ks6.keyleng = 2                  ' �L�[��
                                                ' �L�[�t���O
    TANA_Speck.ks6.keyflag = BtKfExt + BtKfSeg + BtKfChg
    TANA_Speck.ks6.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    TANA_Speck.ks6.reserve = &H0                ' �\��ς�
                                                ' �L�[�P
    TANA_Speck.ks7.keypos = 5                   ' �L�[�|�W�V����
    TANA_Speck.ks7.keyleng = 2                  ' �L�[��
                                                ' �L�[�t���O
    TANA_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfChg
    TANA_Speck.ks7.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    TANA_Speck.ks7.reserve = &H0                ' �\��ς�
                                                ' �L�[�P
    TANA_Speck.ks8.keypos = 7                   ' �L�[�|�W�V����
    TANA_Speck.ks8.keyleng = 2                  ' �L�[��
                                                ' �L�[�t���O
    TANA_Speck.ks8.keyflag = BtKfExt + BtKfChg
    TANA_Speck.ks8.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    TANA_Speck.ks8.reserve = &H0                ' �\��ς�

    sts = BTRV(BtOpCreate, TANA_POS, TANA_Speck, Len(TANA_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�I�}�X�^")
        TANA_Create = True
    End If
End Function

Function TANA_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              �I�}�X�^  �n�o�d�m                                  *
'*                                                                  *
'*      ��  ��:Open Mode(Btrieve�Q��)                               *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String
    
    TANA_Open = False
                                            '�I�}�X�^�t���p�X�捞��
    sts = GetIni("FILE", TANA_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        TANA_Open = True
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, TANA_POS, TANAREC, Len(TANAREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Function
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = TANA_Create()        '�I�}�X�^�쐬
                If sts <> False Then
                    TANA_Open = True
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, TANA_POS, TANAREC, Len(TANAREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�I�}�X�^")
                    TANA_Open = True
                End If
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "�I�}�X�^")
                TANA_Open = True
                Exit Function
        End Select
    Loop
End Function



