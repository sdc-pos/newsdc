Attribute VB_Name = "OSAKA_PSTOCK"
Option Explicit
'********************************************************************
'*
'*              ���o�b�@�z�I���e �t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const OSAKA_PSTOCK_ID$ = "OSAKA_PSTOCK"

'�y�[�W�T�C�Y
Public Const OSAKA_PSTOCK_PG_SIZ% = 2048

'�|�W�V�����E�u���b�N
Public OSAKA_PSTOCK_POS As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
Type OSAKA_PSTOCKREC_Tag
    Soko_No(0 To 1)             As Byte     '�q�ɇ�
    Retu(0 To 1)                As Byte     '�I�ԁ@��
    Ren(0 To 1)                 As Byte     '�I�ԁ@�A
    Dan(0 To 1)                 As Byte     '�I�ԁ@�i
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    NAIGAI(0 To 0)              As Byte     '�����O
    HIN_GAI(0 To 19)            As Byte     '�i�ԁi�O���j
    
    KEIJYO_YM(0 To 5)           As Byte     '�v��N��
        
    NYUKO_QTY(0 To 9)           As Byte     '�������ɐ�
    SYUKO_QTY(0 To 9)           As Byte     '�����o�ɐ�
    ZAIKO_QTY(0 To 9)           As Byte     '�����݌Ɏc��
    FILLER(0 To 47)             As Byte     'FILLER

    Ins_DateTime(0 To 13)       As Byte     '�ް��쐬����

End Type

'�f�[�^�E�o�b�t�@
Public OSAKA_PSTOCKREC          As OSAKA_PSTOCKREC_Tag

'�L�[��`
Type KEY0_OSAKA_PSTOCK                      '�j�d�x�O
    Soko_No(0 To 1)             As Byte     '�q�ɇ�
    Retu(0 To 1)                As Byte     '�I�ԁ@��
    Ren(0 To 1)                 As Byte     '�I�ԁ@�A
    Dan(0 To 1)                 As Byte     '�I�ԁ@�i
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    NAIGAI(0 To 0)              As Byte     '�����O
    HIN_GAI(0 To 19)            As Byte     '�i�ԁi�O���j
End Type



'�L�[�E�f�[�^
Public K0_OSAKA_PSTOCK          As KEY0_OSAKA_PSTOCK

Type OSAKA_PSTOCK_FSpeck
    fs      As BtFileSpeck          ' ̧�� ��߯��\����
    ks0     As BtKeySpeck           ' �� ��߯��\����
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck
    ks4     As BtKeySpeck
    ks5     As BtKeySpeck
    ks6     As BtKeySpeck
End Type

Private OSAKA_PSTOCK_Speck  As OSAKA_PSTOCK_FSpeck
Private Function OSAKA_PSTOCK_Create() As Integer
'********************************************************************
'*
'*              ���o�b�@�z�I���e�@�b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    OSAKA_PSTOCK_Create = True
                                            '�݌Ƀf�[�^�@�t���p�X�捞��
    sts = GetIni("FILE", OSAKA_PSTOCK_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [OSAKA_PSTOCK]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    OSAKA_PSTOCK_Speck.fs.recoleng = Len(OSAKA_PSTOCKREC)           ' ���R�[�h��
    OSAKA_PSTOCK_Speck.fs.PageSize = OSAKA_PSTOCK_PG_SIZ            ' �y�[�W�T�C�Y
    OSAKA_PSTOCK_Speck.fs.idexnumb = 1                              ' �C���f�b�N�X��
    OSAKA_PSTOCK_Speck.fs.fileflag = 0                              ' �t�@�C���t���O
    OSAKA_PSTOCK_Speck.fs.reserve = &H0                             ' �\��ς�
'---------------------------------------------------'
                                                    ' �L�[�O
    OSAKA_PSTOCK_Speck.ks0.keypos = 1                               ' �L�[�|�W�V����
    OSAKA_PSTOCK_Speck.ks0.keyleng = 2                              ' �L�[��
    OSAKA_PSTOCK_Speck.ks0.keyflag = BtKfExt + BtKfChg + BtKfSeg    ' �L�[�t���O
    OSAKA_PSTOCK_Speck.ks0.keytype = Chr(BtKtString)                ' �L�[�^�C�v
    OSAKA_PSTOCK_Speck.ks0.reserve = &H0                            ' �\��ς�
                                                    
    OSAKA_PSTOCK_Speck.ks1.keypos = 3                               ' �L�[�|�W�V����
    OSAKA_PSTOCK_Speck.ks1.keyleng = 2                              ' �L�[��
    OSAKA_PSTOCK_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfSeg    ' �L�[�t���O
    OSAKA_PSTOCK_Speck.ks1.keytype = Chr(BtKtString)                ' �L�[�^�C�v
    OSAKA_PSTOCK_Speck.ks1.reserve = &H0                            ' �\��ς�
                                                    
    OSAKA_PSTOCK_Speck.ks2.keypos = 5                               ' �L�[�|�W�V����
    OSAKA_PSTOCK_Speck.ks2.keyleng = 2                              ' �L�[��
    OSAKA_PSTOCK_Speck.ks2.keyflag = BtKfExt + BtKfChg + BtKfSeg    ' �L�[�t���O
    OSAKA_PSTOCK_Speck.ks2.keytype = Chr(BtKtString)                ' �L�[�^�C�v
    OSAKA_PSTOCK_Speck.ks2.reserve = &H0                            ' �\��ς�
                                                    
    OSAKA_PSTOCK_Speck.ks3.keypos = 7                               ' �L�[�|�W�V����
    OSAKA_PSTOCK_Speck.ks3.keyleng = 2                              ' �L�[��
    OSAKA_PSTOCK_Speck.ks3.keyflag = BtKfExt + BtKfChg + BtKfSeg    ' �L�[�t���O
    OSAKA_PSTOCK_Speck.ks3.keytype = Chr(BtKtString)                ' �L�[�^�C�v
    OSAKA_PSTOCK_Speck.ks3.reserve = &H0                            ' �\��ς�
                                                    
    OSAKA_PSTOCK_Speck.ks4.keypos = 9                               ' �L�[�|�W�V����
    OSAKA_PSTOCK_Speck.ks4.keyleng = 1                              ' �L�[��
    OSAKA_PSTOCK_Speck.ks4.keyflag = BtKfExt + BtKfChg + BtKfSeg    ' �L�[�t���O
    OSAKA_PSTOCK_Speck.ks4.keytype = Chr(BtKtString)                ' �L�[�^�C�v
    OSAKA_PSTOCK_Speck.ks4.reserve = &H0                            ' �\��ς�
                                                    
    OSAKA_PSTOCK_Speck.ks5.keypos = 10                              ' �L�[�|�W�V����
    OSAKA_PSTOCK_Speck.ks5.keyleng = 1                              ' �L�[��
    OSAKA_PSTOCK_Speck.ks5.keyflag = BtKfExt + BtKfChg + BtKfSeg    ' �L�[�t���O
    OSAKA_PSTOCK_Speck.ks5.keytype = Chr(BtKtString)                ' �L�[�^�C�v
    OSAKA_PSTOCK_Speck.ks5.reserve = &H0                            ' �\��ς�
                                                    
    OSAKA_PSTOCK_Speck.ks6.keypos = 11                              ' �L�[�|�W�V����
    OSAKA_PSTOCK_Speck.ks6.keyleng = 20                             ' �L�[��
    OSAKA_PSTOCK_Speck.ks6.keyflag = BtKfExt + BtKfChg              ' �L�[�t���O
    OSAKA_PSTOCK_Speck.ks6.keytype = Chr(BtKtString)                ' �L�[�^�C�v
    OSAKA_PSTOCK_Speck.ks6.reserve = &H0                            ' �\��ς�
                                                    
'---------------------------------------------------'
    sts = BTRV(BtOpCreate, OSAKA_PSTOCK_POS, OSAKA_PSTOCK_Speck, Len(OSAKA_PSTOCK_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���o�b�@�z�I���e")
        Exit Function
    End If
    OSAKA_PSTOCK_Create = False
End Function
Public Function OSAKA_PSTOCK_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���o�b�@�z�I���e�@�n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    OSAKA_PSTOCK_Open = True
                                            '���o�b�@�z�I���e�@�t���p�X�捞��
    sts = GetIni("FILE", OSAKA_PSTOCK_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [OSAKA_PSTOCK]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, OSAKA_PSTOCK_POS, OSAKA_PSTOCKREC, Len(OSAKA_PSTOCKREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = OSAKA_PSTOCK_Create() '���o�b�@�z�I���e�@�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, OSAKA_PSTOCK_POS, OSAKA_PSTOCKREC, Len(OSAKA_PSTOCKREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���o�b�@�z�I���e")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���o�b�@�z�I���e")
                Exit Function
        End Select
    Loop
    OSAKA_PSTOCK_Open = False

End Function

