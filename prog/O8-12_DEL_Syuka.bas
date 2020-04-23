Attribute VB_Name = "O_DEL_SYU"
Option Explicit
'********************************************************************
'*
'*              �폜�ςݏo�ח\��f�[�^  �t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const O_DEL_SYU_ID$ = "O_DEL_SYU"

'�y�[�W�T�C�Y
Public Const O_DEL_SYU_PG_SIZ% = 2048

'�|�W�V�����E�u���b�N
Public O_DEL_SYU_POS    As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type O_DEL_SYUREC_Tag
    WEL_ID(0 To 2)              As Byte     '�g�p�q�@ID
    PRG_ID(0 To 7)              As Byte     '�g�p���v���O����
    KAN_KBN(0 To 0)             As Byte     '�����敪
    DT_SYU(0 To 0)              As Byte     '�f�[�^���
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    KEY_CYU_KBN(0 To 0)         As Byte     '�����敪
    KEY_ID_NO(0 To 7)           As Byte     'ID-NO
    NAIGAI(0 To 0)              As Byte     '�����O
    KEY_HIN_NO(0 To 19)         As Byte     '�i�ڔԍ�
    KEY_MUKE_CODE(0 To 7)       As Byte     '���Ӑ�R�[�h
    KEY_SS_CODE(0 To 7)         As Byte     '������R�[�h
    KEY_SYUKA_YMD(0 To 7)       As Byte     '�o�ד��t
    JGYOBA(0 To 7)              As Byte     '���Ə�
    DATA_KBN(0 To 0)            As Byte     '�f�[�^�敪
    TORI_KBN(0 To 1)            As Byte     '����敪
    ID_NO(0 To 7)               As Byte     'ID-NO
    HIN_NO(0 To 19)             As Byte     '�i�ڔԍ�
    DEN_NO(0 To 9)              As Byte     '�`�[�ԍ�
    SURYO(0 To 6)               As Byte     '�o�ɐ���
    MUKE_CODE(0 To 7)           As Byte     '���Ӑ�R�[�h
    SYUKO_SYUSI(0 To 1)         As Byte     '�o�Ɏ��x
    SYUKA_YMD(0 To 7)           As Byte     '�o�ד��t
    ODER_NO(0 To 11)            As Byte     '�I�[�_�[�ԍ�
    ITEM_NO(0 To 4)             As Byte     '�A�C�e���ԍ�
    MUKE_NAME(0 To 23)          As Byte     '���Ӑ於��
    CYU_KBN(0 To 0)             As Byte     '�����敪
    CYU_KBN_NAME(0 To 9)        As Byte     '�����敪����
    EXPORT_KBN(0 To 0)          As Byte     '�A�o�o�׌����敪
    LABEL_ISSUE_KBN(0 To 0)     As Byte     '�����x�����s�敪
    LABEL_ISSUE_UNIT(0 To 4)    As Byte     '�����x�����s�P�ʐ�
    LABEL_TANKA_KBN(0 To 0)     As Byte     '�����x���P���\���敪
    TANKA(0 To 9)               As Byte     '�P��
    KINGAKU(0 To 9)             As Byte     '���z
    BIKOU2(0 To 19)             As Byte     '���l�Q
    REBATE_KBN(0 To 0)          As Byte     '���x�[�g�敪
    CHOHA_KBN(0 To 0)           As Byte     '���[�敪
    ATAISA_KBN(0 To 0)          As Byte     '�l���敪
    REP_KISHU(0 To 19)          As Byte     '��\�@��
    NS_KANRI_NO(0 To 8)         As Byte     '�m�r�Ǘ��ԍ�
    MTS_HIN_CODE(0 To 10)       As Byte     '�l�s�r���i�R�[�h
    BIKOU1(0 To 39)             As Byte     '���l�P
    CHOKU_KBN(0 To 0)           As Byte     '�����敪
    REBATE_RATE(0 To 4)         As Byte     '���x�[�g��
    HIN_NAME(0 To 19)           As Byte     '�i��
    JGYOBA_GAI(0 To 7)          As Byte     '�ΊO���Ə�
    KISHU_CODE(0 To 2)          As Byte     '�@��R�[�h
    SS_CODE(0 To 7)             As Byte     '������R�[�h
    HIN_NAI(0 To 19)            As Byte     '�i�ԁi�����j
    HTANABAN(0 To 7)            As Byte     '�z�X�g�I��
    PRINT_YMD(0 To 7)           As Byte     '�o�ɕ\������t
    KAN_YMD(0 To 7)             As Byte     '�������t
    KENPIN_YMD(0 To 7)          As Byte     '���i���t
    TOK_KBN(0 To 0)             As Byte     '������敪
    JITU_SURYO(0 To 6)          As Byte     '�o�Ɏ��ѐ���
    INS_NOW(0 To 13)            As Byte     '�捞�ݓ���
    FILLER(0 To 67)             As Byte     'FILLER
End Type

'�f�[�^�E�o�b�t�@
Public O_DEL_SYUREC As O_DEL_SYUREC_Tag

'�L�[��`
Type KEY0_O_DEL_SYU            '�j�d�x�O
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    KEY_CYU_KBN(0 To 0)         As Byte     '�����敪
    KEY_MUKE_CODE(0 To 7)       As Byte     '���Ӑ�R�[�h
    KEY_SS_CODE(0 To 7)         As Byte     '������R�[�h
    KEY_HIN_NO(0 To 19)         As Byte     '�i�ڔԍ�
    KEY_SYUKA_YMD(0 To 7)       As Byte     '�o�ד��t
End Type

Type KEY1_O_DEL_SYU           '�j�d�x�P
    KEY_SYUKA_YMD(0 To 7)       As Byte     '�o�ד��t
End Type

Type KEY2_O_DEL_SYU            '�j�d�x�Q
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    KEY_MUKE_CODE(0 To 7)       As Byte     '���Ӑ�R�[�h
    KEY_SS_CODE(0 To 7)         As Byte     '������R�[�h
    KEY_CYU_KBN(0 To 0)         As Byte     '�����敪
    NAIGAI(0 To 0)              As Byte     '�����O
    KEY_HIN_NO(0 To 19)         As Byte     '�i�ڔԍ�
    KEY_ID_NO(0 To 7)           As Byte     'ID-NO
End Type


'�L�[�E�f�[�^
Public K0_O_DEL_SYU               As KEY0_O_DEL_SYU
Public K1_O_DEL_SYU               As KEY1_O_DEL_SYU
Public K2_O_DEL_SYU               As KEY2_O_DEL_SYU

Type O_DEL_SYU_FSpeck
    fs      As BtFileSpeck                  ' ̧�� ��߯��\����
    ks0     As BtKeySpeck                   ' �� ��߯��\����
    ks1     As BtKeySpeck                   ' �� ��߯��\����
    ks2     As BtKeySpeck                   ' �� ��߯��\����
    ks3     As BtKeySpeck                   ' �� ��߯��\����
    ks4     As BtKeySpeck                   ' �� ��߯��\����
    ks5     As BtKeySpeck                   ' �� ��߯��\����
    ks6     As BtKeySpeck                   ' �� ��߯��\����
    ks7     As BtKeySpeck                   ' �� ��߯��\����
    ks8     As BtKeySpeck                   ' �� ��߯��\����
    ks9     As BtKeySpeck                   ' �� ��߯��\����
    ks10    As BtKeySpeck                   ' �� ��߯��\����
    ks11    As BtKeySpeck                   ' �� ��߯��\����
    ks12    As BtKeySpeck                   ' �� ��߯��\����
    ks13    As BtKeySpeck                   ' �� ��߯��\����
End Type

Private O_DEL_SYU_Speck As O_DEL_SYU_FSpeck

Private Function O_DEL_SYU_Create() As Integer
'********************************************************************
'*
'*              �폜�ςݏo�ח\��f�[�^  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    O_DEL_SYU_Create = True
                                            '�폜�ςݏo�ח\��f�[�^�t���p�X�捞��
    sts = GetIni("FILE", O_DEL_SYU_ID, "CONV200605", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [O_DEL_SYU]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    O_DEL_SYU_Speck.fs.recoleng = Len(O_DEL_SYUREC)               ' ���R�[�h��
    O_DEL_SYU_Speck.fs.PageSize = O_DEL_SYU_PG_SIZ              ' �y�[�W�T�C�Y
    O_DEL_SYU_Speck.fs.idexnumb = 3                           ' �C���f�b�N�X��
    O_DEL_SYU_Speck.fs.fileflag = 0                           ' �t�@�C���t���O
    O_DEL_SYU_Speck.fs.reserve = &H0                          ' �\��ς�
'---------------------------------------------------        �L�[�O
    
    O_DEL_SYU_Speck.ks0.keypos = 14                           ' �L�[�|�W�V����
    O_DEL_SYU_Speck.ks0.keyleng = 1                           ' �L�[��
    O_DEL_SYU_Speck.ks0.keyflag = BtKfExt + BtKfSeg + BtKfDup ' �L�[�t���O
    O_DEL_SYU_Speck.ks0.keytype = Chr(BtKtString)             ' �L�[�^�C�v
    O_DEL_SYU_Speck.ks0.reserve = &H0                         ' �\��ς�
    
    O_DEL_SYU_Speck.ks1.keypos = 15                           ' �L�[�|�W�V����
    O_DEL_SYU_Speck.ks1.keyleng = 1                           ' �L�[��
    O_DEL_SYU_Speck.ks1.keyflag = BtKfExt + BtKfSeg + BtKfDup ' �L�[�t���O
    O_DEL_SYU_Speck.ks1.keytype = Chr(BtKtString)             ' �L�[�^�C�v
    O_DEL_SYU_Speck.ks1.reserve = &H0                         ' �\��ς�
    
    O_DEL_SYU_Speck.ks2.keypos = 45                           ' �L�[�|�W�V����
    O_DEL_SYU_Speck.ks2.keyleng = 8                           ' �L�[��
    O_DEL_SYU_Speck.ks2.keyflag = BtKfExt + BtKfSeg + BtKfDup ' �L�[�t���O
    O_DEL_SYU_Speck.ks2.keytype = Chr(BtKtString)             ' �L�[�^�C�v
    O_DEL_SYU_Speck.ks2.reserve = &H0                         ' �\��ς�
    
    O_DEL_SYU_Speck.ks3.keypos = 53                           ' �L�[�|�W�V����
    O_DEL_SYU_Speck.ks3.keyleng = 8                           ' �L�[��
    O_DEL_SYU_Speck.ks3.keyflag = BtKfExt + BtKfSeg + BtKfDup ' �L�[�t���O
    O_DEL_SYU_Speck.ks3.keytype = Chr(BtKtString)             ' �L�[�^�C�v
    O_DEL_SYU_Speck.ks3.reserve = &H0                         ' �\��ς�
    
    O_DEL_SYU_Speck.ks4.keypos = 25                           ' �L�[�|�W�V����
    O_DEL_SYU_Speck.ks4.keyleng = 20                          ' �L�[��
    O_DEL_SYU_Speck.ks4.keyflag = BtKfExt + BtKfSeg + BtKfDup ' �L�[�t���O
    O_DEL_SYU_Speck.ks4.keytype = Chr(BtKtString)             ' �L�[�^�C�v
    O_DEL_SYU_Speck.ks4.reserve = &H0                         ' �\��ς�
    
    O_DEL_SYU_Speck.ks5.keypos = 61                           ' �L�[�|�W�V����
    O_DEL_SYU_Speck.ks5.keyleng = 8                           ' �L�[��
    O_DEL_SYU_Speck.ks5.keyflag = BtKfExt + BtKfDup           ' �L�[�t���O
    O_DEL_SYU_Speck.ks5.keytype = Chr(BtKtString)             ' �L�[�^�C�v
    O_DEL_SYU_Speck.ks5.reserve = &H0                         ' �\��ς�

'---------------------------------------------------        �L�[�P
    
    O_DEL_SYU_Speck.ks6.keypos = 61                           ' �L�[�|�W�V����
    O_DEL_SYU_Speck.ks6.keyleng = 8                           ' �L�[��
    O_DEL_SYU_Speck.ks6.keyflag = BtKfExt + BtKfDup           ' �L�[�t���O
    O_DEL_SYU_Speck.ks6.keytype = Chr(BtKtString)             ' �L�[�^�C�v
    O_DEL_SYU_Speck.ks6.reserve = &H0                         ' �\��ς�
    
'---------------------------------------------------        �L�[�Q
    
    O_DEL_SYU_Speck.ks7.keypos = 14                           ' �L�[�|�W�V����
    O_DEL_SYU_Speck.ks7.keyleng = 1                           ' �L�[��
    O_DEL_SYU_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfDup ' �L�[�t���O
    O_DEL_SYU_Speck.ks7.keytype = Chr(BtKtString)             ' �L�[�^�C�v
    O_DEL_SYU_Speck.ks7.reserve = &H0                         ' �\��ς�
    
    O_DEL_SYU_Speck.ks8.keypos = 45                           ' �L�[�|�W�V����
    O_DEL_SYU_Speck.ks8.keyleng = 8                           ' �L�[��
    O_DEL_SYU_Speck.ks8.keyflag = BtKfExt + BtKfSeg + BtKfDup ' �L�[�t���O
    O_DEL_SYU_Speck.ks8.keytype = Chr(BtKtString)             ' �L�[�^�C�v
    O_DEL_SYU_Speck.ks8.reserve = &H0                         ' �\��ς�
    
    O_DEL_SYU_Speck.ks9.keypos = 53                           ' �L�[�|�W�V����
    O_DEL_SYU_Speck.ks9.keyleng = 8                           ' �L�[��
    O_DEL_SYU_Speck.ks9.keyflag = BtKfExt + BtKfSeg + BtKfDup ' �L�[�t���O
    O_DEL_SYU_Speck.ks9.keytype = Chr(BtKtString)             ' �L�[�^�C�v
    O_DEL_SYU_Speck.ks9.reserve = &H0                         ' �\��ς�
    
    O_DEL_SYU_Speck.ks10.keypos = 15                          ' �L�[�|�W�V����
    O_DEL_SYU_Speck.ks10.keyleng = 1                          ' �L�[��
    O_DEL_SYU_Speck.ks10.keyflag = BtKfExt + BtKfSeg + BtKfDup ' �L�[�t���O
    O_DEL_SYU_Speck.ks10.keytype = Chr(BtKtString)            ' �L�[�^�C�v
    O_DEL_SYU_Speck.ks10.reserve = &H0                        ' �\��ς�
    
    O_DEL_SYU_Speck.ks11.keypos = 24                          ' �L�[�|�W�V����
    O_DEL_SYU_Speck.ks11.keyleng = 1                          ' �L�[��
    O_DEL_SYU_Speck.ks11.keyflag = BtKfExt + BtKfSeg + BtKfDup ' �L�[�t���O
    O_DEL_SYU_Speck.ks11.keytype = Chr(BtKtString)            ' �L�[�^�C�v
    O_DEL_SYU_Speck.ks11.reserve = &H0                        ' �\��ς�
    
    O_DEL_SYU_Speck.ks12.keypos = 25                          ' �L�[�|�W�V����
    O_DEL_SYU_Speck.ks12.keyleng = 20                         ' �L�[��
    O_DEL_SYU_Speck.ks12.keyflag = BtKfExt + BtKfSeg + BtKfDup ' �L�[�t���O
    O_DEL_SYU_Speck.ks12.keytype = Chr(BtKtString)            ' �L�[�^�C�v
    O_DEL_SYU_Speck.ks12.reserve = &H0                        ' �\��ς�
    
    O_DEL_SYU_Speck.ks13.keypos = 16                          ' �L�[�|�W�V����
    O_DEL_SYU_Speck.ks13.keyleng = 8                          ' �L�[��
    O_DEL_SYU_Speck.ks13.keyflag = BtKfExt + BtKfDup          ' �L�[�t���O
    O_DEL_SYU_Speck.ks13.keytype = Chr(BtKtString)            ' �L�[�^�C�v
    O_DEL_SYU_Speck.ks13.reserve = &H0                        ' �\��ς�
    
    sts = BTRV(BtOpCreate, O_DEL_SYU_POS, O_DEL_SYU_Speck, Len(O_DEL_SYU_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�폜�ςݏo�ח\��f�[�^")
        Exit Function
    End If

    O_DEL_SYU_Create = False

End Function
Function O_DEL_SYU_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*             �폜�ςݏo�ח\��f�[�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    O_DEL_SYU_Open = True
                                            '�폜�ςݏo�ח\��f�[�^�t���p�X�捞��
    sts = GetIni("FILE", O_DEL_SYU_ID, "CONV200605", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [O_DEL_SYU]�ǂݍ��݃G���[ ")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, O_DEL_SYU_POS, O_DEL_SYUREC, Len(O_DEL_SYUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = O_DEL_SYU_Create()        '�폜�ςݏo�ח\��f�[�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, O_DEL_SYU_POS, O_DEL_SYUREC, Len(O_DEL_SYUREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�폜�ςݏo�ח\��f�[�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�폜�ςݏo�ח\��f�[�^")
                Exit Function
        End Select
    Loop
    
    O_DEL_SYU_Open = False

End Function
