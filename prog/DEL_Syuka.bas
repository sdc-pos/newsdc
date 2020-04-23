Attribute VB_Name = "DEL_SYU"
Option Explicit
'********************************************************************
'*
'*              �폜�ςݏo�ח\��f�[�^  �t�@�C����`
'*              �V�@ڲ��đΉ� 2006.05.24
'********************************************************************
'�t�@�C���h�c
Public Const DEL_SYU_ID$ = "DEL_SYU"

'�y�[�W�T�C�Y
Public Const DEL_SYU_PG_SIZ% = 4096

'�|�W�V�����E�u���b�N
Public DEL_SYU_POS    As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type DEL_SYUREC_Tag
    WEL_ID(0 To 2)              As Byte     '�g�p�q�@ID
    PRG_ID(0 To 7)              As Byte     '�g�p���v���O����
    KAN_KBN(0 To 0)             As Byte     '�����敪
    DT_SYU(0 To 0)              As Byte     '�f�[�^���
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    KEY_CYU_KBN(0 To 0)         As Byte     '�����敪
    KEY_ID_NO(0 To 11)          As Byte     'ID-NO
    NAIGAI(0 To 0)              As Byte     '�����O
    KEY_HIN_NO(0 To 19)         As Byte     '�i�ڔԍ�
    KEY_MUKE_CODE(0 To 7)       As Byte     '���Ӑ�R�[�h
    KEY_SS_CODE(0 To 7)         As Byte     '������R�[�h
    KEY_SYUKA_YMD(0 To 7)       As Byte     '�o�ד��t
    '-----------------  νďo���ް��Ұ�ށ@��
    JGYOBA(0 To 7)              As Byte     '���Ə꺰��
    DATA_KBN(0 To 0)            As Byte     '�f�[�^�敪
    TORI_KBN(0 To 1)            As Byte     '����敪
    ID_NO(0 To 11)              As Byte     'ID-NO
    KAIKEI_JGYOBA(0 To 7)       As Byte     '��v�p���Ə꺰��
    SHISAN_JGYOBA(0 To 7)       As Byte     '���Y�Ǘ��p���Ə꺰��
    HIN_NO(0 To 19)             As Byte     '�i�ڔԍ�
    DEN_NO(0 To 9)              As Byte     '�`�[�ԍ�
    SURYO(0 To 6)               As Byte     '�o�א���
    MUKE_CODE(0 To 7)           As Byte     '���Ӑ�R�[�h
    SYUKO_SYUSI(0 To 7)         As Byte     '�݌Ɏ��x
    SHISAN_SYUSI(0 To 7)        As Byte     '���Y�Ǘ��p�݌Ɏ��x����
    HOJYO_SYUSI(0 To 7)         As Byte     '�⏕�݌Ɏ��x����
    SYUKO_YMD(0 To 7)           As Byte     '�o�ɓ�
    TANKA(0 To 9)               As Byte     '���ےP��
    ODER_NO(0 To 11)            As Byte     '�I�[�_�[�ԍ�
    ITEM_NO(0 To 4)             As Byte     '�A�C�e���ԍ�
    ODER_NO_R(0 To 4)           As Byte     '�����Ǘ��ԍ�����
    KOSO_KEITAI(0 To 9)         As Byte     '���`�Ժ���
    SYUKA_YMD(0 To 7)           As Byte     '�o�ח\���
    TANABAN1(0 To 9)            As Byte     '۹����1
    TANABAN2(0 To 9)            As Byte     '۹����2
    TANABAN3(0 To 9)            As Byte     '۹����3
    MUKE_NAME(0 To 23)          As Byte     '���Ӑ於��
    CYU_KBN(0 To 0)             As Byte     '�����敪
    CYU_KBN_NAME(0 To 39)       As Byte     '�����敪����
    ORIGIN1(0 To 9)             As Byte     '���Y��1
    ORIGIN2(0 To 9)             As Byte     '���Y��2
    BIKOU2(0 To 39)             As Byte     '���l2
    HAN_KBN(0 To 0)             As Byte     '�̔��敪
    CHOKU_KBN(0 To 0)           As Byte     '�����w���敪
    UNIT_ID_NO(0 To 11)         As Byte     '�ƯďC���Ǘ��ԍ�
    ZAIKO_HIKIATE(0 To 2)       As Byte     '�݌Ɉ�������
    GOKON_KANRI_NO(0 To 7)      As Byte     '�����Ǘ��ԍ�
    JYUCHU_ZAN(0 To 6)          As Byte     '�󒍎c����
    KYOKYU_KBN(0 To 0)          As Byte     '�����敪
    SHOHIN_SYUSI(0 To 7)        As Byte     '���i���[�i�݌Ɏ��x����
    S_SHISAN_SYUSI(0 To 7)      As Byte     '���i���[�i���Y�Ǘ����x����
    S_HOJYO_SYUSI(0 To 7)       As Byte     '���i���[�i�⏕���x����
    BIKOU1(0 To 39)             As Byte     '���l1
    CHOHA_KBN(0 To 0)           As Byte     '���[�敪
    JYU_HIN_NO(0 To 39)         As Byte     '��t�i�ڔԍ�
    HIN_NAME(0 To 39)           As Byte     '�i��
    HIN_CHANGE_KBN(0 To 0)      As Byte     '�i�ڔԍ��ύX�敪
    MODULE_EXCHANGE(0 To 0)     As Byte     'Ӽޭ�ٌ����敪
    ZAIKO_SYUSI(0 To 7)         As Byte     '�c�݌ɂ܂Ƃߍ݌Ɏ��x����
    ZAN_SHISAN_SYUSI(0 To 7)    As Byte     '�c�݌ɂ܂Ƃߎ��Y�Ǘ����x����
    ZAN_HOJYO_SYUSI(0 To 7)     As Byte     '�c�݌ɂ܂Ƃߕ⏕���x����
    NOUKI_YMD(0 To 7)           As Byte     '�w��[��
    SERVICE_KANRI_NO(0 To 8)    As Byte     '���޽��ЊǗ��ԍ�
    KISHU_CODE(0 To 2)          As Byte     '�@��i�ں���
    ENVIRONMENT_KBN(0 To 0)     As Byte     '����敔�i�敪
    SS_CODE(0 To 7)             As Byte     '��������溰��
    KEPIN_KAIJYO(0 To 0)        As Byte     '���i�����敪
    '-----------------  νďo���ް��Ұ�ށ@��
    HIN_NAI(0 To 19)            As Byte     '�i�ԁi�����j
    HTANABAN(0 To 7)            As Byte     '�z�X�g�I��
    PRINT_YMD(0 To 7)           As Byte     '�o�ɕ\������t
    KAN_YMD(0 To 7)             As Byte     '�������t
    KENPIN_YMD(0 To 7)          As Byte     '���i���t
    TOK_KBN(0 To 0)             As Byte     '������敪
    JITU_SURYO(0 To 6)          As Byte     '�o�Ɏ��ѐ���
    INS_NOW(0 To 13)            As Byte     '�捞�ݓ���
    KENPIN_TANTO_CODE(0 To 4)   As Byte     '���i�S���Һ��� 2006.07.20
    KENPIN_HMS(0 To 5)          As Byte     '���i����       2006.07.20
    
    LK_MUKE_CODE(0 To 7)        As Byte     '����ݸ�p������2006.07.20
        
    FILLER(0 To 47)             As Byte     'FILLER
End Type

'�f�[�^�E�o�b�t�@
Public DEL_SYUREC As DEL_SYUREC_Tag

'�L�[��`
Type KEY0_DEL_SYU            '�j�d�x�O
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    KEY_CYU_KBN(0 To 0)         As Byte     '�����敪
    KEY_MUKE_CODE(0 To 7)       As Byte     '���Ӑ�R�[�h
    KEY_SS_CODE(0 To 7)         As Byte     '������R�[�h
    KEY_HIN_NO(0 To 19)         As Byte     '�i�ڔԍ�
    KEY_SYUKA_YMD(0 To 7)       As Byte     '�o�ד��t
End Type

Type KEY1_DEL_SYU           '�j�d�x�P
    KEY_SYUKA_YMD(0 To 7)       As Byte     '�o�ד��t
End Type

Type KEY2_DEL_SYU            '�j�d�x�Q
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    KEY_MUKE_CODE(0 To 7)       As Byte     '���Ӑ�R�[�h
    KEY_SS_CODE(0 To 7)         As Byte     '������R�[�h
    KEY_CYU_KBN(0 To 0)         As Byte     '�����敪
    NAIGAI(0 To 0)              As Byte     '�����O
    KEY_HIN_NO(0 To 19)         As Byte     '�i�ڔԍ�
    KEY_ID_NO(0 To 11)           As Byte     'ID-NO
End Type


'�L�[�E�f�[�^
Public K0_DEL_SYU               As KEY0_DEL_SYU
Public K1_DEL_SYU               As KEY1_DEL_SYU
Public K2_DEL_SYU               As KEY2_DEL_SYU

Type DEL_SYU_FSpeck
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

Private DEL_SYU_Speck As DEL_SYU_FSpeck

Private Function DEL_SYU_Create() As Integer
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

    DEL_SYU_Create = True
                                            '�폜�ςݏo�ח\��f�[�^�t���p�X�捞��
    sts = GetIni("FILE", DEL_SYU_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [DEL_SYU]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    DEL_SYU_Speck.fs.recoleng = Len(DEL_SYUREC)               ' ���R�[�h��
    DEL_SYU_Speck.fs.PageSize = DEL_SYU_PG_SIZ              ' �y�[�W�T�C�Y
    DEL_SYU_Speck.fs.idexnumb = 3                           ' �C���f�b�N�X��
    DEL_SYU_Speck.fs.fileflag = 0                           ' �t�@�C���t���O
    DEL_SYU_Speck.fs.reserve = &H0                          ' �\��ς�
'---------------------------------------------------        �L�[�O
    
    DEL_SYU_Speck.ks0.keypos = 14                           ' �L�[�|�W�V����
    DEL_SYU_Speck.ks0.keyleng = 1                           ' �L�[��
    DEL_SYU_Speck.ks0.keyflag = BtKfExt + BtKfSeg + BtKfDup ' �L�[�t���O
    DEL_SYU_Speck.ks0.keytype = Chr(BtKtString)             ' �L�[�^�C�v
    DEL_SYU_Speck.ks0.reserve = &H0                         ' �\��ς�
    
    DEL_SYU_Speck.ks1.keypos = 15                           ' �L�[�|�W�V����
    DEL_SYU_Speck.ks1.keyleng = 1                           ' �L�[��
    DEL_SYU_Speck.ks1.keyflag = BtKfExt + BtKfSeg + BtKfDup ' �L�[�t���O
    DEL_SYU_Speck.ks1.keytype = Chr(BtKtString)             ' �L�[�^�C�v
    DEL_SYU_Speck.ks1.reserve = &H0                         ' �\��ς�
    
    DEL_SYU_Speck.ks2.keypos = 49                           ' �L�[�|�W�V����
    DEL_SYU_Speck.ks2.keyleng = 8                           ' �L�[��
    DEL_SYU_Speck.ks2.keyflag = BtKfExt + BtKfSeg + BtKfDup ' �L�[�t���O
    DEL_SYU_Speck.ks2.keytype = Chr(BtKtString)             ' �L�[�^�C�v
    DEL_SYU_Speck.ks2.reserve = &H0                         ' �\��ς�
    
    DEL_SYU_Speck.ks3.keypos = 57                           ' �L�[�|�W�V����
    DEL_SYU_Speck.ks3.keyleng = 8                           ' �L�[��
    DEL_SYU_Speck.ks3.keyflag = BtKfExt + BtKfSeg + BtKfDup ' �L�[�t���O
    DEL_SYU_Speck.ks3.keytype = Chr(BtKtString)             ' �L�[�^�C�v
    DEL_SYU_Speck.ks3.reserve = &H0                         ' �\��ς�
    
    DEL_SYU_Speck.ks4.keypos = 29                           ' �L�[�|�W�V����
    DEL_SYU_Speck.ks4.keyleng = 20                          ' �L�[��
    DEL_SYU_Speck.ks4.keyflag = BtKfExt + BtKfSeg + BtKfDup ' �L�[�t���O
    DEL_SYU_Speck.ks4.keytype = Chr(BtKtString)             ' �L�[�^�C�v
    DEL_SYU_Speck.ks4.reserve = &H0                         ' �\��ς�
    
    DEL_SYU_Speck.ks5.keypos = 65                           ' �L�[�|�W�V����
    DEL_SYU_Speck.ks5.keyleng = 8                           ' �L�[��
    DEL_SYU_Speck.ks5.keyflag = BtKfExt + BtKfDup           ' �L�[�t���O
    DEL_SYU_Speck.ks5.keytype = Chr(BtKtString)             ' �L�[�^�C�v
    DEL_SYU_Speck.ks5.reserve = &H0                         ' �\��ς�

'---------------------------------------------------        �L�[�P
    
    DEL_SYU_Speck.ks6.keypos = 65                           ' �L�[�|�W�V����
    DEL_SYU_Speck.ks6.keyleng = 8                           ' �L�[��
    DEL_SYU_Speck.ks6.keyflag = BtKfExt + BtKfDup           ' �L�[�t���O
    DEL_SYU_Speck.ks6.keytype = Chr(BtKtString)             ' �L�[�^�C�v
    DEL_SYU_Speck.ks6.reserve = &H0                         ' �\��ς�
    
'---------------------------------------------------        �L�[�Q
    
    DEL_SYU_Speck.ks7.keypos = 14                           ' �L�[�|�W�V����
    DEL_SYU_Speck.ks7.keyleng = 1                           ' �L�[��
    DEL_SYU_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfDup ' �L�[�t���O
    DEL_SYU_Speck.ks7.keytype = Chr(BtKtString)             ' �L�[�^�C�v
    DEL_SYU_Speck.ks7.reserve = &H0                         ' �\��ς�
    
    DEL_SYU_Speck.ks8.keypos = 49                           ' �L�[�|�W�V����
    DEL_SYU_Speck.ks8.keyleng = 8                           ' �L�[��
    DEL_SYU_Speck.ks8.keyflag = BtKfExt + BtKfSeg + BtKfDup ' �L�[�t���O
    DEL_SYU_Speck.ks8.keytype = Chr(BtKtString)             ' �L�[�^�C�v
    DEL_SYU_Speck.ks8.reserve = &H0                         ' �\��ς�
    
    DEL_SYU_Speck.ks9.keypos = 57                           ' �L�[�|�W�V����
    DEL_SYU_Speck.ks9.keyleng = 8                           ' �L�[��
    DEL_SYU_Speck.ks9.keyflag = BtKfExt + BtKfSeg + BtKfDup ' �L�[�t���O
    DEL_SYU_Speck.ks9.keytype = Chr(BtKtString)             ' �L�[�^�C�v
    DEL_SYU_Speck.ks9.reserve = &H0                         ' �\��ς�
    
    DEL_SYU_Speck.ks10.keypos = 15                          ' �L�[�|�W�V����
    DEL_SYU_Speck.ks10.keyleng = 1                          ' �L�[��
    DEL_SYU_Speck.ks10.keyflag = BtKfExt + BtKfSeg + BtKfDup ' �L�[�t���O
    DEL_SYU_Speck.ks10.keytype = Chr(BtKtString)            ' �L�[�^�C�v
    DEL_SYU_Speck.ks10.reserve = &H0                        ' �\��ς�
    
    DEL_SYU_Speck.ks11.keypos = 28                          ' �L�[�|�W�V����
    DEL_SYU_Speck.ks11.keyleng = 1                          ' �L�[��
    DEL_SYU_Speck.ks11.keyflag = BtKfExt + BtKfSeg + BtKfDup ' �L�[�t���O
    DEL_SYU_Speck.ks11.keytype = Chr(BtKtString)            ' �L�[�^�C�v
    DEL_SYU_Speck.ks11.reserve = &H0                        ' �\��ς�
    
    DEL_SYU_Speck.ks12.keypos = 29                          ' �L�[�|�W�V����
    DEL_SYU_Speck.ks12.keyleng = 20                         ' �L�[��
    DEL_SYU_Speck.ks12.keyflag = BtKfExt + BtKfSeg + BtKfDup ' �L�[�t���O
    DEL_SYU_Speck.ks12.keytype = Chr(BtKtString)            ' �L�[�^�C�v
    DEL_SYU_Speck.ks12.reserve = &H0                        ' �\��ς�
    
    DEL_SYU_Speck.ks13.keypos = 16                          ' �L�[�|�W�V����
    DEL_SYU_Speck.ks13.keyleng = 12                          ' �L�[��
    DEL_SYU_Speck.ks13.keyflag = BtKfExt + BtKfDup          ' �L�[�t���O
    DEL_SYU_Speck.ks13.keytype = Chr(BtKtString)            ' �L�[�^�C�v
    DEL_SYU_Speck.ks13.reserve = &H0                        ' �\��ς�
    
    sts = BTRV(BtOpCreate, DEL_SYU_POS, DEL_SYU_Speck, Len(DEL_SYU_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�폜�ςݏo�ח\��f�[�^")
        Exit Function
    End If

    DEL_SYU_Create = False

End Function
Function DEL_SYU_Open(Mode As Integer) As Integer
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
    
    DEL_SYU_Open = True
                                            '�폜�ςݏo�ח\��f�[�^�t���p�X�捞��
    sts = GetIni("FILE", DEL_SYU_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [DEL_SYU]�ǂݍ��݃G���[ ")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, DEL_SYU_POS, DEL_SYUREC, Len(DEL_SYUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = DEL_SYU_Create()        '�폜�ςݏo�ח\��f�[�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, DEL_SYU_POS, DEL_SYUREC, Len(DEL_SYUREC), ByVal FullPath, Len(FullPath), Mode)
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
    
    DEL_SYU_Open = False

End Function
