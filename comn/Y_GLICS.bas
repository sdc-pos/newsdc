Attribute VB_Name = "Y_GLICS"
Option Explicit
'********************************************************************
'*                                                                  *
'*              ���ח\��f�[�^  �t�@�C����`                        *
'*                                                                  *
'********************************************************************
'�t�@�C���h�c
Public Const Y_GLICS_ID$ = "Y_GLICS"

'�y�[�W�T�C�Y
Public Const Y_GLICS_PG_SIZ% = 2048

'�|�W�V�����E�u���b�N
Public Y_GLICS_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type Y_GLICSREC_Tag
    KAN_KBN(0 To 0)             As Byte     '�����敪
    DT_SYU(0 To 0)              As Byte     '�f�[�^���
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    NAIGAI(0 To 0)              As Byte     '�����O
    TEXT_NO(0 To 8)             As Byte     '�e�L�X�g��
    
    '-----------------  νē����ް��Ұ�ށ@��
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
    SYUKO_YMD(0 To 7)           As Byte     '�o�ɓ��t
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
    CYU_KBN_NAME(0 To 9)        As Byte     '�����敪����
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
    JYU_HIN_NO(0 To 19)         As Byte     '��t�i�ڔԍ�
    HIN_NAME(0 To 19)           As Byte     '�i��
    HIN_CHANGE_KBN(0 To 0)      As Byte     '�i�ڔԍ��ύX�敪
    MODULE_EXCHANGE(0 To 0)     As Byte     'Ӽޭ�ٌ����敪
    ZAIKO_SYUSI(0 To 7)         As Byte     '�c�݌ɂ܂Ƃߍ݌Ɏ��x����
    ZAN_SHISAN_SYUSI(0 To 7)    As Byte     '�c�݌ɂ܂Ƃߎ��Y�Ǘ����x����
    ZAN_HOJYO_SYUSI(0 To 7)     As Byte     '�c�݌ɂ܂Ƃߕ⏕���x����
    NOUKI_YMD(0 To 7)           As Byte     '�w��[��
    SERVICE_KANRI_NO(0 To 8)    As Byte     '���޽��ЊǗ��ԍ�
    KI_HIN_NO(0 To 2)           As Byte     '�@��i�ں���
    ENVIRONMENT_KBN(0 To 0)     As Byte     '����敔�i�敪
    SS_CODE(0 To 7)             As Byte     '��������溰��
    KEPIN_KAIJYO(0 To 0)        As Byte     '���i�����敪
    '-----------------  νē����ް��Ұ�ށ@��
    
    KAN_DT(0 To 7)              As Byte     '�������t
    BEF_NYU_QTY(0 To 7)         As Byte     '��s���א�
    YOSAN_FROM(0 To 4)          As Byte     '�\�Z�P�ʁi���j
    YOSAN_TO(0 To 4)            As Byte     '�\�Z�P�ʁi��j
    HTANABAN(0 To 7)            As Byte     '�W���I��
    HIN_NAI(0 To 12)            As Byte     '�i�ԁi�����j
    H_SOKO(0 To 7)              As Byte     'νđq�� 2006.10.17
    
    NYU_LIST_OUT(0 To 0)        As Byte     '���ɗ\��o���׸� 2007.06.12
    
    
    
    '-----------------  ��GLICS���ڒǉ� 2007.06.15
    CYOK_KBN(0 To 0)            As Byte     '�����敪
    IO_KBN(0 To 0)              As Byte     '���o�ɋ敪
    PM_KBN(0 To 0)              As Byte     '�ԍ��敪
    DEN_SYU(0 To 0)             As Byte     '�`�[���
    SYUK_CODE(0 To 4)           As Byte     '�x����^�o�א�
    SYUK_NAME(0 To 19)          As Byte     '�x����^�o�א於
    
    
    INS_NOW(0 To 13)            As Byte     '�}���N���������b
    '-----------------  ��GLICS���ڒǉ� 2007.06.15
    
    '----------------   2010.07.08 ��
    GENSANKOKU(0 To 19)         As Byte     '���Y����
    GEN_GENSANKOKU(0 To 19)     As Byte     '�����\�����Y����
    SHIIRE_WORK_CENTER(0 To 7)  As Byte     '���ގd����ܰ�����
    KANKYO_KBN(0 To 2)          As Byte     '����ދ敪
    KANKYO_KBN_ST(0 To 7)       As Byte     '����ދ敪�K�p�J�n
    KANKYO_KBN_SURYO(0 To 9)    As Byte     '����ދ敪����
    ID_NO2(0 To 11)             As Byte     'ID_NO
    AITESAKI_CODE(0 To 15)      As Byte     '����溰��
    JYUCHU_YMD(0 To 7)          As Byte     '�󒍔N����
    SHITEI_NOUKI_YMD(0 To 7)    As Byte     '�w��[���N����
    LIST_OUT_END_F(0 To 0)      As Byte     '����ؽďo��F
    NYUKO_TANABAN(0 To 7)       As Byte     '���ɒI��
    MAEGARI_SURYO(0 To 7)       As Byte     '�O�ؑ��E��
    '----------------   2010.07.08 ��
    
    '----------------   2011.03.23 ��
    MOTO_PROG_ID(0 To 7)        As Byte     '�������v���O����
    MOTO_TEXT_NO(0 To 8)        As Byte     '���e�L�X�g��
    '----------------   2011.03.23 ��
    
    
    
    FILLER(0 To 23)            As Byte      '40-->23    2011.03.23
End Type

'�f�[�^�E�o�b�t�@
Public Y_GLICSREC                  As Y_GLICSREC_Tag

'�L�[��`
Type KEY0_Y_GLICS            '�j�d�x�O
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    SYUKA_YMD(0 To 7)           As Byte     '�o�ד�
    TEXT_NO(0 To 8)             As Byte     '�e�L�X�g��
End Type

Type KEY1_Y_GLICS            '�j�d�x�P
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    KAN_KBN(0 To 0)             As Byte     '�����敪
    NAIGAI(0 To 0)              As Byte     '�����O
    HIN_NO(0 To 19)             As Byte     '�i�ڔԍ�
    SYUKA_YMD(0 To 7)           As Byte     '�o�ד�
    TEXT_NO(0 To 8)             As Byte     '�e�L�X�g��
End Type

Type KEY2_Y_GLICS            '�j�d�x�Q
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    SYUKA_YMD(0 To 7)           As Byte     '�o�ד�
    HIN_NO(0 To 19)             As Byte     '�i�ڔԍ�
    NAIGAI(0 To 0)              As Byte     '�����O
End Type

Type KEY3_Y_GLICS            '�j�d�x�R
    SYUKA_YMD(0 To 7)           As Byte     '�o�ד�
End Type



'�L�[�E�f�[�^
Public K0_Y_GLICS                 As KEY0_Y_GLICS
Public K1_Y_GLICS                 As KEY1_Y_GLICS
Public K2_Y_GLICS                 As KEY2_Y_GLICS
Public K3_Y_GLICS                 As KEY3_Y_GLICS

Private Type Y_GLICS_FSpeck
    fs      As BtFileSpeck              ' ̧�� ��߯��\����
    ks0     As BtKeySpeck               ' �� ��߯��\����
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck
    ks4     As BtKeySpeck
    ks5     As BtKeySpeck
    ks6     As BtKeySpeck
    ks7     As BtKeySpeck
    ks8     As BtKeySpeck
    ks9     As BtKeySpeck
    ks10    As BtKeySpeck
    ks11    As BtKeySpeck
    ks12    As BtKeySpeck
    ks13    As BtKeySpeck
    ks14    As BtKeySpeck
End Type

Private Y_GLICS_Speck As Y_GLICS_FSpeck

Private Function Y_GLICS_Create() As Integer
'********************************************************************
'*                                                                  *
'*              ���ח\��f�[�^  �b�q�d�`�s�d                        *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    Y_GLICS_Create = True
                                            '���ח\��f�[�^�t���p�X�捞��
    sts = GetIni("FILE", Y_GLICS_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [Y_GLICS]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    Y_GLICS_Speck.fs.recoleng = Len(Y_GLICSREC)     ' ���R�[�h��
    Y_GLICS_Speck.fs.PageSize = Y_GLICS_PG_SIZ      ' �y�[�W�T�C�Y
    Y_GLICS_Speck.fs.idexnumb = 4                 ' �C���f�b�N�X��
    Y_GLICS_Speck.fs.fileflag = 0                 ' �t�@�C���t���O
    Y_GLICS_Speck.fs.reserve = &H0                ' �\��ς�
    '-------------------------------------------
                                                ' �L�[�O
    Y_GLICS_Speck.ks0.keypos = 3                  ' �L�[�|�W�V����
    Y_GLICS_Speck.ks0.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    Y_GLICS_Speck.ks0.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_GLICS_Speck.ks0.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    Y_GLICS_Speck.ks0.reserve = &H0               ' �\��ς�
                                                ' �L�[�O
    Y_GLICS_Speck.ks1.keypos = 172                ' �L�[�|�W�V����
    Y_GLICS_Speck.ks1.keyleng = 8                 ' �L�[��
                                                ' �L�[�t���O
    Y_GLICS_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_GLICS_Speck.ks1.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    Y_GLICS_Speck.ks1.reserve = &H0               ' �\��ς�
                                                ' �L�[�O
    Y_GLICS_Speck.ks2.keypos = 5                  ' �L�[�|�W�V����
    Y_GLICS_Speck.ks2.keyleng = 9                 ' �L�[��
    Y_GLICS_Speck.ks2.keyflag = BtKfExt + BtKfChg ' �L�[�t���O
    Y_GLICS_Speck.ks2.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    Y_GLICS_Speck.ks2.reserve = &H0               ' �\��ς�
    '-------------------------------------------
                                                
                                                ' �L�[�P
    Y_GLICS_Speck.ks3.keypos = 3                  ' �L�[�|�W�V����
    Y_GLICS_Speck.ks3.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    Y_GLICS_Speck.ks3.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_GLICS_Speck.ks3.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    Y_GLICS_Speck.ks3.reserve = &H0               ' �\��ς�
                                                ' �L�[�P
    Y_GLICS_Speck.ks4.keypos = 1                  ' �L�[�|�W�V����
    Y_GLICS_Speck.ks4.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    Y_GLICS_Speck.ks4.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_GLICS_Speck.ks4.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    Y_GLICS_Speck.ks4.reserve = &H0               ' �\��ς�
                                                ' �L�[�P
    Y_GLICS_Speck.ks5.keypos = 4                 ' �L�[�|�W�V����
    Y_GLICS_Speck.ks5.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    Y_GLICS_Speck.ks5.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_GLICS_Speck.ks5.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    Y_GLICS_Speck.ks5.reserve = &H0               ' �\��ς�
                                                ' �L�[�P
    Y_GLICS_Speck.ks6.keypos = 53                 ' �L�[�|�W�V����
    Y_GLICS_Speck.ks6.keyleng = 20                ' �L�[��
                                                ' �L�[�t���O
    Y_GLICS_Speck.ks6.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_GLICS_Speck.ks6.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    Y_GLICS_Speck.ks6.reserve = &H0               ' �\��ς�
                                                ' �L�[�P
    Y_GLICS_Speck.ks7.keypos = 172                ' �L�[�|�W�V����
    Y_GLICS_Speck.ks7.keyleng = 8                 ' �L�[��
                                                ' �L�[�t���O
    Y_GLICS_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_GLICS_Speck.ks7.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    Y_GLICS_Speck.ks7.reserve = &H0               ' �\��ς�
                                                ' �L�[�P
    Y_GLICS_Speck.ks8.keypos = 5                ' �L�[�|�W�V����
    Y_GLICS_Speck.ks8.keyleng = 9                 ' �L�[��
                                                ' �L�[�t���O
    Y_GLICS_Speck.ks8.keyflag = BtKfExt + BtKfChg
    Y_GLICS_Speck.ks8.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    Y_GLICS_Speck.ks8.reserve = &H0               ' �\��ς�
    '-------------------------------------------
                                                
                                                ' �L�[�Q
    Y_GLICS_Speck.ks9.keypos = 3                  ' �L�[�|�W�V����
    Y_GLICS_Speck.ks9.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    Y_GLICS_Speck.ks9.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_GLICS_Speck.ks9.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    Y_GLICS_Speck.ks9.reserve = &H0               ' �\��ς�
                                                ' �L�[�Q
    Y_GLICS_Speck.ks10.keypos = 172               ' �L�[�|�W�V����
    Y_GLICS_Speck.ks10.keyleng = 8                ' �L�[��
                                                ' �L�[�t���O
    Y_GLICS_Speck.ks10.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_GLICS_Speck.ks10.keytype = Chr(BtKtString)  ' �L�[�^�C�v
    Y_GLICS_Speck.ks10.reserve = &H0              ' �\��ς�
                                                ' �L�[�Q
    Y_GLICS_Speck.ks11.keypos = 53                ' �L�[�|�W�V����
    Y_GLICS_Speck.ks11.keyleng = 20               ' �L�[��
                                                ' �L�[�t���O
    Y_GLICS_Speck.ks11.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_GLICS_Speck.ks11.keytype = Chr(BtKtString)  ' �L�[�^�C�v
    Y_GLICS_Speck.ks11.reserve = &H0              ' �\��ς�
                                                ' �L�[�Q
    Y_GLICS_Speck.ks12.keypos = 4                 ' �L�[�|�W�V����
    Y_GLICS_Speck.ks12.keyleng = 1                ' �L�[��
                                                ' �L�[�t���O
    Y_GLICS_Speck.ks12.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_GLICS_Speck.ks12.keytype = Chr(BtKtString)  ' �L�[�^�C�v
    Y_GLICS_Speck.ks12.reserve = &H0              ' �\��ς�
                                                ' �L�[�Q
    Y_GLICS_Speck.ks13.keypos = 5                 ' �L�[�|�W�V����
    Y_GLICS_Speck.ks13.keyleng = 9                ' �L�[��
                                                ' �L�[�t���O
    Y_GLICS_Speck.ks13.keyflag = BtKfExt + BtKfChg
    Y_GLICS_Speck.ks13.keytype = Chr(BtKtString)  ' �L�[�^�C�v
    Y_GLICS_Speck.ks13.reserve = &H0              ' �\��ς�
    '-------------------------------------------
                                                
                                                ' �L�[�R
    Y_GLICS_Speck.ks14.keypos = 172               ' �L�[�|�W�V����
    Y_GLICS_Speck.ks14.keyleng = 8                ' �L�[��
                                                ' �L�[�t���O
    Y_GLICS_Speck.ks14.keyflag = BtKfExt + BtKfChg + BtKfDup
    Y_GLICS_Speck.ks14.keytype = Chr(BtKtString)  ' �L�[�^�C�v
    Y_GLICS_Speck.ks14.reserve = &H0              ' �\��ς�
    '-------------------------------------------
    
    sts = BTRV(BtOpCreate, Y_GLICS_POS, Y_GLICS_Speck, Len(Y_GLICS_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���ח\��f�[�^")
        Y_GLICS_Create = True
        Exit Function
    End If

    Y_GLICS_Create = False

End Function

Function Y_GLICS_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              ���ח\��f�[�^  �n�o�d�m                            *
'*                                                                  *
'*      ��  ��:Open Mode(Btrieve�Q��)                               *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    Y_GLICS_Open = True
                                            '���ח\��f�[�^�t���p�X�捞��
    sts = GetIni("FILE", Y_GLICS_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [Y_GLICS]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = Y_GLICS_Create()        '���ח\��f�[�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���ח\��f�[�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���ח\��f�[�^")
                Exit Function
        End Select
    Loop
    
    Y_GLICS_Open = False

End Function


