Attribute VB_Name = "Y_NYU"
Option Explicit
'********************************************************************
'*                                                                  *
'*              ���ח\��f�[�^  �t�@�C����`                        *
'*                                                                  *
'********************************************************************
'�t�@�C���h�c
Public Const Y_NYU_ID$ = "Y_NYU"

'�y�[�W�T�C�Y
Public Const Y_NYU_PG_SIZ% = 2048

'�|�W�V�����E�u���b�N
Public Y_NYU_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type Y_NYUREC_Tag
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
    H_SOKO(0 To 1)              As Byte     'νđq�� 2006.10.17
            
    NYU_LIST_OUT(0 To 0)        As Byte     '���ɗ\��o���׸� 2007.06.12    ���ݖ��g�p 0:�f�[�^�o�͑Ώ� 9:�o�͍�(�������͏o�͑ΏۊO)
    
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
    LIST_OUT_END_F(0 To 0)      As Byte     '���Ɋ֘Aؽďo��F    0:�������Y�����i���ɊǗ�ؽĂ܂��͓��Ɂ^�I������ؽĂ�������
                                                                '9:�������Y�����i���ɊǗ�ؽĂ����Ɂ^�I������ؽĂ�������
    LIST_NYU_KANRI_F(0 To 0)    As Byte     '���ɊǗ�ؽďo��F�@�@�u�������Y�����i���ɊǗ�ؽėp�v 0:����Ώ�(�����) 8:����ΏۊO�@9:�����(0��9)
    LIST_NYU_CHECK_F(0 To 0)    As Byte     '��������ؽďo��F    �u���Ɂ^�I������ؽėp�v�@0:����� 9:�����
    NYUKO_TANABAN(0 To 7)       As Byte     '���ɒI��
    MAEGARI_SURYO(0 To 7)       As Byte     '�O�ؑ��E��
    
    INS_TANTO(0 To 4)           As Byte     '�ǉ��@�S���ҁ@     2009.01.21
    Ins_DateTime(0 To 13)       As Byte     '�ǉ��@����         2009.01.21

    UPD_TANTO(0 To 4)           As Byte     '�X�V�@�S���ҁ@     2005.11.15
    UPD_DATETIME(0 To 13)       As Byte     '�X�V�@����         2005.11.15
    
    '----------------   2010.07.08 ��
    
    '----------------   2011.03.23 ��
    MOTO_PROG_ID(0 To 7)        As Byte     '�������v���O����
    MOTO_TEXT_NO(0 To 8)        As Byte     '���e�L�X�g��
    '----------------   2011.03.23 ��
    
    JITU_SURYO(0 To 6)          As Byte     '���ѐ���           2015.01.21
    
    
    FILLER(0 To 25)             As Byte      '49-->32-->25       2011.03.23-->2015.01.21
End Type

'�f�[�^�E�o�b�t�@
Public Y_NYUREC                  As Y_NYUREC_Tag

'�L�[��`
Type KEY0_Y_NYU            '�j�d�x�O
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    SYUKA_YMD(0 To 7)           As Byte     '�o�ד�
    TEXT_NO(0 To 8)             As Byte     '�e�L�X�g��
End Type

Type KEY1_Y_NYU            '�j�d�x�P
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    KAN_KBN(0 To 0)             As Byte     '�����敪
    NAIGAI(0 To 0)              As Byte     '�����O
    HIN_NO(0 To 19)             As Byte     '�i�ڔԍ�
    SYUKA_YMD(0 To 7)           As Byte     '�o�ד�
    TEXT_NO(0 To 8)             As Byte     '�e�L�X�g��
End Type

Type KEY2_Y_NYU            '�j�d�x�Q
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    SYUKA_YMD(0 To 7)           As Byte     '�o�ד�
    HIN_NO(0 To 19)             As Byte     '�i�ڔԍ�
    NAIGAI(0 To 0)              As Byte     '�����O
    TEXT_NO(0 To 8)             As Byte     '�e�L�X�g��     '2016.06.20
End Type

Type KEY3_Y_NYU            '�j�d�x�R
    SYUKA_YMD(0 To 7)           As Byte     '�o�ד�
End Type

    
Type KEY4_Y_NYU            '�j�d�x�S        2010.07.12
    LIST_OUT_END_F(0 To 0)      As Byte     '����ؽďo��F
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    NAIGAI(0 To 0)              As Byte     '�����O
    HIN_NO(0 To 19)             As Byte     '�i�ڔԍ�
End Type
    


'�L�[�E�f�[�^
Public K0_Y_NYU                 As KEY0_Y_NYU
Public K1_Y_NYU                 As KEY1_Y_NYU
Public K2_Y_NYU                 As KEY2_Y_NYU
Public K3_Y_NYU                 As KEY3_Y_NYU
'2010.07.12
Public K4_Y_NYU                 As KEY4_Y_NYU

Private Type Y_NYU_FSpeck
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

    ks15    As BtKeySpeck       '2010.07.12
    ks16    As BtKeySpeck       '2010.07.12
    ks17    As BtKeySpeck       '2010.07.12
    ks18    As BtKeySpeck       '2010.07.12

End Type

Private Y_NYU_Speck As Y_NYU_FSpeck

Private Function Y_NYU_Create() As Integer
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

    Y_NYU_Create = True
                                            '���ח\��f�[�^�t���p�X�捞��
    sts = GetIni("FILE", Y_NYU_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [Y_NYU]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    Y_NYU_Speck.fs.recoleng = Len(Y_NYUREC)     ' ���R�[�h��
    Y_NYU_Speck.fs.PageSize = Y_NYU_PG_SIZ      ' �y�[�W�T�C�Y
    Y_NYU_Speck.fs.idexnumb = 5                 ' �C���f�b�N�X��
    Y_NYU_Speck.fs.fileflag = 0                 ' �t�@�C���t���O
    Y_NYU_Speck.fs.reserve = &H0                ' �\��ς�
    '-------------------------------------------
                                                ' �L�[�O
    Y_NYU_Speck.ks0.keypos = 3                  ' �L�[�|�W�V����
    Y_NYU_Speck.ks0.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    Y_NYU_Speck.ks0.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_NYU_Speck.ks0.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    Y_NYU_Speck.ks0.reserve = &H0               ' �\��ς�
                                                ' �L�[�O
    Y_NYU_Speck.ks1.keypos = 172                ' �L�[�|�W�V����
    Y_NYU_Speck.ks1.keyleng = 8                 ' �L�[��
                                                ' �L�[�t���O
    Y_NYU_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_NYU_Speck.ks1.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    Y_NYU_Speck.ks1.reserve = &H0               ' �\��ς�
                                                ' �L�[�O
    Y_NYU_Speck.ks2.keypos = 5                  ' �L�[�|�W�V����
    Y_NYU_Speck.ks2.keyleng = 9                 ' �L�[��
    Y_NYU_Speck.ks2.keyflag = BtKfExt + BtKfChg ' �L�[�t���O
    Y_NYU_Speck.ks2.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    Y_NYU_Speck.ks2.reserve = &H0               ' �\��ς�
    '-------------------------------------------
                                                
                                                ' �L�[�P
    Y_NYU_Speck.ks3.keypos = 3                  ' �L�[�|�W�V����
    Y_NYU_Speck.ks3.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    Y_NYU_Speck.ks3.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_NYU_Speck.ks3.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    Y_NYU_Speck.ks3.reserve = &H0               ' �\��ς�
                                                ' �L�[�P
    Y_NYU_Speck.ks4.keypos = 1                  ' �L�[�|�W�V����
    Y_NYU_Speck.ks4.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    Y_NYU_Speck.ks4.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_NYU_Speck.ks4.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    Y_NYU_Speck.ks4.reserve = &H0               ' �\��ς�
                                                ' �L�[�P
    Y_NYU_Speck.ks5.keypos = 4                 ' �L�[�|�W�V����
    Y_NYU_Speck.ks5.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    Y_NYU_Speck.ks5.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_NYU_Speck.ks5.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    Y_NYU_Speck.ks5.reserve = &H0               ' �\��ς�
                                                ' �L�[�P
    Y_NYU_Speck.ks6.keypos = 53                 ' �L�[�|�W�V����
    Y_NYU_Speck.ks6.keyleng = 20                ' �L�[��
                                                ' �L�[�t���O
    Y_NYU_Speck.ks6.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_NYU_Speck.ks6.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    Y_NYU_Speck.ks6.reserve = &H0               ' �\��ς�
                                                ' �L�[�P
    Y_NYU_Speck.ks7.keypos = 172                ' �L�[�|�W�V����
    Y_NYU_Speck.ks7.keyleng = 8                 ' �L�[��
                                                ' �L�[�t���O
    Y_NYU_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_NYU_Speck.ks7.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    Y_NYU_Speck.ks7.reserve = &H0               ' �\��ς�
                                                ' �L�[�P
    Y_NYU_Speck.ks8.keypos = 5                ' �L�[�|�W�V����
    Y_NYU_Speck.ks8.keyleng = 9                 ' �L�[��
                                                ' �L�[�t���O
    Y_NYU_Speck.ks8.keyflag = BtKfExt + BtKfChg
    Y_NYU_Speck.ks8.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    Y_NYU_Speck.ks8.reserve = &H0               ' �\��ς�
    '-------------------------------------------
                                                
                                                ' �L�[�Q
    Y_NYU_Speck.ks9.keypos = 3                  ' �L�[�|�W�V����
    Y_NYU_Speck.ks9.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    Y_NYU_Speck.ks9.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_NYU_Speck.ks9.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    Y_NYU_Speck.ks9.reserve = &H0               ' �\��ς�
                                                ' �L�[�Q
    Y_NYU_Speck.ks10.keypos = 172               ' �L�[�|�W�V����
    Y_NYU_Speck.ks10.keyleng = 8                ' �L�[��
                                                ' �L�[�t���O
    Y_NYU_Speck.ks10.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_NYU_Speck.ks10.keytype = Chr(BtKtString)  ' �L�[�^�C�v
    Y_NYU_Speck.ks10.reserve = &H0              ' �\��ς�
                                                ' �L�[�Q
    Y_NYU_Speck.ks11.keypos = 53                ' �L�[�|�W�V����
    Y_NYU_Speck.ks11.keyleng = 20               ' �L�[��
                                                ' �L�[�t���O
    Y_NYU_Speck.ks11.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_NYU_Speck.ks11.keytype = Chr(BtKtString)  ' �L�[�^�C�v
    Y_NYU_Speck.ks11.reserve = &H0              ' �\��ς�
                                                ' �L�[�Q
    Y_NYU_Speck.ks12.keypos = 4                 ' �L�[�|�W�V����
    Y_NYU_Speck.ks12.keyleng = 1                ' �L�[��
                                                ' �L�[�t���O
    Y_NYU_Speck.ks12.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_NYU_Speck.ks12.keytype = Chr(BtKtString)  ' �L�[�^�C�v
    Y_NYU_Speck.ks12.reserve = &H0              ' �\��ς�
                                                ' �L�[�Q
    Y_NYU_Speck.ks13.keypos = 5                 ' �L�[�|�W�V����
    Y_NYU_Speck.ks13.keyleng = 9                ' �L�[��
                                                ' �L�[�t���O
    Y_NYU_Speck.ks13.keyflag = BtKfExt + BtKfChg
    Y_NYU_Speck.ks13.keytype = Chr(BtKtString)  ' �L�[�^�C�v
    Y_NYU_Speck.ks13.reserve = &H0              ' �\��ς�
    '-------------------------------------------
                                                
                                                ' �L�[�R
    Y_NYU_Speck.ks14.keypos = 172               ' �L�[�|�W�V����
    Y_NYU_Speck.ks14.keyleng = 8                ' �L�[��
                                                ' �L�[�t���O
    Y_NYU_Speck.ks14.keyflag = BtKfExt + BtKfChg + BtKfDup
    Y_NYU_Speck.ks14.keytype = Chr(BtKtString)  ' �L�[�^�C�v
    Y_NYU_Speck.ks14.reserve = &H0              ' �\��ς�
    '-------------------------------------------
    
    
    
    
    '-------------------------------------------    2010.07.12
                                                ' �L�[�S
    Y_NYU_Speck.ks15.keypos = 662               ' �L�[�|�W�V����
    Y_NYU_Speck.ks15.keyleng = 1                ' �L�[��
                                                ' �L�[�t���O
    Y_NYU_Speck.ks15.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_NYU_Speck.ks15.keytype = Chr(BtKtString)  ' �L�[�^�C�v
    Y_NYU_Speck.ks15.reserve = &H0              ' �\��ς�
    
    Y_NYU_Speck.ks16.keypos = 3                 ' �L�[�|�W�V����
    Y_NYU_Speck.ks16.keyleng = 1                ' �L�[��
                                                ' �L�[�t���O
    Y_NYU_Speck.ks16.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_NYU_Speck.ks16.keytype = Chr(BtKtString)  ' �L�[�^�C�v
    Y_NYU_Speck.ks16.reserve = &H0              ' �\��ς�
    
    Y_NYU_Speck.ks17.keypos = 4                 ' �L�[�|�W�V����
    Y_NYU_Speck.ks17.keyleng = 1                ' �L�[��
                                                ' �L�[�t���O
    Y_NYU_Speck.ks17.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_NYU_Speck.ks17.keytype = Chr(BtKtString)  ' �L�[�^�C�v
    Y_NYU_Speck.ks17.reserve = &H0              ' �\��ς�
    
    Y_NYU_Speck.ks18.keypos = 53                ' �L�[�|�W�V����
    Y_NYU_Speck.ks18.keyleng = 20               ' �L�[��
                                                ' �L�[�t���O
    Y_NYU_Speck.ks18.keyflag = BtKfExt + BtKfChg + BtKfDup
    Y_NYU_Speck.ks18.keytype = Chr(BtKtString)  ' �L�[�^�C�v
    Y_NYU_Speck.ks18.reserve = &H0              ' �\��ς�
    
    
    
    
    
    
    
    sts = BTRV(BtOpCreate, Y_NYU_POS, Y_NYU_Speck, Len(Y_NYU_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���ח\��f�[�^")
        Y_NYU_Create = True
        Exit Function
    End If

    Y_NYU_Create = False

End Function

Function Y_NYU_Open(Mode As Integer) As Integer
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
    
    Y_NYU_Open = True
                                            '���ח\��f�[�^�t���p�X�捞��
    sts = GetIni("FILE", Y_NYU_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [Y_NYU]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = Y_NYU_Create()        '���ח\��f�[�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), ByVal FullPath, Len(FullPath), Mode)
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
    
    Y_NYU_Open = False

End Function


