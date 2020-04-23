Attribute VB_Name = "Y_SYU"
Option Explicit
'********************************************************************
'*
'*              �o�ח\��f�[�^  �t�@�C����`
'*              �Vڲ��đΉ�     2006.05.24
'*
'********************************************************************
'�t�@�C���h�c
Public Const Y_SYU_ID$ = "Y_SYU"

'�y�[�W�T�C�Y
Public Const Y_SYU_PG_SIZ% = 4096

'�|�W�V�����E�u���b�N
Public Y_SYU_POS    As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type Y_SYUREC_Tag
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
    KENPIN_TANTO_CODE(0 To 4)   As Byte     '���i�S���Һ��� 2006.08.07
    KENPIN_HMS(0 To 5)          As Byte     '���i����       2006.08.07
    LK_MUKE_CODE(0 To 7)        As Byte     '����ݸ�p������2006.08.07
    LK_SEQ_NO(0 To 11)          As Byte     '����ݸ�p�A��  2006.08.07
    G_KENPIN_F(0 To 0)          As Byte     '��ʌ��i�׸�   2006.08.07
    
    
    KENPIN_SURYO(0 To 6)        As Byte     '���i������     2006.08.07
    
    
    
    
    H_IO_KBN(0 To 0)            As Byte     '���o�ɋ敪(�����A�����敪) 2007.12.11
    H_SOKO_CODE(0 To 0)         As Byte     '�q�ɺ���(�݌Ɏ��xð��ق̑q�ɺ���) 2007.12.11
    
    
    UPD_NOW(0 To 13)            As Byte     '�X�V����   2008.11.28
    
    KAN_HMS(0 To 5)             As Byte     '��������   2011.03.30
    
    FILLER(0 To 5)              As Byte      'FILLER     2011.03.30
End Type

'�f�[�^�E�o�b�t�@
Public Y_SYUREC                 As Y_SYUREC_Tag

'�L�[��`
Type KEY0_Y_SYU            '�j�d�x�O
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    KEY_ID_NO(0 To 11)          As Byte     'ID-NO
End Type

Type KEY1_Y_SYU            '�j�d�x�P
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    KAN_KBN(0 To 0)             As Byte     '�����敪
    KEY_MUKE_CODE(0 To 7)       As Byte     '���Ӑ�R�[�h
    KEY_SS_CODE(0 To 7)         As Byte     '������R�[�h
    KEY_CYU_KBN(0 To 0)         As Byte     '�����敪
    KEY_ID_NO(0 To 11)          As Byte     'ID-NO
    NAIGAI(0 To 0)              As Byte     '�����O
    KEY_HIN_NO(0 To 19)         As Byte     '�i�ڔԍ�
End Type

Type KEY2_Y_SYU            '�j�d�x�Q
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    KEY_CYU_KBN(0 To 0)         As Byte     '�����敪
    KEY_MUKE_CODE(0 To 7)       As Byte     '���Ӑ�R�[�h
    KEY_SS_CODE(0 To 7)         As Byte     '������R�[�h
End Type

Type KEY3_Y_SYU            '�j�d�x�R
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    KEY_CYU_KBN(0 To 0)         As Byte     '�����敪
    KEY_MUKE_CODE(0 To 7)       As Byte     '���Ӑ�R�[�h
    KEY_SS_CODE(0 To 7)         As Byte     '������R�[�h
    NAIGAI(0 To 0)              As Byte     '�����O
    KEY_HIN_NO(0 To 19)         As Byte     '�i�ڔԍ�
    KEY_ID_NO(0 To 11)          As Byte     'ID-NO
End Type

Type KEY4_Y_SYU            '�j�d�x�S
    WEL_ID(0 To 2)              As Byte     '�g�p�q�@ID
    PRG_ID(0 To 7)              As Byte     '�g�p���v���O����
End Type

Type KEY5_Y_SYU            '�j�d�x�T
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    KEY_CYU_KBN(0 To 0)         As Byte     '�����敪
    KEY_MUKE_CODE(0 To 7)       As Byte     '���Ӑ�R�[�h
    KEY_SS_CODE(0 To 7)         As Byte     '������R�[�h
    HTANABAN(0 To 7)            As Byte     '�z�X�g�I��         '2004.06.08
    KEY_SYUKA_YMD(0 To 7)       As Byte     '�o�ד��t
    KEY_HIN_NO(0 To 19)         As Byte     '�i�ڔԍ�           '2004.06.29
End Type

Type KEY6_Y_SYU            '�j�d�x�U
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    KEY_CYU_KBN(0 To 0)         As Byte     '�����敪
    HTANABAN(0 To 7)            As Byte     '�z�X�g�I��
    NAIGAI(0 To 0)              As Byte     '�����O
    KEY_HIN_NO(0 To 19)         As Byte     '�i�ڔԍ�
End Type

Type KEY7_Y_SYU            '�j�d�x�V
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    NAIGAI(0 To 0)              As Byte     '�����O
    KEY_HIN_NO(0 To 19)         As Byte     '�i�ڔԍ�
    KEY_SYUKA_YMD(0 To 7)       As Byte     '�o�ד��t
End Type

Type KEY8_Y_SYU            '�j�d�x�W        2011.08.02
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    NAIGAI(0 To 0)              As Byte     '�����O
    KEY_HIN_NO(0 To 19)         As Byte     '�i�ڔԍ�
    KAN_YMD(0 To 7)             As Byte     '�������t
    KAN_HMS(0 To 13)            As Byte     '��������
End Type

Type KEY9_Y_SYU            '�j�d�x�X        2011.08.02
    KEY_SYUKA_YMD(0 To 7)       As Byte     '�o�ד��t
End Type

Type KEY10_Y_SYU            '�j�d�x10       2012.06.15
    KAN_KBN(0 To 0)             As Byte     '�����敪
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    KEY_SYUKA_YMD(0 To 7)       As Byte     '�o�ד��t
    PRINT_YMD(0 To 7)           As Byte     '�o�ɕ\������t
    DEN_NO(0 To 9)              As Byte     '�`�[�ԍ�
End Type



'�L�[�E�f�[�^
Public K0_Y_SYU                 As KEY0_Y_SYU
Public K1_Y_SYU                 As KEY1_Y_SYU
Public K2_Y_SYU                 As KEY2_Y_SYU
Public K3_Y_SYU                 As KEY3_Y_SYU
Public K4_Y_SYU                 As KEY4_Y_SYU
Public K5_Y_SYU                 As KEY5_Y_SYU
Public K6_Y_SYU                 As KEY6_Y_SYU
Public K7_Y_SYU                 As KEY7_Y_SYU

Public K8_Y_SYU                 As KEY8_Y_SYU       '2011.08.02
Public K9_Y_SYU                 As KEY9_Y_SYU       '2011.08.02

Public K10_Y_SYU                As KEY10_Y_SYU      '2012.06.15


Type Y_SYU_FSpeck
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
    ks14    As BtKeySpeck                   ' �� ��߯��\����
    ks15    As BtKeySpeck                   ' �� ��߯��\����
    ks16    As BtKeySpeck                   ' �� ��߯��\����
    ks17    As BtKeySpeck                   ' �� ��߯��\����
    ks18    As BtKeySpeck                   ' �� ��߯��\����
    ks19    As BtKeySpeck                   ' �� ��߯��\����
    ks20    As BtKeySpeck                   ' �� ��߯��\����
    ks21    As BtKeySpeck                   ' �� ��߯��\����
    ks22    As BtKeySpeck                   ' �� ��߯��\����
    ks23    As BtKeySpeck                   ' �� ��߯��\����
    ks24    As BtKeySpeck                   ' �� ��߯��\����
    ks25    As BtKeySpeck                   ' �� ��߯��\����
    ks26    As BtKeySpeck                   ' �� ��߯��\����
    ks27    As BtKeySpeck                   ' �� ��߯��\����
    ks28    As BtKeySpeck                   ' �� ��߯��\����
    ks29    As BtKeySpeck                   ' �� ��߯��\����
    ks30    As BtKeySpeck                   ' �� ��߯��\����
    ks31    As BtKeySpeck                   ' �� ��߯��\����
    ks32    As BtKeySpeck                   ' �� ��߯��\����
    ks33    As BtKeySpeck                   ' �� ��߯��\����
    ks34    As BtKeySpeck                   ' �� ��߯��\����
    ks35    As BtKeySpeck                   ' �� ��߯��\����
    ks36    As BtKeySpeck                   ' �� ��߯��\����
    ks37    As BtKeySpeck                   ' �� ��߯��\����
    ks38    As BtKeySpeck                   ' �� ��߯��\����

    ks39    As BtKeySpeck                   ' �� ��߯��\����    2011.08.02
    ks40    As BtKeySpeck                   ' �� ��߯��\����    2011.08.02
    ks41    As BtKeySpeck                   ' �� ��߯��\����    2011.08.02
    ks42    As BtKeySpeck                   ' �� ��߯��\����    2011.08.02
    ks43    As BtKeySpeck                   ' �� ��߯��\����    2011.08.02
    ks44    As BtKeySpeck                   ' �� ��߯��\����    2011.08.02

    ks45    As BtKeySpeck                   ' �� ��߯��\����    2012.06.15
    ks46    As BtKeySpeck                   ' �� ��߯��\����    2012.06.15
    ks47    As BtKeySpeck                   ' �� ��߯��\����    2012.06.15
    ks48    As BtKeySpeck                   ' �� ��߯��\����    2012.06.15
    ks49    As BtKeySpeck                   ' �� ��߯��\����    2012.06.15

End Type

Private Y_SYU_Speck As Y_SYU_FSpeck

Private Function Y_SYU_Create() As Integer

'********************************************************************
'*
'*              �o�ח\��f�[�^  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    Y_SYU_Create = True
                                            '�o�ח\��f�[�^�t���p�X�捞��
    sts = GetIni("FILE", Y_SYU_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [Y_SYU]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    Y_SYU_Speck.fs.recoleng = Len(Y_SYUREC)         ' ���R�[�h��
    Y_SYU_Speck.fs.PageSize = Y_SYU_PG_SIZ          ' �y�[�W�T�C�Y
    Y_SYU_Speck.fs.idexnumb = 11                    ' �C���f�b�N�X��    8-->10 2011.08.02   10-->11 2012.06.15
    Y_SYU_Speck.fs.fileflag = 0                     ' �t�@�C���t���O
    Y_SYU_Speck.fs.reserve = &H0                    ' �\��ς�
'---------------------------------------------------' �L�[�O
    Y_SYU_Speck.ks0.keypos = 14                     ' �L�[�|�W�V����
    Y_SYU_Speck.ks0.keyleng = 1                     ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks0.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks0.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_Speck.ks0.reserve = &H0                   ' �\��ς�
    
    Y_SYU_Speck.ks1.keypos = 16                     ' �L�[�|�W�V����
    Y_SYU_Speck.ks1.keyleng = 12                    ' �L�[��
    Y_SYU_Speck.ks1.keyflag = BtKfExt + BtKfChg     ' �L�[�t���O
    Y_SYU_Speck.ks1.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_Speck.ks1.reserve = &H0                   ' �\��ς�

'---------------------------------------------------' �L�[�P
    Y_SYU_Speck.ks2.keypos = 14                     ' �L�[�|�W�V����
    Y_SYU_Speck.ks2.keyleng = 1                     ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks2.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks2.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_Speck.ks2.reserve = &H0                   ' �\��ς�
    
    Y_SYU_Speck.ks3.keypos = 12                     ' �L�[�|�W�V����
    Y_SYU_Speck.ks3.keyleng = 1                     ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks3.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks3.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_Speck.ks3.reserve = &H0                   ' �\��ς�
    
    Y_SYU_Speck.ks4.keypos = 49                     ' �L�[�|�W�V����
    Y_SYU_Speck.ks4.keyleng = 8                     ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks4.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks4.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_Speck.ks4.reserve = &H0                   ' �\��ς�
    
    Y_SYU_Speck.ks5.keypos = 57                     ' �L�[�|�W�V����
    Y_SYU_Speck.ks5.keyleng = 8                     ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks5.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks5.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_Speck.ks5.reserve = &H0                   ' �\��ς�
    
    Y_SYU_Speck.ks6.keypos = 15                     ' �L�[�|�W�V����
    Y_SYU_Speck.ks6.keyleng = 1                     ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks6.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks6.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_Speck.ks6.reserve = &H0                   ' �\��ς�
    
    Y_SYU_Speck.ks7.keypos = 16                     ' �L�[�|�W�V����
    Y_SYU_Speck.ks7.keyleng = 12                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks7.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks7.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_Speck.ks7.reserve = &H0                   ' �\��ς�
    
    Y_SYU_Speck.ks8.keypos = 28                     ' �L�[�|�W�V����
    Y_SYU_Speck.ks8.keyleng = 1                     ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks8.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks8.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_Speck.ks8.reserve = &H0                   ' �\��ς�
    
    Y_SYU_Speck.ks9.keypos = 29                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks9.keyleng = 20                   ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks9.keyflag = BtKfExt + BtKfChg
    Y_SYU_Speck.ks9.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks9.reserve = &H0                  ' �\��ς�
'---------------------------------------------------' �L�[�Q
    Y_SYU_Speck.ks10.keypos = 14                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks10.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks10.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_SYU_Speck.ks10.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks10.reserve = &H0                  ' �\��ς�
    
    Y_SYU_Speck.ks11.keypos = 15                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks11.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks11.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_SYU_Speck.ks11.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks11.reserve = &H0                  ' �\��ς�
    
    Y_SYU_Speck.ks12.keypos = 49                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks12.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks12.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_SYU_Speck.ks12.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks12.reserve = &H0                  ' �\��ς�
    
    Y_SYU_Speck.ks13.keypos = 57                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks13.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks13.keyflag = BtKfExt + BtKfChg + BtKfDup
    Y_SYU_Speck.ks13.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks13.reserve = &H0                  ' �\��ς�
'---------------------------------------------------' �L�[�R
    Y_SYU_Speck.ks14.keypos = 14                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks14.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks14.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks14.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks14.reserve = &H0                  ' �\��ς�
    
    Y_SYU_Speck.ks15.keypos = 15                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks15.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks15.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks15.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks15.reserve = &H0                  ' �\��ς�
    
    Y_SYU_Speck.ks16.keypos = 49                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks16.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks16.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks16.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks16.reserve = &H0                  ' �\��ς�
                                                    
    Y_SYU_Speck.ks17.keypos = 57                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks17.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks17.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks17.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks17.reserve = &H0                  ' �\��ς�
    
    Y_SYU_Speck.ks18.keypos = 28                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks18.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks18.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks18.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks18.reserve = &H0                  ' �\��ς�
    
    Y_SYU_Speck.ks19.keypos = 29                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks19.keyleng = 20                   ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks19.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks19.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks19.reserve = &H0                  ' �\��ς�
    
    Y_SYU_Speck.ks20.keypos = 16                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks20.keyleng = 12                   ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks20.keyflag = BtKfExt + BtKfChg
    Y_SYU_Speck.ks20.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks20.reserve = &H0                  ' �\��ς�
'---------------------------------------------------' �L�[�S
    Y_SYU_Speck.ks21.keypos = 1                     ' �L�[�|�W�V����
    Y_SYU_Speck.ks21.keyleng = 3                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks21.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_SYU_Speck.ks21.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks21.reserve = &H0                  ' �\��ς�
    
    Y_SYU_Speck.ks22.keypos = 4                     ' �L�[�|�W�V����
    Y_SYU_Speck.ks22.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks22.keyflag = BtKfExt + BtKfChg + BtKfDup
    Y_SYU_Speck.ks22.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks22.reserve = &H0                  ' �\��ς�
'---------------------------------------------------' �L�[�T
    Y_SYU_Speck.ks23.keypos = 14                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks23.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks23.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_SYU_Speck.ks23.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks23.reserve = &H0                  ' �\��ς�
    
    Y_SYU_Speck.ks24.keypos = 15                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks24.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks24.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_SYU_Speck.ks24.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks24.reserve = &H0                  ' �\��ς�
    
    Y_SYU_Speck.ks25.keypos = 49                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks25.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks25.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_SYU_Speck.ks25.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks25.reserve = &H0                  ' �\��ς�
    
    Y_SYU_Speck.ks26.keypos = 57                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks26.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks26.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_SYU_Speck.ks26.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks26.reserve = &H0                  ' �\��ς�
    
    
    Y_SYU_Speck.ks27.keypos = 648                   ' �L�[�|�W�V����
    Y_SYU_Speck.ks27.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks27.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_SYU_Speck.ks27.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks27.reserve = &H0                  ' �\��ς�
    
    Y_SYU_Speck.ks28.keypos = 65                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks28.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks28.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_SYU_Speck.ks28.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks28.reserve = &H0                  ' �\��ς�
    
    Y_SYU_Speck.ks29.keypos = 29                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks29.keyleng = 20                   ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks29.keyflag = BtKfExt + BtKfChg + BtKfDup
    Y_SYU_Speck.ks29.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks29.reserve = &H0                  ' �\��ς�
'---------------------------------------------------' �L�[�U
    Y_SYU_Speck.ks30.keypos = 14                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks30.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks30.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_SYU_Speck.ks30.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks30.reserve = &H0                  ' �\��ς�
    
    Y_SYU_Speck.ks31.keypos = 15                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks31.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks31.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_SYU_Speck.ks31.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks31.reserve = &H0                  ' �\��ς�
    
    Y_SYU_Speck.ks32.keypos = 648                   ' �L�[�|�W�V����
    Y_SYU_Speck.ks32.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks32.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_SYU_Speck.ks32.keytype = Chr(BtKtString)      ' �L�[�^�C�v3
    Y_SYU_Speck.ks32.reserve = &H0                  ' �\��ς�
    
    Y_SYU_Speck.ks33.keypos = 28                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks33.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks33.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_SYU_Speck.ks33.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks33.reserve = &H0                  ' �\��ς�
    
    Y_SYU_Speck.ks34.keypos = 29                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks34.keyleng = 20                   ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks34.keyflag = BtKfExt + BtKfChg + BtKfDup
    Y_SYU_Speck.ks34.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks34.reserve = &H0                  ' �\��ς�
'---------------------------------------------------' �L�[�V
    Y_SYU_Speck.ks35.keypos = 14                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks35.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks35.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks35.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks35.reserve = &H0                  ' �\��ς�
    
    Y_SYU_Speck.ks36.keypos = 28                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks36.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks36.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks36.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks36.reserve = &H0                  ' �\��ς�
    
    Y_SYU_Speck.ks37.keypos = 29                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks37.keyleng = 20                   ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks37.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks37.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks37.reserve = &H0                  ' �\��ς�
    
    Y_SYU_Speck.ks38.keypos = 65                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks38.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks38.keyflag = BtKfExt + BtKfDup + BtKfChg
    Y_SYU_Speck.ks38.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks38.reserve = &H0                  ' �\��ς�
    
    
'---------------------------------------------------' �L�[�W    2011.08.02
    Y_SYU_Speck.ks39.keypos = 14                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks39.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks39.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks39.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks39.reserve = &H0                  ' �\��ς�
    
    Y_SYU_Speck.ks40.keypos = 28                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks40.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks40.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks40.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks40.reserve = &H0                  ' �\��ς�
    
    Y_SYU_Speck.ks41.keypos = 29                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks41.keyleng = 20                   ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks41.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks41.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks41.reserve = &H0                  ' �\��ς�
    
    Y_SYU_Speck.ks42.keypos = 664                   ' �L�[�|�W�V����
    Y_SYU_Speck.ks42.keyleng = 8                     ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks42.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks42.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks42.reserve = &H0                  ' �\��ς�
    
    

    Y_SYU_Speck.ks43.keypos = 757                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks43.keyleng = 6                     ' �L�[��
                                                            ' �L�[�t���O
    Y_SYU_Speck.ks43.keyflag = BtKfExt + BtKfDup + BtKfChg
    Y_SYU_Speck.ks43.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks43.reserve = &H0                  ' �\��ς�
    
    
'---------------------------------------------------' �L�[�X    2011.08.02
    Y_SYU_Speck.ks44.keypos = 65                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks44.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks44.keyflag = BtKfExt + BtKfDup + BtKfChg
    Y_SYU_Speck.ks44.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks44.reserve = &H0                  ' �\��ς�
    
    
'---------------------------------------------------' �L�[10    2012.06.15
    
    
    Y_SYU_Speck.ks45.keypos = 12                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks45.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks45.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks45.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks45.reserve = &H0                  ' �\��ς�
    
    
    
    Y_SYU_Speck.ks46.keypos = 14                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks46.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks46.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks46.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks46.reserve = &H0                  ' �\��ς�
    
    
    Y_SYU_Speck.ks47.keypos = 65                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks47.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks47.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks47.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks47.reserve = &H0                  ' �\��ς�
    
    
    Y_SYU_Speck.ks48.keypos = 656                   ' �L�[�|�W�V����
    Y_SYU_Speck.ks48.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks48.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks48.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks48.reserve = &H0                  ' �\��ς�
    
    
    Y_SYU_Speck.ks49.keypos = 132                   ' �L�[�|�W�V����
    Y_SYU_Speck.ks49.keyleng = 10                   ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks49.keyflag = BtKfExt + BtKfDup + BtKfChg
    Y_SYU_Speck.ks49.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks49.reserve = &H0                  ' �\��ς�
    
    
    
    sts = BTRV(BtOpCreate, Y_SYU_POS, Y_SYU_Speck, Len(Y_SYU_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�o�ח\��f�[�^")
        Exit Function
    End If

    Y_SYU_Create = False

End Function

Function Y_SYU_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �o�ח\��f�[�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    Y_SYU_Open = True
                                            '�o�ח\��f�[�^�t���p�X�捞��
    sts = GetIni("FILE", Y_SYU_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [Y_SYU]�ǂݍ��݃G���[ ")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = Y_SYU_Create()        '�o�ח\��f�[�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�o�ח\��f�[�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�o�ח\��f�[�^")
                Exit Function
        End Select
    Loop
    Y_SYU_Open = False
End Function
