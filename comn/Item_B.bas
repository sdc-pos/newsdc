Attribute VB_Name = "ITEM_B"
Option Explicit
'********************************************************************
'*
'*              �i�ڃ}�X�^  �t�@�C����`
'*
'*          CREATE 2004.02.19
'********************************************************************
'�t�@�C���h�c
Public Const ITEM_B_ID$ = "ITEM_B"

'�y�[�W�T�C�Y
Public Const ITEM_B_PG_SIZ% = 4096

'�|�W�V�����E�u���b�N
Public ITEM_B_POS         As POSBLK
'=
'====================================================================
'=          ���R�[�h�������v���V�[�W��     ( Rclr_ITEM_BREC )
'====================================================================
'=
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************


Private Type SHIIRE_TBL_Tag         '�d������`�p��ð���
    CODE(0 To 4)                As Byte     '����
    TANKA(0 To 10)              As Byte     '�P�� 9(8)V99
    TANKA_DT(0 To 7)            As Byte     '�P���ݒ��
    LOT(0 To 7)                 As Byte     'ۯĐ�
    LEAD_TIME(0 To 2)           As Byte     'ذ�����
    LAST_ORDER_DT(0 To 7)       As Byte     '�O�񒍕���
    LAST_ORDER_QTY(0 To 10)     As Byte     '�O�񒍕���
End Type



Private Type BEF_KOUTEI_tag
    BEF_KOUTEI(0 To 5)          As Byte     '�O�H�� 2008.09.19
End Type


Private Type MAIN_KOUTEI_tag
    MAIN_KOUTEI(0 To 5)         As Byte     '��ƍH�� 2008.09.19
End Type

Private Type AFT_KOUTEI_tag
    AFT_KOUTEI(0 To 5)          As Byte     '��H�� 2008.09.19
End Type




'���R�[�h��`
Type ITEM_BREC_Tag
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    NAIGAI(0 To 0)              As Byte     '�����O
    '2005.11.15 �����ύX 13---> 20
    HIN_GAI(0 To 19)            As Byte     '�i�ԁi�O���j
    '2005.11.15 �����ύX 25---> 40
    HIN_NAME(0 To 39)           As Byte     '�i��
    ST_SET_DT(0 To 7)           As Byte     '�W���q�ɐݒ���t
    ST_SOKO(0 To 1)             As Byte     '�W�����ɑq�� �q��
    ST_RETU(0 To 1)             As Byte     '             ��
    ST_REN(0 To 1)              As Byte     '             �A
    ST_DAN(0 To 1)              As Byte     '             �i
    BEF_SOKO(0 To 1)            As Byte     '�O����ɑq�� �q��
    BEF_RETU(0 To 1)            As Byte     '             ��
    BEF_REN(0 To 1)             As Byte     '             �A
    BEF_DAN(0 To 1)             As Byte     '             �i
    LAST_NYU_DT(0 To 7)         As Byte     '�ŏI���ɓ��t
    LAST_SYU_DT(0 To 7)         As Byte     '�ŏI�o�ɓ��t
    '2005.11.15 �����ύX 13---> 20
    HIN_NAI(0 To 19)            As Byte     '�i�ԁi�����j
    BIKOU_SOKO(0 To 1)          As Byte     '���l �z�X�g�q��
    BIKOU_TANA(0 To 7)          As Byte     '���l �z�X�g�I��
    '���g�p�̂��ߍ폜 2005.11.15 SIZAI_CD(0 To 4)    As Byte     '���ރR�[�h
    HOJYU_P(0 To 7)             As Byte     '��[�_�i�댯�݌Ɂj
    AVE_SYUKA(0 To 7)           As Byte     '�����Ϗo�א�
    SAMPLE_QTY(0 To 0)          As Byte     '�T���v����
    LAST_INP_DT(0 To 7)         As Byte     '�ŏI���ד��t
'*------------------------------------------ 2001.02.15 �ǉ� ��
    '���g�p�̂��ߍ폜 2005.11.15 LOCK_F(0 To 0)      As Byte     '�r���t���O
    '���g�p�̂��ߍ폜 2005.11.15 WEL_ID(0 To 2)      As Byte     '�g�p�q�@ID
    '���g�p�̂��ߍ폜 2005.11.15 PRG_ID(0 To 7)      As Byte     '�g�p���v���O����
'*------------------------------------------ 2001.02.15 �ǉ� ��
    LAST_CHK_DT(0 To 7)         As Byte     '�ŏI�ƍ����t2001.06.12
    LAST_CHK_QTY(0 To 7)        As Byte     '�ŏI�ƍ����݌ɐ�2001.06.12
    '���g�p�̂��ߍ폜 2005.11.15 MOTO_JIGYOBU(0 To 0) As Byte    '�������ƕ�     '���g�p2004.02
    BIKOU(0 To 14)              As Byte     '������l
    IRI_QTY(0 To 7)             As Byte     '������萔

    '2005.11.15 �����ύX 13---> 20
    JAN_CODE(0 To 19)           As Byte     'Jan�R�[�h      2004.02
    '2005.11.15 �����ύX 13---> 20
    HIN_CHANGE(0 To 19)         As Byte     '�i�ԓǂݑւ�   2004.02
    GOODS_KBN(0 To 0)           As Byte     '���i���L��     2004.02
    PACKING_NO(0 To 3)          As Byte     '������       2004.02
    RANK(0 To 2)                As Byte     '���݃����N     2004.06
    NEW_RANK(0 To 2)            As Byte     '���݃����N     2004.06
    GLICS1_TANA(0 To 9)         As Byte     '�O���b�N�X�I�ԂP   2005.05
    GLICS2_TANA(0 To 9)         As Byte     '�O���b�N�X�I�ԂQ   2005.05
    GLICS3_TANA(0 To 9)         As Byte     '�O���b�N�X�I�ԂR   2005.05
'*------------------------------------------ 2005.11.15 �ǉ�(�Ɩ��Ǘ�����) ��
    G_SHIIRE_KBN(0 To 1)        As Byte     '�Ɩ��Ǘ��@ �d���敪
    G_HANBAI_KBN(0 To 1)        As Byte     '           �̔��敪
    G_SYUSHI(0 To 2)            As Byte     '           ���x�P��
    G_KUMITATE(0 To 0)          As Byte     '           �g�����i
    G_ST_URITAN(0 To 10)        As Byte     '           �W���e�������P���@9(8)V99
    G_ST_URITAN_DT(0 To 7)      As Byte     '           �W���e�������ݒ��
    G_ST_SHITAN(0 To 10)        As Byte     '           �W���e�������P��  9(8)V99
    G_ST_SHITAN_DT(0 To 7)      As Byte     '           �W���e�������ݒ��
                                            '           �d������
    G_SHIIRE_TBL(0 To 2)        As SHIIRE_TBL_Tag
    G_ZEN_ZAIKO_KIN(0 To 10)    As Byte     '           �O���݌ɋ��z
    G_SHIZAI_KBN(0 To 0)        As Byte     '           ���ދ敪
    G_LABEL_NON(0 To 0)         As Byte     '           ���ٓ\��t���v��Ȃ�
'*------------------------------------------ 2005.11.15 �ǉ�(�Ɩ��Ǘ�����) ��

'*------------------------------------------ 2005.11.15 �ǉ�(���i���ٍ���) ��
    L_HIN_NAME_E(0 To 29)       As Byte     '���i����   �i��
    L_BIKOU(0 To 19)            As Byte     '           ���l
    L_KAISHA_CODE(0 To 1)       As Byte     '           ��ЃR�[�h
    L_KISHU1(0 To 24)           As Byte     '           �@��(1)
    xL_KISHU2(0 To 39)          As Byte     '           �@��(2) ���g�p 2006.01.24
    L_KISHU3(0 To 149)          As Byte     '           �@��(3)(���K�p�@����l)
    L_PAPER(0 To 0)             As Byte     '           ��
    L_PLASTIC(0 To 0)           As Byte     '           �v���X�`�b�N
    L_URIKIN1(0 To 9)           As Byte     '           ���i(1)
    L_URIKIN2(0 To 9)           As Byte     '           ���i(2)
    L_URIKIN3(0 To 9)           As Byte     '           ���i(3)
    L_LABEL(0 To 0)             As Byte     '           �K�p�@������
    L_MAISU(0 To 0)             As Byte     '           ��������
    L_KISHU_BIKOU(0 To 449)     As Byte     '           �K�p�@����l(���@��i�R�j)
    L_SAGYO_SHIJI(0 To 449)     As Byte     '           ��Ǝw��
    L_BIKOU3(0 To 4)            As Byte     '           ���l�R
    L_JGYOBU_CODE(0 To 1)       As Byte     '           ���ƕ��R�[�h
    L_IRI_QTY(0 To 7)           As Byte     '           ���萔
    L_TANA1(0 To 19)            As Byte     '           �I��(1)
    L_TANA2(0 To 19)            As Byte     '           �I��(2)
'*------------------------------------------ 2005.11.15 �ǉ�(���i���ٍ���) ��
    S_TANTO(0 To 1)             As Byte     '���P�^�S���҃R�[�h
    ZAIKO_F(0 To 0)             As Byte     '�݌ɊǗ��ΏۗL�� 1:�Ώ� 0:�ΏۊO

    L_KISHU2(0 To 51)           As Byte     '           �@��(2)

    G_ZEN_ZAIKO_QTY(0 To 7)     As Byte     '           �O���݌ɐ���
    G_LAST_SYUKA_QTY(0 To 7)    As Byte     '           �ŏI�o�א�

    G_S2_ZAI_QTY(0 To 7)        As Byte     'GLICS�݌�(S2) �܈�p
    G_P2_ZAI_QTY(0 To 7)        As Byte     'GLICS�݌�(P2) �܈�p


    K_KEITAI(0 To 9)            As Byte     '���`��


    UNIT_BUHIN(0 To 0)          As Byte     '�Ưĕ��i�敪       2006.07.28
    NAI_BUHIN(0 To 0)           As Byte     '�����������i�敪   2006.07.28
    GAI_BUHIN(0 To 0)           As Byte     '�C�O�������i�敪   2006.07.28
    HYO_TANKA(0 To 9)           As Byte     '�W���P��   2006.07.28

    LAST_CODE(0 To 4)           As Byte     '�ŏI�d����R�[�h   2007.05.29
    LAST_TANKA(0 To 10)         As Byte     '�ŏI�d���P��       2007.05.29

    MAKER_CODE(0 To 7)          As Byte     'Ұ������           2007.06.06
    MAKER_NAME(0 To 39)         As Byte     'Ұ������           2007.06.06


    L_MARK(0 To 0)              As Byte     '�č���ϰ�          2007.11.08

    xSAI_SU(0 To 3)              As Byte     '�ː�               2008.02.14

    D_KEISHIKI(0 To 19)         As Byte     '�`��               2008.02.14
    D_MATERIAL(0 To 19)         As Byte     '�ގ�               2008.02.14
    D_THICKNESS(0 To 9)         As Byte     '����ްف@����      2008.02.14


    D_SIZE_W(0 To 7)            As Byte     '����ްٻ��ށiW�j   2008.02.14
    D_SIZE_D(0 To 7)            As Byte     '����ްٻ��ށiD�j   2008.02.14
    D_SIZE_H(0 To 7)            As Byte     '����ްٻ��ށiH�j   2008.02.14

    D_PRINT(0 To 3)             As Byte     '�������^���Ȃ�    2008.02.14

    S_KOUSU(0 To 7)             As Byte     '���i���@�H��       2008.02.14

    S_KOUSU_GENKA(0 To 10)      As Byte     '���i���@�H������   2008.02.14
    S_KOUSU_BAIKA(0 To 10)      As Byte     '���i���@�H������   2008.02.14
    S_KOUSU_SET_DATE(0 To 7)    As Byte     '���i���@�P���ݒ�� 2008.02.14

    S_SHIZAI_GENKA(0 To 10)     As Byte     '���i���@���ތ���   2008.02.14
    S_SHIZAI_BAIKA(0 To 10)     As Byte     '���i���@���ޔ���   2008.02.14
    S_SHIZAI_SET_DATE(0 To 7)   As Byte     '���i���@�P���ݒ�� 2008.02.14

    SE_USOU_F(0 To 1)           As Byte     '�A�����@�o���׸�   2008.02.14

    USE_TAPE_KIND(0 To 19)      As Byte     '�g�p�e�[�v���     2008.02.14
    USE_TAPE_LNG(0 To 7)        As Byte     '�g�p�e�[�v��       2008.02.14

    H_TANA_MAKE(0 To 0)         As Byte     '�I�ԃ}�[�N         2008.04.02

    SE_TANKA_MEMO(0 To 39)      As Byte     '�����P���@����     2008.04.15

    xGENSANKOKU(0 To 9)         As Byte     '���Y��             2008.06.11-->2009.07.16 ���g�p

    S_GAISO_TANKA(0 To 10)      As Byte     '�O���P�� 9(8)V99   2008.06.12
    S_PPSC_KAKO_KOSU(0 To 7)    As Byte     'PPSC���H�P��9(8)   2008.06.12
    S_BU_KAKO_KOSU(0 To 7)      As Byte     'BU���H�P��9(8)     2008.06.12

    SEI_LOT(0 To 7)             As Byte     '���Y���b�g         2008.07.07
    SEI_RATE(0 To 6)            As Byte     '�����[�g           2008.07.07
    SEI_SYU_KON(0 To 5)         As Byte     '�W������           2008.07.07

    SEI_TANKA_TANTO(0 To 4)     As Byte     '�P���ݒ�S����     2008.07.09

    SHIMUKE_CODE(0 To 1)        As Byte     '�d������           2008.07.09

    SEI_KBN(0 To 0)             As Byte     '�����敪           2008.07.16

    SEI_LABEL_QTY(0 To 1)       As Byte     '���x���\�薇��     2008.07.19

    SEI_SZI_CNT(0 To 1)         As Byte     '���ތ���     �@    2008.08.20�ǉ�
    SEI_DKN_CNT(0 To 1)         As Byte     '��������           2008.08.20�ǉ�

                                            '�O�H��             2008.09.19
    BEF_KOUTEI(0 To 9)          As BEF_KOUTEI_tag
                                            '��ƍH��           2008.09.19
    MAIN_KOUTEI(0 To 9)         As MAIN_KOUTEI_tag
                                            '��H��             2008.09.19
    AFT_KOUTEI(0 To 9)          As AFT_KOUTEI_tag

    SE_IO_TANKA_No(0 To 1)      As Byte     '�I�敪             200.09.19

    STAT(0 To 0)                As Byte     '��ԋ敪           2009.01.21

    INSP_MESSAGE(0 To 39)       As Byte     '�o�׌��iү����     2009.04.17

    S_SEIKYU_F(0 To 0)          As Byte     '���i�������׸�     2009.04.28

    
    
'---------
    
    BEF_S_KOUSU_BAIKA(0 To 10)  As Byte     '���i���@�H������   2009.06.02
    BEF_S_SHIZAI_BAIKA(0 To 10) As Byte     '���i���@���ޔ���   2009.06.02
    BEF_S_GAISO_TANKA(0 To 10)  As Byte     '�O���P�� 9(8)V99   2009.06.02
    BEF_S_PPSC_KAKO_KOSU(0 To 7) As Byte    'PPSC���H�P��9(8)   2009.06.02
    BEF_S_BU_KAKO_KOSU(0 To 7)  As Byte     'BU���H�P��9(8)     2009.06.02
    
    M_BIKOU(0 To 119)           As Byte     '���Ϗ����l         2009.06.02
    SHIYOU_NO(0 To 9)           As Byte     '�d�l����           2009.06.02
    MITSUMORI_KBN(0 To 0)       As Byte     '���ς�敪         2009.06.02
    TANKA_KIRIKAE_DT(0 To 7)    As Byte     '�P���ؑ֓��t       2009.06.02
    KIRIKAE_KBN(0 To 0)         As Byte     '�ؑ֋敪           2009.06.02
    
    
'---------
    
    GENSANKOKU(0 To 19)         As Byte     '���Y��             '2009.07.16
    
    
    
    PLUS_KOUSU(0 To 5)          As Byte     '�v���X���H��       2009.09.17
    
    
    
    KUTI_SU(0 To 3)             As Byte     '����               2010.01.18
    KONPOU_F(0 To 0)            As Byte     '����敪           2010.01.18
    
    SAI_SU(0 To 4)              As Byte     '�ː�               2010.01.18
    
    
    
    TORI_GENSANKOKU(0 To 19)    As Byte     '�捞�ݎ����Y��     2010.07.20
    TORI_GEN_GENSANKOKU(0 To 19) _
                                As Byte     '�捞�ݎ����Y���\�� 2010.07.20
    TORI_SHIIRE_WORK_CENTER(0 To 7) _
                                As Byte     '�d��ܰ��Z���^�[    2010.07.20
        
    
    
    KANKYO_KBN(0 To 2)          As Byte     '����ދ敪       2010.07.27
    KANKYO_KBN_ST(0 To 7)       As Byte     '����ދ敪�K�p�J�n 2010.07.27
    KANKYO_KBN_SURYO(0 To 9)    As Byte     '����ދ敪����   2010.07.27
    
    BEF_L_LABEL(0 To 0)         As Byte     '''''
    
    BEF_1_L_PAPER(0 To 0)       As Byte     '           ��
    BEF_1_L_PLASTIC(0 To 0)     As Byte     '           �v���X�`�b�N
    BEF_2_L_PAPER(0 To 0)       As Byte     '           ��
    BEF_2_L_PLASTIC(0 To 0)     As Byte     '           �v���X�`�b�N
    BEF_3_L_PAPER(0 To 0)       As Byte     '           ��
    BEF_3_L_PLASTIC(0 To 0)     As Byte     '           �v���X�`�b�N
    BEF_4_L_PAPER(0 To 0)       As Byte     '           ��
    BEF_4_L_PLASTIC(0 To 0)     As Byte     '           �v���X�`�b�N
    BEF_LAST_L_PAPER(0 To 0)    As Byte     '           ��
    BEF_LAST_L_PLASTIC(0 To 0)  As Byte     '           �v���X�`�b�N
    
    
    FILLER(0 To 282)            As Byte     'FILLER             2010.07.27    �T�C�Y�ύX

    INS_TANTO(0 To 4)           As Byte     '�ǉ��@�S���ҁ@     2009.01.21
    Ins_DateTime(0 To 13)       As Byte     '�ǉ��@����         2009.01.21

    UPD_TANTO(0 To 4)           As Byte     '�X�V�@�S���ҁ@     2005.11.15
    UPD_DATETIME(0 To 13)       As Byte     '�X�V�@����         2005.11.15

End Type
'�f�[�^�E�o�b�t�@
Public ITEM_BREC As ITEM_BREC_Tag

'�L�[��`

Type KEY0_ITEM_B            '�j�d�x�O
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    '2005.11.15 �����ύX 13---> 20
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
End Type

Type KEY1_ITEM_B            '�j�d�x�P
    LAST_SYU_DT(0 To 7) As Byte     '�ŏI�o�ɓ��t
End Type

Type KEY2_ITEM_B            '�j�d�x�Q
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    '2005.11.15 �����ύX 13---> 20
    HIN_NAI(0 To 19)    As Byte     '�i�ԁi�����j
End Type

Type KEY3_ITEM_B            '�j�d�x�R
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    ST_SET_DT(0 To 7)   As Byte     '�W���q�ɐݒ���t
End Type


Type KEY4_ITEM_B            '�j�d�x�S 2004.02
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    '2005.11.15 �����ύX 13---> 20
    JAN_CODE(0 To 19)   As Byte     'Jan�R�[�h
End Type

Type KEY5_ITEM_B            '�j�d�x�T 2004.02
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    '2005.11.15 �����ύX 13---> 20
    HIN_CHANGE(0 To 19) As Byte     '�i�ԓǂݑւ�
End Type

Type KEY6_ITEM_B            '�j�d�x�U 2004.02
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    ST_SOKO(0 To 1)     As Byte     '�W�����ɑq�� �q��
    ST_RETU(0 To 1)     As Byte     '             ��
    ST_REN(0 To 1)      As Byte     '             �A
    ST_DAN(0 To 1)      As Byte     '             �i
    '2005.11.15 �����ύX 13---> 20
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
End Type



'�L�[�E�f�[�^
Public K0_ITEM_B As KEY0_ITEM_B
Public K1_ITEM_B As KEY1_ITEM_B
Public K2_ITEM_B As KEY2_ITEM_B
Public K3_ITEM_B As KEY3_ITEM_B
Public K4_ITEM_B As KEY4_ITEM_B
Public K5_ITEM_B As KEY5_ITEM_B
Public K6_ITEM_B As KEY6_ITEM_B

Type ITEM_B_FSpeck
    fs      As BtFileSpeck                 ' ̧�� ��߯��\����
    ks0     As BtKeySpeck                 ' �� ��߯��\����
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
    ks15    As BtKeySpeck
    ks16    As BtKeySpeck
    ks17    As BtKeySpeck
    ks18    As BtKeySpeck
    ks19    As BtKeySpeck
    ks20    As BtKeySpeck
    ks21    As BtKeySpeck
End Type

Private ITEM_B_Speck  As ITEM_B_FSpeck

Private Function ITEM_B_Create() As Integer
'********************************************************************
'*
'*              �i�ڃ}�X�^  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    ITEM_B_Create = True
                                            '�i�ڃ}�X�^�t���p�X�捞��
    sts = GetIni("FILE", ITEM_B_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ITEM_B]�ǂݍ��݃G���[ ")
        Exit Function
    End If

    FullPath = RTrim(c)

    ITEM_B_Speck.fs.recoleng = Len(ITEM_BREC)       ' ���R�[�h��
    ITEM_B_Speck.fs.PageSize = ITEM_B_PG_SIZ        ' �y�[�W�T�C�Y
    ITEM_B_Speck.fs.idexnumb = 7                  ' �C���f�b�N�X��
    ITEM_B_Speck.fs.fileflag = 0                  ' �t�@�C���t���O
    ITEM_B_Speck.fs.reserve = &H0                 ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�O
    ITEM_B_Speck.ks0.keypos = 1                   ' �L�[�|�W�V����
    ITEM_B_Speck.ks0.keyleng = 1                  ' �L�[��
    ITEM_B_Speck.ks0.keyflag = BtKfExt + BtKfSeg  ' �L�[�t���O
    ITEM_B_Speck.ks0.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    ITEM_B_Speck.ks0.reserve = &H0                ' �\��ς�

    ITEM_B_Speck.ks1.keypos = 2                   ' �L�[�|�W�V����
    ITEM_B_Speck.ks1.keyleng = 1                  ' �L�[��
    ITEM_B_Speck.ks1.keyflag = BtKfExt + BtKfSeg  ' �L�[�t���O
    ITEM_B_Speck.ks1.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    ITEM_B_Speck.ks1.reserve = &H0                ' �\��ς�

    ITEM_B_Speck.ks2.keypos = 3                   ' �L�[�|�W�V����
    ITEM_B_Speck.ks2.keyleng = 20                 ' �L�[��
    ITEM_B_Speck.ks2.keyflag = BtKfExt            ' �L�[�t���O
    ITEM_B_Speck.ks2.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    ITEM_B_Speck.ks2.reserve = &H0                ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�P
    ITEM_B_Speck.ks3.keypos = 95                  ' �L�[�|�W�V����
    ITEM_B_Speck.ks3.keyleng = 8                  ' �L�[��
                                                ' �L�[�t���O
    ITEM_B_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfChg
    ITEM_B_Speck.ks3.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    ITEM_B_Speck.ks3.reserve = &H0                ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�Q
    ITEM_B_Speck.ks4.keypos = 1                   ' �L�[�|�W�V����
    ITEM_B_Speck.ks4.keyleng = 1                  ' �L�[��
                                                ' �L�[�t���O
    ITEM_B_Speck.ks4.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    ITEM_B_Speck.ks4.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    ITEM_B_Speck.ks4.reserve = &H0                ' �\��ς�

    ITEM_B_Speck.ks5.keypos = 2                   ' �L�[�|�W�V����
    ITEM_B_Speck.ks5.keyleng = 1                  ' �L�[��
                                                ' �L�[�t���O
    ITEM_B_Speck.ks5.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    ITEM_B_Speck.ks5.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    ITEM_B_Speck.ks5.reserve = &H0                ' �\��ς�

    ITEM_B_Speck.ks6.keypos = 103                 ' �L�[�|�W�V����
    ITEM_B_Speck.ks6.keyleng = 20                 ' �L�[��
                                                ' �L�[�t���O
    ITEM_B_Speck.ks6.keyflag = BtKfExt + BtKfDup + BtKfChg
    ITEM_B_Speck.ks6.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    ITEM_B_Speck.ks6.reserve = &H0                ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�R
    ITEM_B_Speck.ks7.keypos = 1                   ' �L�[�|�W�V����
    ITEM_B_Speck.ks7.keyleng = 1                  ' �L�[��
                                                ' �L�[�t���O
    ITEM_B_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    ITEM_B_Speck.ks7.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    ITEM_B_Speck.ks7.reserve = &H0                ' �\��ς�

    ITEM_B_Speck.ks8.keypos = 63                  ' �L�[�|�W�V����
    ITEM_B_Speck.ks8.keyleng = 8                  ' �L�[��
                                                ' �L�[�t���O
    ITEM_B_Speck.ks8.keyflag = BtKfExt + BtKfDup + BtKfChg
    ITEM_B_Speck.ks8.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    ITEM_B_Speck.ks8.reserve = &H0                ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�S
    ITEM_B_Speck.ks9.keypos = 1                   ' �L�[�|�W�V����
    ITEM_B_Speck.ks9.keyleng = 1                  ' �L�[��
                                                ' �L�[�t���O
    ITEM_B_Speck.ks9.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    ITEM_B_Speck.ks9.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    ITEM_B_Speck.ks9.reserve = &H0                ' �\��ς�

    ITEM_B_Speck.ks10.keypos = 2                  ' �L�[�|�W�V����
    ITEM_B_Speck.ks10.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    ITEM_B_Speck.ks10.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    ITEM_B_Speck.ks10.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    ITEM_B_Speck.ks10.reserve = &H0               ' �\��ς�

    ITEM_B_Speck.ks11.keypos = 197                ' �L�[�|�W�V����
    ITEM_B_Speck.ks11.keyleng = 20                ' �L�[��
                                                ' �L�[�t���O
    ITEM_B_Speck.ks11.keyflag = BtKfExt + BtKfDup + BtKfChg
    ITEM_B_Speck.ks11.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    ITEM_B_Speck.ks11.reserve = &H0               ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�T
    ITEM_B_Speck.ks12.keypos = 1                  ' �L�[�|�W�V����
    ITEM_B_Speck.ks12.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    ITEM_B_Speck.ks12.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    ITEM_B_Speck.ks12.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    ITEM_B_Speck.ks12.reserve = &H0               ' �\��ς�

    ITEM_B_Speck.ks13.keypos = 2                  ' �L�[�|�W�V����
    ITEM_B_Speck.ks13.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    ITEM_B_Speck.ks13.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    ITEM_B_Speck.ks13.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    ITEM_B_Speck.ks13.reserve = &H0               ' �\��ς�

    ITEM_B_Speck.ks14.keypos = 217                ' �L�[�|�W�V����
    ITEM_B_Speck.ks14.keyleng = 20                ' �L�[��
                                                ' �L�[�t���O
    ITEM_B_Speck.ks14.keyflag = BtKfExt + BtKfDup + BtKfChg
    ITEM_B_Speck.ks14.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    ITEM_B_Speck.ks14.reserve = &H0               ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�U
    ITEM_B_Speck.ks15.keypos = 1                  ' �L�[�|�W�V����
    ITEM_B_Speck.ks15.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    ITEM_B_Speck.ks15.keyflag = BtKfExt + BtKfSeg + BtKfChg
    ITEM_B_Speck.ks15.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    ITEM_B_Speck.ks15.reserve = &H0               ' �\��ς�

    ITEM_B_Speck.ks16.keypos = 2                  ' �L�[�|�W�V����
    ITEM_B_Speck.ks16.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    ITEM_B_Speck.ks16.keyflag = BtKfExt + BtKfSeg + BtKfChg
    ITEM_B_Speck.ks16.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    ITEM_B_Speck.ks16.reserve = &H0               ' �\��ς�

    ITEM_B_Speck.ks17.keypos = 71                  ' �L�[�|�W�V����
    ITEM_B_Speck.ks17.keyleng = 2                 ' �L�[��
                                                ' �L�[�t���O
    ITEM_B_Speck.ks17.keyflag = BtKfExt + BtKfSeg + BtKfChg
    ITEM_B_Speck.ks17.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    ITEM_B_Speck.ks17.reserve = &H0               ' �\��ς�

    ITEM_B_Speck.ks18.keypos = 73                 ' �L�[�|�W�V����
    ITEM_B_Speck.ks18.keyleng = 2                 ' �L�[��
                                                ' �L�[�t���O
    ITEM_B_Speck.ks18.keyflag = BtKfExt + BtKfSeg + BtKfChg
    ITEM_B_Speck.ks18.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    ITEM_B_Speck.ks18.reserve = &H0               ' �\��ς�

    ITEM_B_Speck.ks19.keypos = 75                 ' �L�[�|�W�V����
    ITEM_B_Speck.ks19.keyleng = 2                 ' �L�[��
                                                ' �L�[�t���O
    ITEM_B_Speck.ks19.keyflag = BtKfExt + BtKfSeg + BtKfChg
    ITEM_B_Speck.ks19.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    ITEM_B_Speck.ks19.reserve = &H0               ' �\��ς�

    ITEM_B_Speck.ks20.keypos = 77                 ' �L�[�|�W�V����
    ITEM_B_Speck.ks20.keyleng = 2                 ' �L�[��
                                                ' �L�[�t���O
    ITEM_B_Speck.ks20.keyflag = BtKfExt + BtKfSeg + BtKfChg
    ITEM_B_Speck.ks20.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    ITEM_B_Speck.ks20.reserve = &H0               ' �\��ς�

    ITEM_B_Speck.ks21.keypos = 3                  ' �L�[�|�W�V����
    ITEM_B_Speck.ks21.keyleng = 20                ' �L�[��
                                                ' �L�[�t���O
    ITEM_B_Speck.ks21.keyflag = BtKfExt + BtKfChg
    ITEM_B_Speck.ks21.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    ITEM_B_Speck.ks21.reserve = &H0               ' �\��ς�
'-----------------------------------------------

    sts = BTRV(BtOpCreate, ITEM_B_POS, ITEM_B_Speck, Len(ITEM_B_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�i�ڃ}�X�^")
        Exit Function
    End If

    ITEM_B_Create = False

End Function

Public Function ITEM_B_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �i�ڃ}�X�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    ITEM_B_Open = True
                                            '�i�ڃ}�X�^�t���p�X�捞��
    sts = GetIni("FILE", ITEM_B_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ITEM_B]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, ITEM_B_POS, ITEM_BREC, Len(ITEM_BREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = ITEM_B_Create()        '�i�ڃ}�X�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ITEM_B_POS, ITEM_BREC, Len(ITEM_BREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�i�ڃ}�X�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�i�ڃ}�X�^")
                Exit Function
        End Select
    Loop

    ITEM_B_Open = False

End Function

Public Sub Rclr_ITEM_BREC()
'********************************************************************
'*
'*              �i�ڃ}�X�^  ���R�[�h������
'*
'********************************************************************
Dim i       As Long


    Call UniCode_Conv(ITEM_BREC.JGYOBU, "")           '���ƕ��敪
    Call UniCode_Conv(ITEM_BREC.NAIGAI, "")           '�����O
    Call UniCode_Conv(ITEM_BREC.HIN_GAI, "")          '�i�ԁi�O���j
    Call UniCode_Conv(ITEM_BREC.HIN_NAME, "")         '�i��
    Call UniCode_Conv(ITEM_BREC.ST_SET_DT, "")        '�W���q�ɐݒ���t
    Call UniCode_Conv(ITEM_BREC.ST_SOKO, "")          '�W�����ɑq�� �q��
    Call UniCode_Conv(ITEM_BREC.ST_RETU, "")          '             ��
    Call UniCode_Conv(ITEM_BREC.ST_REN, "")           '             �A
    Call UniCode_Conv(ITEM_BREC.ST_DAN, "")           '             �i
    Call UniCode_Conv(ITEM_BREC.BEF_SOKO, "")         '�O����ɑq�� �q��
    
    Call UniCode_Conv(ITEM_BREC.BEF_RETU, "")         '             ��
    Call UniCode_Conv(ITEM_BREC.BEF_REN, "")          '             �A
    Call UniCode_Conv(ITEM_BREC.BEF_DAN, "")          '             �i
    Call UniCode_Conv(ITEM_BREC.LAST_NYU_DT, "")      '�ŏI���ɓ��t
    Call UniCode_Conv(ITEM_BREC.LAST_SYU_DT, "")      '�ŏI�o�ɓ��t
    Call UniCode_Conv(ITEM_BREC.HIN_NAI, "")          '�i�ԁi�����j
    Call UniCode_Conv(ITEM_BREC.BIKOU_SOKO, "")       '���l �z�X�g�q��
    Call UniCode_Conv(ITEM_BREC.BIKOU_TANA, "")       '���l �z�X�g�I��
    Call UniCode_Conv(ITEM_BREC.LAST_INP_DT, "")      '�ŏI���ד��t
    Call UniCode_Conv(ITEM_BREC.LAST_CHK_DT, "")      '�ŏI�ƍ����t       2001.06.12
    
    Call UniCode_Conv(ITEM_BREC.BIKOU, "")            '������l
    Call UniCode_Conv(ITEM_BREC.JAN_CODE, "")         'Jan�R�[�h      2004.02
    Call UniCode_Conv(ITEM_BREC.HIN_CHANGE, "")       '�i�ԓǂݑւ�   2004.02
    Call UniCode_Conv(ITEM_BREC.GOODS_KBN, GOODS_ON)  '���i���L��     2004.02
    Call UniCode_Conv(ITEM_BREC.PACKING_NO, "")       '������       2004.02
    Call UniCode_Conv(ITEM_BREC.RANK, "")             '���݃����N     2004.06
    Call UniCode_Conv(ITEM_BREC.NEW_RANK, "")         '���݃����N     2004.06
    Call UniCode_Conv(ITEM_BREC.GLICS1_TANA, "")      '�O���b�N�X�I�ԂP   2005.05
    Call UniCode_Conv(ITEM_BREC.GLICS2_TANA, "")      '�O���b�N�X�I�ԂQ   2005.05
    Call UniCode_Conv(ITEM_BREC.GLICS3_TANA, "")      '�O���b�N�X�I�ԂR   2005.05
    
    Call UniCode_Conv(ITEM_BREC.G_SHIIRE_KBN, "")     '�Ɩ��Ǘ��@ �d���敪
    Call UniCode_Conv(ITEM_BREC.G_HANBAI_KBN, "")     '           �̔��敪
    Call UniCode_Conv(ITEM_BREC.G_SYUSHI, "")         '           ���x�P��
    Call UniCode_Conv(ITEM_BREC.G_KUMITATE, "")       '           �g�����i
    Call UniCode_Conv(ITEM_BREC.G_ST_URITAN_DT, "")   '           �W���e�������ݒ��
    Call UniCode_Conv(ITEM_BREC.G_ST_SHITAN_DT, "")   '           �W���e�������ݒ��
                                                    '           �d������
    For i = 0 To UBound(ITEM_BREC.G_SHIIRE_TBL)
        Call UniCode_Conv(ITEM_BREC.G_SHIIRE_TBL(i).CODE, "")             '����
        Call UniCode_Conv(ITEM_BREC.G_SHIIRE_TBL(i).TANKA_DT, "")         '�P���ݒ��
        Call UniCode_Conv(ITEM_BREC.G_SHIIRE_TBL(i).LEAD_TIME, "")        'ذ�����
        Call UniCode_Conv(ITEM_BREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, "")    '�O�񒍕���
    Next i
    
    Call UniCode_Conv(ITEM_BREC.G_SHIZAI_KBN, "")     '           ���ދ敪
    Call UniCode_Conv(ITEM_BREC.G_LABEL_NON, "")      '           ���ٓ\��t���v��Ȃ�
    
    Call UniCode_Conv(ITEM_BREC.L_HIN_NAME_E, "")     '���i����   �i��
    Call UniCode_Conv(ITEM_BREC.L_BIKOU, "")          '           ���l
    Call UniCode_Conv(ITEM_BREC.L_KAISHA_CODE, "")    '           ��ЃR�[�h
    Call UniCode_Conv(ITEM_BREC.L_KISHU1, "")         '           �@��(1)
    Call UniCode_Conv(ITEM_BREC.xL_KISHU2, "")        '           �@��(2) ���g�p 2006.01.24
    Call UniCode_Conv(ITEM_BREC.L_KISHU3, "")         '           �@��(3)(���K�p�@����l)
    Call UniCode_Conv(ITEM_BREC.L_PAPER, "0")         '           ��
    Call UniCode_Conv(ITEM_BREC.L_PLASTIC, "0")       '           �v���X�`�b�N
    Call UniCode_Conv(ITEM_BREC.L_LABEL, "0")         '           �K�p�@������
    Call UniCode_Conv(ITEM_BREC.L_MAISU, "0")         '           ��������
    Call UniCode_Conv(ITEM_BREC.L_KISHU_BIKOU, "")    '           �K�p�@����l(���@��i�R�j)
    Call UniCode_Conv(ITEM_BREC.L_SAGYO_SHIJI, "")    '           ��Ǝw��
    Call UniCode_Conv(ITEM_BREC.L_BIKOU3, "")         '           ���l�R
    Call UniCode_Conv(ITEM_BREC.L_JGYOBU_CODE, "")    '           ���ƕ��R�[�h
    Call UniCode_Conv(ITEM_BREC.L_TANA1, "")          '           �I��(1)
    Call UniCode_Conv(ITEM_BREC.L_TANA2, "")          '           �I��(2)
    
    Call UniCode_Conv(ITEM_BREC.S_TANTO, "")          '���P�^�S���҃R�[�h
    Call UniCode_Conv(ITEM_BREC.ZAIKO_F, "")          '�݌ɊǗ��ΏۗL�� 1:�Ώ� 0:�ΏۊO
    Call UniCode_Conv(ITEM_BREC.L_KISHU2, "")         '           �@��(2)
    Call UniCode_Conv(ITEM_BREC.K_KEITAI, "")         '���`��
    Call UniCode_Conv(ITEM_BREC.UNIT_BUHIN, "")       '�Ưĕ��i�敪       2006.07.28
    Call UniCode_Conv(ITEM_BREC.NAI_BUHIN, "")        '�����������i�敪   2006.07.28
    Call UniCode_Conv(ITEM_BREC.GAI_BUHIN, "")        '�C�O�������i�敪   2006.07.28
    Call UniCode_Conv(ITEM_BREC.LAST_CODE, "")        '�ŏI�d����R�[�h   2007.05.29
    Call UniCode_Conv(ITEM_BREC.MAKER_CODE, "")       'Ұ������           2007.06.06
    Call UniCode_Conv(ITEM_BREC.MAKER_NAME, "")       'Ұ������           2007.06.06
    
    Call UniCode_Conv(ITEM_BREC.L_MARK, "")           '�č���ϰ�          2007.11.08
    Call UniCode_Conv(ITEM_BREC.SAI_SU, "")           '�ː�               2008.02.14
    Call UniCode_Conv(ITEM_BREC.D_KEISHIKI, "")       '�`��               2008.02.14
    Call UniCode_Conv(ITEM_BREC.D_MATERIAL, "")       '�ގ�               2008.02.14
    Call UniCode_Conv(ITEM_BREC.D_THICKNESS, "")      '����ްف@����      2008.02.14
    Call UniCode_Conv(ITEM_BREC.D_SIZE_W, "")         '����ްٻ��ށiW�j   2008.02.14
    Call UniCode_Conv(ITEM_BREC.D_SIZE_D, "")         '����ްٻ��ށiD�j   2008.02.14
    Call UniCode_Conv(ITEM_BREC.D_SIZE_H, "")         '����ްٻ��ށiH�j   2008.02.14
    Call UniCode_Conv(ITEM_BREC.D_PRINT, "")          '�������^���Ȃ�    2008.02.14
    Call UniCode_Conv(ITEM_BREC.S_KOUSU_SET_DATE, "") '���i���@�P���ݒ�� 2008.02.14
    
    Call UniCode_Conv(ITEM_BREC.S_SHIZAI_SET_DATE, "") '���i���@�P���ݒ�� 2008.02.14
    Call UniCode_Conv(ITEM_BREC.SE_USOU_F, "")        '�A�����@�o���׸�   2008.02.14
    Call UniCode_Conv(ITEM_BREC.USE_TAPE_KIND, "")    '�g�p�e�[�v���     2008.02.14
    Call UniCode_Conv(ITEM_BREC.USE_TAPE_LNG, "")     '�g�p�e�[�v��       2008.02.14
    Call UniCode_Conv(ITEM_BREC.H_TANA_MAKE, "")      '�I�ԃ}�[�N         2008.04.02
    Call UniCode_Conv(ITEM_BREC.SE_TANKA_MEMO, "")    '�����P���@����     2008.04.15
    Call UniCode_Conv(ITEM_BREC.GENSANKOKU, "")       '���Y��             2008.06.11
    Call UniCode_Conv(ITEM_BREC.SEI_LOT, "")          '���Y���b�g         2008.07.07
    Call UniCode_Conv(ITEM_BREC.SEI_SYU_KON, "")      '�W������           2008.07.07
    Call UniCode_Conv(ITEM_BREC.SEI_TANKA_TANTO, "")  '�P���ݒ�S����     2008.07.09
    Call UniCode_Conv(ITEM_BREC.SHIMUKE_CODE, "")     '�d������           2008.07.09
    Call UniCode_Conv(ITEM_BREC.SEI_KBN, "")          '�����敪           2008.07.16
                                            '�O�H��             2008.09.19
    For i = 0 To UBound(ITEM_BREC.BEF_KOUTEI)
        Call UniCode_Conv(ITEM_BREC.BEF_KOUTEI(i).BEF_KOUTEI, "")     '�O�H�� 2008.09.19
    Next i
                                            '��ƍH��           2008.09.19
    For i = 0 To UBound(ITEM_BREC.MAIN_KOUTEI)
        Call UniCode_Conv(ITEM_BREC.MAIN_KOUTEI(i).MAIN_KOUTEI, "")   '��ƍH�� 2008.09.19
    Next i
                                            '��H��             2008.09.19
    For i = 0 To UBound(ITEM_BREC.AFT_KOUTEI)
        Call UniCode_Conv(ITEM_BREC.AFT_KOUTEI(i).AFT_KOUTEI, "")     '��H�� 2008.09.19
    Next i

    Call UniCode_Conv(ITEM_BREC.SE_IO_TANKA_No, "")   '�I�敪             200.09.19
    Call UniCode_Conv(ITEM_BREC.STAT, "")             '��ԋ敪           2009.01.21
    Call UniCode_Conv(ITEM_BREC.INSP_MESSAGE, "")     '�o�׌��iү����     2009.04.17
    Call UniCode_Conv(ITEM_BREC.S_SEIKYU_F, "")       '���i�������׸�     2009.04.28
    Call UniCode_Conv(ITEM_BREC.FILLER, "")           'FILLER             2009.04.28�T�C�Y�ύX
    
    Call UniCode_Conv(ITEM_BREC.INS_TANTO, "")        '�ǉ��@�S���ҁ@     2009.01.21
    Call UniCode_Conv(ITEM_BREC.Ins_DateTime, "")     '�ǉ��@����         2009.01.21
    Call UniCode_Conv(ITEM_BREC.UPD_TANTO, "")        '�X�V�@�S���ҁ@     2005.11.15
    Call UniCode_Conv(ITEM_BREC.UPD_DATETIME, "")     '�X�V�@����         2005.11.15

'-------------------------------------------------------------------------------------------
'               �O�N���A����

                                                    '��[�_�i�댯�݌Ɂj
    Call UniCode_Conv(ITEM_BREC.HOJYU_P, String(UBound(ITEM_BREC.HOJYU_P) + 1, "0"))
                                                    '�����Ϗo�א�
    Call UniCode_Conv(ITEM_BREC.AVE_SYUKA, String(UBound(ITEM_BREC.AVE_SYUKA) + 1, "0"))
                                                    '�T���v����
    Call UniCode_Conv(ITEM_BREC.SAMPLE_QTY, String(UBound(ITEM_BREC.SAMPLE_QTY) + 1, "0"))
                                                    '�ŏI�ƍ����݌ɐ�   2001.06.12
    Call UniCode_Conv(ITEM_BREC.LAST_CHK_QTY, String(UBound(ITEM_BREC.LAST_CHK_QTY) + 1, "0"))
                                                    '������萔
    Call UniCode_Conv(ITEM_BREC.IRI_QTY, String(UBound(ITEM_BREC.IRI_QTY) + 1, "0"))
                                                    '           �W���e�������P���@9(8)V99
    Call UniCode_Conv(ITEM_BREC.G_ST_URITAN, String(UBound(ITEM_BREC.G_ST_URITAN) + 1, "0"))
                                                    '           �W���e�������P��  9(8)V99
    Call UniCode_Conv(ITEM_BREC.G_ST_SHITAN, String(UBound(ITEM_BREC.G_ST_SHITAN) + 1, "0"))

    For i = 0 To UBound(ITEM_BREC.G_SHIIRE_TBL)
                                                                        '�P�� 9(8)V99
        Call UniCode_Conv(ITEM_BREC.G_SHIIRE_TBL(i).TANKA, _
                   String(UBound(ITEM_BREC.G_SHIIRE_TBL(i).TANKA) + 1, "0"))
                                                                        'ۯĐ�
        Call UniCode_Conv(ITEM_BREC.G_SHIIRE_TBL(i).LOT, _
                   String(UBound(ITEM_BREC.G_SHIIRE_TBL(i).LOT) + 1, "0"))
                                                                        '�O�񒍕���
        Call UniCode_Conv(ITEM_BREC.G_SHIIRE_TBL(i).LAST_ORDER_QTY, _
                   String(UBound(ITEM_BREC.G_SHIIRE_TBL(i).LAST_ORDER_QTY) + 1, "0"))
    Next i
                                                    '           �O���݌ɋ��z
    Call UniCode_Conv(ITEM_BREC.G_ZEN_ZAIKO_KIN, String(UBound(ITEM_BREC.G_ZEN_ZAIKO_KIN) + 1, "0"))
                                                    '           ���i(1)
    Call UniCode_Conv(ITEM_BREC.L_URIKIN1, String(UBound(ITEM_BREC.L_URIKIN1) + 1, "0"))
                                                    '           ���i(2)
    Call UniCode_Conv(ITEM_BREC.L_URIKIN2, String(UBound(ITEM_BREC.L_URIKIN2) + 1, "0"))
                                                    '           ���i(3)
    Call UniCode_Conv(ITEM_BREC.L_URIKIN3, String(UBound(ITEM_BREC.L_URIKIN3) + 1, "0"))
                                                    '           ���萔
    Call UniCode_Conv(ITEM_BREC.L_IRI_QTY, String(UBound(ITEM_BREC.L_IRI_QTY) + 1, "0"))
                                                    '           �O���݌ɐ���
    Call UniCode_Conv(ITEM_BREC.G_ZEN_ZAIKO_QTY, String(UBound(ITEM_BREC.G_ZEN_ZAIKO_QTY) + 1, "0"))
                                                    '           �ŏI�o�א�
    Call UniCode_Conv(ITEM_BREC.G_LAST_SYUKA_QTY, String(UBound(ITEM_BREC.G_LAST_SYUKA_QTY) + 1, "0"))
                                                    'GLICS�݌�(S2) �܈�p
    Call UniCode_Conv(ITEM_BREC.G_S2_ZAI_QTY, String(UBound(ITEM_BREC.G_S2_ZAI_QTY) + 1, "0"))
                                                    'GLICS�݌�(P2) �܈�p
    Call UniCode_Conv(ITEM_BREC.G_P2_ZAI_QTY, String(UBound(ITEM_BREC.G_P2_ZAI_QTY) + 1, "0"))
                                                    '�W���P��   2006.07.28
    Call UniCode_Conv(ITEM_BREC.HYO_TANKA, String(UBound(ITEM_BREC.HYO_TANKA) + 1, "0"))
                                                    '�ŏI�d���P��       2007.05.29
    Call UniCode_Conv(ITEM_BREC.LAST_TANKA, String(UBound(ITEM_BREC.LAST_TANKA) + 1, "0"))
                                                    '���i���@�H��       2008.02.14
    Call UniCode_Conv(ITEM_BREC.S_KOUSU, String(UBound(ITEM_BREC.S_KOUSU) + 1, "0"))
                                                    '���i���@�H������   2008.02.14
    Call UniCode_Conv(ITEM_BREC.S_KOUSU_GENKA, String(UBound(ITEM_BREC.S_KOUSU_GENKA) + 1, "0"))
                                                    '���i���@�H������   2008.02.14
    Call UniCode_Conv(ITEM_BREC.S_KOUSU_BAIKA, String(UBound(ITEM_BREC.S_KOUSU_BAIKA) + 1, "0"))
                                                    '���i���@���ތ���   2008.02.14
    Call UniCode_Conv(ITEM_BREC.S_SHIZAI_GENKA, String(UBound(ITEM_BREC.S_SHIZAI_GENKA) + 1, "0"))
                                                    '���i���@���ޔ���   2008.02.14
    Call UniCode_Conv(ITEM_BREC.S_SHIZAI_BAIKA, String(UBound(ITEM_BREC.S_SHIZAI_BAIKA) + 1, "0"))

                                                    '�O���P�� 9(8)V99   2008.06.12
    Call UniCode_Conv(ITEM_BREC.S_GAISO_TANKA, String(UBound(ITEM_BREC.S_GAISO_TANKA) + 1, "0"))
                                                    'PPSC���H�P��9(8)   2008.06.12
    Call UniCode_Conv(ITEM_BREC.S_PPSC_KAKO_KOSU, String(UBound(ITEM_BREC.S_PPSC_KAKO_KOSU) + 1, "0"))
                                                    'BU���H�P��9(8)     2008.06.12
    Call UniCode_Conv(ITEM_BREC.S_BU_KAKO_KOSU, String(UBound(ITEM_BREC.S_BU_KAKO_KOSU) + 1, "0"))
    
                                                    '�����[�g           2008.07.07
    Call UniCode_Conv(ITEM_BREC.SEI_RATE, String(UBound(ITEM_BREC.SEI_RATE) + 1, "0"))
    
                                                    '���x���\�薇��     2008.07.19
    Call UniCode_Conv(ITEM_BREC.SEI_LABEL_QTY, String(UBound(ITEM_BREC.SEI_LABEL_QTY) + 1, "0"))

                                                    '���ތ���     �@    2008.08.20�ǉ�
    Call UniCode_Conv(ITEM_BREC.SEI_SZI_CNT, String(UBound(ITEM_BREC.SEI_SZI_CNT) + 1, "0"))
                                                    '��������           2008.08.20�ǉ�
    Call UniCode_Conv(ITEM_BREC.SEI_DKN_CNT, String(UBound(ITEM_BREC.SEI_DKN_CNT) + 1, "0"))


'-------------------------------------------------------------------------------------------
'               2009.06.02
                                                    '���i���@�H������
    Call UniCode_Conv(ITEM_BREC.BEF_S_KOUSU_BAIKA, String(UBound(ITEM_BREC.BEF_S_KOUSU_BAIKA) + 1, "0"))
                                                    '���i���@���ޔ���
    Call UniCode_Conv(ITEM_BREC.BEF_S_SHIZAI_BAIKA, String(UBound(ITEM_BREC.BEF_S_SHIZAI_BAIKA) + 1, "0"))
                                                    '�O���P��
    Call UniCode_Conv(ITEM_BREC.BEF_S_GAISO_TANKA, String(UBound(ITEM_BREC.BEF_S_GAISO_TANKA) + 1, "0"))
                                                    'PPSC���H�P��
    Call UniCode_Conv(ITEM_BREC.BEF_S_PPSC_KAKO_KOSU, String(UBound(ITEM_BREC.BEF_S_PPSC_KAKO_KOSU) + 1, "0"))
                                                    'BU���H�P��
    Call UniCode_Conv(ITEM_BREC.BEF_S_BU_KAKO_KOSU, String(UBound(ITEM_BREC.BEF_S_BU_KAKO_KOSU) + 1, "0"))
    
    Call UniCode_Conv(ITEM_BREC.M_BIKOU, "")              '���Ϗ����l
    
    Call UniCode_Conv(ITEM_BREC.SHIYOU_NO, "")            '�d�l����
    
    Call UniCode_Conv(ITEM_BREC.MITSUMORI_KBN, "")        '���ς�敪
    
    Call UniCode_Conv(ITEM_BREC.TANKA_KIRIKAE_DT, "")    '�P���ؑ֓��t
    
    Call UniCode_Conv(ITEM_BREC.KIRIKAE_KBN, "")          '�ؑ֋敪
    
'               2009.06.02
'-------------------------------------------------------------------------------------------

End Sub
