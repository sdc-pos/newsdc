Attribute VB_Name = "ITEM"
Option Explicit
'********************************************************************
'*
'*              �i�ڃ}�X�^  �t�@�C����`
'*
'*          CREATE 2004.02.19
'********************************************************************
'�t�@�C���h�c
Public Const ITEM_ID$ = "ITEM"

'�y�[�W�T�C�Y
Public Const ITEM_PG_SIZ% = 4096

'�|�W�V�����E�u���b�N
Public ITEM_POS         As POSBLK
'=
'====================================================================
'=          ���R�[�h�������v���V�[�W��     ( Rclr_ITEMREC )
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
Type ITEMREC_Tag
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


    xK_KEITAI(0 To 9)           As Byte     '���`��(10��) ���g�p 2012.03.13


    UNIT_BUHIN(0 To 0)          As Byte     '�Ưĕ��i�敪       2006.07.28
    NAI_BUHIN(0 To 0)           As Byte     '�����������i�敪   2006.07.28
    GAI_BUHIN(0 To 0)           As Byte     '�C�O�������i�敪   2006.07.28
    HYO_TANKA(0 To 9)           As Byte     '�W���P��   2006.07.28

    LAST_CODE(0 To 4)           As Byte     '�ŏI�d����R�[�h   2007.05.29
    LAST_TANKA(0 To 10)         As Byte     '�ŏI�d���P��       2007.05.29

    MAKER_CODE(0 To 7)          As Byte     'Ұ������           2007.06.06
    MAKER_NAME(0 To 39)         As Byte     'Ұ������           2007.06.06


    L_MARK(0 To 0)              As Byte     '�č���ϰ�          2007.11.08

    xSAI_SU(0 To 3)              As Byte    '�ː�               2008.02.14

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

                                            '�O�H��             2008.09.19  2011.12.12 ���g�p�Ƃ���
    BEF_KOUTEI(0 To 9)          As BEF_KOUTEI_tag
                                            '��ƍH��           2008.09.19
    MAIN_KOUTEI(0 To 9)         As MAIN_KOUTEI_tag
                                            '��H��             2008.09.19  2011.12.12 ���g�p�Ƃ���
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
    
    
    
    PLUS_KOUSU(0 To 5)          As Byte     '�v���X���H��       2009.09.17  2011.12.12 ���g�p�Ƃ���
    
    
    
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
    
    
    BIKOU20(0 To 19)            As Byte     '������l
    
    
    PRT_GENSANKOKU(0 To 0)      As Byte     '���Y���󎚗L��     2010.11.10
    GAISO_IRI_QTY(0 To 7)       As Byte     '�O�������萔 (9(8)) 2010.11.10
    
    
    GOODS_OUT_F(0 To 0)         As Byte     '�u���i���v��v���O�׸� "1":���O    2011.06.30
    
    
    PLN_KOUSU(0 To 10)          As Byte     '�u���i�����сv�p�W���H��           2011.10.02
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   ���Ϗ�����(�i���ú�ذ�Ή�)  2011.12.12
    G_SPTAN(0 To 10)            As Byte     ' �u�������сv���ʒP�� 9(8).99  (�\�� ���g�p)
    
    CATE_ST_KOUTEI(0 To 5)      As Byte     ' �u�������сv�O��H���i�b�j    �W��    9(3).99
    CATE_ST_FUKA(0 To 5)        As Byte     ' �u�������сv�t���H���i�b�j    �W��    9(3).99
    CATE_ST_JITU1(0 To 5)       As Byte     ' �u�������сv ����ƍH���i�b�j �W��    9(3).99
    CATE_ST_YOYU_RITU(0 To 5)   As Byte     ' �u�������сv �]�T���i���j     �W��    9(3).99
    CATE_ST_JITU2(0 To 5)       As Byte     ' �u�������сv ����ƍH���i�b�j �W��    9(3).99
    CATE_ST_TOTAL(0 To 5)       As Byte     ' �u�������сv ��Ǝ��Ԍv�i�b�j �W��    9(3).99
    CATE_ST_FUN(0 To 5)         As Byte     ' �u�������сv ��/�i��/�j   �W��    9(3).99
    CATE_ST_FUN_RATE(0 To 6)    As Byte     ' �u�������сv ��ڰāi�~/���j   �W��    9(4).99
    CATE_ST_KOURYO(0 To 12)     As Byte     ' �u�������сv �H�����i�~/�j  �W��    9(10).99
    
    
    
    
    CATE_AD_KOUTEI(0 To 5)      As Byte     ' �u�������сv�O��H���i�b�j    ����    9(3).99
    CATE_AD_FUKA(0 To 5)        As Byte     ' �u�������сv �t���H���i�b�j   ����    9(3).99
    CATE_AD_JITU1(0 To 5)       As Byte     ' �u�������сv ����ƍH���i�b�j ����    9(3).99
    CATE_AD_YOYU_RITU(0 To 5)   As Byte     ' �u�������сv �]�T���i���j     ����    9(3).99
    CATE_AD_JITU2(0 To 5)       As Byte     ' �u�������сv ����ƍH���i�b�j ����    9(3).99
    CATE_AD_TOTAL(0 To 5)       As Byte     ' �u�������сv ��Ǝ��Ԍv�i�b�j ����    9(3).99
    CATE_AD_FUN(0 To 5)         As Byte     ' �u�������сv  ��/�i��/�j  ����    9(3).99
    CATE_AD_FUN_RATE(0 To 6)    As Byte     ' �u�������сv  ��ڰāi�~/���j  ����    9(4).99
    CATE_AD_KOURYO(0 To 12)     As Byte     ' �u�������сv  �H�����i�~/�j ����    9(10).99
    
    CATEGORY_CODE(0 To 7)       As Byte
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   ���Ϗ�����(�i���ú�ذ�Ή�)  2011.12.12
    CS_TANTO_CD(0 To 7)         As Byte     'CS�S������ 2011.12.12
        
    D_MODEL(0 To 7)             As Byte     '��\�@��i�ں��ށ@ PN�A�g�ž��  2011.12.28
    HINMOKU(0 To 7)             As Byte     '�i�ں��ށ@         PN�A�g�ž��  2011.12.28
    K_KEITAI(0 To 13)           As Byte     '���`��(14��)     2012.03.13
    
    PLN_SAGYOU_KOUSU(0 To 10)   As Byte     '�u���i�����сv��ƍH�� 2012.05.07
    
    
    BEF_S_KOUSU(0 To 7)         As Byte     '���i���@�H��(���J�o���[�O)       2012.09.24
    
    
    ZAIKO_CLR_F(0 To 0)         As Byte     '�I���\�@�݌ɐ���\���e 2012.12.13 (1:��\��(0�Ƃ݂Ȃ�)�@�ȊO:�ʏ�)
    
    SHIIRE_BIKOU(0 To 39)       As Byte     '�d�����l   2018.04.19
    
    
    FILLER(0 To 0)              As Byte     'FILLER             2012.05.07   ���ڒǉ��ɂ��T�C�Y�ύX 49-41 2012.09.24�@41-40�@2012.12.13 40->0 2018.04.19
    

    INS_TANTO(0 To 4)           As Byte     '�ǉ��@�S���ҁ@     2009.01.21
    Ins_DateTime(0 To 13)       As Byte     '�ǉ��@����         2009.01.21
    UPD_TANTO(0 To 4)           As Byte     '�X�V�@�S���ҁ@     2005.11.15
    UPD_DATETIME(0 To 13)       As Byte     '�X�V�@����         2005.11.15

End Type
'�f�[�^�E�o�b�t�@
Public ITEMREC As ITEMREC_Tag

'�L�[��`

Type KEY0_ITEM            '�j�d�x�O
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    '2005.11.15 �����ύX 13---> 20
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
End Type

Type KEY1_ITEM            '�j�d�x�P
    LAST_SYU_DT(0 To 7) As Byte     '�ŏI�o�ɓ��t
End Type

Type KEY2_ITEM            '�j�d�x�Q
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    '2005.11.15 �����ύX 13---> 20
    HIN_NAI(0 To 19)    As Byte     '�i�ԁi�����j
End Type

Type KEY3_ITEM            '�j�d�x�R
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    ST_SET_DT(0 To 7)   As Byte     '�W���q�ɐݒ���t
End Type


Type KEY4_ITEM            '�j�d�x�S 2004.02
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    '2005.11.15 �����ύX 13---> 20
    JAN_CODE(0 To 19)   As Byte     'Jan�R�[�h
End Type

Type KEY5_ITEM            '�j�d�x�T 2004.02
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    '2005.11.15 �����ύX 13---> 20
    HIN_CHANGE(0 To 19) As Byte     '�i�ԓǂݑւ�
End Type

Type KEY6_ITEM            '�j�d�x�U 2004.02
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
Public K0_ITEM As KEY0_ITEM
Public K1_ITEM As KEY1_ITEM
Public K2_ITEM As KEY2_ITEM
Public K3_ITEM As KEY3_ITEM
Public K4_ITEM As KEY4_ITEM
Public K5_ITEM As KEY5_ITEM
Public K6_ITEM As KEY6_ITEM

Type ITEM_FSpeck
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

Private ITEM_Speck  As ITEM_FSpeck

Private Function ITEM_Create() As Integer
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

    ITEM_Create = True
                                            '�i�ڃ}�X�^�t���p�X�捞��
    sts = GetIni("FILE", ITEM_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ITEM]�ǂݍ��݃G���[ ")
        Exit Function
    End If

    FullPath = RTrim(c)

    ITEM_Speck.fs.recoleng = Len(ITEMREC)       ' ���R�[�h��
    ITEM_Speck.fs.PageSize = ITEM_PG_SIZ        ' �y�[�W�T�C�Y
    ITEM_Speck.fs.idexnumb = 7                  ' �C���f�b�N�X��
    ITEM_Speck.fs.fileflag = 0                  ' �t�@�C���t���O
    ITEM_Speck.fs.reserve = &H0                 ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�O
    ITEM_Speck.ks0.keypos = 1                   ' �L�[�|�W�V����
    ITEM_Speck.ks0.keyleng = 1                  ' �L�[��
    ITEM_Speck.ks0.keyflag = BtKfExt + BtKfSeg  ' �L�[�t���O
    ITEM_Speck.ks0.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    ITEM_Speck.ks0.reserve = &H0                ' �\��ς�

    ITEM_Speck.ks1.keypos = 2                   ' �L�[�|�W�V����
    ITEM_Speck.ks1.keyleng = 1                  ' �L�[��
    ITEM_Speck.ks1.keyflag = BtKfExt + BtKfSeg  ' �L�[�t���O
    ITEM_Speck.ks1.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    ITEM_Speck.ks1.reserve = &H0                ' �\��ς�

    ITEM_Speck.ks2.keypos = 3                   ' �L�[�|�W�V����
    ITEM_Speck.ks2.keyleng = 20                 ' �L�[��
    ITEM_Speck.ks2.keyflag = BtKfExt            ' �L�[�t���O
    ITEM_Speck.ks2.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    ITEM_Speck.ks2.reserve = &H0                ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�P
    ITEM_Speck.ks3.keypos = 95                  ' �L�[�|�W�V����
    ITEM_Speck.ks3.keyleng = 8                  ' �L�[��
                                                ' �L�[�t���O
    ITEM_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfChg
    ITEM_Speck.ks3.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    ITEM_Speck.ks3.reserve = &H0                ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�Q
    ITEM_Speck.ks4.keypos = 1                   ' �L�[�|�W�V����
    ITEM_Speck.ks4.keyleng = 1                  ' �L�[��
                                                ' �L�[�t���O
    ITEM_Speck.ks4.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    ITEM_Speck.ks4.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    ITEM_Speck.ks4.reserve = &H0                ' �\��ς�

    ITEM_Speck.ks5.keypos = 2                   ' �L�[�|�W�V����
    ITEM_Speck.ks5.keyleng = 1                  ' �L�[��
                                                ' �L�[�t���O
    ITEM_Speck.ks5.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    ITEM_Speck.ks5.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    ITEM_Speck.ks5.reserve = &H0                ' �\��ς�

    ITEM_Speck.ks6.keypos = 103                 ' �L�[�|�W�V����
    ITEM_Speck.ks6.keyleng = 20                 ' �L�[��
                                                ' �L�[�t���O
    ITEM_Speck.ks6.keyflag = BtKfExt + BtKfDup + BtKfChg
    ITEM_Speck.ks6.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    ITEM_Speck.ks6.reserve = &H0                ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�R
    ITEM_Speck.ks7.keypos = 1                   ' �L�[�|�W�V����
    ITEM_Speck.ks7.keyleng = 1                  ' �L�[��
                                                ' �L�[�t���O
    ITEM_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    ITEM_Speck.ks7.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    ITEM_Speck.ks7.reserve = &H0                ' �\��ς�

    ITEM_Speck.ks8.keypos = 63                  ' �L�[�|�W�V����
    ITEM_Speck.ks8.keyleng = 8                  ' �L�[��
                                                ' �L�[�t���O
    ITEM_Speck.ks8.keyflag = BtKfExt + BtKfDup + BtKfChg
    ITEM_Speck.ks8.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    ITEM_Speck.ks8.reserve = &H0                ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�S
    ITEM_Speck.ks9.keypos = 1                   ' �L�[�|�W�V����
    ITEM_Speck.ks9.keyleng = 1                  ' �L�[��
                                                ' �L�[�t���O
    ITEM_Speck.ks9.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    ITEM_Speck.ks9.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    ITEM_Speck.ks9.reserve = &H0                ' �\��ς�

    ITEM_Speck.ks10.keypos = 2                  ' �L�[�|�W�V����
    ITEM_Speck.ks10.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    ITEM_Speck.ks10.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    ITEM_Speck.ks10.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    ITEM_Speck.ks10.reserve = &H0               ' �\��ς�

    ITEM_Speck.ks11.keypos = 197                ' �L�[�|�W�V����
    ITEM_Speck.ks11.keyleng = 20                ' �L�[��
                                                ' �L�[�t���O
    ITEM_Speck.ks11.keyflag = BtKfExt + BtKfDup + BtKfChg
    ITEM_Speck.ks11.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    ITEM_Speck.ks11.reserve = &H0               ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�T
    ITEM_Speck.ks12.keypos = 1                  ' �L�[�|�W�V����
    ITEM_Speck.ks12.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    ITEM_Speck.ks12.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    ITEM_Speck.ks12.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    ITEM_Speck.ks12.reserve = &H0               ' �\��ς�

    ITEM_Speck.ks13.keypos = 2                  ' �L�[�|�W�V����
    ITEM_Speck.ks13.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    ITEM_Speck.ks13.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    ITEM_Speck.ks13.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    ITEM_Speck.ks13.reserve = &H0               ' �\��ς�

    ITEM_Speck.ks14.keypos = 217                ' �L�[�|�W�V����
    ITEM_Speck.ks14.keyleng = 20                ' �L�[��
                                                ' �L�[�t���O
    ITEM_Speck.ks14.keyflag = BtKfExt + BtKfDup + BtKfChg
    ITEM_Speck.ks14.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    ITEM_Speck.ks14.reserve = &H0               ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�U
    ITEM_Speck.ks15.keypos = 1                  ' �L�[�|�W�V����
    ITEM_Speck.ks15.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    ITEM_Speck.ks15.keyflag = BtKfExt + BtKfSeg + BtKfChg
    ITEM_Speck.ks15.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    ITEM_Speck.ks15.reserve = &H0               ' �\��ς�

    ITEM_Speck.ks16.keypos = 2                  ' �L�[�|�W�V����
    ITEM_Speck.ks16.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    ITEM_Speck.ks16.keyflag = BtKfExt + BtKfSeg + BtKfChg
    ITEM_Speck.ks16.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    ITEM_Speck.ks16.reserve = &H0               ' �\��ς�

    ITEM_Speck.ks17.keypos = 71                  ' �L�[�|�W�V����
    ITEM_Speck.ks17.keyleng = 2                 ' �L�[��
                                                ' �L�[�t���O
    ITEM_Speck.ks17.keyflag = BtKfExt + BtKfSeg + BtKfChg
    ITEM_Speck.ks17.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    ITEM_Speck.ks17.reserve = &H0               ' �\��ς�

    ITEM_Speck.ks18.keypos = 73                 ' �L�[�|�W�V����
    ITEM_Speck.ks18.keyleng = 2                 ' �L�[��
                                                ' �L�[�t���O
    ITEM_Speck.ks18.keyflag = BtKfExt + BtKfSeg + BtKfChg
    ITEM_Speck.ks18.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    ITEM_Speck.ks18.reserve = &H0               ' �\��ς�

    ITEM_Speck.ks19.keypos = 75                 ' �L�[�|�W�V����
    ITEM_Speck.ks19.keyleng = 2                 ' �L�[��
                                                ' �L�[�t���O
    ITEM_Speck.ks19.keyflag = BtKfExt + BtKfSeg + BtKfChg
    ITEM_Speck.ks19.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    ITEM_Speck.ks19.reserve = &H0               ' �\��ς�

    ITEM_Speck.ks20.keypos = 77                 ' �L�[�|�W�V����
    ITEM_Speck.ks20.keyleng = 2                 ' �L�[��
                                                ' �L�[�t���O
    ITEM_Speck.ks20.keyflag = BtKfExt + BtKfSeg + BtKfChg
    ITEM_Speck.ks20.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    ITEM_Speck.ks20.reserve = &H0               ' �\��ς�

    ITEM_Speck.ks21.keypos = 3                  ' �L�[�|�W�V����
    ITEM_Speck.ks21.keyleng = 20                ' �L�[��
                                                ' �L�[�t���O
    ITEM_Speck.ks21.keyflag = BtKfExt + BtKfChg
    ITEM_Speck.ks21.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    ITEM_Speck.ks21.reserve = &H0               ' �\��ς�
'-----------------------------------------------

    sts = BTRV(BtOpCreate, ITEM_POS, ITEM_Speck, Len(ITEM_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�i�ڃ}�X�^")
        Exit Function
    End If

    ITEM_Create = False

End Function

Public Function ITEM_Open(Mode As Integer) As Integer
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

    ITEM_Open = True
                                            '�i�ڃ}�X�^�t���p�X�捞��
    sts = GetIni("FILE", ITEM_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ITEM]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, ITEM_POS, ITEMREC, Len(ITEMREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = ITEM_Create()        '�i�ڃ}�X�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ITEM_POS, ITEMREC, Len(ITEMREC), ByVal FullPath, Len(FullPath), Mode)
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

    ITEM_Open = False

End Function

Public Sub Rclr_ITEMREC()
'********************************************************************
'*
'*              �i�ڃ}�X�^  ���R�[�h������
'*
'********************************************************************
Dim i       As Long


    Call UniCode_Conv(ITEMREC.JGYOBU, "")           '���ƕ��敪
    Call UniCode_Conv(ITEMREC.NAIGAI, "")           '�����O
    Call UniCode_Conv(ITEMREC.HIN_GAI, "")          '�i�ԁi�O���j
    Call UniCode_Conv(ITEMREC.HIN_NAME, "")         '�i��
    Call UniCode_Conv(ITEMREC.ST_SET_DT, "")        '�W���q�ɐݒ���t
    Call UniCode_Conv(ITEMREC.ST_SOKO, "")          '�W�����ɑq�� �q��
    Call UniCode_Conv(ITEMREC.ST_RETU, "")          '             ��
    Call UniCode_Conv(ITEMREC.ST_REN, "")           '             �A
    Call UniCode_Conv(ITEMREC.ST_DAN, "")           '             �i
    Call UniCode_Conv(ITEMREC.BEF_SOKO, "")         '�O����ɑq�� �q��
    
    Call UniCode_Conv(ITEMREC.BEF_RETU, "")         '             ��
    Call UniCode_Conv(ITEMREC.BEF_REN, "")          '             �A
    Call UniCode_Conv(ITEMREC.BEF_DAN, "")          '             �i
    Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")      '�ŏI���ɓ��t
    Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")      '�ŏI�o�ɓ��t
    Call UniCode_Conv(ITEMREC.HIN_NAI, "")          '�i�ԁi�����j
    Call UniCode_Conv(ITEMREC.BIKOU_SOKO, "")       '���l �z�X�g�q��
    Call UniCode_Conv(ITEMREC.BIKOU_TANA, "")       '���l �z�X�g�I��
    Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")      '�ŏI���ד��t
    Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")      '�ŏI�ƍ����t       2001.06.12
    
    Call UniCode_Conv(ITEMREC.BIKOU, "")            '������l
    Call UniCode_Conv(ITEMREC.JAN_CODE, "")         'Jan�R�[�h      2004.02
    Call UniCode_Conv(ITEMREC.HIN_CHANGE, "")       '�i�ԓǂݑւ�   2004.02
    Call UniCode_Conv(ITEMREC.GOODS_KBN, GOODS_ON)  '���i���L��     2004.02
    Call UniCode_Conv(ITEMREC.PACKING_NO, "")       '������       2004.02
    Call UniCode_Conv(ITEMREC.RANK, "")             '���݃����N     2004.06
    Call UniCode_Conv(ITEMREC.NEW_RANK, "")         '���݃����N     2004.06
    Call UniCode_Conv(ITEMREC.GLICS1_TANA, "")      '�O���b�N�X�I�ԂP   2005.05
    Call UniCode_Conv(ITEMREC.GLICS2_TANA, "")      '�O���b�N�X�I�ԂQ   2005.05
    Call UniCode_Conv(ITEMREC.GLICS3_TANA, "")      '�O���b�N�X�I�ԂR   2005.05
    
    Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, "")     '�Ɩ��Ǘ��@ �d���敪
    Call UniCode_Conv(ITEMREC.G_HANBAI_KBN, "")     '           �̔��敪
    Call UniCode_Conv(ITEMREC.G_SYUSHI, "")         '           ���x�P��
    Call UniCode_Conv(ITEMREC.G_KUMITATE, "")       '           �g�����i
    Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, "")   '           �W���e�������ݒ��
    Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, "")   '           �W���e�������ݒ��
                                                    '           �d������
    For i = 0 To UBound(ITEMREC.G_SHIIRE_TBL)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).CODE, "")             '����
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).TANKA_DT, "")         '�P���ݒ��
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LEAD_TIME, "")        'ذ�����
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, "")    '�O�񒍕���
    Next i
    
    Call UniCode_Conv(ITEMREC.G_SHIZAI_KBN, "")     '           ���ދ敪
    Call UniCode_Conv(ITEMREC.G_LABEL_NON, "")      '           ���ٓ\��t���v��Ȃ�
    
    Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, "")     '���i����   �i��
    Call UniCode_Conv(ITEMREC.L_BIKOU, "")          '           ���l
    Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, "")    '           ��ЃR�[�h
    Call UniCode_Conv(ITEMREC.L_KISHU1, "")         '           �@��(1)
    Call UniCode_Conv(ITEMREC.xL_KISHU2, "")        '           �@��(2) ���g�p 2006.01.24
    Call UniCode_Conv(ITEMREC.L_KISHU3, "")         '           �@��(3)(���K�p�@����l)
    Call UniCode_Conv(ITEMREC.L_PAPER, "0")         '           ��
    Call UniCode_Conv(ITEMREC.L_PLASTIC, "0")       '           �v���X�`�b�N
    Call UniCode_Conv(ITEMREC.L_LABEL, "0")         '           �K�p�@������
    Call UniCode_Conv(ITEMREC.L_MAISU, "0")         '           ��������
    Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, "")    '           �K�p�@����l(���@��i�R�j)
    Call UniCode_Conv(ITEMREC.L_SAGYO_SHIJI, "")    '           ��Ǝw��
    Call UniCode_Conv(ITEMREC.L_BIKOU3, "")         '           ���l�R
    Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, "")    '           ���ƕ��R�[�h
    Call UniCode_Conv(ITEMREC.L_TANA1, "")          '           �I��(1)
    Call UniCode_Conv(ITEMREC.L_TANA2, "")          '           �I��(2)
    
    Call UniCode_Conv(ITEMREC.S_TANTO, "")          '���P�^�S���҃R�[�h
    Call UniCode_Conv(ITEMREC.ZAIKO_F, "")          '�݌ɊǗ��ΏۗL�� 1:�Ώ� 0:�ΏۊO
    Call UniCode_Conv(ITEMREC.L_KISHU2, "")         '           �@��(2)
    Call UniCode_Conv(ITEMREC.K_KEITAI, "")         '���`��
    Call UniCode_Conv(ITEMREC.UNIT_BUHIN, "")       '�Ưĕ��i�敪       2006.07.28
    Call UniCode_Conv(ITEMREC.NAI_BUHIN, "")        '�����������i�敪   2006.07.28
    Call UniCode_Conv(ITEMREC.GAI_BUHIN, "")        '�C�O�������i�敪   2006.07.28
    Call UniCode_Conv(ITEMREC.LAST_CODE, "")        '�ŏI�d����R�[�h   2007.05.29
    Call UniCode_Conv(ITEMREC.MAKER_CODE, "")       'Ұ������           2007.06.06
    Call UniCode_Conv(ITEMREC.MAKER_NAME, "")       'Ұ������           2007.06.06
    
    Call UniCode_Conv(ITEMREC.L_MARK, "")           '�č���ϰ�          2007.11.08
    Call UniCode_Conv(ITEMREC.SAI_SU, "")           '�ː�               2008.02.14
    Call UniCode_Conv(ITEMREC.D_KEISHIKI, "")       '�`��               2008.02.14
    Call UniCode_Conv(ITEMREC.D_MATERIAL, "")       '�ގ�               2008.02.14
    Call UniCode_Conv(ITEMREC.D_THICKNESS, "")      '����ްف@����      2008.02.14
    Call UniCode_Conv(ITEMREC.D_SIZE_W, "")         '����ްٻ��ށiW�j   2008.02.14
    Call UniCode_Conv(ITEMREC.D_SIZE_D, "")         '����ްٻ��ށiD�j   2008.02.14
    Call UniCode_Conv(ITEMREC.D_SIZE_H, "")         '����ްٻ��ށiH�j   2008.02.14
    Call UniCode_Conv(ITEMREC.D_PRINT, "")          '�������^���Ȃ�    2008.02.14
    Call UniCode_Conv(ITEMREC.S_KOUSU_SET_DATE, "") '���i���@�P���ݒ�� 2008.02.14
    
    Call UniCode_Conv(ITEMREC.S_SHIZAI_SET_DATE, "") '���i���@�P���ݒ�� 2008.02.14
    Call UniCode_Conv(ITEMREC.SE_USOU_F, "")        '�A�����@�o���׸�   2008.02.14
    Call UniCode_Conv(ITEMREC.USE_TAPE_KIND, "")    '�g�p�e�[�v���     2008.02.14
    Call UniCode_Conv(ITEMREC.USE_TAPE_LNG, "")     '�g�p�e�[�v��       2008.02.14
    Call UniCode_Conv(ITEMREC.H_TANA_MAKE, "")      '�I�ԃ}�[�N         2008.04.02
    Call UniCode_Conv(ITEMREC.SE_TANKA_MEMO, "")    '�����P���@����     2008.04.15
    Call UniCode_Conv(ITEMREC.GENSANKOKU, "")       '���Y��             2008.06.11
    Call UniCode_Conv(ITEMREC.SEI_LOT, "")          '���Y���b�g         2008.07.07
    Call UniCode_Conv(ITEMREC.SEI_SYU_KON, "")      '�W������           2008.07.07
    Call UniCode_Conv(ITEMREC.SEI_TANKA_TANTO, "")  '�P���ݒ�S����     2008.07.09
    Call UniCode_Conv(ITEMREC.SHIMUKE_CODE, "")     '�d������           2008.07.09
    Call UniCode_Conv(ITEMREC.SEI_KBN, "")          '�����敪           2008.07.16
                                            '�O�H��             2008.09.19
    For i = 0 To UBound(ITEMREC.BEF_KOUTEI)
        Call UniCode_Conv(ITEMREC.BEF_KOUTEI(i).BEF_KOUTEI, "")     '�O�H�� 2008.09.19
    Next i
                                            '��ƍH��           2008.09.19
    For i = 0 To UBound(ITEMREC.MAIN_KOUTEI)
        Call UniCode_Conv(ITEMREC.MAIN_KOUTEI(i).MAIN_KOUTEI, "")   '��ƍH�� 2008.09.19
    Next i
                                            '��H��             2008.09.19
    For i = 0 To UBound(ITEMREC.AFT_KOUTEI)
        Call UniCode_Conv(ITEMREC.AFT_KOUTEI(i).AFT_KOUTEI, "")     '��H�� 2008.09.19
    Next i

    Call UniCode_Conv(ITEMREC.SE_IO_TANKA_No, "")   '�I�敪             200.09.19
    Call UniCode_Conv(ITEMREC.STAT, "")             '��ԋ敪           2009.01.21
    Call UniCode_Conv(ITEMREC.INSP_MESSAGE, "")     '�o�׌��iү����     2009.04.17
    Call UniCode_Conv(ITEMREC.S_SEIKYU_F, "")       '���i�������׸�     2009.04.28
    Call UniCode_Conv(ITEMREC.FILLER, "")           'FILLER             2009.04.28�T�C�Y�ύX
    
    Call UniCode_Conv(ITEMREC.INS_TANTO, "")        '�ǉ��@�S���ҁ@     2009.01.21
    Call UniCode_Conv(ITEMREC.Ins_DateTime, "")     '�ǉ��@����         2009.01.21
    Call UniCode_Conv(ITEMREC.UPD_TANTO, "")        '�X�V�@�S���ҁ@     2005.11.15
    Call UniCode_Conv(ITEMREC.UPD_DATETIME, "")     '�X�V�@����         2005.11.15

'-------------------------------------------------------------------------------------------
'               �O�N���A����

                                                    '��[�_�i�댯�݌Ɂj
    Call UniCode_Conv(ITEMREC.HOJYU_P, String(UBound(ITEMREC.HOJYU_P) + 1, "0"))
                                                    '�����Ϗo�א�
    Call UniCode_Conv(ITEMREC.AVE_SYUKA, String(UBound(ITEMREC.AVE_SYUKA) + 1, "0"))
                                                    '�T���v����
    Call UniCode_Conv(ITEMREC.SAMPLE_QTY, String(UBound(ITEMREC.SAMPLE_QTY) + 1, "0"))
                                                    '�ŏI�ƍ����݌ɐ�   2001.06.12
    Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, String(UBound(ITEMREC.LAST_CHK_QTY) + 1, "0"))
                                                    '������萔
    Call UniCode_Conv(ITEMREC.IRI_QTY, String(UBound(ITEMREC.IRI_QTY) + 1, "0"))
                                                    '           �W���e�������P���@9(8)V99
    Call UniCode_Conv(ITEMREC.G_ST_URITAN, String(UBound(ITEMREC.G_ST_URITAN) + 1, "0"))
                                                    '           �W���e�������P��  9(8)V99
    Call UniCode_Conv(ITEMREC.G_ST_SHITAN, String(UBound(ITEMREC.G_ST_SHITAN) + 1, "0"))

    For i = 0 To UBound(ITEMREC.G_SHIIRE_TBL)
                                                                        '�P�� 9(8)V99
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).TANKA, _
                   String(UBound(ITEMREC.G_SHIIRE_TBL(i).TANKA) + 1, "0"))
                                                                        'ۯĐ�
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LOT, _
                   String(UBound(ITEMREC.G_SHIIRE_TBL(i).LOT) + 1, "0"))
                                                                        '�O�񒍕���
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_QTY, _
                   String(UBound(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_QTY) + 1, "0"))
    Next i
                                                    '           �O���݌ɋ��z
    Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, String(UBound(ITEMREC.G_ZEN_ZAIKO_KIN) + 1, "0"))
                                                    '           ���i(1)
    Call UniCode_Conv(ITEMREC.L_URIKIN1, String(UBound(ITEMREC.L_URIKIN1) + 1, "0"))
                                                    '           ���i(2)
    Call UniCode_Conv(ITEMREC.L_URIKIN2, String(UBound(ITEMREC.L_URIKIN2) + 1, "0"))
                                                    '           ���i(3)
    Call UniCode_Conv(ITEMREC.L_URIKIN3, String(UBound(ITEMREC.L_URIKIN3) + 1, "0"))
                                                    '           ���萔
    Call UniCode_Conv(ITEMREC.L_IRI_QTY, String(UBound(ITEMREC.L_IRI_QTY) + 1, "0"))
                                                    '           �O���݌ɐ���
    Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, String(UBound(ITEMREC.G_ZEN_ZAIKO_QTY) + 1, "0"))
                                                    '           �ŏI�o�א�
    Call UniCode_Conv(ITEMREC.G_LAST_SYUKA_QTY, String(UBound(ITEMREC.G_LAST_SYUKA_QTY) + 1, "0"))
                                                    'GLICS�݌�(S2) �܈�p
    Call UniCode_Conv(ITEMREC.G_S2_ZAI_QTY, String(UBound(ITEMREC.G_S2_ZAI_QTY) + 1, "0"))
                                                    'GLICS�݌�(P2) �܈�p
    Call UniCode_Conv(ITEMREC.G_P2_ZAI_QTY, String(UBound(ITEMREC.G_P2_ZAI_QTY) + 1, "0"))
                                                    '�W���P��   2006.07.28
    Call UniCode_Conv(ITEMREC.HYO_TANKA, String(UBound(ITEMREC.HYO_TANKA) + 1, "0"))
                                                    '�ŏI�d���P��       2007.05.29
    Call UniCode_Conv(ITEMREC.LAST_TANKA, String(UBound(ITEMREC.LAST_TANKA) + 1, "0"))
                                                    '���i���@�H��       2008.02.14
    Call UniCode_Conv(ITEMREC.S_KOUSU, String(UBound(ITEMREC.S_KOUSU) + 1, "0"))
                                                    '���i���@�H������   2008.02.14
    Call UniCode_Conv(ITEMREC.S_KOUSU_GENKA, String(UBound(ITEMREC.S_KOUSU_GENKA) + 1, "0"))
                                                    '���i���@�H������   2008.02.14
    Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, String(UBound(ITEMREC.S_KOUSU_BAIKA) + 1, "0"))
                                                    '���i���@���ތ���   2008.02.14
    Call UniCode_Conv(ITEMREC.S_SHIZAI_GENKA, String(UBound(ITEMREC.S_SHIZAI_GENKA) + 1, "0"))
                                                    '���i���@���ޔ���   2008.02.14
    Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, String(UBound(ITEMREC.S_SHIZAI_BAIKA) + 1, "0"))

                                                    '�O���P�� 9(8)V99   2008.06.12
    Call UniCode_Conv(ITEMREC.S_GAISO_TANKA, String(UBound(ITEMREC.S_GAISO_TANKA) + 1, "0"))
                                                    'PPSC���H�P��9(8)   2008.06.12
    Call UniCode_Conv(ITEMREC.S_PPSC_KAKO_KOSU, String(UBound(ITEMREC.S_PPSC_KAKO_KOSU) + 1, "0"))
                                                    'BU���H�P��9(8)     2008.06.12
    Call UniCode_Conv(ITEMREC.S_BU_KAKO_KOSU, String(UBound(ITEMREC.S_BU_KAKO_KOSU) + 1, "0"))
    
                                                    '�����[�g           2008.07.07
    Call UniCode_Conv(ITEMREC.SEI_RATE, String(UBound(ITEMREC.SEI_RATE) + 1, "0"))
    
                                                    '���x���\�薇��     2008.07.19
    Call UniCode_Conv(ITEMREC.SEI_LABEL_QTY, String(UBound(ITEMREC.SEI_LABEL_QTY) + 1, "0"))

                                                    '���ތ���     �@    2008.08.20�ǉ�
    Call UniCode_Conv(ITEMREC.SEI_SZI_CNT, String(UBound(ITEMREC.SEI_SZI_CNT) + 1, "0"))
                                                    '��������           2008.08.20�ǉ�
    Call UniCode_Conv(ITEMREC.SEI_DKN_CNT, String(UBound(ITEMREC.SEI_DKN_CNT) + 1, "0"))


'-------------------------------------------------------------------------------------------
'               2009.06.02
                                                    '���i���@�H������
    Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, String(UBound(ITEMREC.BEF_S_KOUSU_BAIKA) + 1, "0"))
                                                    '���i���@���ޔ���
    Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, String(UBound(ITEMREC.BEF_S_SHIZAI_BAIKA) + 1, "0"))
                                                    '�O���P��
    Call UniCode_Conv(ITEMREC.BEF_S_GAISO_TANKA, String(UBound(ITEMREC.BEF_S_GAISO_TANKA) + 1, "0"))
                                                    'PPSC���H�P��
    Call UniCode_Conv(ITEMREC.BEF_S_PPSC_KAKO_KOSU, String(UBound(ITEMREC.BEF_S_PPSC_KAKO_KOSU) + 1, "0"))
                                                    'BU���H�P��
    Call UniCode_Conv(ITEMREC.BEF_S_BU_KAKO_KOSU, String(UBound(ITEMREC.BEF_S_BU_KAKO_KOSU) + 1, "0"))
    
    Call UniCode_Conv(ITEMREC.M_BIKOU, "")              '���Ϗ����l
    
    Call UniCode_Conv(ITEMREC.SHIYOU_NO, "")            '�d�l����
    
    Call UniCode_Conv(ITEMREC.MITSUMORI_KBN, "")        '���ς�敪
    
    Call UniCode_Conv(ITEMREC.TANKA_KIRIKAE_DT, "")    '�P���ؑ֓��t
    
    Call UniCode_Conv(ITEMREC.KIRIKAE_KBN, "")          '�ؑ֋敪
    
'               2009.06.02
'-------------------------------------------------------------------------------------------









'-------------------------------------------------------------------------------------------
'               2009.10.14

                                                
    Call UniCode_Conv(ITEMREC.GENSANKOKU, "")           '���Y��             2009.07.16
    
    Call UniCode_Conv(ITEMREC.PLUS_KOUSU, "")           '�v���X���H��       2009.09.17
    
    Call UniCode_Conv(ITEMREC.KUTI_SU, "")              '����               2009.07.16
    
    Call UniCode_Conv(ITEMREC.KONPOU_F, "")             '����敪           2009.07.16
    
    Call UniCode_Conv(ITEMREC.SAI_SU, "")               '�ː�               2009.07.16
    
    Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, "")      '�捞�ݎ����Y��     2009.07.16
    Call UniCode_Conv(ITEMREC.TORI_GEN_GENSANKOKU, "")  '                   2009.07.16
    Call UniCode_Conv(ITEMREC.TORI_SHIIRE_WORK_CENTER, "")  'TORI_SHIIRE_WORK_CENTER                   2009.07.16
    
    Call UniCode_Conv(ITEMREC.KANKYO_KBN, "")           '����ދ敪     2009.07.16
    Call UniCode_Conv(ITEMREC.KANKYO_KBN_ST, "")        '����ދ敪�K�p�J�n     2009.07.16
    Call UniCode_Conv(ITEMREC.KANKYO_KBN_SURYO, "")     '����ދ敪����     2009.07.16
    Call UniCode_Conv(ITEMREC.BIKOU20, "")              '������l     2009.07.16
    
    Call UniCode_Conv(ITEMREC.BEF_L_LABEL, "")          '
    Call UniCode_Conv(ITEMREC.BEF_1_L_PAPER, "")          '
    Call UniCode_Conv(ITEMREC.BEF_1_L_PLASTIC, "")          '
    
    Call UniCode_Conv(ITEMREC.BEF_2_L_PAPER, "")          '
    Call UniCode_Conv(ITEMREC.BEF_2_L_PLASTIC, "")          '
    
    Call UniCode_Conv(ITEMREC.BEF_3_L_PAPER, "")          '
    Call UniCode_Conv(ITEMREC.BEF_3_L_PLASTIC, "")          '
    
    Call UniCode_Conv(ITEMREC.BEF_4_L_PAPER, "")          '
    Call UniCode_Conv(ITEMREC.BEF_4_L_PLASTIC, "")          '
    
    Call UniCode_Conv(ITEMREC.BEF_LAST_L_PAPER, "")          '
    Call UniCode_Conv(ITEMREC.BEF_LAST_L_PLASTIC, "")          '
    
'               2009.10.14
'-------------------------------------------------------------------------------------------
    

'-------------------------------------------------------------------------------------------
'               2010.11.11
                                                        '���Y���󎚗L��     2010.11.10
    Call UniCode_Conv(ITEMREC.PRT_GENSANKOKU, "")
                                                        '�O�������萔 (999) 2010.11.10
    Call UniCode_Conv(ITEMREC.GAISO_IRI_QTY, String(UBound(ITEMREC.GAISO_IRI_QTY) + 1, "0"))

'               2010.11.11
'-------------------------------------------------------------------------------------------


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.01.13
    Call UniCode_Conv(ITEMREC.GOODS_OUT_F, "")          '�u���i���v��v���O�׸� "1":���O
    Call UniCode_Conv(ITEMREC.PLN_KOUSU, "")            '�u���i�����сv�p�W���H��
    
    Call UniCode_Conv(ITEMREC.G_SPTAN, "")              ' �u�������сv���ʒP�� 9(8).99  (�\�� ���g�p)
    
    Call UniCode_Conv(ITEMREC.CATE_ST_KOUTEI, "")       ' �u�������сv�O��H���i�b�j    �W��    9(3).99
    Call UniCode_Conv(ITEMREC.CATE_ST_FUKA, "")         ' �u�������сv�t���H���i�b�j    �W��    9(3).99
    Call UniCode_Conv(ITEMREC.CATE_ST_JITU1, "")        ' �u�������сv ����ƍH���i�b�j �W��    9(3).99
    Call UniCode_Conv(ITEMREC.CATE_ST_YOYU_RITU, "")    ' �u�������сv �]�T���i���j     �W��    9(3).99
    Call UniCode_Conv(ITEMREC.CATE_ST_JITU2, "")        ' �u�������сv ����ƍH���i�b�j �W��    9(3).99
    Call UniCode_Conv(ITEMREC.CATE_ST_TOTAL, "")        ' �u�������сv ��Ǝ��Ԍv�i�b�j �W��    9(3).99
    Call UniCode_Conv(ITEMREC.CATE_ST_FUN, "")          ' �u�������сv ��/�i��/�j   �W��    9(3).99
    Call UniCode_Conv(ITEMREC.CATE_ST_FUN_RATE, "")     ' �u�������сv ��ڰāi�~/���j   �W��    9(4).99
    Call UniCode_Conv(ITEMREC.CATE_ST_KOURYO, "")       ' �u�������сv �H�����i�~/�j  �W��    9(10).99
    
    
    
    
    Call UniCode_Conv(ITEMREC.CATE_AD_KOUTEI, "")       ' �u�������сv�O��H���i�b�j    ����    9(3).99
    Call UniCode_Conv(ITEMREC.CATE_AD_FUKA, "")         ' �u�������сv �t���H���i�b�j   ����    9(3).99
    Call UniCode_Conv(ITEMREC.CATE_AD_JITU1, "")        ' �u�������сv ����ƍH���i�b�j ����    9(3).99
    Call UniCode_Conv(ITEMREC.CATE_AD_YOYU_RITU, "")    ' �u�������сv �]�T���i���j     ����    9(3).99
    Call UniCode_Conv(ITEMREC.CATE_AD_JITU2, "")        ' �u�������сv ����ƍH���i�b�j ����    9(3).99
    Call UniCode_Conv(ITEMREC.CATE_AD_TOTAL, "")        ' �u�������сv ��Ǝ��Ԍv�i�b�j ����    9(3).99
    Call UniCode_Conv(ITEMREC.CATE_AD_FUN, "")          ' �u�������сv  ��/�i��/�j  ����    9(3).99
    Call UniCode_Conv(ITEMREC.CATE_AD_FUN_RATE, "")     ' �u�������сv  ��ڰāi�~/���j  ����    9(4).99
    Call UniCode_Conv(ITEMREC.CATE_AD_KOURYO, "")       ' �u�������сv  �H�����i�~/�j ����    9(10).99
    
    Call UniCode_Conv(ITEMREC.CATEGORY_CODE, "")
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   ���Ϗ�����(�i���ú�ذ�Ή�)  2011.12.12
    Call UniCode_Conv(ITEMREC.CS_TANTO_CD, "")          'CS�S������ 2011.12.12
        
    Call UniCode_Conv(ITEMREC.D_MODEL, "")              '��\�@��i�ں��ށ@ PN�A�g�ž��  2011.12.28
    Call UniCode_Conv(ITEMREC.HINMOKU, "")              '�i�ں��ށ@         PN�A�g�ž��  2011.12.28
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.01.13
    Call UniCode_Conv(ITEMREC.xK_KEITAI, "")            ' ���@���`��  2012.03.13

    Call UniCode_Conv(ITEMREC.PLN_SAGYOU_KOUSU, "")     '�u���i�����сv��ƍH�� 2012.05.07

    Call UniCode_Conv(ITEMREC.BEF_S_KOUSU, "")

    Call UniCode_Conv(ITEMREC.ZAIKO_CLR_F, "")          '�I���\�@�݌ɐ���\���e 2012.12.13 (1:��\��(0�Ƃ݂Ȃ�)�@�ȊO:�ʏ�)

    Call UniCode_Conv(ITEMREC.SHIIRE_BIKOU, "")         '�d���@���l�@       2018.04.17
    
        
    Call UniCode_Conv(ITEMREC.INS_TANTO, "")            '�ǉ��@�S���ҁ@     2009.01.21
    Call UniCode_Conv(ITEMREC.Ins_DateTime, "")         '�ǉ��@����         2009.01.21
    Call UniCode_Conv(ITEMREC.UPD_TANTO, "")            '�X�V�@�S���ҁ@     2005.11.15
    Call UniCode_Conv(ITEMREC.UPD_DATETIME, "")         '�X�V�@����         2005.11.15
    
End Sub
