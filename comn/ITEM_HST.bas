Attribute VB_Name = "ITEM_HST"
Option Explicit
'********************************************************************
'*
'*              �i�ڒP���ύX����  �t�@�C����`
'*
'*          CREATE 2008.07.19
'********************************************************************
'�t�@�C���h�c
Public Const ITEM_HST_ID$ = "ITEM_HST"

'�y�[�W�T�C�Y
Public Const ITEM_HST_PG_SIZ% = 4096

'�|�W�V�����E�u���b�N
Public ITEM_HST_POS         As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************


Private Type SHIIRE_TBL_Tag         '�d������`�p��ð���
    CODE(0 To 4)            As Byte     '����
    TANKA(0 To 10)          As Byte     '�P�� 9(8)V99
    TANKA_DT(0 To 7)        As Byte     '�P���ݒ��
    LOT(0 To 7)             As Byte     'ۯĐ�
    LEAD_TIME(0 To 2)       As Byte     'ذ�����
    LAST_ORDER_DT(0 To 7)   As Byte     '�O�񒍕���
    LAST_ORDER_QTY(0 To 10)  As Byte    '�O�񒍕���
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
Type ITEM_HSTREC_Tag
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
    G_SPTAN(0 To 10)            As Byte     ' �u�������сv���ʒP�� 9(8).99
    
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
        
    FILLER(0 To 90)            As Byte      'FILLER   2011.12.12  ���ڒǉ��ɂ��T�C�Y�ύX

    INS_TANTO(0 To 4)           As Byte     '�ǉ��@�S���ҁ@     2009.01.21
    Ins_DateTime(0 To 13)       As Byte     '�ǉ��@����         2009.01.21

    UPD_TANTO(0 To 4)           As Byte     '�X�V�@�S���ҁ@     2005.11.15
    UPD_DATETIME(0 To 13)       As Byte     '�X�V�@����         2005.11.15

End Type
'�f�[�^�E�o�b�t�@
Public ITEM_HSTREC As ITEM_HSTREC_Tag

'�L�[��`

Type KEY0_ITEM_HST                  '�j�d�x�O
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
End Type

Type KEY1_ITEM_HST                  '�j�d�x�P
    TANKA_KIRIKAE_DT(0 To 7)    As Byte     '�P���ؑ֓��t       2009.06.02
End Type




'�L�[�E�f�[�^
Public K0_ITEM_HST      As KEY0_ITEM_HST

Type ITEM_HST_FSpeck
    fs      As BtFileSpeck                 ' ̧�� ��߯��\����
    ks0     As BtKeySpeck                 ' �� ��߯��\����
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck

    ks3     As BtKeySpeck
End Type

Private ITEM_HST_Speck  As ITEM_HST_FSpeck
Private Function ITEM_HST_Create() As Integer
'********************************************************************
'*
'*              �i�ڒP���ύX����  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    ITEM_HST_Create = True
                                            '�i�ڒP���ύX����   �t���p�X�捞��
    sts = GetIni("FILE", ITEM_HST_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ITEM_HST]�ǂݍ��݃G���[ ")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    ITEM_HST_Speck.fs.recoleng = Len(ITEM_HSTREC)   ' ���R�[�h��
    ITEM_HST_Speck.fs.PageSize = ITEM_HST_PG_SIZ    ' �y�[�W�T�C�Y
    ITEM_HST_Speck.fs.idexnumb = 2                  ' �C���f�b�N�X��
    ITEM_HST_Speck.fs.fileflag = 0                  ' �t�@�C���t���O
    ITEM_HST_Speck.fs.reserve = &H0                 ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�O
    ITEM_HST_Speck.ks0.keypos = 1                   ' �L�[�|�W�V����
    ITEM_HST_Speck.ks0.keyleng = 1                  ' �L�[��
                                                    ' �L�[�t���O
    ITEM_HST_Speck.ks0.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    ITEM_HST_Speck.ks0.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    ITEM_HST_Speck.ks0.reserve = &H0                ' �\��ς�
                                                
    ITEM_HST_Speck.ks1.keypos = 2                   ' �L�[�|�W�V����
    ITEM_HST_Speck.ks1.keyleng = 1                  ' �L�[��
                                                    ' �L�[�t���O
    ITEM_HST_Speck.ks1.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    ITEM_HST_Speck.ks1.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    ITEM_HST_Speck.ks1.reserve = &H0                ' �\��ς�
                                                
    ITEM_HST_Speck.ks2.keypos = 3                   ' �L�[�|�W�V����
    ITEM_HST_Speck.ks2.keyleng = 20                 ' �L�[��
                                                    ' �L�[�t���O
    ITEM_HST_Speck.ks2.keyflag = BtKfExt + BtKfDup + BtKfChg
    ITEM_HST_Speck.ks2.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    ITEM_HST_Speck.ks2.reserve = &H0                ' �\��ς�
'-----------------------------------------------

'-----------------------------------------------
                                                ' �L�[�P

    ITEM_HST_Speck.ks3.keypos = 2627                   ' �L�[�|�W�V����
    ITEM_HST_Speck.ks3.keyleng = 8                 ' �L�[��
                                                    ' �L�[�t���O
    ITEM_HST_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfChg
    ITEM_HST_Speck.ks3.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    ITEM_HST_Speck.ks3.reserve = &H0                ' �\��ς�
'-----------------------------------------------



    sts = BTRV(BtOpCreate, ITEM_HST_POS, ITEM_HST_Speck, Len(ITEM_HST_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�i�ڒP���ύX����")
        Exit Function
    End If

    ITEM_HST_Create = False

End Function

Public Function ITEM_HST_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �i�ڒP���ύX����  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    ITEM_HST_Open = True
    
    sts = GetIni("FILE", ITEM_HST_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ITEM_HST]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, ITEM_HST_POS, ITEM_HSTREC, Len(ITEM_HSTREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = ITEM_HST_Create()        '�i�ڒP���ύX���� �쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ITEM_HST_POS, ITEM_HSTREC, Len(ITEM_HSTREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�i�ڒP���ύX����")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�i�ڒP���ύX����")
                Exit Function
        End Select
    Loop

    ITEM_HST_Open = False

End Function


