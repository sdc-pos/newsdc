Attribute VB_Name = "L_ITEM"
Option Explicit
'********************************************************************
'*
'*              �i�ڃ}�X�^  �t�@�C����`
'*
'*          CREATE 2004.02.19
'********************************************************************
'�t�@�C���h�c
Public Const L_ITEM_ID$ = "L_ITEM"

'�y�[�W�T�C�Y
Public Const L_ITEM_PG_SIZ% = 4096

'�|�W�V�����E�u���b�N
Public L_ITEM_POS         As POSBLK
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


'���R�[�h��`
Type L_ITEMREC_Tag
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    '2005.11.15 �����ύX 13---> 20
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
    '2005.11.15 �����ύX 25---> 40
    HIN_NAME(0 To 39)   As Byte     '�i��
    ST_SET_DT(0 To 7)   As Byte     '�W���q�ɐݒ���t
    ST_SOKO(0 To 1)     As Byte     '�W�����ɑq�� �q��
    ST_RETU(0 To 1)     As Byte     '             ��
    ST_REN(0 To 1)      As Byte     '             �A
    ST_DAN(0 To 1)      As Byte     '             �i
    BEF_SOKO(0 To 1)    As Byte     '�O����ɑq�� �q��
    BEF_RETU(0 To 1)    As Byte     '             ��
    BEF_REN(0 To 1)     As Byte     '             �A
    BEF_DAN(0 To 1)     As Byte     '             �i
    LAST_NYU_DT(0 To 7) As Byte     '�ŏI���ɓ��t
    LAST_SYU_DT(0 To 7) As Byte     '�ŏI�o�ɓ��t
    '2005.11.15 �����ύX 13---> 20
    HIN_NAI(0 To 19)    As Byte     '�i�ԁi�����j
    BIKOU_SOKO(0 To 1)  As Byte     '���l �z�X�g�q��
    BIKOU_TANA(0 To 7)  As Byte     '���l �z�X�g�I��
    '���g�p�̂��ߍ폜 2005.11.15 SIZAI_CD(0 To 4)    As Byte     '���ރR�[�h
    HOJYU_P(0 To 7)     As Byte     '��[�_�i�댯�݌Ɂj
    AVE_SYUKA(0 To 7)   As Byte     '�����Ϗo�א�
    SAMPLE_QTY(0 To 0)  As Byte     '�T���v����
    LAST_INP_DT(0 To 7) As Byte     '�ŏI���ד��t
'*------------------------------------------ 2001.02.15 �ǉ� ��
    '���g�p�̂��ߍ폜 2005.11.15 LOCK_F(0 To 0)      As Byte     '�r���t���O
    '���g�p�̂��ߍ폜 2005.11.15 WEL_ID(0 To 2)      As Byte     '�g�p�q�@ID
    '���g�p�̂��ߍ폜 2005.11.15 PRG_ID(0 To 7)      As Byte     '�g�p���v���O����
'*------------------------------------------ 2001.02.15 �ǉ� ��
    LAST_CHK_DT(0 To 7) As Byte     '�ŏI�ƍ����t2001.06.12
    LAST_CHK_QTY(0 To 7) As Byte    '�ŏI�ƍ����݌ɐ�2001.06.12
    '���g�p�̂��ߍ폜 2005.11.15 MOTO_JIGYOBU(0 To 0) As Byte    '�������ƕ�     '���g�p2004.02
    BIKOU(0 To 14)      As Byte     '������l
    IRI_QTY(0 To 7)     As Byte     '������萔
    
    '2005.11.15 �����ύX 13---> 20
    JAN_CODE(0 To 19)   As Byte     'Jan�R�[�h      2004.02
    '2005.11.15 �����ύX 13---> 20
    HIN_CHANGE(0 To 19) As Byte     '�i�ԓǂݑւ�   2004.02
    GOODS_KBN(0 To 0)   As Byte     '���i���L��     2004.02
    PACKING_NO(0 To 3)  As Byte     '������       2004.02
    RANK(0 To 2)        As Byte     '���݃����N     2004.06
    NEW_RANK(0 To 2)    As Byte     '���݃����N     2004.06
    GLICS1_TANA(0 To 9) As Byte     '�O���b�N�X�I�ԂP   2005.05
    GLICS2_TANA(0 To 9) As Byte     '�O���b�N�X�I�ԂQ   2005.05
    GLICS3_TANA(0 To 9) As Byte     '�O���b�N�X�I�ԂR   2005.05
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
    xL_KISHU2(0 To 39)           As Byte     '           �@��(2)
    L_KISHU3(0 To 149)          As Byte     '           �@��(3)
    L_PAPER(0 To 0)             As Byte     '           ��
    L_PLASTIC(0 To 0)           As Byte     '           �v���X�`�b�N
    L_URIKIN1(0 To 9)           As Byte     '           ���i(1)
    L_URIKIN2(0 To 9)           As Byte     '           ���i(2)
    L_URIKIN3(0 To 9)           As Byte     '           ���i(3)
    L_LABEL(0 To 0)             As Byte     '           �K�p�@������
    L_MAISU(0 To 0)             As Byte     '           ��������
    L_KISHU_BIKOU(0 To 449)     As Byte     '           �K�p�@����l
    L_SAGYO_SHIJI(0 To 449)     As Byte     '           ��Ǝw��
    L_BIKOU3(0 To 4)            As Byte     '           ���l�R
    L_JGYOBU_CODE(0 To 1)       As Byte     '           ���ƕ��R�[�h
    L_IRI_QTY(0 To 7)           As Byte     '           ���萔
    L_TANA1(0 To 19)            As Byte     '           �I��(1)
    L_TANA2(0 To 19)            As Byte     '           �I��(2)
'*------------------------------------------ 2005.11.15 �ǉ�(���i���ٍ���) ��
    S_TANTO(0 To 1)             As Byte     '���P�^�S���҃R�[�h
    ZAIKO_F(0 To 0)             As Byte     '�݌ɊǗ��ΏۗL�� 0:�Ώ� 1:�ΏۊO
    
    
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
    
    
    SAI_SU(0 To 3)              As Byte     '�ː�               2008.02.14
    
    D_KEISHIKI(0 To 19)         As Byte     '�`��               2008.02.14
    D_MATERIAL(0 To 19)         As Byte     '�ގ�               2008.02.14
    D_THICKNESS(0 To 9)         As Byte     '����ްف@����      2008.02.14
    
    
    D_SIZE_W(0 To 7)            As Byte     '����ްٻ��ށiW�j   2008.02.14
    D_SIZE_D(0 To 7)            As Byte     '����ްٻ��ށiD�j   2008.02.14
    D_SIZE_H(0 To 7)            As Byte     '����ްٻ��ށiH�j   2008.02.14
        
    D_PRINT(0 To 3)            As Byte      '�������^���Ȃ�   2008.02.14
            
        
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
    
    
    GENSANKOKU(0 To 9)          As Byte     '���Y��             2008.06.11
    
    S_GAISO_TANKA(0 To 10)      As Byte     '�O���P�� 9(8)V99   2008.06.12
    S_PPSC_KAKO_KOSU(0 To 7)    As Byte     'PPSC���H�P��9(8)   2008.06.12
    S_BU_KAKO_KOSU(0 To 7)      As Byte     'BU���H�P��9(8)   2008.06.12
    
    FILLER(0 To 865)           As Byte     'FILLER
    
    
    
    

    UPD_TANTO(0 To 4)           As Byte     '�X�V�@�S���ҁ@ 2005.11.15
    UPD_DATETIME(0 To 13)       As Byte     '�X�V�@����     2005.11.15

End Type
'�f�[�^�E�o�b�t�@
Public L_ITEMREC As L_ITEMREC_Tag

'�L�[��`

Type KEY0_L_ITEM            '�j�d�x�O
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    '2005.11.15 �����ύX 13---> 20
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
End Type

Type KEY1_L_ITEM            '�j�d�x�P
    LAST_SYU_DT(0 To 7) As Byte     '�ŏI�o�ɓ��t
End Type

Type KEY2_L_ITEM            '�j�d�x�Q
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    '2005.11.15 �����ύX 13---> 20
    HIN_NAI(0 To 19)    As Byte     '�i�ԁi�����j
End Type

Type KEY3_L_ITEM            '�j�d�x�R
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    ST_SET_DT(0 To 7)   As Byte     '�W���q�ɐݒ���t
End Type


Type KEY4_L_ITEM            '�j�d�x�S 2004.02
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    '2005.11.15 �����ύX 13---> 20
    JAN_CODE(0 To 19)   As Byte     'Jan�R�[�h
End Type

Type KEY5_L_ITEM            '�j�d�x�T 2004.02
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    '2005.11.15 �����ύX 13---> 20
    HIN_CHANGE(0 To 19) As Byte     '�i�ԓǂݑւ�
End Type

Type KEY6_L_ITEM            '�j�d�x�U 2004.02
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
Public K0_L_ITEM As KEY0_L_ITEM
Public K1_L_ITEM As KEY1_L_ITEM
Public K2_L_ITEM As KEY2_L_ITEM
Public K3_L_ITEM As KEY3_L_ITEM
Public K4_L_ITEM As KEY4_L_ITEM
Public K5_L_ITEM As KEY5_L_ITEM
Public K6_L_ITEM As KEY6_L_ITEM

Type L_ITEM_FSpeck
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

Private L_ITEM_Speck  As L_ITEM_FSpeck
Private Function L_ITEM_Create() As Integer
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

    L_ITEM_Create = True
                                            '�i�ڃ}�X�^�t���p�X�捞��
    sts = GetIni("FILE", L_ITEM_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [L_ITEM]�ǂݍ��݃G���[ ")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    L_ITEM_Speck.fs.recoleng = Len(L_ITEMREC)   ' ���R�[�h��
    L_ITEM_Speck.fs.PageSize = ITEM_PG_SIZ      ' �y�[�W�T�C�Y
    L_ITEM_Speck.fs.idexnumb = 7                  ' �C���f�b�N�X��
    L_ITEM_Speck.fs.fileflag = 0                  ' �t�@�C���t���O
    L_ITEM_Speck.fs.reserve = &H0                 ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�O
    L_ITEM_Speck.ks0.keypos = 1                   ' �L�[�|�W�V����
    L_ITEM_Speck.ks0.keyleng = 1                  ' �L�[��
    L_ITEM_Speck.ks0.keyflag = BtKfExt + BtKfSeg  ' �L�[�t���O
    L_ITEM_Speck.ks0.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    L_ITEM_Speck.ks0.reserve = &H0                ' �\��ς�
                                                
    L_ITEM_Speck.ks1.keypos = 2                   ' �L�[�|�W�V����
    L_ITEM_Speck.ks1.keyleng = 1                  ' �L�[��
    L_ITEM_Speck.ks1.keyflag = BtKfExt + BtKfSeg  ' �L�[�t���O
    L_ITEM_Speck.ks1.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    L_ITEM_Speck.ks1.reserve = &H0                ' �\��ς�
                                                
    L_ITEM_Speck.ks2.keypos = 3                   ' �L�[�|�W�V����
    L_ITEM_Speck.ks2.keyleng = 20                 ' �L�[��
    L_ITEM_Speck.ks2.keyflag = BtKfExt            ' �L�[�t���O
    L_ITEM_Speck.ks2.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    L_ITEM_Speck.ks2.reserve = &H0                ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�P
    L_ITEM_Speck.ks3.keypos = 95                  ' �L�[�|�W�V����
    L_ITEM_Speck.ks3.keyleng = 8                  ' �L�[��
                                                ' �L�[�t���O
    L_ITEM_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfChg
    L_ITEM_Speck.ks3.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    L_ITEM_Speck.ks3.reserve = &H0                ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�Q
    L_ITEM_Speck.ks4.keypos = 1                   ' �L�[�|�W�V����
    L_ITEM_Speck.ks4.keyleng = 1                  ' �L�[��
                                                ' �L�[�t���O
    L_ITEM_Speck.ks4.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    L_ITEM_Speck.ks4.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    L_ITEM_Speck.ks4.reserve = &H0                ' �\��ς�
                                                    
    L_ITEM_Speck.ks5.keypos = 2                   ' �L�[�|�W�V����
    L_ITEM_Speck.ks5.keyleng = 1                  ' �L�[��
                                                ' �L�[�t���O
    L_ITEM_Speck.ks5.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    L_ITEM_Speck.ks5.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    L_ITEM_Speck.ks5.reserve = &H0                ' �\��ς�
                                                
    L_ITEM_Speck.ks6.keypos = 103                 ' �L�[�|�W�V����
    L_ITEM_Speck.ks6.keyleng = 20                 ' �L�[��
                                                ' �L�[�t���O
    L_ITEM_Speck.ks6.keyflag = BtKfExt + BtKfDup + BtKfChg
    L_ITEM_Speck.ks6.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    L_ITEM_Speck.ks6.reserve = &H0                ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�R
    L_ITEM_Speck.ks7.keypos = 1                   ' �L�[�|�W�V����
    L_ITEM_Speck.ks7.keyleng = 1                  ' �L�[��
                                                ' �L�[�t���O
    L_ITEM_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    L_ITEM_Speck.ks7.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    L_ITEM_Speck.ks7.reserve = &H0                ' �\��ς�
                                                
    L_ITEM_Speck.ks8.keypos = 63                  ' �L�[�|�W�V����
    L_ITEM_Speck.ks8.keyleng = 8                  ' �L�[��
                                                ' �L�[�t���O
    L_ITEM_Speck.ks8.keyflag = BtKfExt + BtKfDup + BtKfChg
    L_ITEM_Speck.ks8.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    L_ITEM_Speck.ks8.reserve = &H0                ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�S
    L_ITEM_Speck.ks9.keypos = 1                   ' �L�[�|�W�V����
    L_ITEM_Speck.ks9.keyleng = 1                  ' �L�[��
                                                ' �L�[�t���O
    L_ITEM_Speck.ks9.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    L_ITEM_Speck.ks9.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    L_ITEM_Speck.ks9.reserve = &H0                ' �\��ς�
                                                
    L_ITEM_Speck.ks10.keypos = 2                  ' �L�[�|�W�V����
    L_ITEM_Speck.ks10.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    L_ITEM_Speck.ks10.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    L_ITEM_Speck.ks10.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    L_ITEM_Speck.ks10.reserve = &H0               ' �\��ς�
                                                
    L_ITEM_Speck.ks11.keypos = 197                ' �L�[�|�W�V����
    L_ITEM_Speck.ks11.keyleng = 20                ' �L�[��
                                                ' �L�[�t���O
    L_ITEM_Speck.ks11.keyflag = BtKfExt + BtKfDup + BtKfChg
    L_ITEM_Speck.ks11.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    L_ITEM_Speck.ks11.reserve = &H0               ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�T
    L_ITEM_Speck.ks12.keypos = 1                  ' �L�[�|�W�V����
    L_ITEM_Speck.ks12.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    L_ITEM_Speck.ks12.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    L_ITEM_Speck.ks12.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    L_ITEM_Speck.ks12.reserve = &H0               ' �\��ς�
                                                
    L_ITEM_Speck.ks13.keypos = 2                  ' �L�[�|�W�V����
    L_ITEM_Speck.ks13.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    L_ITEM_Speck.ks13.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    L_ITEM_Speck.ks13.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    L_ITEM_Speck.ks13.reserve = &H0               ' �\��ς�
                                                
    L_ITEM_Speck.ks14.keypos = 217                ' �L�[�|�W�V����
    L_ITEM_Speck.ks14.keyleng = 20                ' �L�[��
                                                ' �L�[�t���O
    L_ITEM_Speck.ks14.keyflag = BtKfExt + BtKfDup + BtKfChg
    L_ITEM_Speck.ks14.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    L_ITEM_Speck.ks14.reserve = &H0               ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�U
    L_ITEM_Speck.ks15.keypos = 1                  ' �L�[�|�W�V����
    L_ITEM_Speck.ks15.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    L_ITEM_Speck.ks15.keyflag = BtKfExt + BtKfSeg + BtKfChg
    L_ITEM_Speck.ks15.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    L_ITEM_Speck.ks15.reserve = &H0               ' �\��ς�

    L_ITEM_Speck.ks16.keypos = 2                  ' �L�[�|�W�V����
    L_ITEM_Speck.ks16.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    L_ITEM_Speck.ks16.keyflag = BtKfExt + BtKfSeg + BtKfChg
    L_ITEM_Speck.ks16.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    L_ITEM_Speck.ks16.reserve = &H0               ' �\��ς�

    L_ITEM_Speck.ks17.keypos = 71                  ' �L�[�|�W�V����
    L_ITEM_Speck.ks17.keyleng = 2                 ' �L�[��
                                                ' �L�[�t���O
    L_ITEM_Speck.ks17.keyflag = BtKfExt + BtKfSeg + BtKfChg
    L_ITEM_Speck.ks17.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    L_ITEM_Speck.ks17.reserve = &H0               ' �\��ς�

    L_ITEM_Speck.ks18.keypos = 73                 ' �L�[�|�W�V����
    L_ITEM_Speck.ks18.keyleng = 2                 ' �L�[��
                                                ' �L�[�t���O
    L_ITEM_Speck.ks18.keyflag = BtKfExt + BtKfSeg + BtKfChg
    L_ITEM_Speck.ks18.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    L_ITEM_Speck.ks18.reserve = &H0               ' �\��ς�

    L_ITEM_Speck.ks19.keypos = 75                 ' �L�[�|�W�V����
    L_ITEM_Speck.ks19.keyleng = 2                 ' �L�[��
                                                ' �L�[�t���O
    L_ITEM_Speck.ks19.keyflag = BtKfExt + BtKfSeg + BtKfChg
    L_ITEM_Speck.ks19.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    L_ITEM_Speck.ks19.reserve = &H0               ' �\��ς�

    L_ITEM_Speck.ks20.keypos = 77                 ' �L�[�|�W�V����
    L_ITEM_Speck.ks20.keyleng = 2                 ' �L�[��
                                                ' �L�[�t���O
    L_ITEM_Speck.ks20.keyflag = BtKfExt + BtKfSeg + BtKfChg
    L_ITEM_Speck.ks20.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    L_ITEM_Speck.ks20.reserve = &H0               ' �\��ς�

    L_ITEM_Speck.ks21.keypos = 3                  ' �L�[�|�W�V����
    L_ITEM_Speck.ks21.keyleng = 20                ' �L�[��
                                                ' �L�[�t���O
    L_ITEM_Speck.ks21.keyflag = BtKfExt + BtKfChg
    L_ITEM_Speck.ks21.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    L_ITEM_Speck.ks21.reserve = &H0               ' �\��ς�
'-----------------------------------------------

    sts = BTRV(BtOpCreate, L_ITEM_POS, L_ITEM_Speck, Len(L_ITEM_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "L�i�ڃ}�X�^")
        Exit Function
    End If

    L_ITEM_Create = False

End Function

Public Function L_ITEM_Open(Mode As Integer) As Integer
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
    
    L_ITEM_Open = True
                                            '�i�ڃ}�X�^�t���p�X�捞��
    sts = GetIni("FILE", L_ITEM_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [L_ITEM]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, L_ITEM_POS, L_ITEMREC, Len(L_ITEMREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = L_ITEM_Create()        '�i�ڃ}�X�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, L_ITEM_POS, L_ITEMREC, Len(L_ITEMREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "L_�i�ڃ}�X�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "L_�i�ڃ}�X�^")
                Exit Function
        End Select
    Loop

    L_ITEM_Open = False

End Function


