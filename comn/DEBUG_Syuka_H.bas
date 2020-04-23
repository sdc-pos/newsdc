Attribute VB_Name = "DEBUG_SYU_H"
Option Explicit
'********************************************************************
'*
'*              �o�ח\��iνāj�f�[�^  �t�@�C����`
'*              ���o�b��p    2006.12.02
'*
'********************************************************************
'�t�@�C���h�c
Public Const DEBUG_SYU_H_ID$ = "DEBUG_SYU_H"

'�y�[�W�T�C�Y
Public Const DEBUG_SYU_H_PG_SIZ% = 4096

'�|�W�V�����E�u���b�N
Public DEBUG_SYU_H_POS    As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type DEBUG_SYU_HREC_Tag
    ID_NO(0 To 11)              As Byte     'ID_NO(�L�� :�`�[�� 7 ��+�ǔ� 2��)
    SYUKA_NO(0 To 2)            As Byte     '��
    SYUKA_YMD(0 To 7)           As Byte     '�o�ח\���
    OKURISAKI(0 To 39)          As Byte     '����於
    URIDEN(0 To 0)              As Byte     '���`
    DEN_NO(0 To 6)              As Byte     '�`�[��
    SEQ_NO(0 To 0)              As Byte     '�ǔ�
    HIN_NO(0 To 19)             As Byte     '�i��
    SURYO(0 To 6)               As Byte     '�o�א���
    ODER_NO(0 To 9)             As Byte     '�����ԍ�
    MUKE_CODE(0 To 7)           As Byte     '���Ӑ�R�[�h
    MUKE_NAME(0 To 39)          As Byte     '���Ӑ於��
    BIKOU(0 To 99)              As Byte     '���l
    UNSOU_KAISHA(0 To 39)       As Byte     '�^����Ж�
    
    INS_NOW(0 To 13)            As Byte     '�捞�ݓ���
    PRINT_NOW(0 To 13)          As Byte     '�o�����و������

    DATA_CNT(0 To 4)            As Byte     '�ް�������
    
    OKURI_NO(0 To 19)           As Byte     '�����
    KENPIN_NOW(0 To 13)         As Byte     '���i����
    KENPIN_TANTO_CODE(0 To 4)   As Byte     '���i�S���Һ���

    xKUTI_SU(0 To 1)            As Byte     '����   '2007.02.01 ���g�p
    
    KYOSEI_END(0 To 0)          As Byte     '���������׸�

    CANCEL_F(0 To 0)            As Byte     '��ݾ��׸�
    
    INPUT_BIKOU(0 To 59)        As Byte     '���͔��l
    
    INS_BIN(0 To 1)             As Byte     '��
    
    KUTI_SU(0 To 3)             As Byte     '����   '2007.02.01 �����ύX�ɂ��V��
    
    
    
    JGYOBU(0 To 0)              As Byte     '���ƕ�     2007.03.14
    NAIGAI(0 To 0)              As Byte     '�����O     2007.03.14
    
    SYU_NO(0 To 11)             As Byte     '�o�ɕ\��   2007.03.14
    J_SURYO(0 To 6)             As Byte     '�o�Ɏ��ѐ� 2007.03.14
    
    
    COL_OKURISAKI_CD(0 To 19)   As Byte     '�W�񑗂��CD   2007.07.07
    OKURISAKI_CD(0 To 8)        As Byte     '�����CD       2007.07.07
    
    JYUSHO(0 To 159)            As Byte     '�Z��       2009.11.19
    
    TEL_NO(0 To 19)             As Byte     '�d�b�ԍ�   2010.01.21
    YUBIN_NO(0 To 7)            As Byte     '�X�֔ԍ�   2010.01.21
        
    JURYO(0 To 5)               As Byte     '�d��       2010.01.21
    SAI_SU(0 To 5)              As Byte     '�ː�       2010.01.21
    
    
    OKURI_NO_SEQ(0 To 2)        As Byte     '����󇂁@�}�ԁ@2010.01.21
    
    
    KONPOU_F(0 To 0)            As Byte     '����敪       2010.01.18
    KUTI_SU_TAN(0 To 5)         As Byte     '����(�P��)     2010.01.21
    SAI_SU_TAN(0 To 5)          As Byte     '�ː�(�P��)     2010.01.21
    
    OKURI_NO_SEQ_TO(0 To 2)     As Byte     '����󇂁@�}�ԁ@2010.01.21
    
    
    SAI_SU_TAN_SAV(0 To 5)      As Byte     '�ː�(�P��:�C���s��)    2010.11.01
    SAI_SU_CALC(0 To 5)         As Byte     '�ː��v�Z�l(����P��)   2010.11.01
    
    
    KUTI_SU_CALC(0 To 5)        As Byte     '�����v�Z�l(����P��)   2010.11.9
    
    SEK_KEN_NO(0 To 5)          As Byte     '���Ǉ��@�@�@���Ǘ���(��)   2011.04.30
    SEK_HIN_NO(0 To 5)          As Byte     '�i�Ǉ��@�@�@���Ǘ���(��)   2011.04.30
    
    SEK_SHOGO_TANTO(0 To 9)     As Byte     '�����ް��ƍ��S��       2011.05.02
    SEK_SHOGO_DATETIME(0 To 13) As Byte     '�����ް��ƍ�����       2011.05.02
    
    
    CNT_BARA_SU(0 To 6)         As Byte     '���i���с@�o��     2012.10.02
    CNT_HAKO_SU(0 To 6)         As Byte     '���i���с@��       2012.10.02
    GAISO_IRI_QTY(0 To 7)       As Byte     '�O�����萔         2012.10.02
    
    
    Y_HIN_CHK_CNT(0 To 5)       As Byte     '�i�ԓǍ��݉�     2012.10.02
    J_HIN_CHK_CNT(0 To 5)       As Byte     '�i�ԓǍ��ݍς݉� 2012.10.02
    
    KEN_HINBAN(0 To 19)         As Byte     '���i���i��         2012.10.02
    
    FILLER(0 To 159)            As Byte     'FILLER             2012.10.02 (157


    INS_TANTO(0 To 9)           As Byte     '�ǉ��@�S����       2011.05.06
    Ins_DateTime(0 To 13)       As Byte     '�ǉ��@����         2011.05.06
    UPD_TANTO(0 To 9)           As Byte     '�X�V�@�S����       2011.05.06
    UPD_DATETIME(0 To 13)       As Byte     '�X�V�@����         2011.05.06



End Type

'�f�[�^�E�o�b�t�@
Public DEBUG_SYU_HREC               As DEBUG_SYU_HREC_Tag

'�L�[��`
Type KEY0_DEBUG_SYU_H            '�j�d�x�O
    DEN_NO(0 To 6)              As Byte     '�`�[��
    SEQ_NO(0 To 0)              As Byte     '�ǔ�
End Type

Type KEY1_DEBUG_SYU_H            '�j�d�x�P
    PRINT_NOW(0 To 13)          As Byte     '�o�����و������
    INS_NOW(0 To 13)            As Byte     '�捞�ݓ���
    DATA_CNT(0 To 4)            As Byte     '�ް�������
End Type

Type KEY2_DEBUG_SYU_H            '�j�d�x�Q
    OKURI_NO(0 To 19)           As Byte     '�����
End Type

Type KEY3_DEBUG_SYU_H            '�j�d�x�R
    SYUKA_YMD(0 To 7)           As Byte     '�o�ח\���
End Type

Type KEY4_DEBUG_SYU_H            '�j�d�x�S
    ID_NO(0 To 11)              As Byte     'ID_NO(�L�� :�`�[�� 7 ��+�ǔ� 2��)
End Type

Type KEY5_DEBUG_SYU_H            '�j�d�x�T      2007.03.14
    INS_BIN(0 To 1)             As Byte     '��
    SYUKA_YMD(0 To 7)           As Byte     '�o�ח\���
    JGYOBU(0 To 0)              As Byte     '���ƕ�
    NAIGAI(0 To 0)              As Byte     '�����O
    HIN_NO(0 To 19)             As Byte     '�i��
End Type

Type KEY6_DEBUG_SYU_H            '�j�d�x�U      2007.03.14
    SYU_NO(0 To 11)             As Byte     '�o�ɕ\��
End Type

Type KEY7_DEBUG_SYU_H            '�j�d�x�V      2007.03.14
    SYUKA_YMD(0 To 7)           As Byte     '�o�ח\���
    INS_BIN(0 To 1)             As Byte     '��
    SYU_NO(0 To 11)             As Byte     '�o�ɕ\��
End Type


Type KEY8_DEBUG_SYU_H            '�j�d�x�W      2011.04.30
    SEK_KEN_NO(0 To 5)          As Byte     '���Ǉ��@�@�@���Ǘ���(��)   2011.04.30
    SEK_HIN_NO(0 To 5)          As Byte     '�i�Ǉ��@�@�@���Ǘ���(��)   2011.04.30
End Type



'�L�[�E�f�[�^
Public K0_DEBUG_SYU_H               As KEY0_DEBUG_SYU_H
Public K1_DEBUG_SYU_H               As KEY1_DEBUG_SYU_H
Public K2_DEBUG_SYU_H               As KEY2_DEBUG_SYU_H
Public K3_DEBUG_SYU_H               As KEY3_DEBUG_SYU_H
Public K4_DEBUG_SYU_H               As KEY4_DEBUG_SYU_H
Public K5_DEBUG_SYU_H               As KEY5_DEBUG_SYU_H
Public K6_DEBUG_SYU_H               As KEY6_DEBUG_SYU_H
Public K7_DEBUG_SYU_H               As KEY7_DEBUG_SYU_H

Public K8_DEBUG_SYU_H               As KEY8_DEBUG_SYU_H     '2011.04.30

Type DEBUG_SYU_H_FSpeck
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

    ks17    As BtKeySpeck                   ' �� ��߯��\����    2011.04.30
    ks18    As BtKeySpeck                   ' �� ��߯��\����    2011.04.30



End Type

Private DEBUG_SYU_H_Speck As DEBUG_SYU_H_FSpeck

Private Function DEBUG_SYU_H_Create() As Integer

'********************************************************************
'*
'*              �o�ח\��(νĲҰ��)�f�[�^  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    DEBUG_SYU_H_Create = True
                                            '�o�ח\��f�[�^�t���p�X�捞��
    sts = GetIni(App.EXEName, DEBUG_SYU_H_ID, App.EXEName, c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [DEBUG_SYU_H]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    DEBUG_SYU_H_Speck.fs.recoleng = Len(DEBUG_SYU_HREC)     ' ���R�[�h��
    DEBUG_SYU_H_Speck.fs.PageSize = DEBUG_SYU_H_PG_SIZ      ' �y�[�W�T�C�Y
    DEBUG_SYU_H_Speck.fs.idexnumb = 9                   ' �C���f�b�N�X��
    DEBUG_SYU_H_Speck.fs.fileflag = 0                   ' �t�@�C���t���O
    DEBUG_SYU_H_Speck.fs.reserve = &H0                  ' �\��ς�
'---------------------------------------------------' �L�[�O
    DEBUG_SYU_H_Speck.ks0.keypos = 65                   ' �L�[�|�W�V����
    DEBUG_SYU_H_Speck.ks0.keyleng = 7                   ' �L�[��
                                                    ' �L�[�t���O
    DEBUG_SYU_H_Speck.ks0.keyflag = BtKfExt + BtKfSeg + BtKfDup
    DEBUG_SYU_H_Speck.ks0.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    DEBUG_SYU_H_Speck.ks0.reserve = &H0                 ' �\��ς�
    
    DEBUG_SYU_H_Speck.ks1.keypos = 72                   ' �L�[�|�W�V����
    DEBUG_SYU_H_Speck.ks1.keyleng = 1                   ' �L�[��
                                                    ' �L�[�t���O
    DEBUG_SYU_H_Speck.ks1.keyflag = BtKfExt + BtKfDup
    DEBUG_SYU_H_Speck.ks1.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    DEBUG_SYU_H_Speck.ks1.reserve = &H0                 ' �\��ς�
'---------------------------------------------------' �L�[�O
    
'---------------------------------------------------' �L�[�P
    DEBUG_SYU_H_Speck.ks2.keypos = 312                  ' �L�[�|�W�V����
    DEBUG_SYU_H_Speck.ks2.keyleng = 14                  ' �L�[��
                                                    ' �L�[�t���O
    DEBUG_SYU_H_Speck.ks2.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    DEBUG_SYU_H_Speck.ks2.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    DEBUG_SYU_H_Speck.ks2.reserve = &H0                 ' �\��ς�

    
    DEBUG_SYU_H_Speck.ks3.keypos = 298                  ' �L�[�|�W�V����
    DEBUG_SYU_H_Speck.ks3.keyleng = 14                  ' �L�[��
                                                    ' �L�[�t���O
    DEBUG_SYU_H_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    DEBUG_SYU_H_Speck.ks3.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    DEBUG_SYU_H_Speck.ks3.reserve = &H0                 ' �\��ς�
    
    DEBUG_SYU_H_Speck.ks4.keypos = 326                  ' �L�[�|�W�V����
    DEBUG_SYU_H_Speck.ks4.keyleng = 5                   ' �L�[��
                                                    ' �L�[�t���O
    DEBUG_SYU_H_Speck.ks4.keyflag = BtKfExt + BtKfDup + BtKfChg
    DEBUG_SYU_H_Speck.ks4.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    DEBUG_SYU_H_Speck.ks4.reserve = &H0                 ' �\��ς�
'---------------------------------------------------' �L�[�P
'---------------------------------------------------' �L�[�Q
    DEBUG_SYU_H_Speck.ks5.keypos = 331                  ' �L�[�|�W�V����
    DEBUG_SYU_H_Speck.ks5.keyleng = 20                  ' �L�[��
                                                    ' �L�[�t���O
    DEBUG_SYU_H_Speck.ks5.keyflag = BtKfExt + BtKfDup + BtKfChg
    DEBUG_SYU_H_Speck.ks5.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    DEBUG_SYU_H_Speck.ks5.reserve = &H0                 ' �\��ς�
'---------------------------------------------------' �L�[�Q
'---------------------------------------------------' �L�[�R
    DEBUG_SYU_H_Speck.ks6.keypos = 16                    ' �L�[�|�W�V����
    DEBUG_SYU_H_Speck.ks6.keyleng = 8                   ' �L�[��
                                                    ' �L�[�t���O
    DEBUG_SYU_H_Speck.ks6.keyflag = BtKfExt + BtKfDup + BtKfChg
    DEBUG_SYU_H_Speck.ks6.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    DEBUG_SYU_H_Speck.ks6.reserve = &H0                 ' �\��ς�
'---------------------------------------------------' �L�[�R
    
'---------------------------------------------------' �L�[�S
    DEBUG_SYU_H_Speck.ks7.keypos = 1                    ' �L�[�|�W�V����
    DEBUG_SYU_H_Speck.ks7.keyleng = 12                  ' �L�[��
                                                    ' �L�[�t���O
    DEBUG_SYU_H_Speck.ks7.keyflag = BtKfExt + BtKfChg + BtKfDup
    DEBUG_SYU_H_Speck.ks7.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    DEBUG_SYU_H_Speck.ks7.reserve = &H0                 ' �\��ς�
'---------------------------------------------------' �L�[�S

'---------------------------------------------------' �L�[�T
    DEBUG_SYU_H_Speck.ks8.keypos = 434                  ' �L�[�|�W�V����
    DEBUG_SYU_H_Speck.ks8.keyleng = 2                   ' �L�[��
                                                    ' �L�[�t���O
    DEBUG_SYU_H_Speck.ks8.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    DEBUG_SYU_H_Speck.ks8.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    DEBUG_SYU_H_Speck.ks8.reserve = &H0                 ' �\��ς�

    DEBUG_SYU_H_Speck.ks9.keypos = 16                   ' �L�[�|�W�V����
    DEBUG_SYU_H_Speck.ks9.keyleng = 8                   ' �L�[��
                                                    ' �L�[�t���O
    DEBUG_SYU_H_Speck.ks9.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    DEBUG_SYU_H_Speck.ks9.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    DEBUG_SYU_H_Speck.ks9.reserve = &H0                 ' �\��ς�


    DEBUG_SYU_H_Speck.ks10.keypos = 440                  ' �L�[�|�W�V����
    DEBUG_SYU_H_Speck.ks10.keyleng = 1                   ' �L�[��
                                                    ' �L�[�t���O
    DEBUG_SYU_H_Speck.ks10.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    DEBUG_SYU_H_Speck.ks10.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    DEBUG_SYU_H_Speck.ks10.reserve = &H0                 ' �\��ς�

    DEBUG_SYU_H_Speck.ks11.keypos = 441                 ' �L�[�|�W�V����
    DEBUG_SYU_H_Speck.ks11.keyleng = 1                  ' �L�[��
                                                    ' �L�[�t���O
    DEBUG_SYU_H_Speck.ks11.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    DEBUG_SYU_H_Speck.ks11.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    DEBUG_SYU_H_Speck.ks11.reserve = &H0                ' �\��ς�


    DEBUG_SYU_H_Speck.ks12.keypos = 73                  ' �L�[�|�W�V����
    DEBUG_SYU_H_Speck.ks12.keyleng = 20                 ' �L�[��
                                                    ' �L�[�t���O
    DEBUG_SYU_H_Speck.ks12.keyflag = BtKfExt + BtKfChg + BtKfDup
    DEBUG_SYU_H_Speck.ks12.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    DEBUG_SYU_H_Speck.ks12.reserve = &H0

'---------------------------------------------------' �L�[�T

'---------------------------------------------------' �L�[�U
    DEBUG_SYU_H_Speck.ks13.keypos = 442                 ' �L�[�|�W�V����
    DEBUG_SYU_H_Speck.ks13.keyleng = 12                 ' �L�[��
                                                    ' �L�[�t���O
    DEBUG_SYU_H_Speck.ks13.keyflag = BtKfExt + BtKfChg + BtKfDup
    DEBUG_SYU_H_Speck.ks13.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    DEBUG_SYU_H_Speck.ks13.reserve = &H0

'---------------------------------------------------' �L�[�U
    
'---------------------------------------------------' �L�[�V
    DEBUG_SYU_H_Speck.ks14.keypos = 16                  ' �L�[�|�W�V����
    DEBUG_SYU_H_Speck.ks14.keyleng = 8                  ' �L�[��
                                                    ' �L�[�t���O
    DEBUG_SYU_H_Speck.ks14.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    DEBUG_SYU_H_Speck.ks14.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    DEBUG_SYU_H_Speck.ks14.reserve = &H0

    DEBUG_SYU_H_Speck.ks15.keypos = 434                 ' �L�[�|�W�V����
    DEBUG_SYU_H_Speck.ks15.keyleng = 2                  ' �L�[��
                                                    ' �L�[�t���O
    DEBUG_SYU_H_Speck.ks15.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    DEBUG_SYU_H_Speck.ks15.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    DEBUG_SYU_H_Speck.ks15.reserve = &H0

    DEBUG_SYU_H_Speck.ks16.keypos = 442                 ' �L�[�|�W�V����
    DEBUG_SYU_H_Speck.ks16.keyleng = 12                 ' �L�[��
                                                    ' �L�[�t���O
    DEBUG_SYU_H_Speck.ks16.keyflag = BtKfExt + BtKfChg + BtKfDup
    DEBUG_SYU_H_Speck.ks16.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    DEBUG_SYU_H_Speck.ks16.reserve = &H0

'---------------------------------------------------' �L�[�V
    
    
'---------------------------------------------------' �L�[�V
    DEBUG_SYU_H_Speck.ks14.keypos = 16                  ' �L�[�|�W�V����
    DEBUG_SYU_H_Speck.ks14.keyleng = 8                  ' �L�[��
                                                    ' �L�[�t���O
    DEBUG_SYU_H_Speck.ks14.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    DEBUG_SYU_H_Speck.ks14.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    DEBUG_SYU_H_Speck.ks14.reserve = &H0

    DEBUG_SYU_H_Speck.ks15.keypos = 434                 ' �L�[�|�W�V����
    DEBUG_SYU_H_Speck.ks15.keyleng = 2                  ' �L�[��
                                                    ' �L�[�t���O
    DEBUG_SYU_H_Speck.ks15.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    DEBUG_SYU_H_Speck.ks15.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    DEBUG_SYU_H_Speck.ks15.reserve = &H0

    DEBUG_SYU_H_Speck.ks16.keypos = 442                 ' �L�[�|�W�V����
    DEBUG_SYU_H_Speck.ks16.keyleng = 12                 ' �L�[��
                                                    ' �L�[�t���O
    DEBUG_SYU_H_Speck.ks16.keyflag = BtKfExt + BtKfChg + BtKfDup
    DEBUG_SYU_H_Speck.ks16.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    DEBUG_SYU_H_Speck.ks16.reserve = &H0

'---------------------------------------------------' �L�[�W    2011.04.30
    DEBUG_SYU_H_Speck.ks17.keypos = 727                 ' �L�[�|�W�V����
    DEBUG_SYU_H_Speck.ks17.keyleng = 6                  ' �L�[��
                                                    ' �L�[�t���O
    DEBUG_SYU_H_Speck.ks17.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    DEBUG_SYU_H_Speck.ks17.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    DEBUG_SYU_H_Speck.ks17.reserve = &H0

    DEBUG_SYU_H_Speck.ks18.keypos = 733                 ' �L�[�|�W�V����
    DEBUG_SYU_H_Speck.ks18.keyleng = 6                  ' �L�[��
                                                    ' �L�[�t���O
    DEBUG_SYU_H_Speck.ks18.keyflag = BtKfExt + BtKfChg + BtKfDup
    DEBUG_SYU_H_Speck.ks18.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    DEBUG_SYU_H_Speck.ks18.reserve = &H0
'---------------------------------------------------' �L�[�W    2011.04.30
    
    
    sts = BTRV(BtOpCreate, DEBUG_SYU_H_POS, DEBUG_SYU_H_Speck, Len(DEBUG_SYU_H_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�o�ח\��(νĲҰ��)�f�[�^")
        Exit Function
    End If

    DEBUG_SYU_H_Create = False

End Function

Function DEBUG_SYU_H_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �o�ח\��(νĲҰ��)�f�[�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    DEBUG_SYU_H_Open = True
                                            '�o�ח\��f�[�^�t���p�X�捞��
    sts = GetIni(App.EXEName, DEBUG_SYU_H_ID, App.EXEName, c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [DEBUG_SYU_H]�ǂݍ��݃G���[ ")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, DEBUG_SYU_H_POS, DEBUG_SYU_HREC, Len(DEBUG_SYU_HREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = DEBUG_SYU_H_Create()        '�o�ח\��f�[�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, DEBUG_SYU_H_POS, DEBUG_SYU_HREC, Len(DEBUG_SYU_HREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�o�ח\��(νĲҰ��)�f�[�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�o�ח\��(νĲҰ��)�f�[�^")
                Exit Function
        End Select
    Loop
    DEBUG_SYU_H_Open = False
End Function
