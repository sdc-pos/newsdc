Attribute VB_Name = "O_Y_SYU"
Option Explicit
'********************************************************************
'*
'*              �o�ח\��f�[�^  �t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const O_Y_SYU_ID$ = "O_Y_SYU"

'�y�[�W�T�C�Y
Public Const O_Y_SYU_PG_SIZ% = 2048

'�|�W�V�����E�u���b�N
Public O_Y_SYU_POS    As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type O_Y_SYUREC_Tag
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
Public O_Y_SYUREC                 As O_Y_SYUREC_Tag

'�L�[��`
Type KEY0_O_Y_SYU            '�j�d�x�O
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
'    KEY_CYU_KBN(0 To 0)         As Byte     '�����敪
    KEY_ID_NO(0 To 7)           As Byte     'ID-NO
End Type

Type KEY1_O_Y_SYU            '�j�d�x�P
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    KAN_KBN(0 To 0)             As Byte     '�����敪
    KEY_MUKE_CODE(0 To 7)       As Byte     '���Ӑ�R�[�h
    KEY_SS_CODE(0 To 7)         As Byte     '������R�[�h
    KEY_CYU_KBN(0 To 0)         As Byte     '�����敪
    KEY_ID_NO(0 To 7)           As Byte     'ID-NO
    NAIGAI(0 To 0)              As Byte     '�����O
    KEY_HIN_NO(0 To 19)         As Byte     '�i�ڔԍ�
End Type

Type KEY2_O_Y_SYU            '�j�d�x�Q
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    KEY_CYU_KBN(0 To 0)         As Byte     '�����敪
    KEY_MUKE_CODE(0 To 7)       As Byte     '���Ӑ�R�[�h
    KEY_SS_CODE(0 To 7)         As Byte     '������R�[�h
End Type

Type KEY3_O_Y_SYU            '�j�d�x�R
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    KEY_CYU_KBN(0 To 0)         As Byte     '�����敪
    KEY_MUKE_CODE(0 To 7)       As Byte     '���Ӑ�R�[�h
    KEY_SS_CODE(0 To 7)         As Byte     '������R�[�h
    NAIGAI(0 To 0)              As Byte     '�����O
    KEY_HIN_NO(0 To 19)         As Byte     '�i�ڔԍ�
    KEY_ID_NO(0 To 7)           As Byte     'ID-NO
End Type

Type KEY4_O_Y_SYU            '�j�d�x�S
    WEL_ID(0 To 2)              As Byte     '�g�p�q�@ID
    PRG_ID(0 To 7)              As Byte     '�g�p���v���O����
End Type

Type KEY5_O_Y_SYU            '�j�d�x�T
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    KEY_CYU_KBN(0 To 0)         As Byte     '�����敪
    KEY_MUKE_CODE(0 To 7)       As Byte     '���Ӑ�R�[�h
    KEY_SS_CODE(0 To 7)         As Byte     '������R�[�h
    HTANABAN(0 To 7)            As Byte     '�z�X�g�I��         '2004.06.08
    KEY_SYUKA_YMD(0 To 7)       As Byte     '�o�ד��t
    KEY_HIN_NO(0 To 19)         As Byte     '�i�ڔԍ�           '2004.06.29
End Type

Type KEY6_O_Y_SYU            '�j�d�x�U
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    KEY_CYU_KBN(0 To 0)         As Byte     '�����敪
    HTANABAN(0 To 7)            As Byte     '�z�X�g�I��
    NAIGAI(0 To 0)              As Byte     '�����O
    KEY_HIN_NO(0 To 19)         As Byte     '�i�ڔԍ�
End Type

Type KEY7_O_Y_SYU            '�j�d�x�V
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    NAIGAI(0 To 0)              As Byte     '�����O
    KEY_HIN_NO(0 To 19)         As Byte     '�i�ڔԍ�
    KEY_SYUKA_YMD(0 To 7)       As Byte     '�o�ד��t
End Type

'�L�[�E�f�[�^
Public K0_O_Y_SYU                 As KEY0_O_Y_SYU
Public K1_O_Y_SYU                 As KEY1_O_Y_SYU
Public K2_O_Y_SYU                 As KEY2_O_Y_SYU
Public K3_O_Y_SYU                 As KEY3_O_Y_SYU
Public K4_O_Y_SYU                 As KEY4_O_Y_SYU
Public K5_O_Y_SYU                 As KEY5_O_Y_SYU
Public K6_O_Y_SYU                 As KEY6_O_Y_SYU
Public K7_O_Y_SYU                 As KEY7_O_Y_SYU

Type O_Y_SYU_FSpeck
    fs      As BtFileSpeck                  ' ̧�� ��߯��\����
    ks0     As BtKeySpeck                   ' �� ��߯��\����
'    ks1     As BtKeySpeck                   ' �� ��߯��\����
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
    ks39    As BtKeySpeck                   ' �� ��߯��\����
End Type

Private O_Y_SYU_Speck As O_Y_SYU_FSpeck

Private Function O_Y_SYU_Create() As Integer
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

    O_Y_SYU_Create = True
                                            '�o�ח\��f�[�^�t���p�X�捞��
    sts = GetIni("FILE", O_Y_SYU_ID, "CONV200605", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [O_Y_SYU]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    O_Y_SYU_Speck.fs.recoleng = Len(O_Y_SYUREC)         ' ���R�[�h��
    O_Y_SYU_Speck.fs.PageSize = O_Y_SYU_PG_SIZ          ' �y�[�W�T�C�Y
    O_Y_SYU_Speck.fs.idexnumb = 8                     ' �C���f�b�N�X��
    O_Y_SYU_Speck.fs.fileflag = 0                     ' �t�@�C���t���O
    O_Y_SYU_Speck.fs.reserve = &H0                    ' �\��ς�
'---------------------------------------------------' �L�[�O
    O_Y_SYU_Speck.ks0.keypos = 14                     ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks0.keyleng = 1                     ' �L�[��
    O_Y_SYU_Speck.ks0.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    O_Y_SYU_Speck.ks0.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    O_Y_SYU_Speck.ks0.reserve = &H0                   ' �\��ς�
    
'    O_Y_SYU_Speck.ks1.keypos = 15                     ' �L�[�|�W�V����
'    O_Y_SYU_Speck.ks1.keyleng = 1                     ' �L�[��
'    O_Y_SYU_Speck.ks1.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
'    O_Y_SYU_Speck.ks1.keytype = Chr(BtKtString)       ' �L�[�^�C�v
'    O_Y_SYU_Speck.ks1.reserve = &H0                   ' �\��ς�
    
    O_Y_SYU_Speck.ks2.keypos = 16                     ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks2.keyleng = 8                     ' �L�[��
    O_Y_SYU_Speck.ks2.keyflag = BtKfExt               ' �L�[�t���O
    O_Y_SYU_Speck.ks2.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    O_Y_SYU_Speck.ks2.reserve = &H0                   ' �\��ς�

'---------------------------------------------------' �L�[�P
    O_Y_SYU_Speck.ks3.keypos = 14                     ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks3.keyleng = 1                     ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks3.keyflag = BtKfExt + BtKfChg + BtKfSeg
    O_Y_SYU_Speck.ks3.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    O_Y_SYU_Speck.ks3.reserve = &H0                   ' �\��ς�
    
    O_Y_SYU_Speck.ks4.keypos = 12                     ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks4.keyleng = 1                     ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks4.keyflag = BtKfExt + BtKfChg + BtKfSeg
    O_Y_SYU_Speck.ks4.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    O_Y_SYU_Speck.ks4.reserve = &H0                   ' �\��ς�
    
    O_Y_SYU_Speck.ks5.keypos = 45                     ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks5.keyleng = 8                     ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks5.keyflag = BtKfExt + BtKfChg + BtKfSeg
    O_Y_SYU_Speck.ks5.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    O_Y_SYU_Speck.ks5.reserve = &H0                   ' �\��ς�
    
    O_Y_SYU_Speck.ks6.keypos = 53                     ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks6.keyleng = 8                     ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks6.keyflag = BtKfExt + BtKfChg + BtKfSeg
    O_Y_SYU_Speck.ks6.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    O_Y_SYU_Speck.ks6.reserve = &H0                   ' �\��ς�
    
    O_Y_SYU_Speck.ks7.keypos = 15                     ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks7.keyleng = 1                     ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks7.keyflag = BtKfExt + BtKfChg + BtKfSeg
    O_Y_SYU_Speck.ks7.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    O_Y_SYU_Speck.ks7.reserve = &H0                   ' �\��ς�
    
    O_Y_SYU_Speck.ks8.keypos = 16                     ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks8.keyleng = 8                     ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks8.keyflag = BtKfExt + BtKfChg + BtKfSeg
    O_Y_SYU_Speck.ks8.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    O_Y_SYU_Speck.ks8.reserve = &H0                   ' �\��ς�
    
    O_Y_SYU_Speck.ks9.keypos = 24                     ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks9.keyleng = 1                     ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks9.keyflag = BtKfExt + BtKfChg + BtKfSeg
    O_Y_SYU_Speck.ks9.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    O_Y_SYU_Speck.ks9.reserve = &H0                   ' �\��ς�
    
    O_Y_SYU_Speck.ks10.keypos = 25                     ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks10.keyleng = 20                     ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks10.keyflag = BtKfExt + BtKfChg
    O_Y_SYU_Speck.ks10.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    O_Y_SYU_Speck.ks10.reserve = &H0                   ' �\��ς�
'---------------------------------------------------' �L�[�Q
    O_Y_SYU_Speck.ks11.keypos = 14                    ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks11.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks11.keyflag = BtKfExt + BtKfDup + BtKfSeg
    O_Y_SYU_Speck.ks11.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_Y_SYU_Speck.ks11.reserve = &H0                  ' �\��ς�
    
    O_Y_SYU_Speck.ks12.keypos = 15                    ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks12.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks12.keyflag = BtKfExt + BtKfDup + BtKfSeg
    O_Y_SYU_Speck.ks12.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_Y_SYU_Speck.ks12.reserve = &H0                  ' �\��ς�
    
    O_Y_SYU_Speck.ks13.keypos = 45                    ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks13.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks13.keyflag = BtKfExt + BtKfDup + BtKfSeg
    O_Y_SYU_Speck.ks13.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_Y_SYU_Speck.ks13.reserve = &H0                  ' �\��ς�
    
    O_Y_SYU_Speck.ks14.keypos = 53                    ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks14.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks14.keyflag = BtKfExt + BtKfDup
    O_Y_SYU_Speck.ks14.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_Y_SYU_Speck.ks14.reserve = &H0                  ' �\��ς�
'---------------------------------------------------' �L�[�R
    O_Y_SYU_Speck.ks15.keypos = 14                    ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks15.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks15.keyflag = BtKfExt + BtKfSeg
    O_Y_SYU_Speck.ks15.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_Y_SYU_Speck.ks15.reserve = &H0                  ' �\��ς�
    
    O_Y_SYU_Speck.ks16.keypos = 15                    ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks16.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks16.keyflag = BtKfExt + BtKfSeg
    O_Y_SYU_Speck.ks16.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_Y_SYU_Speck.ks16.reserve = &H0                  ' �\��ς�
    
    O_Y_SYU_Speck.ks17.keypos = 45                    ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks17.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks17.keyflag = BtKfExt + BtKfSeg
    O_Y_SYU_Speck.ks17.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_Y_SYU_Speck.ks17.reserve = &H0                  ' �\��ς�
                                                    
    O_Y_SYU_Speck.ks18.keypos = 53                    ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks18.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks18.keyflag = BtKfExt + BtKfSeg
    O_Y_SYU_Speck.ks18.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_Y_SYU_Speck.ks18.reserve = &H0                  ' �\��ς�
    
    O_Y_SYU_Speck.ks19.keypos = 24                    ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks19.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks19.keyflag = BtKfExt + BtKfSeg
    O_Y_SYU_Speck.ks19.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_Y_SYU_Speck.ks19.reserve = &H0                  ' �\��ς�
    
    O_Y_SYU_Speck.ks20.keypos = 25                    ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks20.keyleng = 20                   ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks20.keyflag = BtKfExt + BtKfSeg
    O_Y_SYU_Speck.ks20.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_Y_SYU_Speck.ks20.reserve = &H0                  ' �\��ς�
    
    O_Y_SYU_Speck.ks21.keypos = 16                    ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks21.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks21.keyflag = BtKfExt
    O_Y_SYU_Speck.ks21.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_Y_SYU_Speck.ks21.reserve = &H0                  ' �\��ς�
'---------------------------------------------------' �L�[�S
    O_Y_SYU_Speck.ks22.keypos = 1                     ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks22.keyleng = 3                    ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks22.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    O_Y_SYU_Speck.ks22.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_Y_SYU_Speck.ks22.reserve = &H0                  ' �\��ς�
    
    O_Y_SYU_Speck.ks23.keypos = 4                     ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks23.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks23.keyflag = BtKfExt + BtKfChg + BtKfDup
    O_Y_SYU_Speck.ks23.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_Y_SYU_Speck.ks23.reserve = &H0                  ' �\��ς�
'---------------------------------------------------' �L�[�T
    O_Y_SYU_Speck.ks24.keypos = 14                    ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks24.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks24.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    O_Y_SYU_Speck.ks24.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_Y_SYU_Speck.ks24.reserve = &H0                  ' �\��ς�
    
    O_Y_SYU_Speck.ks25.keypos = 15                    ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks25.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks25.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    O_Y_SYU_Speck.ks25.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_Y_SYU_Speck.ks25.reserve = &H0                  ' �\��ς�
    
    O_Y_SYU_Speck.ks26.keypos = 45                    ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks26.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks26.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    O_Y_SYU_Speck.ks26.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_Y_SYU_Speck.ks26.reserve = &H0                  ' �\��ς�
    
    O_Y_SYU_Speck.ks27.keypos = 53                    ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks27.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks27.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    O_Y_SYU_Speck.ks27.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_Y_SYU_Speck.ks27.reserve = &H0                  ' �\��ς�
    
    
    O_Y_SYU_Speck.ks28.keypos = 391                   ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks28.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks28.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    O_Y_SYU_Speck.ks28.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_Y_SYU_Speck.ks28.reserve = &H0                  ' �\��ς�
    
    O_Y_SYU_Speck.ks29.keypos = 61                    ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks29.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks29.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    O_Y_SYU_Speck.ks29.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_Y_SYU_Speck.ks29.reserve = &H0                  ' �\��ς�
    
    O_Y_SYU_Speck.ks30.keypos = 25                    ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks30.keyleng = 20                   ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks30.keyflag = BtKfExt + BtKfDup + BtKfChg
    O_Y_SYU_Speck.ks30.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_Y_SYU_Speck.ks30.reserve = &H0                  ' �\��ς�
'---------------------------------------------------' �L�[�U
    O_Y_SYU_Speck.ks31.keypos = 14                    ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks31.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks31.keyflag = BtKfExt + BtKfDup + BtKfSeg
    O_Y_SYU_Speck.ks31.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_Y_SYU_Speck.ks31.reserve = &H0                  ' �\��ς�
    
    O_Y_SYU_Speck.ks32.keypos = 15                    ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks32.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks32.keyflag = BtKfExt + BtKfDup + BtKfSeg
    O_Y_SYU_Speck.ks32.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_Y_SYU_Speck.ks32.reserve = &H0                  ' �\��ς�
    
    O_Y_SYU_Speck.ks33.keypos = 391                   ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks33.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks33.keyflag = BtKfExt + BtKfDup + BtKfSeg
    O_Y_SYU_Speck.ks33.keytype = Chr(BtKtString)      ' �L�[�^�C�v3
    O_Y_SYU_Speck.ks33.reserve = &H0                  ' �\��ς�
    
    O_Y_SYU_Speck.ks34.keypos = 24                    ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks34.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks34.keyflag = BtKfExt + BtKfDup + BtKfSeg
    O_Y_SYU_Speck.ks34.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_Y_SYU_Speck.ks34.reserve = &H0                  ' �\��ς�
    
    O_Y_SYU_Speck.ks35.keypos = 25                    ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks35.keyleng = 20                   ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks35.keyflag = BtKfExt + BtKfDup
    O_Y_SYU_Speck.ks35.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_Y_SYU_Speck.ks35.reserve = &H0                  ' �\��ς�
'---------------------------------------------------' �L�[�V
    O_Y_SYU_Speck.ks36.keypos = 14                    ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks36.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks36.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    O_Y_SYU_Speck.ks36.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_Y_SYU_Speck.ks36.reserve = &H0                  ' �\��ς�
    
    O_Y_SYU_Speck.ks37.keypos = 24                    ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks37.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks37.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    O_Y_SYU_Speck.ks37.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_Y_SYU_Speck.ks37.reserve = &H0                  ' �\��ς�
    
    O_Y_SYU_Speck.ks38.keypos = 25                    ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks38.keyleng = 20                   ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks38.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    O_Y_SYU_Speck.ks38.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_Y_SYU_Speck.ks38.reserve = &H0                  ' �\��ς�
    
    O_Y_SYU_Speck.ks39.keypos = 61                    ' �L�[�|�W�V����
    O_Y_SYU_Speck.ks39.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    O_Y_SYU_Speck.ks39.keyflag = BtKfExt + BtKfDup + BtKfChg
    O_Y_SYU_Speck.ks39.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_Y_SYU_Speck.ks39.reserve = &H0                  ' �\��ς�
    
    sts = BTRV(BtOpCreate, O_Y_SYU_POS, O_Y_SYU_Speck, Len(O_Y_SYU_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�o�ח\��f�[�^")
        Exit Function
    End If

    O_Y_SYU_Create = False

End Function

Function O_Y_SYU_Open(Mode As Integer) As Integer
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
    
    O_Y_SYU_Open = True
                                            '�o�ח\��f�[�^�t���p�X�捞��
    sts = GetIni("FILE", O_Y_SYU_ID, "CONV200605", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [O_Y_SYU]�ǂݍ��݃G���[ ")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, O_Y_SYU_POS, O_Y_SYUREC, Len(O_Y_SYUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = O_Y_SYU_Create()        '�o�ח\��f�[�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, O_Y_SYU_POS, O_Y_SYUREC, Len(O_Y_SYUREC), ByVal FullPath, Len(FullPath), Mode)
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
    O_Y_SYU_Open = False
End Function
