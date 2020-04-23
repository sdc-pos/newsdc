Attribute VB_Name = "O_Y_NYU"
Option Explicit
'********************************************************************
'*                                                                  *
'*              ���ח\��f�[�^  �t�@�C����`                        *
'*                                                                  *
'********************************************************************
'�t�@�C���h�c
Public Const O_Y_NYU_ID$ = "O_Y_NYU"

'�y�[�W�T�C�Y
Public Const O_Y_NYU_PG_SIZ% = 2048

'�|�W�V�����E�u���b�N
Public O_Y_NYU_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type O_Y_NYUREC_Tag
    KAN_KBN(0 To 0)             As Byte     '�����敪
    DT_SYU(0 To 0)              As Byte     '�f�[�^���
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    NAIGAI(0 To 0)              As Byte     '�����O
    TEXT_NO(0 To 8)             As Byte     '�e�L�X�g��
    JGYOBA(0 To 7)              As Byte     '���Ə�
    DATA_KBN(0 To 0)            As Byte     '�f�[�^�敪
    TORI_KBN(0 To 1)            As Byte     '����敪
    ID_NO(0 To 7)               As Byte     'ID-NO
    HIN_NO(0 To 19)             As Byte     '�i�ڔԍ�
    DEN_NO(0 To 9)              As Byte     '�`�[�ԍ�
    SURYO(0 To 6)               As Byte     '�o�ɐ���
    MUKE_CODE(0 To 7)           As Byte     '�o�ɐ�
    SYUKO_SYUSI(0 To 1)         As Byte     '�o�Ɏ��x
    SYUKO_YMD(0 To 7)           As Byte     '�o�ɓ��t
    TANKA(0 To 9)               As Byte     '�P��
    ODER_NO(0 To 11)            As Byte     '�I�[�_�[�ԍ�
    ITEM_NO(0 To 4)             As Byte     '�A�C�e���ԍ�
    ODER_R_NO(0 To 4)           As Byte     '�I�[�_�[����
    KOSO_KEITAI(0 To 9)         As Byte     '���`��
    SYUKA_YMD(0 To 7)           As Byte     '�o�ד�
    TANABAN1(0 To 9)            As Byte     '�I�ԂP
    TANABAN2(0 To 9)            As Byte     '�I�ԂQ
    TANABAN3(0 To 9)            As Byte     '�I�ԂR
    MUKE_NAME(0 To 23)          As Byte     '�o�ɐ於��
    CYU_KBN(0 To 0)             As Byte     '�����敪
    CYU_KBN_NAME(0 To 9)        As Byte     '�����敪����
    ORIGIN1(0 To 9)             As Byte     '���Y���P
    ORIGIN2(0 To 9)             As Byte     '���Y���Q
    BIKOU2(0 To 39)             As Byte     '���l�Q
    HAN_KBN(0 To 0)             As Byte     '�̔��敪
    CHOKU_KBN(0 To 0)           As Byte     '�����敪
    UNIT_ID_NO(0 To 7)          As Byte     '�ƯďC��ID-NO
    ZAIKO_HIKIATE(0 To 2)       As Byte     '�݌Ɉ�������
    GOKON_KANRI_NO(0 To 8)      As Byte     '�����Ǘ��ԍ�
    JUCHU_ZAN(0 To 6)           As Byte     '�󒍎c����
    KYOKYU_KBN(0 To 0)          As Byte     '�����敪
    SHOHIN_SYUSI(0 To 1)        As Byte     '���i���[������x
    BIKOU1(0 To 39)             As Byte     '���l�P
    CHOHA_KBN(0 To 0)           As Byte     '���[�敪
    JYU_HIN_NO(0 To 19)         As Byte     '�󒍕i�ڔԍ�
    HIN_NAME(0 To 19)           As Byte     '�i��
    HIN_CHANGE_KBN(0 To 0)      As Byte     '�i�ԕύX�敪
    MODULE_EXCHANGE(0 To 0)     As Byte     '���W���[�������敪
    ZAIKO_SYUSI(0 To 1)         As Byte     '�c�݌ɂ܂Ƃߍ݌Ɏ��x�R�[�h
    NOUKI_YMD(0 To 7)           As Byte     '�w��[��
    SERVICE_KANRI_NO(0 To 8)    As Byte     '�T�[�r�X��ЊǗ��ԍ�
    KI_HIN_NO(0 To 2)           As Byte     '�@��i�ڃR�[�h
    ENVIRONMENT_KBN(0 To 0)     As Byte     '���K�i���i�敪
    KAN_DT(0 To 7)              As Byte     '�������t
    BEF_NYU_QTY(0 To 7)         As Byte     '��s���א�
    YOSAN_FROM(0 To 4)          As Byte     '�\�Z�P�ʁi���j
    YOSAN_TO(0 To 4)            As Byte     '�\�Z�P�ʁi��j
    HTANABAN(0 To 7)            As Byte     '�W���I��
    HIN_NAI(0 To 12)            As Byte     '�i�ԁi�����j
    FILLER(0 To 64)             As Byte
End Type

'�f�[�^�E�o�b�t�@
Public O_Y_NYUREC                  As O_Y_NYUREC_Tag

'�L�[��`
Type KEY0_O_Y_NYU            '�j�d�x�O
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    SYUKA_YMD(0 To 7)           As Byte     '�o�ד�
    TEXT_NO(0 To 8)             As Byte     '�e�L�X�g��
End Type

Type KEY1_O_Y_NYU            '�j�d�x�P
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    KAN_KBN(0 To 0)             As Byte     '�����敪
    NAIGAI(0 To 0)              As Byte     '�����O
    HIN_NO(0 To 19)             As Byte     '�i�ڔԍ�
    SYUKA_YMD(0 To 7)           As Byte     '�o�ד�
    TEXT_NO(0 To 8)             As Byte     '�e�L�X�g��
End Type

Type KEY2_O_Y_NYU            '�j�d�x�Q
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    SYUKA_YMD(0 To 7)           As Byte     '�o�ד�
    HIN_NO(0 To 19)             As Byte     '�i�ڔԍ�
    NAIGAI(0 To 0)              As Byte     '�����O
End Type

Type KEY3_O_Y_NYU            '�j�d�x�R
    SYUKA_YMD(0 To 7)           As Byte     '�o�ד�
End Type



'�L�[�E�f�[�^
Public K0_O_Y_NYU                 As KEY0_O_Y_NYU
Public K1_O_Y_NYU                 As KEY1_O_Y_NYU
Public K2_O_Y_NYU                 As KEY2_O_Y_NYU
Public K3_O_Y_NYU                 As KEY3_O_Y_NYU

Private Type O_Y_NYU_FSpeck
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

Private O_Y_NYU_Speck As O_Y_NYU_FSpeck

Private Function O_Y_NYU_Create() As Integer
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

    O_Y_NYU_Create = True
                                            '���ח\��f�[�^�t���p�X�捞��
    sts = GetIni("FILE", O_Y_NYU_ID, "CONV200605", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [O_Y_NYU]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    O_Y_NYU_Speck.fs.recoleng = Len(O_Y_NYUREC)     ' ���R�[�h��
    O_Y_NYU_Speck.fs.PageSize = O_Y_NYU_PG_SIZ      ' �y�[�W�T�C�Y
    O_Y_NYU_Speck.fs.idexnumb = 4                 ' �C���f�b�N�X��
    O_Y_NYU_Speck.fs.fileflag = 0                 ' �t�@�C���t���O
    O_Y_NYU_Speck.fs.reserve = &H0                ' �\��ς�
    '-------------------------------------------
                                                ' �L�[�O
    O_Y_NYU_Speck.ks0.keypos = 3                  ' �L�[�|�W�V����
    O_Y_NYU_Speck.ks0.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    O_Y_NYU_Speck.ks0.keyflag = BtKfExt + BtKfSeg
    O_Y_NYU_Speck.ks0.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    O_Y_NYU_Speck.ks0.reserve = &H0               ' �\��ς�
                                                ' �L�[�O
    O_Y_NYU_Speck.ks1.keypos = 130                ' �L�[�|�W�V����
    O_Y_NYU_Speck.ks1.keyleng = 8                 ' �L�[��
                                                ' �L�[�t���O
    O_Y_NYU_Speck.ks1.keyflag = BtKfExt + BtKfSeg
    O_Y_NYU_Speck.ks1.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    O_Y_NYU_Speck.ks1.reserve = &H0               ' �\��ς�
                                                ' �L�[�O
    O_Y_NYU_Speck.ks2.keypos = 5                  ' �L�[�|�W�V����
    O_Y_NYU_Speck.ks2.keyleng = 9                 ' �L�[��
    O_Y_NYU_Speck.ks2.keyflag = BtKfExt           ' �L�[�t���O
    O_Y_NYU_Speck.ks2.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    O_Y_NYU_Speck.ks2.reserve = &H0               ' �\��ς�
    '-------------------------------------------
                                                
                                                ' �L�[�P
    O_Y_NYU_Speck.ks3.keypos = 3                  ' �L�[�|�W�V����
    O_Y_NYU_Speck.ks3.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    O_Y_NYU_Speck.ks3.keyflag = BtKfExt + BtKfSeg + BtKfChg
    O_Y_NYU_Speck.ks3.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    O_Y_NYU_Speck.ks3.reserve = &H0               ' �\��ς�
                                                ' �L�[�P
    O_Y_NYU_Speck.ks4.keypos = 1                  ' �L�[�|�W�V����
    O_Y_NYU_Speck.ks4.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    O_Y_NYU_Speck.ks4.keyflag = BtKfExt + BtKfSeg + BtKfChg
    O_Y_NYU_Speck.ks4.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    O_Y_NYU_Speck.ks4.reserve = &H0               ' �\��ς�
                                                ' �L�[�P
    O_Y_NYU_Speck.ks5.keypos = 4                 ' �L�[�|�W�V����
    O_Y_NYU_Speck.ks5.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    O_Y_NYU_Speck.ks5.keyflag = BtKfExt + BtKfSeg + BtKfChg
    O_Y_NYU_Speck.ks5.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    O_Y_NYU_Speck.ks5.reserve = &H0               ' �\��ς�
                                                ' �L�[�P
    O_Y_NYU_Speck.ks6.keypos = 33                 ' �L�[�|�W�V����
    O_Y_NYU_Speck.ks6.keyleng = 20                ' �L�[��
                                                ' �L�[�t���O
    O_Y_NYU_Speck.ks6.keyflag = BtKfExt + BtKfSeg + BtKfChg
    O_Y_NYU_Speck.ks6.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    O_Y_NYU_Speck.ks6.reserve = &H0               ' �\��ς�
                                                ' �L�[�P
    O_Y_NYU_Speck.ks7.keypos = 130                ' �L�[�|�W�V����
    O_Y_NYU_Speck.ks7.keyleng = 8                 ' �L�[��
                                                ' �L�[�t���O
    O_Y_NYU_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfChg
    O_Y_NYU_Speck.ks7.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    O_Y_NYU_Speck.ks7.reserve = &H0               ' �\��ς�
                                                ' �L�[�P
    O_Y_NYU_Speck.ks8.keypos = 5                ' �L�[�|�W�V����
    O_Y_NYU_Speck.ks8.keyleng = 9                 ' �L�[��
                                                ' �L�[�t���O
    O_Y_NYU_Speck.ks8.keyflag = BtKfExt + BtKfChg
    O_Y_NYU_Speck.ks8.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    O_Y_NYU_Speck.ks8.reserve = &H0               ' �\��ς�
    '-------------------------------------------
                                                
                                                ' �L�[�Q
    O_Y_NYU_Speck.ks9.keypos = 3                  ' �L�[�|�W�V����
    O_Y_NYU_Speck.ks9.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    O_Y_NYU_Speck.ks9.keyflag = BtKfExt + BtKfSeg
    O_Y_NYU_Speck.ks9.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    O_Y_NYU_Speck.ks9.reserve = &H0               ' �\��ς�
                                                ' �L�[�Q
    O_Y_NYU_Speck.ks10.keypos = 130               ' �L�[�|�W�V����
    O_Y_NYU_Speck.ks10.keyleng = 8                ' �L�[��
                                                ' �L�[�t���O
    O_Y_NYU_Speck.ks10.keyflag = BtKfExt + BtKfSeg
    O_Y_NYU_Speck.ks10.keytype = Chr(BtKtString)  ' �L�[�^�C�v
    O_Y_NYU_Speck.ks10.reserve = &H0              ' �\��ς�
                                                ' �L�[�Q
    O_Y_NYU_Speck.ks11.keypos = 33                ' �L�[�|�W�V����
    O_Y_NYU_Speck.ks11.keyleng = 20               ' �L�[��
                                                ' �L�[�t���O
    O_Y_NYU_Speck.ks11.keyflag = BtKfExt + BtKfSeg
    O_Y_NYU_Speck.ks11.keytype = Chr(BtKtString)  ' �L�[�^�C�v
    O_Y_NYU_Speck.ks11.reserve = &H0              ' �\��ς�
                                                ' �L�[�Q
    O_Y_NYU_Speck.ks12.keypos = 4                 ' �L�[�|�W�V����
    O_Y_NYU_Speck.ks12.keyleng = 1                ' �L�[��
                                                ' �L�[�t���O
    O_Y_NYU_Speck.ks12.keyflag = BtKfExt + BtKfSeg
    O_Y_NYU_Speck.ks12.keytype = Chr(BtKtString)  ' �L�[�^�C�v
    O_Y_NYU_Speck.ks12.reserve = &H0              ' �\��ς�
                                                ' �L�[�Q
    O_Y_NYU_Speck.ks13.keypos = 5                 ' �L�[�|�W�V����
    O_Y_NYU_Speck.ks13.keyleng = 9                ' �L�[��
                                                ' �L�[�t���O
    O_Y_NYU_Speck.ks13.keyflag = BtKfExt
    O_Y_NYU_Speck.ks13.keytype = Chr(BtKtString)  ' �L�[�^�C�v
    O_Y_NYU_Speck.ks13.reserve = &H0              ' �\��ς�
    '-------------------------------------------
                                                
                                                ' �L�[�R
    O_Y_NYU_Speck.ks14.keypos = 130                ' �L�[�|�W�V����
    O_Y_NYU_Speck.ks14.keyleng = 8                ' �L�[��
                                                ' �L�[�t���O
    O_Y_NYU_Speck.ks14.keyflag = BtKfExt + BtKfDup
    O_Y_NYU_Speck.ks14.keytype = Chr(BtKtString)  ' �L�[�^�C�v
    O_Y_NYU_Speck.ks14.reserve = &H0              ' �\��ς�
    '-------------------------------------------
    
    sts = BTRV(BtOpCreate, O_Y_NYU_POS, O_Y_NYU_Speck, Len(O_Y_NYU_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���ח\��f�[�^")
        O_Y_NYU_Create = True
        Exit Function
    End If

    O_Y_NYU_Create = False

End Function

Function O_Y_NYU_Open(Mode As Integer) As Integer
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
    
    O_Y_NYU_Open = True
                                            '���ח\��f�[�^�t���p�X�捞��
    sts = GetIni("FILE", O_Y_NYU_ID, "CONV200605", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [O_Y_NYU]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, O_Y_NYU_POS, O_Y_NYUREC, Len(O_Y_NYUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = O_Y_NYU_Create()        '���ח\��f�[�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, O_Y_NYU_POS, O_Y_NYUREC, Len(O_Y_NYUREC), ByVal FullPath, Len(FullPath), Mode)
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
    
    O_Y_NYU_Open = False

End Function


