Attribute VB_Name = "ITEM_O"
Option Explicit
'********************************************************************
'*
'*              ��㎖�@���ϗp�i�ڃ}�X�^  �t�@�C����`
'*
'*          CREATE 2016.05.24
'********************************************************************
'�t�@�C���h�c
Public Const ITEM_O_ID$ = "ITEM_O"

'�y�[�W�T�C�Y
Public Const ITEM_O_PG_SIZ% = 4096

'�|�W�V�����E�u���b�N
Public ITEM_O_POS               As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************



'���R�[�h��`
Type ITEM_O_REC_Tag
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    NAIGAI(0 To 0)              As Byte     '�����O
    HIN_GAI(0 To 19)            As Byte     '�i��(�O��)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.11.07
    SEQ_NO(0 To 2)              As Byte     'SEQ_NO
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.11.07
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.09.28
    KO_JGYOBU(0 To 0)           As Byte     '�q�@���ƕ��敪
    KO_NAIGAI(0 To 0)           As Byte     '�q�@�����O
    KO_HIN_GAI(0 To 19)         As Byte     '�q�@�i��(�O��)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.09.28
    
    
    NAKANISHI_TANI(0 To 3)      As Byte     '�����H���@�P��
    NAKANISHI_KIN(0 To 10)      As Byte     '�����H���@���z

    SHOHIN_TANI(0 To 3)         As Byte     '���i���H���@�P��
    SHOHIN_KIN(0 To 10)         As Byte     '���i���H���@���z

    PF_KAKOU_TANI(0 To 3)       As Byte     'PF���H�@�P��
    PF_KAKOU_KIN(0 To 10)       As Byte     'PF���H�@���z

    PE_KAKOU_TANI(0 To 3)       As Byte     'PE���H�@�P��
    PE_KAKOU_KIN(0 To 10)       As Byte     'PE���H�@���z

    PE_SHIZAI_TANI(0 To 3)      As Byte     'PF���ށ@�P��
    PE_SHIZAI_KIN(0 To 10)      As Byte     'PF���ށ@���z

    HINBAN_LABEL_TANI(0 To 3)   As Byte     '�i�ԕ\�����ف@�P��
    HINBAN_LABEL_KIN(0 To 10)   As Byte     '�i�ԕ\�����ف@���z

    KOUJI_SETSU_TANI(0 To 3)    As Byte     '�ݒu�H���������@�P��
    KOUJI_SETSU_KIN(0 To 10)    As Byte     '�ݒu�H���������@���z

    KONPOU_TANI(0 To 3)         As Byte     '����ށ@�P��
    KONPOU_KIN(0 To 10)         As Byte     '����ށ@���z

    FUKU_SHIZAI_TANI(0 To 3)    As Byte     '�����ށ@�P��
    FUKU_SHIZAI_KIN(0 To 10)    As Byte     '�����ށ@���z

    KONPOU_ASSY_TANI(0 To 3)    As Byte     '����ASSY�@�P��
    KONPOU_ASSY_KIN(0 To 10)    As Byte     '����ASSY�@���z

    KANRI_TANI(0 To 3)          As Byte     '�Ǘ���@�P��
    KANRI_KIN(0 To 10)          As Byte     '�Ǘ���@���z
    
    GOUKEI_KIN(0 To 10)         As Byte     '���v���z
    
        
    
    INPUT_TANTO_CODE(0 To 4)    As Byte     '���͒S���Һ���
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.09.28
    BUZAI_TANTO_NAME(0 To 19)   As Byte     '���ޒS���Җ�
    T_HIN_NAME(0 To 39)         As Byte     '��o�i��
    TANI(0 To 3)                As Byte     '�P��
    T_TANKA(0 To 10)            As Byte     '��o�P��
    T_KINGAKU(0 To 10)          As Byte     '��o���z
    NAKANISHI_T_KIN(0 To 10)    As Byte     '�����H���@��o���z
    SHOHIN_T_KIN(0 To 10)       As Byte     '���i���H���@��o���z
    PF_KAKOU_T_KIN(0 To 10)     As Byte     'PF���H�@��o���z
    PE_KAKOU_T_KIN(0 To 10)     As Byte     'PE���H�@��o���z
    PE_SHIZAI_T_KIN(0 To 10)    As Byte     'PF���ށ@��o���z
    HINBAN_LABEL_T_KIN(0 To 10) As Byte     '�i�ԕ\�����ف@��o���z
    KOUJI_SETSU_T_KIN(0 To 10)  As Byte     '�ݒu�H���������@��o���z
    KONPOU_T_KIN(0 To 10)       As Byte     '����ށ@��o���z
    FUKU_SHIZAI_T_KIN(0 To 10)  As Byte     '�����ށ@��o���z
    KONPOU_ASSY_T_KIN(0 To 10)  As Byte     '����ASSY�@��o���z
    KANRI_T_KIN(0 To 10)        As Byte     '�Ǘ���@��o���z
    GOUKEI_T_KIN(0 To 10)       As Byte     '��o���v���z

    NAKANISHI_F(0 To 0)         As Byte     '�����H�� ���Ϗ��\���׸�
    SHOHIN_F(0 To 0)            As Byte     '���i���H�� ���Ϗ��\���׸�
    PF_KAKOU_F(0 To 0)          As Byte     'PF���H ���Ϗ��\���׸�
    PE_KAKOU_F(0 To 0)          As Byte     'PE���H ���Ϗ��\���׸�
    PE_SHIZAI_F(0 To 0)         As Byte     'PF���� ���Ϗ��\���׸�
    HINBAN_LABEL_F(0 To 0)      As Byte     '�i�ԕ\������ ���Ϗ��\���׸�
    KOUJI_SETSU_F(0 To 0)       As Byte     '�ݒu�H�������� ���Ϗ��\���׸�
    KONPOU_F(0 To 0)            As Byte     '����� ���Ϗ��\���׸�
    FUKU_SHIZAI_F(0 To 0)       As Byte     '������ ���Ϗ��\���׸�
    KONPOU_ASSY_F(0 To 0)       As Byte     '����ASSY ���Ϗ��\���׸�
    KANRI_F(0 To 0)             As Byte     '�Ǘ��� ���Ϗ��\���׸�

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.09.28
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.11.07
    KO_QTY(0 To 5)              As Byte     '����
    NAKANISHI_QTY(0 To 5)       As Byte     '�����H���@����
    SHOHIN_QTY(0 To 5)          As Byte     '���i���H���@����
    PF_KAKOU_QTY(0 To 5)        As Byte     'PF���H�@����
    PE_KAKOU_QTY(0 To 5)        As Byte     'PE���H�@����
    PE_SHIZAI_QTY(0 To 5)       As Byte     'PF���ށ@����
    HINBAN_LABEL_QTY(0 To 5)    As Byte     '�i�ԕ\�����ف@����
    KOUJI_SETSU_QTY(0 To 5)     As Byte     '�ݒu�H���������@����
    KONPOU_QTY(0 To 5)          As Byte     '����ށ@����
    FUKU_SHIZAI_QTY(0 To 5)     As Byte     '�����ށ@����
    KONPOU_ASSY_QTY(0 To 5)     As Byte     '����ASSY�@����
    KANRI_QTY(0 To 5)           As Byte     '�Ǘ���@����
    
    
    NAKANISHI_T_TAN(0 To 10)    As Byte     '�����H���@��o�P��
    SHOHIN_T_TAN(0 To 10)       As Byte     '���i���H���@��o�P��
    PF_KAKOU_T_TAN(0 To 10)     As Byte     'PF���H�@��o�P��
    PE_KAKOU_T_TAN(0 To 10)     As Byte     'PE���H�@��o�P��
    PE_SHIZAI_T_TAN(0 To 10)    As Byte     'PF���ށ@��o�P��
    HINBAN_LABEL_T_TAN(0 To 10) As Byte     '�i�ԕ\�����ف@��o�P��
    KOUJI_SETSU_T_TAN(0 To 10)  As Byte     '�ݒu�H���������@��o�P��
    KONPOU_T_TAN(0 To 10)       As Byte     '����ށ@��o�P��
    FUKU_SHIZAI_T_TAN(0 To 10)  As Byte     '�����ށ@��o�P��
    KONPOU_ASSY_T_TAN(0 To 10)  As Byte     '����ASSY�@��o�P��
    KANRI_T_TAN(0 To 10)        As Byte     '�Ǘ���@��o�P��
    
    KO_SYUBETSU(0 To 1)         As Byte         '�q�@���
    
    FILLER(0 To 323)            As Byte
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.11.07
    
    INS_TANTO(0 To 9)           As Byte     '�ǉ��@�S����
    Ins_DateTime(0 To 13)       As Byte     '�ǉ��@����
    UPD_TANTO(0 To 9)           As Byte     '�X�V�@�S����
    UPD_DATETIME(0 To 13)       As Byte     '�X�V�@����

End Type
'�f�[�^�E�o�b�t�@
Public ITEM_O_REC               As ITEM_O_REC_Tag

'�L�[��`

Type KEY0_ITEM_O                '�j�d�x�O
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    NAIGAI(0 To 0)              As Byte     '�����O
    HIN_GAI(0 To 19)            As Byte     '�i��(�O��)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.11.07
    SEQ_NO(0 To 2)              As Byte     'SEQ_NO
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.11.07

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.09.28-->2017.11.07
'    KO_JGYOBU(0 To 0)           As Byte     '�q�@���ƕ��敪
'    KO_NAIGAI(0 To 0)           As Byte     '�q�@�����O
'    KO_HIN_GAI(0 To 19)         As Byte     '�q�@�i��(�O��)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.09.28-->2017.11.07

End Type


Type KEY1_ITEM_O                '�j�d�x�O
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    NAIGAI(0 To 0)              As Byte     '�����O
    HIN_GAI(0 To 19)            As Byte     '�i��(�O��)

    KO_JGYOBU(0 To 0)           As Byte     '�q�@���ƕ��敪
    KO_NAIGAI(0 To 0)           As Byte     '�q�@�����O
    KO_HIN_GAI(0 To 19)         As Byte     '�q�@�i��(�O��)

    SEQ_NO(0 To 2)              As Byte     'SEQ_NO

End Type




'�L�[�E�f�[�^
Public K0_ITEM_O                As KEY0_ITEM_O
Public K1_ITEM_O                As KEY1_ITEM_O

Type ITEM_O_FSpeck
    fs      As BtFileSpeck                  ' ̧�� ��߯��\����
    ks0     As BtKeySpeck                   ' �� ��߯��\����
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck                   ' �� ��߯��\����    2017.09.28
    
    ks4     As BtKeySpeck                   ' �� ��߯��\����    2017.09.28
    ks5     As BtKeySpeck                   ' �� ��߯��\����    2017.09.28
    ks6     As BtKeySpeck                   ' �� ��߯��\����    2017.09.28
    ks7     As BtKeySpeck                   ' �� ��߯��\����    2017.09.28
    ks8     As BtKeySpeck                   ' �� ��߯��\����    2017.09.28
    ks9     As BtKeySpeck                   ' �� ��߯��\����    2017.09.28
    ks10     As BtKeySpeck                   ' �� ��߯��\����    2017.09.28

End Type

Private ITEM_O_Speck            As ITEM_O_FSpeck

Private Function ITEM_O_Create() As Integer
'********************************************************************
'*
'*              ��㎖�@���ϗp�i�ڃ}�X�^  CREATE
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************

Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    ITEM_O_Create = True
                                        '��㎖�@���ϗp�i�ڃ}�X�^ �t���p�X�捞��
    sts = GetIni("FILE", ITEM_O_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ITEM_O]�ǂݍ��݃G���[ ")
        Exit Function
    End If

    FullPath = RTrim(c)

    ITEM_O_Speck.fs.recoleng = Len(ITEM_O_REC)          ' ���R�[�h��
    ITEM_O_Speck.fs.PageSize = ITEM_O_PG_SIZ            ' �y�[�W�T�C�Y
    ITEM_O_Speck.fs.idexnumb = 1                        ' �C���f�b�N�X��
    ITEM_O_Speck.fs.fileflag = 0                        ' �t�@�C���t���O
    ITEM_O_Speck.fs.reserve = &H0                       ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�O
    ITEM_O_Speck.ks0.keypos = 1                         ' �L�[�|�W�V����
    ITEM_O_Speck.ks0.keyleng = 1                        ' �L�[��
    ITEM_O_Speck.ks0.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    ITEM_O_Speck.ks0.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    ITEM_O_Speck.ks0.reserve = &H0                      ' �\��ς�

    ITEM_O_Speck.ks1.keypos = 2                         ' �L�[�|�W�V����
    ITEM_O_Speck.ks1.keyleng = 1                        ' �L�[��
    ITEM_O_Speck.ks1.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    ITEM_O_Speck.ks1.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    ITEM_O_Speck.ks1.reserve = &H0                      ' �\��ς�

    ITEM_O_Speck.ks2.keypos = 3                         ' �L�[�|�W�V����
    ITEM_O_Speck.ks2.keyleng = 20                       ' �L�[��
    ITEM_O_Speck.ks2.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    ITEM_O_Speck.ks2.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    ITEM_O_Speck.ks2.reserve = &H0                      ' �\��ς�


    ITEM_O_Speck.ks3.keypos = 23                        ' �L�[�|�W�V����
    ITEM_O_Speck.ks3.keyleng = 3                        ' �L�[��
    ITEM_O_Speck.ks3.keyflag = BtKfExt                  ' �L�[�t���O
    ITEM_O_Speck.ks3.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    ITEM_O_Speck.ks3.reserve = &H0                      ' �\��ς�



'-----------------------------------------------
    sts = BTRV(BtOpCreate, ITEM_O_POS, ITEM_O_Speck, Len(ITEM_O_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "��㎖�@���ϗp�i�ڃ}�X�^")
        Exit Function
    End If

    ITEM_O_Create = False

End Function

Public Function ITEM_O_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ��㎖�@���ϗp�i�ڃ}�X�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    ITEM_O_Open = True
                                            '��㎖�@���ϗp�i�ڃ}�X�^ �t���p�X�捞��
    sts = GetIni("FILE", ITEM_O_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ITEM_O]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, ITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = ITEM_O_Create()    '��㎖�@���ϗp�i�ڃ}�X�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "��㎖�@���ϗp�i�ڃ}�X�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "��㎖�@���ϗp�i�ڃ}�X�^")
                Exit Function
        End Select
    Loop

    ITEM_O_Open = False

End Function

Public Sub Rclr_ITEM_O_REC()

'********************************************************************
'*
'*              ��㎖�@���ϗp�i�ڃ}�X�^  ���R�[�h������
'*
'********************************************************************

    Call UniCode_Conv(ITEM_O_REC.JGYOBU, "")            '���ƕ��敪
    Call UniCode_Conv(ITEM_O_REC.NAIGAI, "")            '�����O
    Call UniCode_Conv(ITEM_O_REC.HIN_GAI, "")           '�i�ԁi�O���j


    Call UniCode_Conv(ITEM_O_REC.KO_JGYOBU, "")         '���ƕ��敪     2017.09.28
    Call UniCode_Conv(ITEM_O_REC.KO_NAIGAI, "")         '�����O         2017.09.28
    Call UniCode_Conv(ITEM_O_REC.KO_HIN_GAI, "")        '�i�ԁi�O���j   2017.09.28


    Call UniCode_Conv(ITEM_O_REC.NAKANISHI_TANI, "")    '�����H���@�P��
    Call UniCode_Conv(ITEM_O_REC.NAKANISHI_KIN, "")     '�����H���@���z

    Call UniCode_Conv(ITEM_O_REC.SHOHIN_TANI, "")       '���i���H���@�P��
    Call UniCode_Conv(ITEM_O_REC.SHOHIN_KIN, "")        '���i���H���@���z

    Call UniCode_Conv(ITEM_O_REC.PF_KAKOU_TANI, "")     'PF���H�@�P��
    Call UniCode_Conv(ITEM_O_REC.PF_KAKOU_KIN, "")      'PF���H�@���z

    Call UniCode_Conv(ITEM_O_REC.PE_KAKOU_TANI, "")     'PE���H�@�P��
    Call UniCode_Conv(ITEM_O_REC.PE_KAKOU_KIN, "")      'PE���H�@���z

    Call UniCode_Conv(ITEM_O_REC.PE_SHIZAI_TANI, "")    'PF���ށ@�P��
    Call UniCode_Conv(ITEM_O_REC.PE_SHIZAI_KIN, "")     'PF���ށ@���z

    Call UniCode_Conv(ITEM_O_REC.HINBAN_LABEL_TANI, "") '�i�ԕ\�����ف@�P��
    Call UniCode_Conv(ITEM_O_REC.HINBAN_LABEL_KIN, "") '�i�ԕ\�����ف@���z

    Call UniCode_Conv(ITEM_O_REC.KOUJI_SETSU_TANI, "")  '�ݒu�H���������@�P��
    Call UniCode_Conv(ITEM_O_REC.KOUJI_SETSU_KIN, "")  '�ݒu�H���������@���z

    Call UniCode_Conv(ITEM_O_REC.KONPOU_TANI, "")       '����ށ@�P��
    Call UniCode_Conv(ITEM_O_REC.KONPOU_KIN, "")        '����ށ@���z

    Call UniCode_Conv(ITEM_O_REC.FUKU_SHIZAI_TANI, "")  '�����ށ@�P��
    Call UniCode_Conv(ITEM_O_REC.FUKU_SHIZAI_KIN, "")   '�����ށ@���z

    Call UniCode_Conv(ITEM_O_REC.KONPOU_ASSY_TANI, "")  '����ASSY�@�P��
    Call UniCode_Conv(ITEM_O_REC.KONPOU_ASSY_KIN, "")   '����ASSY�@���z

    Call UniCode_Conv(ITEM_O_REC.KANRI_TANI, "")        '�Ǘ���@�P��
    Call UniCode_Conv(ITEM_O_REC.KANRI_KIN, "")         '�Ǘ���@���z
    
    Call UniCode_Conv(ITEM_O_REC.GOUKEI_KIN, "")        '���v�@���z
    Call UniCode_Conv(ITEM_O_REC.INPUT_TANTO_CODE, "")  '���͒S���Һ���
    
    
 '>>>>>>>>>>>>>>>>>>>>> 2017.09.27
    Call UniCode_Conv(ITEM_O_REC.BUZAI_TANTO_NAME, "")  '���ޒS���Җ�
    Call UniCode_Conv(ITEM_O_REC.T_HIN_NAME, "")        '��o�i��
    Call UniCode_Conv(ITEM_O_REC.TANI, "")              '�P��
    Call UniCode_Conv(ITEM_O_REC.T_TANKA, "")           '��o�P��
    Call UniCode_Conv(ITEM_O_REC.T_KINGAKU, "")         '��o���z
    Call UniCode_Conv(ITEM_O_REC.NAKANISHI_T_KIN, "")   '�����H���@��o���z
    Call UniCode_Conv(ITEM_O_REC.SHOHIN_T_KIN, "")      '���i���H���@��o���z
    Call UniCode_Conv(ITEM_O_REC.PF_KAKOU_T_KIN, "")    'PF���H�@��o���z
    Call UniCode_Conv(ITEM_O_REC.PE_KAKOU_T_KIN, "")    'PE���H�@��o���z
    Call UniCode_Conv(ITEM_O_REC.PE_SHIZAI_T_KIN, "")   'PF���ށ@��o���z
    Call UniCode_Conv(ITEM_O_REC.HINBAN_LABEL_T_KIN, "") '�i�ԕ\�����ف@��o���z
    Call UniCode_Conv(ITEM_O_REC.KOUJI_SETSU_T_KIN, "") '�ݒu�H���������@��o���z
    Call UniCode_Conv(ITEM_O_REC.KONPOU_T_KIN, "")      '����ށ@��o���z
    Call UniCode_Conv(ITEM_O_REC.FUKU_SHIZAI_T_KIN, "") '�����ށ@��o���z
    Call UniCode_Conv(ITEM_O_REC.KONPOU_ASSY_T_KIN, "") '����ASSY�@��o���z
    Call UniCode_Conv(ITEM_O_REC.KANRI_T_KIN, "")       '�Ǘ���@��o���z
    Call UniCode_Conv(ITEM_O_REC.GOUKEI_T_KIN, "")      '��o���v���z

    Call UniCode_Conv(ITEM_O_REC.NAKANISHI_F, "")       '�����H�� ���Ϗ��\���׸�
    Call UniCode_Conv(ITEM_O_REC.SHOHIN_F, "")          '���i���H�� ���Ϗ��\���׸�
    Call UniCode_Conv(ITEM_O_REC.PF_KAKOU_F, "")        'PF���H ���Ϗ��\���׸�
    Call UniCode_Conv(ITEM_O_REC.PE_KAKOU_F, "")        'PE���H ���Ϗ��\���׸�
    Call UniCode_Conv(ITEM_O_REC.PE_SHIZAI_F, "")       'PF���� ���Ϗ��\���׸�
    Call UniCode_Conv(ITEM_O_REC.HINBAN_LABEL_F, "")    '�i�ԕ\������ ���Ϗ��\���׸�
    Call UniCode_Conv(ITEM_O_REC.KOUJI_SETSU_F, "")     '�ݒu�H�������� ���Ϗ��\���׸�
    Call UniCode_Conv(ITEM_O_REC.KONPOU_F, "")          '����� ���Ϗ��\���׸�
    Call UniCode_Conv(ITEM_O_REC.FUKU_SHIZAI_F, "")     '������ ���Ϗ��\���׸�
    Call UniCode_Conv(ITEM_O_REC.KONPOU_ASSY_F, "")     '����ASSY ���Ϗ��\���׸�
    Call UniCode_Conv(ITEM_O_REC.KANRI_F, "")           '�Ǘ��� ���Ϗ��\���׸�
 '>>>>>>>>>>>>>>>>>>>>> 2017.09.27
    
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.11.07
    Call UniCode_Conv(ITEM_O_REC.KO_QTY, "")            '����
    Call UniCode_Conv(ITEM_O_REC.NAKANISHI_QTY, "")     '�����H���@����
    Call UniCode_Conv(ITEM_O_REC.SHOHIN_QTY, "")        '���i���H���@����
    Call UniCode_Conv(ITEM_O_REC.PF_KAKOU_QTY, "")      'PF���H�@����
    Call UniCode_Conv(ITEM_O_REC.PE_KAKOU_QTY, "")      'PE���H�@����
    Call UniCode_Conv(ITEM_O_REC.PE_SHIZAI_QTY, "")     'PF���ށ@����
    Call UniCode_Conv(ITEM_O_REC.HINBAN_LABEL_QTY, "")  '�i�ԕ\�����ف@����
    Call UniCode_Conv(ITEM_O_REC.KOUJI_SETSU_QTY, "")   '�ݒu�H���������@����
    Call UniCode_Conv(ITEM_O_REC.KONPOU_QTY, "")        '����ށ@����
    Call UniCode_Conv(ITEM_O_REC.FUKU_SHIZAI_QTY, "")   '�����ށ@����
    Call UniCode_Conv(ITEM_O_REC.KONPOU_ASSY_QTY, "")   '����ASSY�@����
    Call UniCode_Conv(ITEM_O_REC.KANRI_QTY, "")         '�Ǘ���@����
    
    
    Call UniCode_Conv(ITEM_O_REC.NAKANISHI_T_TAN, "")   '�����H���@��o�P��
    Call UniCode_Conv(ITEM_O_REC.SHOHIN_T_TAN, "")      '���i���H���@��o�P��
    Call UniCode_Conv(ITEM_O_REC.PF_KAKOU_T_TAN, "")    'PF���H�@��o�P��
    Call UniCode_Conv(ITEM_O_REC.PE_KAKOU_T_TAN, "")    'PE���H�@��o�P��
    Call UniCode_Conv(ITEM_O_REC.PE_SHIZAI_T_TAN, "")   'PF���ށ@��o�P��
    Call UniCode_Conv(ITEM_O_REC.HINBAN_LABEL_T_TAN, "") '�i�ԕ\�����ف@��o�P��
    Call UniCode_Conv(ITEM_O_REC.KOUJI_SETSU_T_TAN, "") '�ݒu�H����+�����@��o�P��
    Call UniCode_Conv(ITEM_O_REC.KONPOU_T_TAN, "")      '����ށ@��o�P��
    Call UniCode_Conv(ITEM_O_REC.FUKU_SHIZAI_T_TAN, "") '�����ށ@��o�P��
    Call UniCode_Conv(ITEM_O_REC.KONPOU_ASSY_T_TAN, "") '����ASSY�@��o�P��
    Call UniCode_Conv(ITEM_O_REC.KANRI_T_TAN, "")       '�Ǘ���@��o�P��
    
    
    
    Call UniCode_Conv(ITEM_O_REC.FILLER, "")
    
    Call UniCode_Conv(ITEM_O_REC.INS_TANTO, "")         '�ǉ��@�S����
    Call UniCode_Conv(ITEM_O_REC.Ins_DateTime, "")      '�ǉ��@����
    Call UniCode_Conv(ITEM_O_REC.UPD_TANTO, "")         '�X�V�@�S����
    Call UniCode_Conv(ITEM_O_REC.UPD_DATETIME, "")      '�X�V�@����



End Sub
