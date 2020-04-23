Attribute VB_Name = "DEL_Syuka_TEI"
Option Explicit
'********************************************************************
'*
'*              �폜�ςݓ@�ʒ����f�[�^  �t�@�C����`
'*
'*          CREATE 2011.06.23
'********************************************************************
'�t�@�C���h�c
Public Const DEL_SYU_TEI_ID$ = "DEL_SYU_TEI"

'�y�[�W�T�C�Y
Public Const DEL_SYU_TEI_PG_SIZ% = 4096

'�|�W�V�����E�u���b�N
Public DEL_SYU_TEI_POS            As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type DEL_SYU_TEI_REC_Tag
    SND_YMD(0 To 7)                 As Byte         '�f�[�^�쐬��
    SND_HMS(0 To 5)                 As Byte         '�f�[�^�쐬����
    SEQ_NO(0 To 4)                  As Byte         '�A��
    JUC_YMD(0 To 7)                 As Byte         '�󒍓�
    NOU_CD(0 To 3)                  As Byte         '�[�������
    NOU_NM(0 To 19)                 As Byte         '�[������ꖼ
    TOK_CD(0 To 7)                  As Byte         '���Ӑ溰��
    CHO_CD(0 To 7)                  As Byte         '���[�溰��
    THINB_CD(0 To 19)               As Byte         '���Ӑ�i�ԁ@���i��(��)
    HINB_CD(0 To 19)                As Byte         '�i�ԁ@      ���i��(��)
    CHU_CD(0 To 9)                  As Byte         '�������@    ���w�}��(��)
    SYU_JUN(0 To 9)                 As Byte         '�o�׏��ԁ@  ���w�}��(���E��)
    TEI_NM(0 To 29)                 As Byte         '�@���@      ���w�}��(���E�E)
    JUC_SUU(0 To 7)                 As Byte         '�󒍐���
    SYU_YMD(0 To 7)                 As Byte         '�o�׊m���
    NOU_YMD(0 To 7)                 As Byte         '�[����
    KEN_NO(0 To 5)                  As Byte         '���Ǉ��@�@�@���Ǘ���(��)
    HIN_NO(0 To 5)                  As Byte         '�i�Ǉ��@�@�@���Ǘ���(��)
    TANP_KB(0 To 0)                 As Byte         '�P�i�敪
    YOBI1_NM(0 To 54)               As Byte         '�\��
    GSEQ_NO(0 To 4)                 As Byte         '�ް�������
    TEI_LABELID(0 To 12)            As Byte         '�@������ID(���������w�}��(��)+����)    2011.04.25
    HAKO_NO(0 To 2)                 As Byte         '����                                   2011.04.25
    JITU_SUU(0 To 7)                As Byte         '���o�ɐ�(�����ւ̏o�ɐ� ���ݖ��g�p)  2011.04.26
    JITU_TANTO(0 To 9)              As Byte         '�o�Ɂ@�S����(���ݖ��g�p)               2011.04.26
    JITU_DATETIME(0 To 13)          As Byte         '�o�Ɂ@����(���ݖ��g�p)                 2011.04.26
    KONPO_TANTO(0 To 9)             As Byte         '����@�S����                           2011.04.26
    KONPO_DATETIME(0 To 13)         As Byte         '����@����                             2011.04.26
    SHOGO_TANTO(0 To 9)             As Byte         '�����ް��ƍ��S��                       2011.05.02
    SHOGO_DATETIME(0 To 13)         As Byte         '�����ް��ƍ�����                       2011.05.02
    
    L_KENKAN(0 To 11)               As Byte         '���ǖ��� long                          2011.05.06
    L_TEI_NAME(0 To 49)             As Byte         '�@��2 50                               2011.05.06
    L_TOK_NAME(0 To 49)             As Byte         '���Ӑ於 50                            2011.05.06
    L_SOTO_NO(0 To 9)               As Byte         '�O���ԍ� 50 �� 10                      2011.05.06
    L_UCHI_NO(0 To 9)               As Byte         '�����ԍ� 50 �� 10                      2011.05.06
    L_WIDTH(0 To 9)                 As Byte         '����(��) 10                            2011.05.06
    L_HEIGHT(0 To 9)                As Byte         '����     20                            2011.05.06
    L_CONTENT(0 To 9)               As Byte         '�̐�     30                            2011.05.06
    L_KNo(0 To 1)                   As Byte         '�H��No 2 32                            2011.05.06
    L_SERIES1(0 To 19)              As Byte         '�i�ԃV���[�Y 20  52                    2011.05.06
    L_SERIES2(0 To 19)              As Byte         '�i�ԃV���[�Y 2                         2011.05.06
    L_PAGE(0 To 4)                  As Byte         '�y�[�W�ԍ�                             2011.05.06
    
    KUTI_SU(0 To 3)                 As Byte         '���� 9999  (�@������ID���ɓ����l)      2011.05.10
    SAI_SU(0 To 5)                  As Byte         '�ː� 999.99 (�@������ID���ɓ����l)     2011.05.10
    
    KONPO_ID(0 To 19)               As Byte         '����ID                                 2011.05.10
    
    
    KENPIN_TANTO(0 To 9)             As Byte        '���i�S����                             2011.05.12
    KENPIN_DATETIME(0 To 13)         As Byte        '���i����                               2011.05.12
    
    
    SYUGO_KONPO_TANTO(0 To 9)       As Byte         '�W������S����                         2011.05.12
    SYUGO_KONPO_DATETIME(0 To 13)   As Byte         '�W���������                           2011.05.12
    
    
    
    FILLER(0 To 338)                As Byte         'FILLER                                 2011.05.12
    INS_TANTO(0 To 9)               As Byte         '�ǉ��@�S����
    Ins_DateTime(0 To 13)           As Byte         '�ǉ��@����
    UPD_TANTO(0 To 9)               As Byte         '�X�V�@�S����
    UPD_DATETIME(0 To 13)           As Byte         '�X�V�@����



End Type
'�f�[�^�E�o�b�t�@
Public DEL_SYU_TEI_REC                As DEL_SYU_TEI_REC_Tag

'�L�[��`

Type KEY0_DEL_SYU_TEI                 '�j�d�x�O
    
    SND_YMD(0 To 7)                 As Byte         '�f�[�^�쐬��
    SND_HMS(0 To 5)                 As Byte         '�f�[�^�쐬����
    SEQ_NO(0 To 4)                  As Byte         '�A��

End Type


Type KEY1_DEL_SYU_TEI                 '�j�d�x�P
    
    TEI_LABELID(0 To 12)            As Byte         '�@������ID(���������w�}��(��)+����)

End Type


Type KEY2_DEL_SYU_TEI                 '�j�d�x�Q
    
    KEN_NO(0 To 5)                  As Byte         '���Ǉ��@�@�@���Ǘ���(��)
    HIN_NO(0 To 5)                  As Byte         '���Ǉ��@�@�@���Ǘ���(��)

End Type


Type KEY3_DEL_SYU_TEI                 '�j�d�x�R
    
    KONPO_ID(0 To 19)               As Byte         '����ID     2011.05.10

End Type





'�L�[�E�f�[�^
Public K0_DEL_SYU_TEI                 As KEY0_DEL_SYU_TEI
Public K1_DEL_SYU_TEI                 As KEY1_DEL_SYU_TEI
Public K2_DEL_SYU_TEI                 As KEY2_DEL_SYU_TEI

Public K3_DEL_SYU_TEI                 As KEY3_DEL_SYU_TEI   '2011.05.12


Private Type DEL_SYU_TEI_FSpeck
    fs      As BtFileSpeck              ' ̧�� ��߯��\����
    ks0     As BtKeySpeck               ' �� ��߯��\����
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck
    ks4     As BtKeySpeck
    ks5     As BtKeySpeck

    ks6     As BtKeySpeck                               '2011.05.12

End Type

Private DEL_SYU_TEI_Speck  As DEL_SYU_TEI_FSpeck

Private Function DEL_SYU_TEI_Create() As Integer
'********************************************************************
'*
'*              �@�ʒ����f�[�^  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    DEL_SYU_TEI_Create = True
                                            '���Y�}�X�^�t���p�X�捞��
    sts = GetIni("FILE", DEL_SYU_TEI_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [DEL_SYU_TEI]�ǂݍ��݃G���[ ")
        Exit Function
    End If

    FullPath = RTrim(c)

    DEL_SYU_TEI_Speck.fs.recoleng = Len(DEL_SYU_TEI_REC)    ' ���R�[�h��
    DEL_SYU_TEI_Speck.fs.PageSize = DEL_SYU_TEI_PG_SIZ      ' �y�[�W�T�C�Y
    DEL_SYU_TEI_Speck.fs.idexnumb = 4                       ' �C���f�b�N�X��
    DEL_SYU_TEI_Speck.fs.fileflag = 0                       ' �t�@�C���t���O
    DEL_SYU_TEI_Speck.fs.reserve = &H0                      ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�O
    DEL_SYU_TEI_Speck.ks0.keypos = 1                        ' �L�[�|�W�V����
    DEL_SYU_TEI_Speck.ks0.keyleng = 8                       ' �L�[��
                                                            ' �L�[�t���O
    DEL_SYU_TEI_Speck.ks0.keyflag = BtKfExt + BtKfDup + BtKfSeg
    DEL_SYU_TEI_Speck.ks0.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    DEL_SYU_TEI_Speck.ks0.reserve = &H0                     ' �\��ς�

    DEL_SYU_TEI_Speck.ks1.keypos = 9                        ' �L�[�|�W�V����
    DEL_SYU_TEI_Speck.ks1.keyleng = 6                       ' �L�[��
                                                            ' �L�[�t���O
    DEL_SYU_TEI_Speck.ks1.keyflag = BtKfExt + BtKfDup + BtKfSeg
    DEL_SYU_TEI_Speck.ks1.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    DEL_SYU_TEI_Speck.ks1.reserve = &H0                     ' �\��ς�

    DEL_SYU_TEI_Speck.ks2.keypos = 15                       ' �L�[�|�W�V����
    DEL_SYU_TEI_Speck.ks2.keyleng = 5                       ' �L�[��
                                                            ' �L�[�t���O
    DEL_SYU_TEI_Speck.ks2.keyflag = BtKfExt + BtKfDup
    DEL_SYU_TEI_Speck.ks2.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    DEL_SYU_TEI_Speck.ks2.reserve = &H0                     ' �\��ς�




'-----------------------------------------------
                                                ' �L�[�P
    DEL_SYU_TEI_Speck.ks3.keypos = 255                      ' �L�[�|�W�V����
    DEL_SYU_TEI_Speck.ks3.keyleng = 13                      ' �L�[��
                                                            ' �L�[�t���O
    DEL_SYU_TEI_Speck.ks3.keyflag = BtKfExt + BtKfChg + BtKfDup
    DEL_SYU_TEI_Speck.ks3.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    DEL_SYU_TEI_Speck.ks3.reserve = &H0                     ' �\��ς�



'-----------------------------------------------
                                                ' �L�[�Q
    DEL_SYU_TEI_Speck.ks4.keypos = 182                      ' �L�[�|�W�V����
    DEL_SYU_TEI_Speck.ks4.keyleng = 6                       ' �L�[��
                                                            ' �L�[�t���O
    DEL_SYU_TEI_Speck.ks4.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    DEL_SYU_TEI_Speck.ks4.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    DEL_SYU_TEI_Speck.ks4.reserve = &H0                     ' �\��ς�

    DEL_SYU_TEI_Speck.ks5.keypos = 188                      ' �L�[�|�W�V����
    DEL_SYU_TEI_Speck.ks5.keyleng = 6                       ' �L�[��
                                                            ' �L�[�t���O
    DEL_SYU_TEI_Speck.ks5.keyflag = BtKfExt + BtKfChg + BtKfDup
    DEL_SYU_TEI_Speck.ks5.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    DEL_SYU_TEI_Speck.ks5.reserve = &H0                     ' �\��ς�


'-----------------------------------------------
                                                ' �L�[�R
    DEL_SYU_TEI_Speck.ks6.keypos = 570                      ' �L�[�|�W�V����
    DEL_SYU_TEI_Speck.ks6.keyleng = 20                      ' �L�[��
                                                            ' �L�[�t���O
    DEL_SYU_TEI_Speck.ks6.keyflag = BtKfExt + BtKfChg + BtKfDup
    DEL_SYU_TEI_Speck.ks6.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    DEL_SYU_TEI_Speck.ks6.reserve = &H0                     ' �\��ς�



'-----------------------------------------------

    sts = BTRV(BtOpCreate, DEL_SYU_TEI_POS, DEL_SYU_TEI_Speck, Len(DEL_SYU_TEI_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�@�ʒ����f�[�^")
        Exit Function
    End If

    DEL_SYU_TEI_Create = False

End Function

Public Function DEL_SYU_TEI_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �@�ʒ����f�[�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    DEL_SYU_TEI_Open = True
                                            '�@�ʒ����f�[�^ �t���p�X�捞��
    sts = GetIni("FILE", DEL_SYU_TEI_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [DEL_SYU_TEI]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, DEL_SYU_TEI_POS, DEL_SYU_TEI_REC, Len(DEL_SYU_TEI_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = DEL_SYU_TEI_Create()        '�@�ʒ����f�[�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, DEL_SYU_TEI_POS, DEL_SYU_TEI_REC, Len(DEL_SYU_TEI_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�@�ʒ����f�[�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�@�ʒ����f�[�^")
                Exit Function
        End Select
    Loop

    DEL_SYU_TEI_Open = False

End Function

