Attribute VB_Name = "IDO"
Option Explicit
'********************************************************************
'*
'*              �݌Ɉړ����@�t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const IDO_ID$ = "IDO"

'�y�[�W�T�C�Y
Public Const IDO_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public IDO_POS As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type IDOREC_Tag
    JITU_DT(0 To 7)                     As Byte     '���ѓ��t
    JITU_TM(0 To 5)                     As Byte     '���ю���
    JGYOBU(0 To 0)                      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)                      As Byte     '�����O
    HIN_GAI(0 To 19)                    As Byte     '�i�ԁi�O���j
    RIRK_ID(0 To 1)                     As Byte     '�������
    SUMI_JITU_QTY(0 To 7)               As Byte     '���ѐ���(���i���ς�)
    MI_JITU_QTY(0 To 7)                 As Byte     '���ѐ���(�����i)
    FROM_SOKO(0 To 1)                   As Byte     'From �q�ɇ�
    FROM_RETU(0 To 1)                   As Byte     '   �@��
    FROM_REN(0 To 1)                    As Byte     '   �@�A
    FROM_DAN(0 To 1)                    As Byte     '   �@�i
    TO_SOKO(0 To 1)                     As Byte     '�s�n �q�ɇ�
    TO_RETU(0 To 1)                     As Byte     '   �@��
    TO_REN(0 To 1)                      As Byte     '   �@�A
    TO_DAN(0 To 1)                      As Byte     '   �@�i
    DEN_DT(0 To 7)                      As Byte     '�`�[���t
    DEN_NO(0 To 9)                      As Byte     '�`�[��
    PRG_ID(0 To 7)                      As Byte     '�o�͌��v���O����
    HIN_NAI(0 To 19)                    As Byte     '�i�ԁi�����j
    NYUKA_DT(0 To 7)                    As Byte     '���ד��t
    NYUKO_DT(0 To 7)                    As Byte     '���ɓ��t
    WEL_ID(0 To 2)                      As Byte     '�Ώے[����
    RIRK_NAME(0 To 9)                   As Byte     '������ʖ���
    HIN_NAME(0 To 39)                   As Byte     '�i��
    SUMI_HIN_Zaiko_Qty(0 To 7)          As Byte     '�i�ڕʍ݌ɐ��i���i���ς݁j
    MI_HIN_Zaiko_Qty(0 To 7)            As Byte     '�i�ڕʍ݌ɐ��i�����i�j
    SUMI_FROM_TANA_Zaiko_Qty(0 To 7)    As Byte     'FROM�I�ʕi�ڕʍ݌ɐ�
    SUMI_TO_TANA_Zaiko_Qty(0 To 7)      As Byte     'TO�I�ʕi�ڕʍ݌ɐ�
    MI_FROM_TANA_Zaiko_Qty(0 To 7)      As Byte     'FROM�I�ʕi�ڕʍ݌ɐ�
    MI_TO_TANA_Zaiko_Qty(0 To 7)        As Byte     'TO�I�ʕi�ڕʍ݌ɐ�
    TOKU_MARK(0 To 0)                   As Byte     '������}�[�N
    MEMO(0 To 59)                       As Byte     '����
    TANTO_CODE(0 To 4)                  As Byte     '�S���҃R�[�h
    TANTO_NAME(0 To 19)                 As Byte     '�S���Җ���
    MUKE_CODE(0 To 7)                   As Byte     '���Ӑ�R�[�h
    MUKE_NAME(0 To 39)                  As Byte     '���Ӑ於��
    SS_CODE(0 To 7)                     As Byte     '������R�[�h
    SS_NAME(0 To 39)                    As Byte     '�����於��
    MUKE_DNAME(0 To 9)                  As Byte     '���Ӑ旪��
    MUKE_CHG_CD(0 To 1)                 As Byte     '������Ǒւ��R�[�h
    SUM_KBN(0 To 0)                     As Byte     '�W�v�敪
'    ID_NO(0 To 7)                       As Byte     'ID-NO
    ID_NO(0 To 11)                      As Byte     'ID-NO (8����12��)      2006/05/24

    Ins_DateTime(0 To 13)               As Byte     '�}������2004.12.09

    '���ޏ����̈גǉ�2005.01.05
    SHIIRE_CODE(0 To 4)                 As Byte     '�d���溰��
    SHIIRE_TANKA(0 To 10)               As Byte     '�d���P��(9(8)V99)
    KEIJYO_YM(0 To 5)                   As Byte     '�v��N��(YYYYMM)
    '���ޏ����̈גǉ�2005.01.05

    BIN_NO(0 To 1)                      As Byte     '��(01:1�� 02:2�� 03:3��)   2007/05/15


    '----------------   2010.07.08 ��
    GENSANKOKU(0 To 19)                 As Byte     '���Y����
    SHIIRE_WORK_CENTER(0 To 7)          As Byte     '���ގd����ܰ�����
    ID_NO2(0 To 11)                     As Byte     'ID_NO
    YOSAN_FROM(0 To 4)                  As Byte     '�\�Z�P�ʁi���j
    YOSAN_TO(0 To 4)                    As Byte     '�\�Z�P�ʁi��j
    '----------------   2010.07.08 ��


'    FILLER(0 To 167)                     As Byte
'    FILLER(0 To 163)                     As Byte    '                      2006/05/24
    FILLER(0 To 111)                     As Byte    '                           2007/05/15

End Type

'�f�[�^�E�o�b�t�@
Public IDOREC   As IDOREC_Tag

'�L�[��`
Type KEY0_IDO            '�j�d�x�O
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    JITU_DT(0 To 7)             As Byte     '���ѓ��t
    JITU_TM(0 To 5)             As Byte     '���ю���
End Type

Type KEY1_IDO            '�j�d�x�P
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    NAIGAI(0 To 0)              As Byte     '�����O
    HIN_GAI(0 To 19)            As Byte     '�i�ԁi�O���j
    JITU_DT(0 To 7)             As Byte     '���ѓ��t
    JITU_TM(0 To 5)             As Byte     '���ю���
End Type




'�L�[�E�f�[�^
Public K0_IDO                   As KEY0_IDO
Public K1_IDO                   As KEY1_IDO

Type IDO_FSpeck
    fs      As BtFileSpeck          ' ̧�� ��߯��\����
    ks0     As BtKeySpeck           ' �� ��߯��\����
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck
    ks4     As BtKeySpeck
    ks5     As BtKeySpeck
    ks6     As BtKeySpeck
    ks7     As BtKeySpeck
    ks8     As BtKeySpeck
End Type

Private IDO_Speck               As IDO_FSpeck
Private Function IDO_Create() As Integer
'********************************************************************
'*
'*              �݌Ɉړ����@�b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    IDO_Create = True
                                            '�݌Ɉړ����t���p�X�捞��
    sts = GetIni("FILE", IDO_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [IDO]�ǂݍ��݃G���[")
        Exit Function
    End If
     
    FullPath = RTrim(c)
    
    IDO_Speck.fs.recoleng = Len(IDOREC)         ' ���R�[�h��
    IDO_Speck.fs.PageSize = IDO_PG_SIZ          ' �y�[�W�T�C�Y
    IDO_Speck.fs.idexnumb = 2                   ' �C���f�b�N�X��
    IDO_Speck.fs.fileflag = 0                   ' �t�@�C���t���O
    IDO_Speck.fs.reserve = &H0                  ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�O
    IDO_Speck.ks0.keypos = 15                   ' �L�[�|�W�V����
                                                ' �L�[��
    IDO_Speck.ks0.keyleng = 1
                                                ' �L�[�t���O
    IDO_Speck.ks0.keyflag = BtKfExt + BtKfDup + BtKfSeg
    IDO_Speck.ks0.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    IDO_Speck.ks0.reserve = &H0                 ' �\��ς�
    
    IDO_Speck.ks1.keypos = 1                    ' �L�[�|�W�V����
    IDO_Speck.ks1.keyleng = 8                   ' �L�[��
                                                ' �L�[�t���O
    IDO_Speck.ks1.keyflag = BtKfExt + BtKfDup + BtKfSeg
    IDO_Speck.ks1.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    IDO_Speck.ks1.reserve = &H0                 ' �\��ς�
    
    IDO_Speck.ks2.keypos = 9                    ' �L�[�|�W�V����
    IDO_Speck.ks2.keyleng = 6                   ' �L�[��
    IDO_Speck.ks2.keyflag = BtKfExt + BtKfDup   ' �L�[�t���O
    IDO_Speck.ks2.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    IDO_Speck.ks2.reserve = &H0                 ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�P
    IDO_Speck.ks3.keypos = 15                   ' �L�[�|�W�V����
    IDO_Speck.ks3.keyleng = 1                   ' �L�[��
                                                ' �L�[�t���O
    IDO_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfSeg
    IDO_Speck.ks3.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    IDO_Speck.ks3.reserve = &H0                 ' �\��ς�

    IDO_Speck.ks4.keypos = 16                   ' �L�[�|�W�V����
    IDO_Speck.ks4.keyleng = 1                   ' �L�[��
                                                ' �L�[�t���O
    IDO_Speck.ks4.keyflag = BtKfExt + BtKfDup + BtKfSeg
    IDO_Speck.ks4.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    IDO_Speck.ks4.reserve = &H0                 ' �\��ς�

    IDO_Speck.ks5.keypos = 17                   ' �L�[�|�W�V����
    IDO_Speck.ks5.keyleng = 20                  ' �L�[��
                                                ' �L�[�t���O
    IDO_Speck.ks5.keyflag = BtKfExt + BtKfDup + BtKfSeg
    IDO_Speck.ks5.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    IDO_Speck.ks5.reserve = &H0                 ' �\��ς�

    IDO_Speck.ks6.keypos = 1                    ' �L�[�|�W�V����
    IDO_Speck.ks6.keyleng = 8                   ' �L�[��
                                                ' �L�[�t���O
    IDO_Speck.ks6.keyflag = BtKfExt + BtKfDup + BtKfSeg
    IDO_Speck.ks6.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    IDO_Speck.ks6.reserve = &H0                 ' �\��ς�

    IDO_Speck.ks7.keypos = 9                    ' �L�[�|�W�V����
    IDO_Speck.ks7.keyleng = 6                   ' �L�[��
                                                ' �L�[�t���O
    IDO_Speck.ks7.keyflag = BtKfExt + BtKfDup
    IDO_Speck.ks7.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    IDO_Speck.ks7.reserve = &H0                 ' �\��ς�
'-----------------------------------------------

    sts = BTRV(BtOpCreate, IDO_POS, IDO_Speck, Len(IDO_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�݌Ɉړ���")
        Exit Function
    End If

    IDO_Create = False

End Function

Public Function IDO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �݌Ɉړ����@�n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    IDO_Open = True
                                            '�݌Ɉړ����t���p�X�捞��
    sts = GetIni("FILE", IDO_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [IDO]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, IDO_POS, IDOREC, Len(IDOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = IDO_Create()        '�݌Ɉړ����쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, IDO_POS, IDOREC, Len(IDOREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�݌Ɉړ���")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�݌Ɉړ���")
                Exit Function
        End Select
    Loop
    IDO_Open = False
End Function


