Attribute VB_Name = "PLN_tmpP_COMP"
Option Explicit
'********************************************************************
'*
'*              ���ޏ��v�ʒ��ԃt�@�C�� �t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const PLN_tmpP_COMP_ID$ = "PLN_tmpP_COMP"

'�y�[�W�T�C�Y
Public Const PLN_tmpP_COMP_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public PLN_tmpP_COMP_POS As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
Type PLN_tmpP_COMP_REC_Tag
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j
    KO_SYUBETSU(0 To 1)     As Byte     '���
    KO_JGYOBU(0 To 0)       As Byte     '���ƕ��敪
    KO_NAIGAI(0 To 0)       As Byte     '�����O
    KO_HIN_GAI(0 To 19)     As Byte     '�i�ԁi�O���j
    YOTEI_DT(0 To 7)        As Byte     '���i���\����t
    YOTEI_QTY(0 To 7)       As Byte     '���i���\�萔
    KO_QTY(0 To 5)          As Byte     '�q�@����(999V99)
    USE_QTY(0 To 5)         As Byte     '�q�@�K�v��
    DATA_KBN(0 To 0)        As Byte     '�ް��敪
    INS_TANTO(0 To 9)       As Byte     '�ǉ��@�S����
    Ins_DateTime(0 To 13)   As Byte     '�ǉ��@����         YYYYMMDDhhmmss

End Type

'�f�[�^�E�o�b�t�@
Public PLN_tmpP_COMP_REC    As PLN_tmpP_COMP_REC_Tag

'�L�[��`
Type KEY0_PLN_tmpP_COMP                 '�j�d�x�O
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j
    KO_SYUBETSU(0 To 1)     As Byte     '���
    KO_JGYOBU(0 To 0)       As Byte     '���ƕ��敪
    KO_NAIGAI(0 To 0)       As Byte     '�����O
    KO_HIN_GAI(0 To 19)     As Byte     '�i�ԁi�O���j
    YOTEI_DT(0 To 7)        As Byte     '���i���\����t
End Type

Type KEY1_PLN_tmpP_COMP                 '�j�d�x�P
    YOTEI_DT(0 To 7)        As Byte     '���i���\����t
    KO_SYUBETSU(0 To 1)     As Byte     '���
    KO_JGYOBU(0 To 0)       As Byte     '���ƕ��敪
    KO_NAIGAI(0 To 0)       As Byte     '�����O
    KO_HIN_GAI(0 To 19)     As Byte     '�i�ԁi�O���j
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j
End Type

'�L�[�E�f�[�^
Public K0_PLN_tmpP_COMP     As KEY0_PLN_tmpP_COMP
Public K1_PLN_tmpP_COMP     As KEY1_PLN_tmpP_COMP

Type PLN_tmpP_COMP_FSpeck
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
    ks9     As BtKeySpeck
    ks10    As BtKeySpeck
    ks11    As BtKeySpeck
    ks12    As BtKeySpeck
    ks13    As BtKeySpeck
    ks14    As BtKeySpeck
    ks15    As BtKeySpeck
End Type

Private PLN_tmpP_COMP_Speck  As PLN_tmpP_COMP_FSpeck
Private Function PLN_tmpP_COMP_Create() As Integer
'********************************************************************
'*
'*              ���ޏ��v�ʒ��ԃt�@�C���@�b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    PLN_tmpP_COMP_Create = True
                                            '���ޏ��v�ʒ��ԃt�@�C���@�t���p�X�捞��
    sts = GetIni("FILE", PLN_tmpP_COMP_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [PLN_tmpP_COMP]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    PLN_tmpP_COMP_Speck.fs.recoleng = Len(PLN_tmpP_COMP_REC)    ' ���R�[�h��
    PLN_tmpP_COMP_Speck.fs.PageSize = PLN_tmpP_COMP_PG_SIZ      ' �y�[�W�T�C�Y
    PLN_tmpP_COMP_Speck.fs.idexnumb = 2                         ' �C���f�b�N�X��
    PLN_tmpP_COMP_Speck.fs.fileflag = 0                         ' �t�@�C���t���O
    PLN_tmpP_COMP_Speck.fs.reserve = &H0                        ' �\��ς�
'---------------------------------------------------'
                                                    ' �L�[�O
    PLN_tmpP_COMP_Speck.ks0.keypos = 1                          ' �L�[�|�W�V����
    PLN_tmpP_COMP_Speck.ks0.keyleng = 1                         ' �L�[��
    PLN_tmpP_COMP_Speck.ks0.keyflag = BtKfExt + BtKfSeg         ' �L�[�t���O
    PLN_tmpP_COMP_Speck.ks0.keytype = Chr(BtKtString)           ' �L�[�^�C�v
    PLN_tmpP_COMP_Speck.ks0.reserve = &H0                       ' �\��ς�
                                                    
    PLN_tmpP_COMP_Speck.ks1.keypos = 2                          ' �L�[�|�W�V����
    PLN_tmpP_COMP_Speck.ks1.keyleng = 1                         ' �L�[��
    PLN_tmpP_COMP_Speck.ks1.keyflag = BtKfExt + BtKfSeg         ' �L�[�t���O
    PLN_tmpP_COMP_Speck.ks1.keytype = Chr(BtKtString)           ' �L�[�^�C�v
    PLN_tmpP_COMP_Speck.ks1.reserve = &H0                       ' �\��ς�
                                                    
    PLN_tmpP_COMP_Speck.ks2.keypos = 3                          ' �L�[�|�W�V����
    PLN_tmpP_COMP_Speck.ks2.keyleng = 20                        ' �L�[��
    PLN_tmpP_COMP_Speck.ks2.keyflag = BtKfExt + BtKfSeg         ' �L�[�t���O
    PLN_tmpP_COMP_Speck.ks2.keytype = Chr(BtKtString)           ' �L�[�^�C�v
    PLN_tmpP_COMP_Speck.ks2.reserve = &H0                       ' �\��ς�
                                                    
                                                    
    PLN_tmpP_COMP_Speck.ks3.keypos = 23                         ' �L�[�|�W�V����
    PLN_tmpP_COMP_Speck.ks3.keyleng = 2                         ' �L�[��
    PLN_tmpP_COMP_Speck.ks3.keyflag = BtKfExt + BtKfSeg         ' �L�[�t���O
    PLN_tmpP_COMP_Speck.ks3.keytype = Chr(BtKtString)           ' �L�[�^�C�v
    PLN_tmpP_COMP_Speck.ks3.reserve = &H0                       ' �\��ς�
                                                    
    PLN_tmpP_COMP_Speck.ks4.keypos = 25                         ' �L�[�|�W�V����
    PLN_tmpP_COMP_Speck.ks4.keyleng = 1                         ' �L�[��
    PLN_tmpP_COMP_Speck.ks4.keyflag = BtKfExt + BtKfSeg         ' �L�[�t���O
    PLN_tmpP_COMP_Speck.ks4.keytype = Chr(BtKtString)           ' �L�[�^�C�v
    PLN_tmpP_COMP_Speck.ks4.reserve = &H0                       ' �\��ς�
    
    PLN_tmpP_COMP_Speck.ks5.keypos = 26                         ' �L�[�|�W�V����
    PLN_tmpP_COMP_Speck.ks5.keyleng = 1                         ' �L�[��
    PLN_tmpP_COMP_Speck.ks5.keyflag = BtKfExt + BtKfSeg         ' �L�[�t���O
    PLN_tmpP_COMP_Speck.ks5.keytype = Chr(BtKtString)           ' �L�[�^�C�v
    PLN_tmpP_COMP_Speck.ks5.reserve = &H0                       ' �\��ς�
    
    PLN_tmpP_COMP_Speck.ks6.keypos = 27                         ' �L�[�|�W�V����
    PLN_tmpP_COMP_Speck.ks6.keyleng = 20                        ' �L�[��
    PLN_tmpP_COMP_Speck.ks6.keyflag = BtKfExt + BtKfSeg         ' �L�[�t���O
    PLN_tmpP_COMP_Speck.ks6.keytype = Chr(BtKtString)           ' �L�[�^�C�v
    PLN_tmpP_COMP_Speck.ks6.reserve = &H0                       ' �\��ς�
                                                    
    PLN_tmpP_COMP_Speck.ks7.keypos = 47                         ' �L�[�|�W�V����
    PLN_tmpP_COMP_Speck.ks7.keyleng = 8                         ' �L�[��
    PLN_tmpP_COMP_Speck.ks7.keyflag = BtKfExt                   ' �L�[�t���O
    PLN_tmpP_COMP_Speck.ks7.keytype = Chr(BtKtString)           ' �L�[�^�C�v
    PLN_tmpP_COMP_Speck.ks7.reserve = &H0                       ' �\��ς�
                                                    
'---------------------------------------------------'
                                                    ' �L�[�P
    PLN_tmpP_COMP_Speck.ks8.keypos = 47                        ' �L�[�|�W�V����
    PLN_tmpP_COMP_Speck.ks8.keyleng = 8                        ' �L�[��
    PLN_tmpP_COMP_Speck.ks8.keyflag = BtKfExt + BtKfSeg         ' �L�[�t���O
    PLN_tmpP_COMP_Speck.ks8.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    PLN_tmpP_COMP_Speck.ks8.reserve = &H0                      ' �\��ς�
    
    
    PLN_tmpP_COMP_Speck.ks9.keypos = 23                         ' �L�[�|�W�V����
    PLN_tmpP_COMP_Speck.ks9.keyleng = 2                         ' �L�[��
    PLN_tmpP_COMP_Speck.ks9.keyflag = BtKfExt + BtKfSeg         ' �L�[�t���O
    PLN_tmpP_COMP_Speck.ks9.keytype = Chr(BtKtString)           ' �L�[�^�C�v
    PLN_tmpP_COMP_Speck.ks9.reserve = &H0                       ' �\��ς�
                                                    
    PLN_tmpP_COMP_Speck.ks10.keypos = 25                         ' �L�[�|�W�V����
    PLN_tmpP_COMP_Speck.ks10.keyleng = 1                         ' �L�[��
    PLN_tmpP_COMP_Speck.ks10.keyflag = BtKfExt + BtKfSeg         ' �L�[�t���O
    PLN_tmpP_COMP_Speck.ks10.keytype = Chr(BtKtString)           ' �L�[�^�C�v
    PLN_tmpP_COMP_Speck.ks10.reserve = &H0                       ' �\��ς�
    
    PLN_tmpP_COMP_Speck.ks11.keypos = 26                        ' �L�[�|�W�V����
    PLN_tmpP_COMP_Speck.ks11.keyleng = 1                        ' �L�[��
    PLN_tmpP_COMP_Speck.ks11.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    PLN_tmpP_COMP_Speck.ks11.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    PLN_tmpP_COMP_Speck.ks11.reserve = &H0                      ' �\��ς�
    
    PLN_tmpP_COMP_Speck.ks12.keypos = 27                        ' �L�[�|�W�V����
    PLN_tmpP_COMP_Speck.ks12.keyleng = 20                       ' �L�[��
    PLN_tmpP_COMP_Speck.ks12.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    PLN_tmpP_COMP_Speck.ks12.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    PLN_tmpP_COMP_Speck.ks12.reserve = &H0                      ' �\��ς�
                                                    
    PLN_tmpP_COMP_Speck.ks13.keypos = 1                         ' �L�[�|�W�V����
    PLN_tmpP_COMP_Speck.ks13.keyleng = 1                        ' �L�[��
    PLN_tmpP_COMP_Speck.ks13.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    PLN_tmpP_COMP_Speck.ks13.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    PLN_tmpP_COMP_Speck.ks13.reserve = &H0                      ' �\��ς�
                                                    
    PLN_tmpP_COMP_Speck.ks14.keypos = 2                         ' �L�[�|�W�V����
    PLN_tmpP_COMP_Speck.ks14.keyleng = 1                        ' �L�[��
    PLN_tmpP_COMP_Speck.ks14.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    PLN_tmpP_COMP_Speck.ks14.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    PLN_tmpP_COMP_Speck.ks14.reserve = &H0                      ' �\��ς�
                                                    
    PLN_tmpP_COMP_Speck.ks15.keypos = 3                         ' �L�[�|�W�V����
    PLN_tmpP_COMP_Speck.ks15.keyleng = 20                       ' �L�[��
    PLN_tmpP_COMP_Speck.ks15.keyflag = BtKfExt                  ' �L�[�t���O
    PLN_tmpP_COMP_Speck.ks15.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    PLN_tmpP_COMP_Speck.ks15.reserve = &H0                      ' �\��ς�
                                                    
                                                    
                                                    
'---------------------------------------------------'
    sts = BTRV(BtOpCreate, PLN_tmpP_COMP_POS, PLN_tmpP_COMP_Speck, Len(PLN_tmpP_COMP_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���ޏ��v�ʒ��ԃt�@�C��")
        Exit Function
    End If
    PLN_tmpP_COMP_Create = False
End Function
Public Function PLN_tmpP_COMP_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���ޏ��v�ʒ��ԃt�@�C���@�n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    PLN_tmpP_COMP_Open = True
                                            '���ޏ��v�ʒ��ԃt�@�C���@�t���p�X�捞��
    sts = GetIni("FILE", PLN_tmpP_COMP_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [PLN_P_COMP]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, PLN_tmpP_COMP_POS, PLN_tmpP_COMP_REC, Len(PLN_tmpP_COMP_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = PLN_tmpP_COMP_Create()   '���ޏ��v�ʒ��ԃt�@�C���@�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, PLN_tmpP_COMP_POS, PLN_tmpP_COMP_REC, Len(PLN_tmpP_COMP_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���ޏ��v�ʒ��ԃt�@�C��")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���ޏ��v�ʒ��ԃt�@�C��")
                Exit Function
        End Select
    Loop
    PLN_tmpP_COMP_Open = False

End Function
