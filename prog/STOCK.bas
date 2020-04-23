Attribute VB_Name = "STOCK"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �I�����f�[�^  �t�@�C����`                          *
'*                                                                  *
'********************************************************************
'�t�@�C���h�c
Public Const STOCK_ID$ = "STOCK"

'�y�[�W�T�C�Y
Public Const STOCK_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public STOCK_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type STOCKREC_Tag
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j
    
    ST_SOKO(0 To 1)         As Byte     '�W�����ɑq��
    ST_RETU(0 To 1)         As Byte     '�W�����ɑq��
    ST_REN(0 To 1)          As Byte     '�W�����ɑq��
    ST_DAN(0 To 1)          As Byte     '�W�����ɑq��
    
    
    
    HOST_ZAIKO(0 To 7)      As Byte     '�������_�݌�
    POS_ZAIKO(0 To 7)       As Byte     '�o�n�r���݌�
    ST_ZAIKO(0 To 7)        As Byte     '�W���I�ԍ݌�
    
    EE1_LOCATION(0 To 7)    As Byte     '�ʒu���P
    EE1_ZAIKO(0 To 7)       As Byte     '�݌�
    EE2_LOCATION(0 To 7)    As Byte     '�ʒu���Q
    EE2_ZAIKO(0 To 7)       As Byte     '�݌�
    EE3_LOCATION(0 To 7)    As Byte     '�ʒu���R
    EE3_ZAIKO(0 To 7)       As Byte     '�݌�
    
    ETC_ZAIKO(0 To 7)       As Byte     '���̑��݌�
    CHECK_MARK(0 To 0)      As Byte     '�ƍ��}�[�N
    PRINT_YMD(0 To 7)       As Byte     '������t
    INPUT_YMD(0 To 7)       As Byte     '���͓��t
    
    SAI_QTY(0 To 8)         As Byte     '���ِ��@2004.06.29
    
    BU_ZAI_QTY(0 To 7)      As Byte     'BU�݌�     2007.08.22
    PPSC_ZAI_QTY(0 To 7)    As Byte     'PPSC�݌�   2007.08.22
    
    
    
    FILLER(0 To 7)          As Byte
    
End Type
'�f�[�^�E�o�b�t�@
Public STOCKREC As STOCKREC_Tag

'�L�[��`

Type KEY0_STOCK             '�j�d�x�O
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j
End Type

Type KEY1_STOCK             '�j�d�x�P
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    
    ST_SOKO(0 To 1)         As Byte     '�W�����ɑq��
    ST_RETU(0 To 1)         As Byte     '�W�����ɑq��
    ST_REN(0 To 1)          As Byte     '�W�����ɑq��
    ST_DAN(0 To 1)          As Byte     '�W�����ɑq��
    
    
    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j
End Type

Type KEY2_STOCK             '�j�d�x�Q
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    ST_SOKO(0 To 1)         As Byte     '�W�����ɑq��   2007.08.22
    CHECK_MARK(0 To 0)      As Byte     '�ƍ��}�[�N
End Type



'�SBU�p�@KEY��`
Type KEY3_STOCK             '�j�d�x�R
    NAIGAI(0 To 0)          As Byte     '�����O
    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j
End Type

Type KEY4_STOCK             '�j�d�x�S
    NAIGAI(0 To 0)          As Byte     '�����O
    
    ST_SOKO(0 To 1)         As Byte     '�W�����ɑq��
    ST_RETU(0 To 1)         As Byte     '�W�����ɑq��
    ST_REN(0 To 1)          As Byte     '�W�����ɑq��
    ST_DAN(0 To 1)          As Byte     '�W�����ɑq��
    
    
    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j
End Type

Type KEY5_STOCK             '�j�d�x�T
    NAIGAI(0 To 0)          As Byte     '�����O
    ST_SOKO(0 To 1)         As Byte     '�W�����ɑq��
    CHECK_MARK(0 To 0)      As Byte     '�ƍ��}�[�N
End Type




'�L�[�E�f�[�^
Public K0_STOCK     As KEY0_STOCK
Public K1_STOCK     As KEY1_STOCK
Public K2_STOCK     As KEY2_STOCK

Public K3_STOCK     As KEY3_STOCK
Public K4_STOCK     As KEY4_STOCK
Public K5_STOCK     As KEY5_STOCK


Private Type STOCK_FSpeck
    
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
    ks15    As BtKeySpeck
    ks16    As BtKeySpeck
    ks17    As BtKeySpeck
    ks18    As BtKeySpeck

End Type

Private STOCK_Speck As STOCK_FSpeck
Private Function STOCK_Create() As Integer
'********************************************************************
'*                                                                  *
'*              �I�����f�[�^  �b�q�d�`�s�d                          *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    STOCK_Create = True
                                        '�I�����f�[�^�t���p�X�捞��
    sts = GetIni("FILE", STOCK_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [STOCK]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    STOCK_Speck.fs.recoleng = Len(STOCKREC)     ' ���R�[�h��
    STOCK_Speck.fs.PageSize = STOCK_PG_SIZ      ' �y�[�W�T�C�Y
    
    STOCK_Speck.fs.idexnumb = 6                 ' �C���f�b�N�X��    �SBU�Ή�3-->6
    
    STOCK_Speck.fs.fileflag = 0                 ' �t�@�C���t���O
    STOCK_Speck.fs.reserve = &H0                ' �\��ς�
'------------------------------------------------
                                                ' �L�[�O
    STOCK_Speck.ks0.keypos = 1                  ' �L�[�|�W�V����
    STOCK_Speck.ks0.keyleng = 1                 ' �L�[��
    STOCK_Speck.ks0.keyflag = BtKfExt + BtKfSeg ' �L�[�t���O
    STOCK_Speck.ks0.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    STOCK_Speck.ks0.reserve = &H0               ' �\��ς�
                                                
    STOCK_Speck.ks1.keypos = 2                  ' �L�[�|�W�V����
    STOCK_Speck.ks1.keyleng = 1                 ' �L�[��
    STOCK_Speck.ks1.keyflag = BtKfExt + BtKfSeg ' �L�[�t���O
    STOCK_Speck.ks1.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    STOCK_Speck.ks1.reserve = &H0               ' �\��ς�
                                                
    STOCK_Speck.ks2.keypos = 3                  ' �L�[�|�W�V����
    STOCK_Speck.ks2.keyleng = 20                ' �L�[��
    STOCK_Speck.ks2.keyflag = BtKfExt           ' �L�[�t���O
    STOCK_Speck.ks2.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    STOCK_Speck.ks2.reserve = &H0               ' �\��ς�
'------------------------------------------------
                                                ' �L�[�P
    STOCK_Speck.ks3.keypos = 1                  ' �L�[�|�W�V����
    STOCK_Speck.ks3.keyleng = 1                 ' �L�[��
    STOCK_Speck.ks3.keyflag = BtKfExt + BtKfSeg ' �L�[�t���O
    STOCK_Speck.ks3.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    STOCK_Speck.ks3.reserve = &H0               ' �\��ς�
    
    STOCK_Speck.ks4.keypos = 2                  ' �L�[�|�W�V����
    STOCK_Speck.ks4.keyleng = 1                 ' �L�[��
    STOCK_Speck.ks4.keyflag = BtKfExt + BtKfSeg ' �L�[�t���O
    STOCK_Speck.ks4.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    STOCK_Speck.ks4.reserve = &H0               ' �\��ς�
                                                
    STOCK_Speck.ks5.keypos = 23                 ' �L�[�|�W�V����
    STOCK_Speck.ks5.keyleng = 8                 ' �L�[��
    STOCK_Speck.ks5.keyflag = BtKfExt + BtKfSeg ' �L�[�t���O
    STOCK_Speck.ks5.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    STOCK_Speck.ks5.reserve = &H0               ' �\��ς�
                                                
    STOCK_Speck.ks6.keypos = 3                  ' �L�[�|�W�V����
    STOCK_Speck.ks6.keyleng = 20                ' �L�[��
    STOCK_Speck.ks6.keyflag = BtKfExt           ' �L�[�t���O
    STOCK_Speck.ks6.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    STOCK_Speck.ks6.reserve = &H0               ' �\��ς�
'------------------------------------------------
                                                ' �L�[�Q
    STOCK_Speck.ks7.keypos = 1                  ' �L�[�|�W�V����
    STOCK_Speck.ks7.keyleng = 1                 ' �L�[��
    STOCK_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfChg + BtKfDup
    STOCK_Speck.ks7.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    STOCK_Speck.ks7.reserve = &H0               ' �\��ς�
    
    STOCK_Speck.ks8.keypos = 2                  ' �L�[�|�W�V����
    STOCK_Speck.ks8.keyleng = 1                 ' �L�[��
    STOCK_Speck.ks8.keyflag = BtKfExt + BtKfSeg + BtKfChg + BtKfDup
    STOCK_Speck.ks8.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    STOCK_Speck.ks8.reserve = &H0               ' �\��ς�
    
                                                
    STOCK_Speck.ks9.keypos = 23                 ' �L�[�|�W�V����
    STOCK_Speck.ks9.keyleng = 2                 ' �L�[��
    STOCK_Speck.ks9.keyflag = BtKfExt + BtKfSeg + BtKfChg + BtKfDup
    STOCK_Speck.ks9.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    STOCK_Speck.ks9.reserve = &H0               ' �\��ς�
    
    
    STOCK_Speck.ks10.keypos = 111                ' �L�[�|�W�V����
    STOCK_Speck.ks10.keyleng = 1                 ' �L�[��
    STOCK_Speck.ks10.keyflag = BtKfExt + BtKfChg + BtKfDup
    STOCK_Speck.ks10.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    STOCK_Speck.ks10.reserve = &H0               ' �\��ς�
'------------------------------------------------
    
    
    
    
    
    
    
    
    
'------------------------------------------------
                                                ' �L�[�R
                                                
    STOCK_Speck.ks11.keypos = 2                 ' �L�[�|�W�V����
    STOCK_Speck.ks11.keyleng = 1                ' �L�[��
    STOCK_Speck.ks11.keyflag = BtKfExt + BtKfSeg + BtKfChg + BtKfDup
    STOCK_Speck.ks11.keytype = Chr(BtKtString)  ' �L�[�^�C�v
    STOCK_Speck.ks11.reserve = &H0              ' �\��ς�
                                                
    STOCK_Speck.ks12.keypos = 3                 ' �L�[�|�W�V����
    STOCK_Speck.ks12.keyleng = 20               ' �L�[��
    STOCK_Speck.ks12.keyflag = BtKfExt + BtKfChg + BtKfDup
    STOCK_Speck.ks12.keytype = Chr(BtKtString)  ' �L�[�^�C�v
    STOCK_Speck.ks12.reserve = &H0              ' �\��ς�
'------------------------------------------------
                                                ' �L�[�S
    
    STOCK_Speck.ks13.keypos = 2                  ' �L�[�|�W�V����
    STOCK_Speck.ks13.keyleng = 1                 ' �L�[��
    STOCK_Speck.ks13.keyflag = BtKfExt + BtKfSeg + BtKfChg + BtKfDup
    STOCK_Speck.ks13.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    STOCK_Speck.ks13.reserve = &H0               ' �\��ς�
                                                
    STOCK_Speck.ks14.keypos = 23                 ' �L�[�|�W�V����
    STOCK_Speck.ks14.keyleng = 8                 ' �L�[��
    STOCK_Speck.ks14.keyflag = BtKfExt + BtKfSeg + BtKfChg + BtKfDup
    STOCK_Speck.ks14.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    STOCK_Speck.ks14.reserve = &H0               ' �\��ς�
                                                
    STOCK_Speck.ks15.keypos = 3                  ' �L�[�|�W�V����
    STOCK_Speck.ks15.keyleng = 20                ' �L�[��
    STOCK_Speck.ks15.keyflag = BtKfExt + BtKfChg + BtKfDup
    STOCK_Speck.ks15.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    STOCK_Speck.ks15.reserve = &H0               ' �\��ς�
'------------------------------------------------
                                                ' �L�[�T
    
    STOCK_Speck.ks16.keypos = 2                  ' �L�[�|�W�V����
    STOCK_Speck.ks16.keyleng = 1                 ' �L�[��
    STOCK_Speck.ks16.keyflag = BtKfExt + BtKfSeg + BtKfChg + BtKfDup
    STOCK_Speck.ks16.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    STOCK_Speck.ks16.reserve = &H0               ' �\��ς�
    
                                                
    STOCK_Speck.ks17.keypos = 23                 ' �L�[�|�W�V����
    STOCK_Speck.ks17.keyleng = 2                 ' �L�[��
    STOCK_Speck.ks17.keyflag = BtKfExt + BtKfSeg + BtKfChg + BtKfDup
    STOCK_Speck.ks17.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    STOCK_Speck.ks17.reserve = &H0               ' �\��ς�

    
    STOCK_Speck.ks18.keypos = 111                ' �L�[�|�W�V����
    STOCK_Speck.ks18.keyleng = 1                 ' �L�[��
    STOCK_Speck.ks18.keyflag = BtKfExt + BtKfChg + BtKfDup
    STOCK_Speck.ks18.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    STOCK_Speck.ks18.reserve = &H0               ' �\��ς�
'------------------------------------------------
    
    
    
    
    
    
    sts = BTRV(BtOpCreate, STOCK_POS, STOCK_Speck, Len(STOCK_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�I�����f�[�^")
        Exit Function
    End If

    STOCK_Create = False

End Function

Function STOCK_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              �I�����f�[�^  �n�o�d�m                              *
'*                                                                  *
'*      ��  ��:Open Mode(Btrieve�Q��)                               *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    STOCK_Open = True
                                    '�I�����f�[�^�t���p�X�捞��
    sts = GetIni("FILE", STOCK_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [STOCK]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, STOCK_POS, STOCKREC, Len(STOCKREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = STOCK_Create()        '�I�����f�[�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, STOCK_POS, STOCKREC, Len(STOCKREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�I�����f�[�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�I�����f�[�^")
                Exit Function
        End Select
    Loop
    
    STOCK_Open = False

End Function
