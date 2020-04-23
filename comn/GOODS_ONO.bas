Attribute VB_Name = "GOODS_ONO"
Option Explicit
'********************************************************************
'*
'*              ���i���W�v�t�@�C���i�ꎞ�t�@�C���j �t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const GOODS_ONO_ID$ = "GOODS_ONO"

'�y�[�W�T�C�Y
Public Const GOODS_ONO_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public GOODS_ONO_POS As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
Type GOODS_ONOREC_Tag
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j
    ST_SOKO(0 To 1)         As Byte     '�W���I�� �q��
    ST_RETU(0 To 1)         As Byte     '�W���I�� ��
    ST_REN(0 To 1)          As Byte     '�W���I�� �A
    ST_DAN(0 To 1)          As Byte     '�W���I�� �i
    PACKING_NO(0 To 3)      As Byte     '����
    Sumi_QTY(0 To 7)        As Byte     '���i���ςݍ݌ɐ�
    Mi_QTY(0 To 7)          As Byte     '�����i�݌ɐ�
    AVE_SYUKA(0 To 7)       As Byte     '���Ϗo�א�
    SUMI_PERCENT(0 To 7)    As Byte     '���O���i����
End Type

'�f�[�^�E�o�b�t�@
Public GOODS_ONOREC         As GOODS_ONOREC_Tag

'�L�[��`
Type KEY0_GOODS_ONO                 '�j�d�x�O
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j
End Type


Type KEY1_GOODS_ONO                 '�j�d�x�P
    
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    AVE_SYUKA(0 To 7)       As Byte     '���Ϗo�א�
    Sumi_QTY(0 To 7)        As Byte     '���i���ςݍ݌ɐ�
    Mi_QTY(0 To 7)          As Byte     '�����i�݌ɐ�
    SUMI_PERCENT(0 To 7)    As Byte     '���O���i����
    ST_SOKO(0 To 1)         As Byte     '�W���I��
    ST_RETU(0 To 1)         As Byte     '�W���I�� ��
    ST_REN(0 To 1)          As Byte     '�W���I�� �A
    ST_DAN(0 To 1)          As Byte     '�W���I�� �i
    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j

End Type


'�L�[�E�f�[�^
Public K0_GOODS_ONO     As KEY0_GOODS_ONO
Public K1_GOODS_ONO     As KEY1_GOODS_ONO

Type GOODS_ONO_FSpeck
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
End Type

Private GOODS_ONO_Speck As GOODS_ONO_FSpeck
Private Function GOODS_ONO_Create() As Integer
'********************************************************************
'*
'*              ���i���W�v�t�@�C���@�b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    GOODS_ONO_Create = True
                                            '���i���W�v�t�@�C���@�t���p�X�捞��
    sts = GetIni("FILE", GOODS_ONO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [GOODS_ONO]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    GOODS_ONO_Speck.fs.recoleng = Len(GOODS_ONOREC) ' ���R�[�h��
    GOODS_ONO_Speck.fs.PageSize = GOODS_ONO_PG_SIZ  ' �y�[�W�T�C�Y
    GOODS_ONO_Speck.fs.idexnumb = 2                 ' �C���f�b�N�X��
    GOODS_ONO_Speck.fs.fileflag = 0                 ' �t�@�C���t���O
    GOODS_ONO_Speck.fs.reserve = &H0                ' �\��ς�
'---------------------------------------------------'
                                                    ' �L�[�O
    GOODS_ONO_Speck.ks0.keypos = 1                  ' �L�[�|�W�V����
    GOODS_ONO_Speck.ks0.keyleng = 1                 ' �L�[��
    GOODS_ONO_Speck.ks0.keyflag = BtKfExt + BtKfSeg ' �L�[�t���O
    GOODS_ONO_Speck.ks0.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    GOODS_ONO_Speck.ks0.reserve = &H0               ' �\��ς�
                                                    
    GOODS_ONO_Speck.ks1.keypos = 2                  ' �L�[�|�W�V����
    GOODS_ONO_Speck.ks1.keyleng = 1                 ' �L�[��
    GOODS_ONO_Speck.ks1.keyflag = BtKfExt + BtKfSeg ' �L�[�t���O
    GOODS_ONO_Speck.ks1.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    GOODS_ONO_Speck.ks1.reserve = &H0               ' �\��ς�
                                                    
                                                    
    GOODS_ONO_Speck.ks2.keypos = 3                  ' �L�[�|�W�V����
    GOODS_ONO_Speck.ks2.keyleng = 20                ' �L�[��
    GOODS_ONO_Speck.ks2.keyflag = BtKfExt           ' �L�[�t���O
    GOODS_ONO_Speck.ks2.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    GOODS_ONO_Speck.ks2.reserve = &H0               ' �\��ς�


'---------------------------------------------------'
                                                    ' �L�[�P
    GOODS_ONO_Speck.ks3.keypos = 1                      ' �L�[�|�W�V����
    GOODS_ONO_Speck.ks3.keyleng = 1                     ' �L�[��
    GOODS_ONO_Speck.ks3.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    GOODS_ONO_Speck.ks3.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    GOODS_ONO_Speck.ks3.reserve = &H0                   ' �\��ς�
                                                    
    GOODS_ONO_Speck.ks4.keypos = 2                      ' �L�[�|�W�V����
    GOODS_ONO_Speck.ks4.keyleng = 1                     ' �L�[��
    GOODS_ONO_Speck.ks4.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    GOODS_ONO_Speck.ks4.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    GOODS_ONO_Speck.ks4.reserve = &H0                   ' �\��ς�
    
    GOODS_ONO_Speck.ks5.keypos = 51                    ' �L�[�|�W�V����
    GOODS_ONO_Speck.ks5.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    GOODS_ONO_Speck.ks5.keyflag = BtKfExt + BtKfSeg + BtKfDec
    GOODS_ONO_Speck.ks5.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    GOODS_ONO_Speck.ks5.reserve = &H0                  ' �\��ς�
                                                    
    GOODS_ONO_Speck.ks6.keypos = 35                    ' �L�[�|�W�V����
    GOODS_ONO_Speck.ks6.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    GOODS_ONO_Speck.ks6.keyflag = BtKfExt + BtKfSeg
    GOODS_ONO_Speck.ks6.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    GOODS_ONO_Speck.ks6.reserve = &H0                  ' �\��ς�
                                                    
    GOODS_ONO_Speck.ks7.keypos = 43                    ' �L�[�|�W�V����
    GOODS_ONO_Speck.ks7.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    GOODS_ONO_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfDec
    GOODS_ONO_Speck.ks7.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    GOODS_ONO_Speck.ks7.reserve = &H0                  ' �\��ς�
                                                    
    GOODS_ONO_Speck.ks8.keypos = 59                    ' �L�[�|�W�V����
    GOODS_ONO_Speck.ks8.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    GOODS_ONO_Speck.ks8.keyflag = BtKfExt + BtKfSeg
    GOODS_ONO_Speck.ks8.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    GOODS_ONO_Speck.ks8.reserve = &H0                  ' �\��ς�
    
    GOODS_ONO_Speck.ks9.keypos = 23                     ' �L�[�|�W�V����
    GOODS_ONO_Speck.ks9.keyleng = 2                     ' �L�[��
    GOODS_ONO_Speck.ks9.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    GOODS_ONO_Speck.ks9.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    GOODS_ONO_Speck.ks9.reserve = &H0                   ' �\��ς�
    
    GOODS_ONO_Speck.ks10.keypos = 25                     ' �L�[�|�W�V����
    GOODS_ONO_Speck.ks10.keyleng = 2                     ' �L�[��
    GOODS_ONO_Speck.ks10.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    GOODS_ONO_Speck.ks10.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    GOODS_ONO_Speck.ks10.reserve = &H0                   ' �\��ς�
    
    GOODS_ONO_Speck.ks11.keypos = 27                     ' �L�[�|�W�V����
    GOODS_ONO_Speck.ks11.keyleng = 2                     ' �L�[��
    GOODS_ONO_Speck.ks11.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    GOODS_ONO_Speck.ks11.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    GOODS_ONO_Speck.ks11.reserve = &H0                   ' �\��ς�
    
    GOODS_ONO_Speck.ks12.keypos = 29                    ' �L�[�|�W�V����
    GOODS_ONO_Speck.ks12.keyleng = 2                    ' �L�[��
    GOODS_ONO_Speck.ks12.keyflag = BtKfExt + BtKfSeg    ' �L�[�t���O
    GOODS_ONO_Speck.ks12.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    GOODS_ONO_Speck.ks12.reserve = &H0                  ' �\��ς�
                                                    
                                                    
    GOODS_ONO_Speck.ks13.keypos = 3                      ' �L�[�|�W�V����
    GOODS_ONO_Speck.ks13.keyleng = 20                    ' �L�[��
    GOODS_ONO_Speck.ks13.keyflag = BtKfExt               ' �L�[�t���O
    GOODS_ONO_Speck.ks13.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    GOODS_ONO_Speck.ks13.reserve = &H0                   ' �\��ς�

'---------------------------------------------------'
    sts = BTRV(BtOpCreate, GOODS_ONO_POS, GOODS_ONO_Speck, Len(GOODS_ONO_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���i���W�v�t�@�C��")
        Exit Function
    End If
    
    GOODS_ONO_Create = False

End Function
Public Function GOODS_ONO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���i���W�v�t�@�C���@�n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    GOODS_ONO_Open = True
                                            '���i���W�v�t�@�C���@�t���p�X�捞��
    sts = GetIni("FILE", GOODS_ONO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [GOODS_ONO]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, GOODS_ONO_POS, GOODS_ONOREC, Len(GOODS_ONOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = GOODS_ONO_Create()    '���i���W�v�t�@�C���@�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, GOODS_ONO_POS, GOODS_ONOREC, Len(GOODS_ONOREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���i���W�v�t�@�C��")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���i���W�v�t�@�C��")
                Exit Function
        End Select
    Loop
    GOODS_ONO_Open = False

End Function

