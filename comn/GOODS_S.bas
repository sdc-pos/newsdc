Attribute VB_Name = "GOODS_S"
Option Explicit
'********************************************************************
'*
'*              ���i���W�v�t�@�C���i�ꎞ�t�@�C���j �t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const GOODS_S_ID$ = "GOODS_S"

'�y�[�W�T�C�Y
Public Const GOODS_S_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public GOODS_S_POS As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
Private Type GOODS_SREC_Tag
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j
    Soko_No(0 To 1)         As Byte     '���z�q�ɔԍ��i�݌ɒ��j
    ST_SOKO(0 To 1)         As Byte     '�W���I�� �q��
    ST_RETU(0 To 1)         As Byte     '�W���I�� ��
    ST_REN(0 To 1)          As Byte     '�W���I�� �A
    ST_DAN(0 To 1)          As Byte     '�W���I�� �i
    PACKING_NO(0 To 3)      As Byte     '����
    SOKO_QTY(0 To 7)        As Byte     '���z�q�ɕ��݌�
    Sumi_QTY(0 To 7)        As Byte     '���i���ςݍ݌ɐ�
    Mi_QTY(0 To 7)          As Byte     '�����i�݌ɐ�
    AVE_SYUKA(0 To 7)       As Byte     '���Ϗo�א�
    SUMI_PERCENT(0 To 7)    As Byte     '���O���i����

    KOSOU(0 To 19)          As Byte     '���� 2008.03.03
    GAISOU(0 To 19)         As Byte     '�O���� 2008.03.03


End Type

'�f�[�^�E�o�b�t�@
Public GOODS_SREC             As GOODS_SREC_Tag

'�L�[��`
Type KEY0_GOODS_S                   '�j�d�x�O
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j
    Soko_No(0 To 1)         As Byte     '���z�q�ɔԍ��i�݌ɒ��j
End Type

Type KEY1_GOODS_S                   '�j�d�x�P
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    Soko_No(0 To 1)         As Byte     '���z�q�ɔԍ��i�݌ɒ��j
    SUMI_PERCENT(0 To 7)    As Byte     '���O���i����
    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j
End Type


'�L�[�E�f�[�^
Public K0_GOODS_S         As KEY0_GOODS_S
Public K1_GOODS_S         As KEY1_GOODS_S

Type GOODS_S_FSpeck
    fs      As BtFileSpeck          ' ̧�� ��߯��\����
    ks0     As BtKeySpeck           ' �� ��߯��\����
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck
    ks4     As BtKeySpeck
    ks5     As BtKeySpeck
End Type

Private GOODS_S_Speck As GOODS_S_FSpeck
Private Function GOODS_S_Create() As Integer
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

    GOODS_S_Create = True
                                            '���i���W�v�t�@�C���@�t���p�X�捞��
    sts = GetIni("FILE", GOODS_S_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [GOODS_S]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    GOODS_S_Speck.fs.recoleng = Len(GOODS_SREC)     ' ���R�[�h��
    GOODS_S_Speck.fs.PageSize = GOODS_S_PG_SIZ      ' �y�[�W�T�C�Y
    GOODS_S_Speck.fs.idexnumb = 2                   ' �C���f�b�N�X��
    GOODS_S_Speck.fs.fileflag = 0                   ' �t�@�C���t���O
    GOODS_S_Speck.fs.reserve = &H0                  ' �\��ς�
'---------------------------------------------------'
                                                    ' �L�[�O
    GOODS_S_Speck.ks0.keypos = 1                    ' �L�[�|�W�V����
    GOODS_S_Speck.ks0.keyleng = 24                  ' �L�[��
    GOODS_S_Speck.ks0.keyflag = BtKfExt             ' �L�[�t���O
    GOODS_S_Speck.ks0.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    GOODS_S_Speck.ks0.reserve = &H0                 ' �\��ς�
'---------------------------------------------------'
                                                    ' �L�[�P
    GOODS_S_Speck.ks1.keypos = 1                    ' �L�[�|�W�V����
    GOODS_S_Speck.ks1.keyleng = 1                   ' �L�[��
    GOODS_S_Speck.ks1.keyflag = BtKfExt + BtKfSeg   ' �L�[�t���O
    GOODS_S_Speck.ks1.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    GOODS_S_Speck.ks1.reserve = &H0                 ' �\��ς�
                                                    
    GOODS_S_Speck.ks2.keypos = 2                    ' �L�[�|�W�V����
    GOODS_S_Speck.ks2.keyleng = 1                   ' �L�[��
    GOODS_S_Speck.ks2.keyflag = BtKfExt + BtKfSeg   ' �L�[�t���O
    GOODS_S_Speck.ks2.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    GOODS_S_Speck.ks2.reserve = &H0                 ' �\��ς�
                                                    
    GOODS_S_Speck.ks3.keypos = 23                   ' �L�[�|�W�V����
    GOODS_S_Speck.ks3.keyleng = 2                   ' �L�[��
    GOODS_S_Speck.ks3.keyflag = BtKfExt + BtKfSeg   ' �L�[�t���O
    GOODS_S_Speck.ks3.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    GOODS_S_Speck.ks3.reserve = &H0                 ' �\��ς�
                                                    
    GOODS_S_Speck.ks4.keypos = 69                   ' �L�[�|�W�V����
    GOODS_S_Speck.ks4.keyleng = 8                   ' �L�[��
    GOODS_S_Speck.ks4.keyflag = BtKfExt + BtKfSeg   ' �L�[�t���O
    GOODS_S_Speck.ks4.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    GOODS_S_Speck.ks4.reserve = &H0                 ' �\��ς�
                                                    
    GOODS_S_Speck.ks5.keypos = 3                    ' �L�[�|�W�V����
    GOODS_S_Speck.ks5.keyleng = 20                  ' �L�[��
    GOODS_S_Speck.ks5.keyflag = BtKfExt             ' �L�[�t���O
    GOODS_S_Speck.ks5.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    GOODS_S_Speck.ks5.reserve = &H0                 ' �\��ς�
'---------------------------------------------------'
    sts = BTRV(BtOpCreate, GOODS_S_POS, GOODS_S_Speck, Len(GOODS_S_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���i���W�v�t�@�C��")
        Exit Function
    End If
    
    GOODS_S_Create = False

End Function
Public Function GOODS_S_Open(Mode As Integer) As Integer
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
    
    GOODS_S_Open = True
                                            '���i���W�v�t�@�C���@�t���p�X�捞��
    sts = GetIni("FILE", GOODS_S_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [GOODS_S]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    
    
    sts = BTRV(BtOpClose, GOODS_S_POS, GOODS_SREC, Len(GOODS_SREC), K0_GOODS_S, Len(K0_GOODS_S), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���i���W�v�t�@�C��")
        End If
    End If
    
    
    On Error Resume Next    '2007.11.14
    Kill (FullPath)         '2007.11.14
    On Error GoTo 0         '2007.11.14
    
    
    
    
    Do
        sts = BTRV(BtOpOpen, GOODS_S_POS, GOODS_SREC, Len(GOODS_SREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = GOODS_S_Create()      '���i���W�v�t�@�C���@�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, GOODS_S_POS, GOODS_SREC, Len(GOODS_SREC), ByVal FullPath, Len(FullPath), Mode)
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
    GOODS_S_Open = False

End Function

