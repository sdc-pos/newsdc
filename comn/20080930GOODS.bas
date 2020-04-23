Attribute VB_Name = "GOODS"
Option Explicit
'********************************************************************
'*
'*              ���i���W�v�t�@�C���i�ꎞ�t�@�C���j �t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const GOODS_ID$ = "GOODS"

'�y�[�W�T�C�Y
Public Const GOODS_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public GOODS_POS As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
Type GOODSREC_Tag
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
Public GOODSREC             As GOODSREC_Tag

'�L�[��`
Type KEY0_GOODS                    '�j�d�x�O
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    ST_SOKO(0 To 1)         As Byte     '�W���I��
    SUMI_PERCENT(0 To 7)    As Byte     '���O���i����
    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j
End Type


Type KEY1_GOODS                    '�j�d�x�P
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    ST_SOKO(0 To 1)         As Byte     '�W���I��
    ST_RETU(0 To 1)         As Byte     '�W���I�� ��
    ST_REN(0 To 1)          As Byte     '�W���I�� �A
    ST_DAN(0 To 1)          As Byte     '�W���I�� �i
    SUMI_PERCENT(0 To 7)    As Byte     '���O���i����
    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j
End Type

Type KEY2_GOODS                    '�j�d�x�Q    2007.11.14
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j
End Type



'�L�[�E�f�[�^
Public K0_GOODS         As KEY0_GOODS
Public K1_GOODS         As KEY1_GOODS
Public K2_GOODS         As KEY2_GOODS

Type GOODS_FSpeck
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
End Type

Private GOODS_Speck As GOODS_FSpeck
Private Function GOODS_Create() As Integer
'********************************************************************
'*
'*              ���i���W�v�t�@�C���@�b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*      2007.11.14  :KEY2(���ƕ�+�����O+�i��(�O))�@�ǉ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    GOODS_Create = True
                                            '���i���W�v�t�@�C���@�t���p�X�捞��
    sts = GetIni("FILE", GOODS_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [GOODS]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    GOODS_Speck.fs.recoleng = Len(GOODSREC)         ' ���R�[�h��
    GOODS_Speck.fs.PageSize = GOODS_PG_SIZ          ' �y�[�W�T�C�Y
    GOODS_Speck.fs.idexnumb = 3                     ' �C���f�b�N�X��
    GOODS_Speck.fs.fileflag = 0                     ' �t�@�C���t���O
    GOODS_Speck.fs.reserve = &H0                    ' �\��ς�
'---------------------------------------------------'
                                                    ' �L�[�O
    GOODS_Speck.ks0.keypos = 1                      ' �L�[�|�W�V����
    GOODS_Speck.ks0.keyleng = 1                     ' �L�[��
    GOODS_Speck.ks0.keyflag = BtKfExt + BtKfChg + BtKfSeg   ' �L�[�t���O
    GOODS_Speck.ks0.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    GOODS_Speck.ks0.reserve = &H0                   ' �\��ς�
                                                    
    GOODS_Speck.ks1.keypos = 2                      ' �L�[�|�W�V����
    GOODS_Speck.ks1.keyleng = 1                     ' �L�[��
    GOODS_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfSeg     ' �L�[�t���O
    GOODS_Speck.ks1.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    GOODS_Speck.ks1.reserve = &H0                   ' �\��ς�
                                                    
    GOODS_Speck.ks2.keypos = 23                     ' �L�[�|�W�V����
    GOODS_Speck.ks2.keyleng = 2                     ' �L�[��
    GOODS_Speck.ks2.keyflag = BtKfExt + BtKfChg + BtKfSeg     ' �L�[�t���O
    GOODS_Speck.ks2.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    GOODS_Speck.ks2.reserve = &H0                   ' �\��ς�
                                                    
    GOODS_Speck.ks3.keypos = 59                     ' �L�[�|�W�V����
    GOODS_Speck.ks3.keyleng = 8                     ' �L�[��
    GOODS_Speck.ks3.keyflag = BtKfExt + BtKfChg + BtKfSeg     ' �L�[�t���O
    GOODS_Speck.ks3.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    GOODS_Speck.ks3.reserve = &H0                   ' �\��ς�
                                                    
    GOODS_Speck.ks4.keypos = 3                      ' �L�[�|�W�V����
    GOODS_Speck.ks4.keyleng = 20                    ' �L�[��
    GOODS_Speck.ks4.keyflag = BtKfExt + BtKfChg               ' �L�[�t���O
    GOODS_Speck.ks4.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    GOODS_Speck.ks4.reserve = &H0                   ' �\��ς�


'---------------------------------------------------'
                                                    ' �L�[�P
    GOODS_Speck.ks5.keypos = 1                      ' �L�[�|�W�V����
    GOODS_Speck.ks5.keyleng = 1                     ' �L�[��
    GOODS_Speck.ks5.keyflag = BtKfExt + BtKfChg + BtKfSeg     ' �L�[�t���O
    GOODS_Speck.ks5.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    GOODS_Speck.ks5.reserve = &H0                   ' �\��ς�
                                                    
    GOODS_Speck.ks6.keypos = 2                      ' �L�[�|�W�V����
    GOODS_Speck.ks6.keyleng = 1                     ' �L�[��
    GOODS_Speck.ks6.keyflag = BtKfExt + BtKfChg + BtKfSeg     ' �L�[�t���O
    GOODS_Speck.ks6.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    GOODS_Speck.ks6.reserve = &H0                   ' �\��ς�
                                                    
    GOODS_Speck.ks7.keypos = 23                     ' �L�[�|�W�V����
    GOODS_Speck.ks7.keyleng = 8                     ' �L�[��
    GOODS_Speck.ks7.keyflag = BtKfExt + BtKfChg + BtKfSeg     ' �L�[�t���O
    GOODS_Speck.ks7.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    GOODS_Speck.ks7.reserve = &H0                   ' �\��ς�
                                                    
    GOODS_Speck.ks8.keypos = 59                     ' �L�[�|�W�V����
    GOODS_Speck.ks8.keyleng = 8                     ' �L�[��
    GOODS_Speck.ks8.keyflag = BtKfExt + BtKfChg + BtKfSeg     ' �L�[�t���O
    GOODS_Speck.ks8.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    GOODS_Speck.ks8.reserve = &H0                   ' �\��ς�
                                                    
    GOODS_Speck.ks9.keypos = 3                      ' �L�[�|�W�V����
    GOODS_Speck.ks9.keyleng = 20                    ' �L�[��
    GOODS_Speck.ks9.keyflag = BtKfExt + BtKfChg               ' �L�[�t���O
    GOODS_Speck.ks9.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    GOODS_Speck.ks9.reserve = &H0                   ' �\��ς�

'---------------------------------------------------'
    
    
    
'---------------------------------------------------'
                                                    ' �L�[�P
    GOODS_Speck.ks10.keypos = 1                      ' �L�[�|�W�V����
    GOODS_Speck.ks10.keyleng = 1                     ' �L�[��
    GOODS_Speck.ks10.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    GOODS_Speck.ks10.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    GOODS_Speck.ks10.reserve = &H0                   ' �\��ς�
                                                    
    GOODS_Speck.ks11.keypos = 2                      ' �L�[�|�W�V����
    GOODS_Speck.ks11.keyleng = 1                     ' �L�[��
    GOODS_Speck.ks11.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    GOODS_Speck.ks11.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    GOODS_Speck.ks11.reserve = &H0                   ' �\��ς�
    
    GOODS_Speck.ks12.keypos = 3                      ' �L�[�|�W�V����
    GOODS_Speck.ks12.keyleng = 20                    ' �L�[��
    GOODS_Speck.ks12.keyflag = BtKfExt               ' �L�[�t���O
    GOODS_Speck.ks12.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    GOODS_Speck.ks12.reserve = &H0                   ' �\��ς�
'---------------------------------------------------'
    
    
    
    sts = BTRV(BtOpCreate, GOODS_POS, GOODS_Speck, Len(GOODS_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���i���W�v�t�@�C��")
        Exit Function
    End If
    
    GOODS_Create = False

End Function
Public Function GOODS_Open(Mode As Integer) As Integer
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
    
    GOODS_Open = True
                                            '���i���W�v�t�@�C���@�t���p�X�捞��
    sts = GetIni("FILE", GOODS_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [GOODS]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    
    
    sts = BTRV(BtOpClose, GOODS_POS, GOODSREC, Len(GOODSREC), K0_GOODS, Len(K0_GOODS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���i���W�v�t�@�C��")
        End If
    End If
    
    
    On Error Resume Next    '2007.11.14
    Kill (FullPath)         '2007.11.14
    On Error GoTo 0         '2007.11.14
    
    
    Do
        sts = BTRV(BtOpOpen, GOODS_POS, GOODSREC, Len(GOODSREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = GOODS_Create()        '���i���W�v�t�@�C���@�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, GOODS_POS, GOODSREC, Len(GOODSREC), ByVal FullPath, Len(FullPath), Mode)
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
    GOODS_Open = False

End Function

