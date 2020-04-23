Attribute VB_Name = "SUMZ"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �݌ɏW�v�f�[�^�@�t�@�C����`                          *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
'�t�@�C���h�c
Public Const SUMZ_ID$ = "SUMZ"

'�y�[�W�T�C�Y
Public Const SUMZ_PG_SIZ% = 2048

'�|�W�V�����E�u���b�N
Public SUMZ_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                              *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type SUMZREC_Tag
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j
    ST_SOKO(0 To 1)         As Byte     '�W�����ɑq�� �q��
    ST_RETU(0 To 1)         As Byte     '             ��
    ST_REN(0 To 1)          As Byte     '             �A
    ST_DAN(0 To 1)          As Byte     '             �i
    T_Zai_Qty(0 To 7)       As Byte     '�݌ɑ���(����)
    ZEN_Zai_Qty(0 To 7)     As Byte     '�݌ɑ���(�O��)
    SYK_E_QTY(0 To 7)       As Byte     '�o�ɍςݐ�
    NYUKA_YQTY(0 To 7)      As Byte     '���ח\�萔
    HS_ZAIQTY(0 To 7)       As Byte     'νč݌ɐ�(����)
    ZEN_HS_ZAIQTY(0 To 7)   As Byte     'νč݌ɐ�(�O��)
    SAI_QTY(0 To 7)         As Byte     '���ِ�
    SUM_DT(0 To 7)          As Byte     '�W�v���t
    
    BU_ZAI_QTY(0 To 7)      As Byte     'BU�݌�
    PPSC_ZAI_QTY(0 To 7)    As Byte     'PPSC�݌�
    
    
    ZEN_SAI_QTY(0 To 7)     As Byte     '�O�����ِ� 2009.02.09
    SAI_YMD(0 To 7)         As Byte     '���ٔ����� 2009.02.09
    FILLER(0 To 1)          As Byte     'FILLER     2009.02.09
End Type

'�f�[�^�E�o�b�t�@
Public SUMZREC As SUMZREC_Tag

'�L�[��`
Private Type KEY0_SUMZ            '�j�d�x�O
    JGYOBU(0 To 0) As Byte          '���ƕ��敪
    NAIGAI(0 To 0) As Byte          '�����O
    HIN_GAI(0 To 19) As Byte        '�i�ԁi�O���j
End Type

Private Type KEY1_SUMZ            '�j�d�x�P
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    ST_SOKO(0 To 1)     As Byte     '�W�����ɑq�� �q��
    ST_RETU(0 To 1)     As Byte     '             ��
    ST_REN(0 To 1)      As Byte     '             �A
    ST_DAN(0 To 1)      As Byte     '             �i
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
End Type

'�L�[�E�f�[�^
Public K0_SUMZ As KEY0_SUMZ
Public K1_SUMZ As KEY1_SUMZ

Private Type SUMZ_FSpeck
    fs As BtFileSpeck               ' ̧�� ��߯��\����
    ks0 As BtKeySpeck               ' �� ��߯��\����
    ks1 As BtKeySpeck               ' �� ��߯��\����
    ks2 As BtKeySpeck               ' �� ��߯��\����
    ks3 As BtKeySpeck               ' �� ��߯��\����
    ks4 As BtKeySpeck               ' �� ��߯��\����
    ks5 As BtKeySpeck               ' �� ��߯��\����
    ks6 As BtKeySpeck               ' �� ��߯��\����
    ks7 As BtKeySpeck               ' �� ��߯��\����
    ks8 As BtKeySpeck               ' �� ��߯��\����
    ks9 As BtKeySpeck               ' �� ��߯��\����
End Type

Private SUMZ_Speck As SUMZ_FSpeck

Private Function SUMZ_Create() As Integer
'********************************************************************
'*                                                                  *
'*              �݌ɏW�v�f�[�^�@�b�q�d�`�s�d                        *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    SUMZ_Create = True
                                            '�݌ɏW�v�f�[�^�t���p�X�捞��
    sts = GetIni("FILE", SUMZ_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI[SUMZ] �ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    SUMZ_Speck.fs.recoleng = Len(SUMZREC)       ' ���R�[�h��
    SUMZ_Speck.fs.PageSize = SUMZ_PG_SIZ        ' �y�[�W�T�C�Y
    SUMZ_Speck.fs.idexnumb = 2                  ' �C���f�b�N�X��
    SUMZ_Speck.fs.fileflag = 0                  ' �t�@�C���t���O
    SUMZ_Speck.fs.reserve = &H0                 ' �\��ς�
'-----------------------------------------------' �L�[�O
    SUMZ_Speck.ks0.keypos = 1                   ' �L�[�|�W�V����
    SUMZ_Speck.ks0.keyleng = 1                  ' �L�[��
    SUMZ_Speck.ks0.keyflag = BtKfExt + BtKfSeg  ' �L�[�t���O
    SUMZ_Speck.ks0.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    SUMZ_Speck.ks0.reserve = &H0                ' �\��ς�

    SUMZ_Speck.ks1.keypos = 2                   ' �L�[�|�W�V����
    SUMZ_Speck.ks1.keyleng = 1                  ' �L�[��
    SUMZ_Speck.ks1.keyflag = BtKfExt + BtKfSeg  ' �L�[�t���O
    SUMZ_Speck.ks1.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    SUMZ_Speck.ks1.reserve = &H0                ' �\��ς�

    SUMZ_Speck.ks2.keypos = 3                   ' �L�[�|�W�V����
    SUMZ_Speck.ks2.keyleng = 20                 ' �L�[��
    SUMZ_Speck.ks2.keyflag = BtKfExt            ' �L�[�t���O
    SUMZ_Speck.ks2.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    SUMZ_Speck.ks2.reserve = &H0                ' �\��ς�
'-----------------------------------------------' �L�[�P
    SUMZ_Speck.ks3.keypos = 1                   ' �L�[�|�W�V����
    SUMZ_Speck.ks3.keyleng = 1                  ' �L�[��
                                                ' �L�[�t���O
    SUMZ_Speck.ks3.keyflag = BtKfExt + BtKfChg + BtKfSeg
    SUMZ_Speck.ks3.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    SUMZ_Speck.ks3.reserve = &H0                ' �\��ς�
    
    SUMZ_Speck.ks4.keypos = 2                   ' �L�[�|�W�V����
    SUMZ_Speck.ks4.keyleng = 1                  ' �L�[��
                                                ' �L�[�t���O
    SUMZ_Speck.ks4.keyflag = BtKfExt + BtKfChg + BtKfSeg
    SUMZ_Speck.ks4.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    SUMZ_Speck.ks4.reserve = &H0                ' �\��ς�
    
    SUMZ_Speck.ks5.keypos = 23                  ' �L�[�|�W�V����
    SUMZ_Speck.ks5.keyleng = 2                  ' �L�[��
                                                ' �L�[�t���O
    SUMZ_Speck.ks5.keyflag = BtKfExt + BtKfChg + BtKfSeg
    SUMZ_Speck.ks5.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    SUMZ_Speck.ks5.reserve = &H0                ' �\��ς�
    
    SUMZ_Speck.ks6.keypos = 25                  ' �L�[�|�W�V����
    SUMZ_Speck.ks6.keyleng = 2                  ' �L�[��
                                                ' �L�[�t���O
    SUMZ_Speck.ks6.keyflag = BtKfExt + BtKfChg + BtKfSeg
    SUMZ_Speck.ks6.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    SUMZ_Speck.ks6.reserve = &H0                ' �\��ς�
    
    SUMZ_Speck.ks7.keypos = 27                  ' �L�[�|�W�V����
    SUMZ_Speck.ks7.keyleng = 2                  ' �L�[��
                                                ' �L�[�t���O
    SUMZ_Speck.ks7.keyflag = BtKfExt + BtKfChg + BtKfSeg
    SUMZ_Speck.ks7.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    SUMZ_Speck.ks7.reserve = &H0                ' �\��ς�
    
    SUMZ_Speck.ks8.keypos = 29                  ' �L�[�|�W�V����
    SUMZ_Speck.ks8.keyleng = 2                  ' �L�[��
                                                ' �L�[�t���O
    SUMZ_Speck.ks8.keyflag = BtKfExt + BtKfChg + BtKfSeg
    SUMZ_Speck.ks8.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    SUMZ_Speck.ks8.reserve = &H0                ' �\��ς�
    
    SUMZ_Speck.ks9.keypos = 3                   ' �L�[�|�W�V����
    SUMZ_Speck.ks9.keyleng = 20                 ' �L�[��
    SUMZ_Speck.ks9.keyflag = BtKfExt + BtKfChg  ' �L�[�t���O
    SUMZ_Speck.ks9.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    SUMZ_Speck.ks9.reserve = &H0                ' �\��ς�

    sts = BTRV(BtOpCreate, SUMZ_POS, SUMZ_Speck, Len(SUMZ_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�݌ɏW�v�f�[�^")
        Exit Function
    End If
    
    SUMZ_Create = False

End Function

Function SUMZ_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              �݌ɏW�v�f�[�^�@�n�o�d�m                            *
'*                                                                  *
'*      ��  ��:Open Mode(Btrieve�Q��)                               *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    SUMZ_Open = True
                                            '�݌ɏW�v�f�[�^�t���p�X�捞��
    sts = GetIni("FILE", SUMZ_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI[SUMZ] �ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, SUMZ_POS, SUMZREC, Len(SUMZREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = SUMZ_Create()        '�݌ɏW�v�f�[�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, SUMZ_POS, SUMZREC, Len(SUMZREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�݌ɏW�v�f�[�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�݌ɏW�v�f�[�^")
                Exit Function
        End Select
    Loop

    SUMZ_Open = False
End Function


