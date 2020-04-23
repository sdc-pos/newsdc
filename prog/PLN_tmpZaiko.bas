Attribute VB_Name = "PLN_tmpZaiko"
Option Explicit
'********************************************************************
'*
'*              ���ޏ��v�ʊm�F��ʒ��ԃt�@�C�� �t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const PLN_tmpZaiko_ID$ = "PLN_tmpZaiko"

'�y�[�W�T�C�Y
Public Const PLN_tmpZaiko_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public PLN_tmpZaiko_POS As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
Type PLN_tmpZaikoREC_Tag
    SYUBETSU(0 To 1)        As Byte     '���
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j
    RIREKI_DT(0 To 7)       As Byte     '�N����
    DATA_KBN(0 To 0)        As Byte     '�ް��敪
    ST_ZAIKO_QTY(0 To 5)    As Byte     '�J�n���݌ɐ�
    SYOUHI_QTY(0 To 5)      As Byte     '����
    NYUKA_QTY(0 To 5)       As Byte     '����
    ZAIKO_QTY(0 To 5)       As Byte     '�݌Ɏc
    INS_TANTO(0 To 9)       As Byte     '�ǉ��@�S����
    Ins_DateTime(0 To 13)   As Byte     '�ǉ��@����         YYYYMMDDhhmmss

End Type

'�f�[�^�E�o�b�t�@
Public PLN_tmpZaikoREC      As PLN_tmpZaikoREC_Tag

'�L�[��`
Type KEY0_PLN_tmpZaiko              '�j�d�x�O
    SYUBETSU(0 To 1)        As Byte     '���
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j
    RIREKI_DT(0 To 7)       As Byte     '�N����
End Type

Type KEY1_PLN_tmpZaiko              '�j�d�x�P
    RIREKI_DT(0 To 7)       As Byte     '�N����
    SYUBETSU(0 To 1)        As Byte     '���
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j
End Type

'�L�[�E�f�[�^
Public K0_PLN_tmpZaiko      As KEY0_PLN_tmpZaiko
Public K1_PLN_tmpZaiko      As KEY1_PLN_tmpZaiko

Type PLN_tmpZaiko_FSpeck
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
End Type

Private PLN_tmpZaiko_Speck  As PLN_tmpZaiko_FSpeck
Private Function PLN_tmpZaiko_Create() As Integer
'********************************************************************
'*
'*              ���ޏ��v�ʊm�F��ʒ��ԃt�@�C���@�b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    PLN_tmpZaiko_Create = True
                                            '�݌Ƀf�[�^�@�t���p�X�捞��
    sts = GetIni("FILE", PLN_tmpZaiko_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [PLN_tmpZaiko]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    PLN_tmpZaiko_Speck.fs.recoleng = Len(PLN_tmpZaikoREC)   ' ���R�[�h��
    PLN_tmpZaiko_Speck.fs.PageSize = PLN_tmpZaiko_PG_SIZ    ' �y�[�W�T�C�Y
    PLN_tmpZaiko_Speck.fs.idexnumb = 2                      ' �C���f�b�N�X��
    PLN_tmpZaiko_Speck.fs.fileflag = 0                      ' �t�@�C���t���O
    PLN_tmpZaiko_Speck.fs.reserve = &H0                     ' �\��ς�
'---------------------------------------------------'
                                                    ' �L�[�O
    PLN_tmpZaiko_Speck.ks0.keypos = 1                       ' �L�[�|�W�V����
    PLN_tmpZaiko_Speck.ks0.keyleng = 2                      ' �L�[��
    PLN_tmpZaiko_Speck.ks0.keyflag = BtKfExt + BtKfSeg      ' �L�[�t���O
    PLN_tmpZaiko_Speck.ks0.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    PLN_tmpZaiko_Speck.ks0.reserve = &H0                    ' �\��ς�
                                                    
    PLN_tmpZaiko_Speck.ks1.keypos = 3                       ' �L�[�|�W�V����
    PLN_tmpZaiko_Speck.ks1.keyleng = 1                      ' �L�[��
    PLN_tmpZaiko_Speck.ks1.keyflag = BtKfExt + BtKfSeg      ' �L�[�t���O
    PLN_tmpZaiko_Speck.ks1.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    PLN_tmpZaiko_Speck.ks1.reserve = &H0                    ' �\��ς�
                                                    
    PLN_tmpZaiko_Speck.ks2.keypos = 4                       ' �L�[�|�W�V����
    PLN_tmpZaiko_Speck.ks2.keyleng = 1                      ' �L�[��
    PLN_tmpZaiko_Speck.ks2.keyflag = BtKfExt + BtKfSeg      ' �L�[�t���O
    PLN_tmpZaiko_Speck.ks2.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    PLN_tmpZaiko_Speck.ks2.reserve = &H0                    ' �\��ς�
                                                    
    PLN_tmpZaiko_Speck.ks3.keypos = 5                       ' �L�[�|�W�V����
    PLN_tmpZaiko_Speck.ks3.keyleng = 20                     ' �L�[��
    PLN_tmpZaiko_Speck.ks3.keyflag = BtKfExt + BtKfSeg      ' �L�[�t���O
    PLN_tmpZaiko_Speck.ks3.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    PLN_tmpZaiko_Speck.ks3.reserve = &H0                    ' �\��ς�
                                                    
    PLN_tmpZaiko_Speck.ks4.keypos = 25                      ' �L�[�|�W�V����
    PLN_tmpZaiko_Speck.ks4.keyleng = 8                      ' �L�[��
    PLN_tmpZaiko_Speck.ks4.keyflag = BtKfExt                ' �L�[�t���O
    PLN_tmpZaiko_Speck.ks4.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    PLN_tmpZaiko_Speck.ks4.reserve = &H0                    ' �\��ς�
                                                    
'---------------------------------------------------'
                                                    ' �L�[�P
    PLN_tmpZaiko_Speck.ks5.keypos = 25                      ' �L�[�|�W�V����
    PLN_tmpZaiko_Speck.ks5.keyleng = 8                      ' �L�[��
                                                            ' �L�[�t���O
    PLN_tmpZaiko_Speck.ks5.keyflag = BtKfExt + BtKfSeg + BtKfDup
    PLN_tmpZaiko_Speck.ks5.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    PLN_tmpZaiko_Speck.ks5.reserve = &H0                    ' �\��ς�
    
    PLN_tmpZaiko_Speck.ks6.keypos = 1                       ' �L�[�|�W�V����
    PLN_tmpZaiko_Speck.ks6.keyleng = 2                      ' �L�[��
                                                            ' �L�[�t���O
    PLN_tmpZaiko_Speck.ks6.keyflag = BtKfExt + BtKfSeg + BtKfDup
    PLN_tmpZaiko_Speck.ks6.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    PLN_tmpZaiko_Speck.ks6.reserve = &H0                    ' �\��ς�
    
    
    PLN_tmpZaiko_Speck.ks7.keypos = 3                       ' �L�[�|�W�V����
    PLN_tmpZaiko_Speck.ks7.keyleng = 1                      ' �L�[��
                                                            ' �L�[�t���O
    PLN_tmpZaiko_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfDup
    PLN_tmpZaiko_Speck.ks7.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    PLN_tmpZaiko_Speck.ks7.reserve = &H0                    ' �\��ς�
                                                    
    PLN_tmpZaiko_Speck.ks8.keypos = 4                       ' �L�[�|�W�V����
    PLN_tmpZaiko_Speck.ks8.keyleng = 1                      ' �L�[��
                                                            ' �L�[�t���O
    PLN_tmpZaiko_Speck.ks8.keyflag = BtKfExt + BtKfSeg + BtKfDup
    PLN_tmpZaiko_Speck.ks8.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    PLN_tmpZaiko_Speck.ks8.reserve = &H0                    ' �\��ς�
                                                    
    PLN_tmpZaiko_Speck.ks9.keypos = 5                       ' �L�[�|�W�V����
    PLN_tmpZaiko_Speck.ks9.keyleng = 20                     ' �L�[��
                                                            ' �L�[�t���O
    PLN_tmpZaiko_Speck.ks9.keyflag = BtKfExt + BtKfDup
    PLN_tmpZaiko_Speck.ks9.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    PLN_tmpZaiko_Speck.ks9.reserve = &H0                    ' �\��ς�
'---------------------------------------------------'
    sts = BTRV(BtOpCreate, PLN_tmpZaiko_POS, PLN_tmpZaiko_Speck, Len(PLN_tmpZaiko_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���ޏ��v�ʊm�F��ʒ��ԃt�@�C��")
        Exit Function
    End If
    PLN_tmpZaiko_Create = False
End Function
Public Function PLN_tmpZaiko_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���ޏ��v�ʊm�F��ʒ��ԃt�@�C���@�n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    PLN_tmpZaiko_Open = True
                                            '���ޏ��v�ʊm�F��ʒ��ԃt�@�C���@�t���p�X�捞��
    sts = GetIni("FILE", PLN_tmpZaiko_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [PLN_tmpZaiko]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, PLN_tmpZaiko_POS, PLN_tmpZaikoREC, Len(PLN_tmpZaikoREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = PLN_tmpZaiko_Create() '���ޏ��v�ʊm�F��ʒ��ԃt�@�C���@�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, PLN_tmpZaiko_POS, PLN_tmpZaikoREC, Len(PLN_tmpZaikoREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���ޏ��v�ʊm�F��ʒ��ԃt�@�C��")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���ޏ��v�ʊm�F��ʒ��ԃt�@�C��")
                Exit Function
        End Select
    Loop
    PLN_tmpZaiko_Open = False

End Function

