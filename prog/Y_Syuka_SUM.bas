Attribute VB_Name = "Y_SYU_SUM"
Option Explicit
'********************************************************************
'*
'*              �o�ח\��i���PC�o�ɕ\�p�j�f�[�^  �t�@�C����`
'*              ���o�b��p    2007.03.14
'*
'********************************************************************
'�t�@�C���h�c
Public Const Y_SYU_SUM_ID$ = "Y_SYU_SUM"

'�y�[�W�T�C�Y
Public Const Y_SYU_SUM_PG_SIZ% = 4096

'�|�W�V�����E�u���b�N
Public Y_SYU_SUM_POS    As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type Y_SYU_SUMREC_Tag
    SYUKA_YMD(0 To 7)           As Byte     '�o�ד��t
    INS_BIN(0 To 1)             As Byte     '��
    ST_SOKO(0 To 1)             As Byte     '�W���I��     �q��
    ST_RETU(0 To 1)             As Byte     '             ��
    ST_REN(0 To 1)              As Byte     '             �A
    ST_DAN(0 To 1)              As Byte     '             �i
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    NAIGAI(0 To 0)              As Byte     '�����O
    HIN_NO(0 To 19)             As Byte     '�i�ڔԍ�
    
    Y_SURYO(0 To 6)             As Byte     '�o�ח\�萔��
    J_SURYO(0 To 6)             As Byte     '�o�׎��ѐ���
    
    SYU_NO(0 To 11)             As Byte     '�o�ɕ\��
    DATA_CNT(0 To 3)            As Byte     '����
    
    ST_ZAIKO_QTY(0 To 7)        As Byte     '�W���I�ԍ݌ɐ�
    
    BETU_SOKO(0 To 1)           As Byte     '�ʒu�I��     �q��
    BETU_RETU(0 To 1)           As Byte     '             ��
    BETU_REN(0 To 1)            As Byte     '             �A
    BETU_DAN(0 To 1)            As Byte     '             �i
    
    BETU_ZAIKO_QTY(0 To 7)      As Byte     '�ʒu�݌ɐ�
    
    SYO_ZAIKO_QTY(0 To 7)       As Byte     '���i�����݌ɐ�
    NYU_ZAIKO_QTY(0 To 7)       As Byte     '���בq�ɍ݌ɐ�
    
    INS_NOW(0 To 13)            As Byte     '�ް��쐬����
    
    FILLER(0 To 39)             As Byte     'FILLER
End Type

'�f�[�^�E�o�b�t�@
Public Y_SYU_SUMREC             As Y_SYU_SUMREC_Tag

'�L�[��`
Type KEY0_Y_SYU_SUM         '�j�d�x�O
    SYUKA_YMD(0 To 7)           As Byte     '�o�ד��t
    INS_BIN(0 To 1)             As Byte     '��
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    NAIGAI(0 To 0)              As Byte     '�����O
    HIN_NO(0 To 19)             As Byte     '�i�ڔԍ�
End Type

Type KEY1_Y_SYU_SUM         '�j�d�x�P
    SYUKA_YMD(0 To 7)           As Byte     '�o�ד��t
    INS_BIN(0 To 1)             As Byte     '��
    ST_SOKO(0 To 1)             As Byte     '�W���I��     �q��
    ST_RETU(0 To 1)             As Byte     '             ��
    ST_REN(0 To 1)              As Byte     '             �A
    ST_DAN(0 To 1)              As Byte     '             �i
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    NAIGAI(0 To 0)              As Byte     '�����O
    HIN_NO(0 To 19)             As Byte     '�i�ڔԍ�
End Type

Type KEY2_Y_SYU_SUM         '�j�d�x�Q
    SYUKA_YMD(0 To 7)           As Byte     '�o�ד��t
    INS_BIN(0 To 1)             As Byte     '��
End Type

Type KEY3_Y_SYU_SUM         '�j�d�x�R
    INS_BIN(0 To 1)             As Byte     '��
    SYU_NO(0 To 11)              As Byte     '�o�ɕ\��
End Type


'�L�[�E�f�[�^
Public K0_Y_SYU_SUM             As KEY0_Y_SYU_SUM
Public K1_Y_SYU_SUM             As KEY1_Y_SYU_SUM
Public K2_Y_SYU_SUM             As KEY2_Y_SYU_SUM
Public K3_Y_SYU_SUM             As KEY3_Y_SYU_SUM

Type Y_SYU_SUM_FSpeck
    fs      As BtFileSpeck                  ' ̧�� ��߯��\����
    ks0     As BtKeySpeck                   ' �� ��߯��\����
    ks1     As BtKeySpeck                   ' �� ��߯��\����
    ks2     As BtKeySpeck                   ' �� ��߯��\����
    ks3     As BtKeySpeck                   ' �� ��߯��\����
    ks4     As BtKeySpeck                   ' �� ��߯��\����
    
    ks5     As BtKeySpeck                   ' �� ��߯��\����
    ks6     As BtKeySpeck                   ' �� ��߯��\����
    ks7     As BtKeySpeck                   ' �� ��߯��\����
    ks8     As BtKeySpeck                   ' �� ��߯��\����
    ks9     As BtKeySpeck                   ' �� ��߯��\����
    ks10    As BtKeySpeck                   ' �� ��߯��\����
    ks11    As BtKeySpeck                   ' �� ��߯��\����
    ks12    As BtKeySpeck                   ' �� ��߯��\����
    ks13    As BtKeySpeck                   ' �� ��߯��\����
    
    ks14    As BtKeySpeck                   ' �� ��߯��\����
    ks15    As BtKeySpeck                   ' �� ��߯��\����
    
    ks16    As BtKeySpeck                   ' �� ��߯��\����
    ks17    As BtKeySpeck                   ' �� ��߯��\����



End Type

Private Y_SYU_SUM_Speck     As Y_SYU_SUM_FSpeck

Private Function Y_SYU_SUM_Create(Optional F_NAME As String = " ") As Integer
'********************************************************************
'*
'*              �o�ח\��(���PC�o�ɕ\�p)�f�[�^  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

Dim Ret         As Integer

    Y_SYU_SUM_Create = True
                                            '�o�ח\��f�[�^�t���p�X�捞��
    sts = GetIni("FILE", Y_SYU_SUM_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [Y_SYU_SUM]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    If Trim(F_NAME) = "" Then
        FullPath = RTrim(c)
    Else
        Ret = InStr(1, Trim(c), ".") - 1
        FullPath = Left(Trim(c), Ret) & Trim(F_NAME) & Right(Trim(c), Len(Trim(c)) - Ret)
    End If

    Y_SYU_SUM_Speck.fs.recoleng = Len(Y_SYU_SUMREC)     ' ���R�[�h��
    Y_SYU_SUM_Speck.fs.PageSize = Y_SYU_SUM_PG_SIZ      ' �y�[�W�T�C�Y
    Y_SYU_SUM_Speck.fs.idexnumb = 4                     ' �C���f�b�N�X��
    Y_SYU_SUM_Speck.fs.fileflag = 0                     ' �t�@�C���t���O
    Y_SYU_SUM_Speck.fs.reserve = &H0                    ' �\��ς�
'---------------------------------------------------' �L�[�O
    Y_SYU_SUM_Speck.ks0.keypos = 1                      ' �L�[�|�W�V����
    Y_SYU_SUM_Speck.ks0.keyleng = 8                     ' �L�[��
                                                        ' �L�[�t���O
    Y_SYU_SUM_Speck.ks0.keyflag = BtKfExt + BtKfSeg
    Y_SYU_SUM_Speck.ks0.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_SUM_Speck.ks0.reserve = &H0                   ' �\��ς�
    
    Y_SYU_SUM_Speck.ks1.keypos = 9                      ' �L�[�|�W�V����
    Y_SYU_SUM_Speck.ks1.keyleng = 2                     ' �L�[��
                                                        ' �L�[�t���O
    Y_SYU_SUM_Speck.ks1.keyflag = BtKfExt + BtKfSeg
    Y_SYU_SUM_Speck.ks1.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_SUM_Speck.ks1.reserve = &H0                   ' �\��ς�
    
    Y_SYU_SUM_Speck.ks2.keypos = 19                     ' �L�[�|�W�V����
    Y_SYU_SUM_Speck.ks2.keyleng = 1                     ' �L�[��
                                                        ' �L�[�t���O
    Y_SYU_SUM_Speck.ks2.keyflag = BtKfExt + BtKfSeg
    Y_SYU_SUM_Speck.ks2.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_SUM_Speck.ks2.reserve = &H0                   ' �\��ς�
    
    Y_SYU_SUM_Speck.ks3.keypos = 20                     ' �L�[�|�W�V����
    Y_SYU_SUM_Speck.ks3.keyleng = 1                     ' �L�[��
                                                        ' �L�[�t���O
    Y_SYU_SUM_Speck.ks3.keyflag = BtKfExt + BtKfSeg
    Y_SYU_SUM_Speck.ks3.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_SUM_Speck.ks3.reserve = &H0                   ' �\��ς�
    
    Y_SYU_SUM_Speck.ks4.keypos = 21                     ' �L�[�|�W�V����
    Y_SYU_SUM_Speck.ks4.keyleng = 20                    ' �L�[��
                                                        ' �L�[�t���O
    Y_SYU_SUM_Speck.ks4.keyflag = BtKfExt
    Y_SYU_SUM_Speck.ks4.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_SUM_Speck.ks4.reserve = &H0                   ' �\��ς�
'---------------------------------------------------' �L�[�O
    
'---------------------------------------------------' �L�[�P
    Y_SYU_SUM_Speck.ks5.keypos = 1                      ' �L�[�|�W�V����
    Y_SYU_SUM_Speck.ks5.keyleng = 8                     ' �L�[��
                                                        ' �L�[�t���O
    Y_SYU_SUM_Speck.ks5.keyflag = BtKfExt + BtKfSeg
    Y_SYU_SUM_Speck.ks5.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_SUM_Speck.ks5.reserve = &H0                   ' �\��ς�
    
    Y_SYU_SUM_Speck.ks6.keypos = 9                      ' �L�[�|�W�V����
    Y_SYU_SUM_Speck.ks6.keyleng = 2                     ' �L�[��
                                                        ' �L�[�t���O
    Y_SYU_SUM_Speck.ks6.keyflag = BtKfExt + BtKfSeg
    Y_SYU_SUM_Speck.ks6.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_SUM_Speck.ks6.reserve = &H0                   ' �\��ς�
    
    Y_SYU_SUM_Speck.ks7.keypos = 11                     ' �L�[�|�W�V����
    Y_SYU_SUM_Speck.ks7.keyleng = 2                     ' �L�[��
                                                        ' �L�[�t���O
    Y_SYU_SUM_Speck.ks7.keyflag = BtKfExt + BtKfSeg
    Y_SYU_SUM_Speck.ks7.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_SUM_Speck.ks7.reserve = &H0                   ' �\��ς�
    
    Y_SYU_SUM_Speck.ks8.keypos = 13                     ' �L�[�|�W�V����
    Y_SYU_SUM_Speck.ks8.keyleng = 2                     ' �L�[��
                                                        ' �L�[�t���O
    Y_SYU_SUM_Speck.ks8.keyflag = BtKfExt + BtKfSeg
    Y_SYU_SUM_Speck.ks8.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_SUM_Speck.ks8.reserve = &H0                   ' �\��ς�
    
    Y_SYU_SUM_Speck.ks9.keypos = 15                     ' �L�[�|�W�V����
    Y_SYU_SUM_Speck.ks9.keyleng = 2                     ' �L�[��
                                                        ' �L�[�t���O
    Y_SYU_SUM_Speck.ks9.keyflag = BtKfExt + BtKfSeg
    Y_SYU_SUM_Speck.ks9.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_SUM_Speck.ks9.reserve = &H0                   ' �\��ς�
    
    Y_SYU_SUM_Speck.ks10.keypos = 17                    ' �L�[�|�W�V����
    Y_SYU_SUM_Speck.ks10.keyleng = 2                    ' �L�[��
                                                        ' �L�[�t���O
    Y_SYU_SUM_Speck.ks10.keyflag = BtKfExt + BtKfSeg
    Y_SYU_SUM_Speck.ks10.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_SUM_Speck.ks10.reserve = &H0                  ' �\��ς�
    
    Y_SYU_SUM_Speck.ks11.keypos = 19                    ' �L�[�|�W�V����
    Y_SYU_SUM_Speck.ks11.keyleng = 1                    ' �L�[��
                                                        ' �L�[�t���O
    Y_SYU_SUM_Speck.ks11.keyflag = BtKfExt + BtKfSeg
    Y_SYU_SUM_Speck.ks11.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_SUM_Speck.ks11.reserve = &H0                  ' �\��ς�
    
    Y_SYU_SUM_Speck.ks12.keypos = 20                    ' �L�[�|�W�V����
    Y_SYU_SUM_Speck.ks12.keyleng = 1                    ' �L�[��
                                                        ' �L�[�t���O
    Y_SYU_SUM_Speck.ks12.keyflag = BtKfExt + BtKfSeg
    Y_SYU_SUM_Speck.ks12.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_SUM_Speck.ks12.reserve = &H0                  ' �\��ς�
    
    Y_SYU_SUM_Speck.ks13.keypos = 21                    ' �L�[�|�W�V����
    Y_SYU_SUM_Speck.ks13.keyleng = 20                   ' �L�[��
                                                        ' �L�[�t���O
    Y_SYU_SUM_Speck.ks13.keyflag = BtKfExt
    Y_SYU_SUM_Speck.ks13.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_SUM_Speck.ks13.reserve = &H0                  ' �\��ς�
'---------------------------------------------------' �L�[�P
    
'---------------------------------------------------' �L�[�Q
    Y_SYU_SUM_Speck.ks14.keypos = 1                     ' �L�[�|�W�V����
    Y_SYU_SUM_Speck.ks14.keyleng = 8                    ' �L�[��
                                                        ' �L�[�t���O
    Y_SYU_SUM_Speck.ks14.keyflag = BtKfExt + BtKfDup + BtKfSeg
    Y_SYU_SUM_Speck.ks14.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_SUM_Speck.ks14.reserve = &H0                  ' �\��ς�
    
    Y_SYU_SUM_Speck.ks15.keypos = 9                     ' �L�[�|�W�V����
    Y_SYU_SUM_Speck.ks15.keyleng = 2                    ' �L�[��
                                                        ' �L�[�t���O
    Y_SYU_SUM_Speck.ks15.keyflag = BtKfExt + BtKfDup
    Y_SYU_SUM_Speck.ks15.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_SUM_Speck.ks15.reserve = &H0                  ' �\��ς�
'---------------------------------------------------' �L�[�Q
'---------------------------------------------------' �L�[�R
    Y_SYU_SUM_Speck.ks16.keypos = 9                     ' �L�[�|�W�V����
    Y_SYU_SUM_Speck.ks16.keyleng = 2                    ' �L�[��
                                                        ' �L�[�t���O
    Y_SYU_SUM_Speck.ks16.keyflag = BtKfExt + BtKfDup + BtKfSeg
    Y_SYU_SUM_Speck.ks16.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_SUM_Speck.ks16.reserve = &H0                  ' �\��ς�

    Y_SYU_SUM_Speck.ks17.keypos = 55                    ' �L�[�|�W�V����
    Y_SYU_SUM_Speck.ks17.keyleng = 12                   ' �L�[��
                                                        ' �L�[�t���O
    Y_SYU_SUM_Speck.ks17.keyflag = BtKfExt + BtKfDup
    Y_SYU_SUM_Speck.ks17.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_SUM_Speck.ks17.reserve = &H0                  ' �\��ς�

'---------------------------------------------------' �L�[�R
    sts = BTRV(BtOpCreate, Y_SYU_SUM_POS, Y_SYU_SUM_Speck, Len(Y_SYU_SUM_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�o�ח\��(���PC�o�ɕ\�p)�f�[�^")
        Exit Function
    End If

    Y_SYU_SUM_Create = False

End Function

Function Y_SYU_SUM_Open(Mode As Integer, Optional F_NAME As String = " ") As Integer
'********************************************************************
'*
'*              �o�ח\��(���PC�o�ɕ\�p)�f�[�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
Dim Ret         As Integer
    
    Y_SYU_SUM_Open = True
                                            '�o�ח\��f�[�^�t���p�X�捞��
    sts = GetIni("FILE", Y_SYU_SUM_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [Y_SYU_SUM]�ǂݍ��݃G���[ ")
        Exit Function
    End If
    
    If Trim(F_NAME) = "" Then
        FullPath = RTrim(c)
    Else
        Ret = InStr(1, Trim(c), ".") - 1
        FullPath = Left(Trim(c), Ret) & Trim(F_NAME) & Right(Trim(c), Len(Trim(c)) - Ret)
    End If
    
''    On Error Resume Next
''    Kill (FullPath)
''    On Error GoTo 0
    
    Do
        sts = BTRV(BtOpOpen, Y_SYU_SUM_POS, Y_SYU_SUMREC, Len(Y_SYU_SUMREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = Y_SYU_SUM_Create(F_NAME)      '�o�ח\��f�[�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, Y_SYU_SUM_POS, Y_SYU_SUMREC, Len(Y_SYU_SUMREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�o�ח\��(���PC�o�ɕ\�p)�f�[�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�o�ח\��(���PC�o�ɕ\�p)�f�[�^")
                Exit Function
        End Select
    Loop
    Y_SYU_SUM_Open = False
End Function
