Attribute VB_Name = "tmpZAIKO_F110010"
Option Explicit
'********************************************************************
'*
'*              �݌Ƀf�[�^�i�ꎞ�f�[�^�j �t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const tmpZAIKO_ID$ = "tmpZAIKO"

'�y�[�W�T�C�Y
Public Const tmpZAIKO_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public tmpZAIKO_POS As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
Type tmpZAIKOREC_Tag
    SOKO_NO(0 To 1)     As Byte     '�q�ɇ�
    Retu(0 To 1)        As Byte     '�I�ԁ@��
    Ren(0 To 1)         As Byte     '�I�ԁ@�A
    Dan(0 To 1)         As Byte     '�I�ԁ@�i
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    '2005.12.05 13-->20
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
    GOODS_ON(0 To 0)    As Byte     '���i���^�����i��
    NYUKA_DT(0 To 7)    As Byte     '���ד��t
    NYUKO_DT(0 To 7)    As Byte     '���ɓ��t
    '2005.12.05 13-->20
    HIN_NAI(0 To 19)    As Byte     '�i�ԁi�����j
    YUKO_Z_QTY(0 To 7)  As Byte     '�L���݌ɐ�
    LOCK_F(0 To 0)      As Byte     '�r���t���O
    WEL_ID(0 To 2)      As Byte     '�g�p�q�@ID
    PRG_ID(0 To 7)      As Byte     '�g�p���v���O����
    GOODS_YMD(0 To 7)   As Byte     '���i�����t
    
    '2005.12.05 ���ڒǉ�
    SHIIRE_CODE(0 To 4) As Byte     '�d���溰��
    SHIIRE_TANKA(0 To 10) As Byte   '�d���P��(9(8)V99)
    KEIJYO_YM(0 To 5)   As Byte     '�v��N��
    '2005.12.05 ���ڒǉ�
    
    '----------------   2010.07.08 ��
    GENSANKOKU(0 To 19)         As Byte     '���Y����
    SHIIRE_WORK_CENTER(0 To 7)  As Byte     '���ގd����ܰ�����
    ID_NO2(0 To 11)             As Byte     'ID_NO
    YOSAN_FROM(0 To 4)          As Byte     '�\�Z�P�ʁi���j
    YOSAN_TO(0 To 4)            As Byte     '�\�Z�P�ʁi��j
    '----------------   2010.07.08 ��
    
    
    FILLER(0 To 24)     As Byte     'FILLER
End Type

'�f�[�^�E�o�b�t�@
Public tmpZAIKOREC      As tmpZAIKOREC_Tag

'�L�[��`

Type KEY0_tmpZAIKO                  '�j�d�x�O
    SOKO_NO(0 To 1)     As Byte     '�q�ɇ�
    Retu(0 To 1)        As Byte     '�I�ԁ@��
    Ren(0 To 1)         As Byte     '�I�ԁ@�A
    Dan(0 To 1)         As Byte     '�I�ԁ@�i
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
    NYUKA_DT(0 To 7)    As Byte     '���ד��t
End Type

Type KEY1_tmpZAIKO                  '�j�d�x�P
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
    NYUKA_DT(0 To 7)    As Byte     '���ד��t
    SOKO_NO(0 To 1)     As Byte     '�q�ɇ�
    Retu(0 To 1)        As Byte     '�I�ԁ@��
    Ren(0 To 1)         As Byte     '�I�ԁ@�A
    Dan(0 To 1)         As Byte     '�I�ԁ@�i
End Type

'�L�[�E�f�[�^
Public K0_tmpZAIKO      As KEY0_tmpZAIKO
Public K1_tmpZAIKO      As KEY1_tmpZAIKO

Type tmpZAIKO_FSpeck
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

Private tmpZAIKO_Speck As tmpZAIKO_FSpeck
Private Function tmpZAIKO_Create() As Integer
'********************************************************************
'*
'*              �݌Ƀf�[�^�i�ꎞ�f�[�^�j�@�b�q�d�`�s�d
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

    tmpZAIKO_Create = True
                                            '�݌Ƀf�[�^�@�t���p�X�捞��
    sts = GetIni("FILE", tmpZAIKO_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [tmpZAIKO]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    Ret = InStrRev(Trim(c), ".") - 1
    
    FullPath = Left(Trim(c), Ret) & "F110010" & Right(Trim(c), Len(Trim(c)) - Ret)
    
    
 



    tmpZAIKO_Speck.fs.recoleng = Len(tmpZAIKOREC)   ' ���R�[�h��
    tmpZAIKO_Speck.fs.PageSize = tmpZAIKO_PG_SIZ    ' �y�[�W�T�C�Y
    tmpZAIKO_Speck.fs.idexnumb = 2                  ' �C���f�b�N�X��
    tmpZAIKO_Speck.fs.fileflag = 0                  ' �t�@�C���t���O
    tmpZAIKO_Speck.fs.reserve = &H0                 ' �\��ς�
'---------------------------------------------------' �L�[�O
    tmpZAIKO_Speck.ks0.keypos = 1                   ' �L�[�|�W�V����
    tmpZAIKO_Speck.ks0.keyleng = 2                  ' �L�[��
                                                    ' �L�[�t���O
    tmpZAIKO_Speck.ks0.keyflag = BtKfExt + BtKfSeg
    tmpZAIKO_Speck.ks0.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    tmpZAIKO_Speck.ks0.reserve = &H0                ' �\��ς�
                                                    
    tmpZAIKO_Speck.ks1.keypos = 3                   ' �L�[�|�W�V����
    tmpZAIKO_Speck.ks1.keyleng = 2                  ' �L�[��
                                                    ' �L�[�t���O
    tmpZAIKO_Speck.ks1.keyflag = BtKfExt + BtKfSeg
    tmpZAIKO_Speck.ks1.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    tmpZAIKO_Speck.ks1.reserve = &H0                ' �\��ς�
                                                    
    tmpZAIKO_Speck.ks2.keypos = 5                   ' �L�[�|�W�V����
    tmpZAIKO_Speck.ks2.keyleng = 2                  ' �L�[��
                                                    ' �L�[�t���O
    tmpZAIKO_Speck.ks2.keyflag = BtKfExt + BtKfSeg
    tmpZAIKO_Speck.ks2.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    tmpZAIKO_Speck.ks2.reserve = &H0                ' �\��ς�
                                                    
    tmpZAIKO_Speck.ks3.keypos = 7                   ' �L�[�|�W�V����
    tmpZAIKO_Speck.ks3.keyleng = 2                  ' �L�[��
                                                    ' �L�[�t���O
    tmpZAIKO_Speck.ks3.keyflag = BtKfExt + BtKfSeg
    tmpZAIKO_Speck.ks3.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    tmpZAIKO_Speck.ks3.reserve = &H0                ' �\��ς�
                                                    
    tmpZAIKO_Speck.ks4.keypos = 9                   ' �L�[�|�W�V����
    tmpZAIKO_Speck.ks4.keyleng = 1                  ' �L�[��
                                                    ' �L�[�t���O
    tmpZAIKO_Speck.ks4.keyflag = BtKfExt + BtKfSeg
    tmpZAIKO_Speck.ks4.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    tmpZAIKO_Speck.ks4.reserve = &H0                ' �\��ς�
                                                    
    tmpZAIKO_Speck.ks5.keypos = 10                  ' �L�[�|�W�V����
    tmpZAIKO_Speck.ks5.keyleng = 1                  ' �L�[��
                                                    ' �L�[�t���O
    tmpZAIKO_Speck.ks5.keyflag = BtKfExt + BtKfSeg
    tmpZAIKO_Speck.ks5.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    tmpZAIKO_Speck.ks5.reserve = &H0                ' �\��ς�
                                                    
    tmpZAIKO_Speck.ks6.keypos = 11                  ' �L�[�|�W�V����
    tmpZAIKO_Speck.ks6.keyleng = 20                 ' �L�[��
                                                    ' �L�[�t���O
    tmpZAIKO_Speck.ks6.keyflag = BtKfExt + BtKfSeg
    tmpZAIKO_Speck.ks6.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    tmpZAIKO_Speck.ks6.reserve = &H0                ' �\��ς�
    
    tmpZAIKO_Speck.ks7.keypos = 32                  ' �L�[�|�W�V����
    tmpZAIKO_Speck.ks7.keyleng = 8                  ' �L�[��
                                                    ' �L�[�t���O
    tmpZAIKO_Speck.ks7.keyflag = BtKfExt
    tmpZAIKO_Speck.ks7.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    tmpZAIKO_Speck.ks7.reserve = &H0                ' �\��ς�
'---------------------------------------------------' �L�[�P
    tmpZAIKO_Speck.ks8.keypos = 9                   ' �L�[�|�W�V����
    tmpZAIKO_Speck.ks8.keyleng = 1                  ' �L�[��
                                                    ' �L�[�t���O
    tmpZAIKO_Speck.ks8.keyflag = BtKfExt + BtKfSeg
    tmpZAIKO_Speck.ks8.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    tmpZAIKO_Speck.ks8.reserve = &H0                ' �\��ς�
                                                    
    tmpZAIKO_Speck.ks9.keypos = 10                 ' �L�[�|�W�V����
    tmpZAIKO_Speck.ks9.keyleng = 1                 ' �L�[��
                                                    ' �L�[�t���O
    tmpZAIKO_Speck.ks9.keyflag = BtKfExt + BtKfSeg
    tmpZAIKO_Speck.ks9.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    tmpZAIKO_Speck.ks9.reserve = &H0               ' �\��ς�
                                                    
    tmpZAIKO_Speck.ks10.keypos = 11                 ' �L�[�|�W�V����
    tmpZAIKO_Speck.ks10.keyleng = 20                ' �L�[��
                                                    ' �L�[�t���O
    tmpZAIKO_Speck.ks10.keyflag = BtKfExt + BtKfSeg
    tmpZAIKO_Speck.ks10.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    tmpZAIKO_Speck.ks10.reserve = &H0               ' �\��ς�
                                                    
    tmpZAIKO_Speck.ks11.keypos = 32                 ' �L�[�|�W�V����
    tmpZAIKO_Speck.ks11.keyleng = 8                 ' �L�[��
                                                    ' �L�[�t���O
    tmpZAIKO_Speck.ks11.keyflag = BtKfExt + BtKfSeg
    tmpZAIKO_Speck.ks11.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    tmpZAIKO_Speck.ks11.reserve = &H0               ' �\��ς�
                                                    
    tmpZAIKO_Speck.ks12.keypos = 1                  ' �L�[�|�W�V����
    tmpZAIKO_Speck.ks12.keyleng = 2                 ' �L�[��
                                                    ' �L�[�t���O
    tmpZAIKO_Speck.ks12.keyflag = BtKfExt + BtKfSeg
    tmpZAIKO_Speck.ks12.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    tmpZAIKO_Speck.ks12.reserve = &H0               ' �\��ς�
                                                    
    tmpZAIKO_Speck.ks13.keypos = 3                  ' �L�[�|�W�V����
    tmpZAIKO_Speck.ks13.keyleng = 2                 ' �L�[��
                                                    ' �L�[�t���O
    tmpZAIKO_Speck.ks13.keyflag = BtKfExt + BtKfSeg
    tmpZAIKO_Speck.ks13.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    tmpZAIKO_Speck.ks13.reserve = &H0               ' �\��ς�
                                                    
    tmpZAIKO_Speck.ks14.keypos = 5                  ' �L�[�|�W�V����
    tmpZAIKO_Speck.ks14.keyleng = 2                 ' �L�[��
                                                    ' �L�[�t���O
    tmpZAIKO_Speck.ks14.keyflag = BtKfExt + BtKfSeg
    tmpZAIKO_Speck.ks14.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    tmpZAIKO_Speck.ks14.reserve = &H0               ' �\��ς�
                                                    
    tmpZAIKO_Speck.ks15.keypos = 7                  ' �L�[�|�W�V����
    tmpZAIKO_Speck.ks15.keyleng = 2                 ' �L�[��
                                                    ' �L�[�t���O
    tmpZAIKO_Speck.ks15.keyflag = BtKfExt
    tmpZAIKO_Speck.ks15.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    tmpZAIKO_Speck.ks15.reserve = &H0               ' �\��ς�
'---------------------------------------------------'
    sts = BTRV(BtOpCreate, tmpZAIKO_POS, tmpZAIKO_Speck, Len(tmpZAIKO_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�݌Ƀf�[�^�i�ꎞ�f�[�^�j")
        Exit Function
    End If
    tmpZAIKO_Create = False
End Function
Public Function tmpZAIKO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �݌Ƀf�[�^�i�ꎞ�f�[�^�j�@�n�o�d�m
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
    
    tmpZAIKO_Open = True
                                            '�݌Ƀf�[�^�@�t���p�X�捞��
    sts = GetIni("FILE", tmpZAIKO_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [tmpZAIKO]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    
    Ret = InStrRev(Trim(c), ".") - 1
    
    FullPath = Left(Trim(c), Ret) & "F110010" & Right(Trim(c), Len(Trim(c)) - Ret)
    
    
    Do
        sts = BTRV(BtOpOpen, tmpZAIKO_POS, tmpZAIKOREC, Len(tmpZAIKOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = tmpZAIKO_Create()        '�݌Ƀf�[�^�@�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, tmpZAIKO_POS, tmpZAIKOREC, Len(tmpZAIKOREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�݌Ƀf�[�^�i�ꎞ�f�[�^�j")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�݌Ƀf�[�^�i�ꎞ�f�[�^�j")
                Exit Function
        End Select
    Loop
    tmpZAIKO_Open = False

End Function

