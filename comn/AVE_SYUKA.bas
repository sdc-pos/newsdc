Attribute VB_Name = "AVE_SYUKA"
Option Explicit
'********************************************************************
'*
'*              ���Ϗo�א��@�t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const AVE_SYUKA_ID$ = "AVE_SYUKA"

'�y�[�W�T�C�Y
Public Const AVE_SYUKA_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public AVE_SYUKA_POS As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type AVE_SYUKAREC_Tag
    JGYOBU(0 To 0)          As Byte         '���ƕ��敪
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '�i�ڃR�[�h�i�O���j
    ST_LOCATION(0 To 7)     As Byte         '�W���I��
    UPDATE_YMD(0 To 7)      As Byte         '�W�v�N����
    ZEN3_YM(0 To 5)         As Byte         '�O�X�X�N��         2011.07.01 ���g�p ��߰��ر�
    ZEN3_SYUKA(0 To 7)      As Byte         '�O�X�X���o�א�     2011.07.01 ���g�p 0�ر�
    ZEN2_YM(0 To 5)         As Byte         '�O�X�N��           2011.07.01 ���g�p ��߰��ر�
    ZEN2_SYUKA(0 To 7)      As Byte         '�O�X���o�א�       2011.07.01 ���g�p 0�ر�
    ZEN1_YM(0 To 5)         As Byte         '�O�N��             2011.07.01 ���g�p ��߰��ر�
    ZEN1_SYUKA(0 To 7)      As Byte         '�O���o�א�         2011.07.01 �S�o�ׂƂ��Ďg�p
    AVE_SYUKA(0 To 7)       As Byte         '���Ϗo�א�
    Two_Year_SYUKA(0 To 7)  As Byte         '�ߋ��Q�N�Ԏ���

'-------------------------------------------' 2011.07.01 ��
    TOTAL_CNT(0 To 7)           As Byte     '���o�׌���
    TOTAL_AVE_CNT(0 To 7)       As Byte     '���ϑ��o�׌���


    S_SYUKA_QTY1(0 To 7)        As Byte     '���Y�v��o�א�(1)
    S_SYUKA_CNT1(0 To 7)        As Byte     '���Y�v��o�׌���(1)
    S_AVE_SYUKA_QTY1(0 To 7)    As Byte     '���ϐ��Y�v��o�א�(1)
    S_AVE_SYUKA_CNT1(0 To 7)    As Byte     '���ϐ��Y�v��o�׌���(1)

    S_SYUKA_QTY2(0 To 7)        As Byte     '���Y�v��o�א�(2)
    S_SYUKA_CNT2(0 To 7)        As Byte     '���Y�v��o�׌���(2)
    S_AVE_SYUKA_QTY2(0 To 7)    As Byte     '���ϐ��Y�v��o�א�(2)
    S_AVE_SYUKA_CNT2(0 To 7)    As Byte     '���ϐ��Y�v��o�׌���(2)


    NAI_BUHIN(0 To 0)           As Byte     '�����������i�敪
    HIN_NAME(0 To 39)           As Byte     '�i��


    FILLER(0 To 38)             As Byte     'FILLER
'-------------------------------------------' 2011.07.01�@��





End Type

'�f�[�^�E�o�b�t�@
Public AVE_SYUKAREC         As AVE_SYUKAREC_Tag

'�L�[��`
Type KEY0_AVE_SYUKA         '�j�d�x�O
    JGYOBU(0 To 0)          As Byte         '���ƕ��敪
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '�i�ڃR�[�h�i�O���j
End Type

Type KEY1_AVE_SYUKA         '�j�d�x�O
    ST_LOCATION(0 To 7)     As Byte         '�W���I��
    JGYOBU(0 To 0)          As Byte         '���ƕ��敪
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '�i�ڃR�[�h�i�O���j
End Type

'�L�[�E�f�[�^
Public K0_AVE_SYUKA         As KEY0_AVE_SYUKA
Public K1_AVE_SYUKA         As KEY1_AVE_SYUKA

Type AVE_SYUKA_FSpeck
    fs      As BtFileSpeck                  '̧�� ��߯��\����
    ks0     As BtKeySpeck                   '�� ��߯��\����
    ks1     As BtKeySpeck                   '�� ��߯��\����
    ks2     As BtKeySpeck                   '�� ��߯��\����
    ks3     As BtKeySpeck                   '�� ��߯��\����
    ks4     As BtKeySpeck                   '�� ��߯��\����
    ks5     As BtKeySpeck                   '�� ��߯��\����
    ks6     As BtKeySpeck                   '�� ��߯��\����
    ks7     As BtKeySpeck                   '�� ��߯��\����
End Type

Private AVE_SYUKA_Speck As AVE_SYUKA_FSpeck

Private Function AVE_SYUKA_Create() As Integer
'********************************************************************
'*
'*              �����Ϗo�א��@�b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    AVE_SYUKA_Create = True
                                            '�����Ϗo�א��t���p�X�捞��
    sts = GetIni("FILE", AVE_SYUKA_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [AVE_SYUKA]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    AVE_SYUKA_Speck.fs.recoleng = Len(AVE_SYUKAREC)     ' ���R�[�h��
    AVE_SYUKA_Speck.fs.PageSize = AVE_SYUKA_PG_SIZ      ' �y�[�W�T�C�Y
    AVE_SYUKA_Speck.fs.idexnumb = 2                     ' �C���f�b�N�X��
    AVE_SYUKA_Speck.fs.fileflag = 0                     ' �t�@�C���t���O
    AVE_SYUKA_Speck.fs.reserve = &H0                    ' �\��ς�
                                                    
'---------------------------------------------------
                                                        ' �L�[�O
    AVE_SYUKA_Speck.ks0.keypos = 1                      ' �L�[�|�W�V����
    AVE_SYUKA_Speck.ks0.keyleng = 1                     ' �L�[��
    AVE_SYUKA_Speck.ks0.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    AVE_SYUKA_Speck.ks0.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    AVE_SYUKA_Speck.ks0.reserve = &H0                   ' �\��ς�

    AVE_SYUKA_Speck.ks1.keypos = 2                      ' �L�[�|�W�V����
    AVE_SYUKA_Speck.ks1.keyleng = 1                     ' �L�[��
    AVE_SYUKA_Speck.ks1.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    AVE_SYUKA_Speck.ks1.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    AVE_SYUKA_Speck.ks1.reserve = &H0                   ' �\��ς�

    AVE_SYUKA_Speck.ks2.keypos = 3                      ' �L�[�|�W�V����
    AVE_SYUKA_Speck.ks2.keyleng = 20                    ' �L�[��
    AVE_SYUKA_Speck.ks2.keyflag = BtKfExt               ' �L�[�t���O
    AVE_SYUKA_Speck.ks2.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    AVE_SYUKA_Speck.ks2.reserve = &H0                   ' �\��ς�
'---------------------------------------------------
                                                        ' �L�[�P
    AVE_SYUKA_Speck.ks3.keypos = 23                     ' �L�[�|�W�V����
    AVE_SYUKA_Speck.ks3.keyleng = 8                     ' �L�[��
    AVE_SYUKA_Speck.ks3.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    AVE_SYUKA_Speck.ks3.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    AVE_SYUKA_Speck.ks3.reserve = &H0                   ' �\��ς�

    AVE_SYUKA_Speck.ks4.keypos = 1                      ' �L�[�|�W�V����
    AVE_SYUKA_Speck.ks4.keyleng = 1                     ' �L�[��
    AVE_SYUKA_Speck.ks4.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    AVE_SYUKA_Speck.ks4.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    AVE_SYUKA_Speck.ks4.reserve = &H0                   ' �\��ς�

    AVE_SYUKA_Speck.ks5.keypos = 2                      ' �L�[�|�W�V����
    AVE_SYUKA_Speck.ks5.keyleng = 1                     ' �L�[��
    AVE_SYUKA_Speck.ks5.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    AVE_SYUKA_Speck.ks5.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    AVE_SYUKA_Speck.ks5.reserve = &H0                   ' �\��ς�

    AVE_SYUKA_Speck.ks6.keypos = 3                      ' �L�[�|�W�V����
    AVE_SYUKA_Speck.ks6.keyleng = 20                    ' �L�[��
    AVE_SYUKA_Speck.ks6.keyflag = BtKfExt               ' �L�[�t���O
    AVE_SYUKA_Speck.ks6.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    AVE_SYUKA_Speck.ks6.reserve = &H0                   ' �\��ς�

    sts = BTRV(BtOpCreate, AVE_SYUKA_POS, AVE_SYUKA_Speck, Len(AVE_SYUKA_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�����Ϗo�א�")
        Exit Function
    End If

    AVE_SYUKA_Create = False

End Function

Public Function AVE_SYUKA_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �����Ϗo�א��@�n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    AVE_SYUKA_Open = True
                                            '�����Ϗo�א��t���p�X�捞��
    sts = GetIni("FILE", AVE_SYUKA_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [AVE_SYUKA]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = AVE_SYUKA_Create()    '�����Ϗo�א��쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�����Ϗo�א�")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�����Ϗo�א�")
                Exit Function
        End Select
    Loop

    AVE_SYUKA_Open = False

End Function
