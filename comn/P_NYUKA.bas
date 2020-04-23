Attribute VB_Name = "P_NYU"
Option Explicit
'********************************************************************
'*
'*              ���ޓ��׃`�F�b�N�f�[�^�i�O�؃f�[�^�j�@�t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const P_NYU_ID$ = "P_NYU"

'�y�[�W�T�C�Y
Public Const P_NYU_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public P_NYU_POS    As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type P_NYUREC_Tag
    JGYOBU(0 To 0)          As Byte         '���ƕ��敪
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '�i�ԁi�O���j
    NYUKA_DT(0 To 7)        As Byte         '���ד�
    NYUKA_QTY(0 To 7)       As Byte         '�O�ؐ�
    SOUSAI_DT(0 To 7)       As Byte         '�ŐV���E���t
    SOUSAI_QTY(0 To 7)      As Byte         '���E����
    WS_ID(0 To 2)           As Byte         '�o�^�[��
    
    SHIIRE_CODE(0 To 4)     As Byte         '�d���溰��
    SHIIRE_TANKA(0 To 10)   As Byte         '�d���P��(9(8)V99)
    
    FILLER(0 To 40)         As Byte         'FILLER
    UPD_DATETIME(0 To 13)   As Byte         '�X�V����
End Type

'�f�[�^�E�o�b�t�@
Public P_NYUREC         As P_NYUREC_Tag

'�L�[��`
Type KEY0_P_NYU                         '�j�d�x�O
    JGYOBU(0 To 0)          As Byte         '���ƕ��敪
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '�i�ԁi�O���j
    NYUKA_DT(0 To 7)        As Byte         '���ד�
End Type


'�L�[��`
Type KEY1_P_NYU                         '�j�d�x�P
    JGYOBU(0 To 0)          As Byte         '���ƕ��敪
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '�i�ԁi�O���j
    SHIIRE_CODE(0 To 4)     As Byte         '�d���溰��
    SHIIRE_TANKA(0 To 10)   As Byte         '�d���P��(9(8)V99)
End Type


'�L�[�E�f�[�^
Public K0_P_NYU         As KEY0_P_NYU
Public K1_P_NYU         As KEY1_P_NYU

Type P_NYU_FSpeck
    fs              As BtFileSpeck      '̧�� ��߯��\����
    ks0             As BtKeySpeck       '�� ��߯��\����
    ks1             As BtKeySpeck       '�� ��߯��\����
    ks2             As BtKeySpeck       '�� ��߯��\����
    ks3             As BtKeySpeck       '�� ��߯��\����
    ks4             As BtKeySpeck       '�� ��߯��\����
    ks5             As BtKeySpeck       '�� ��߯��\����
    ks6             As BtKeySpeck       '�� ��߯��\����
    ks7             As BtKeySpeck       '�� ��߯��\����
    ks8             As BtKeySpeck       '�� ��߯��\����
End Type

Private P_NYU_Speck     As P_NYU_FSpeck

Private Function P_NYU_Create() As Integer
'********************************************************************
'*                                                                  *
'*              ���ޑO�؃f�[�^�@�b�q�d�`�s�d                        *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    P_NYU_Create = True
                                            '���ޑO�؃f�[�^�t���p�X�捞��
    sts = GetIni("FILE", P_NYU_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_NYU]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    P_NYU_Speck.fs.recoleng = Len(P_NYUREC)     ' ���R�[�h��
    P_NYU_Speck.fs.PageSize = P_NYU_PG_SIZ      ' �y�[�W�T�C�Y
    P_NYU_Speck.fs.idexnumb = 2                 ' �C���f�b�N�X��
    P_NYU_Speck.fs.fileflag = 0                 ' �t�@�C���t���O
    P_NYU_Speck.fs.reserve = &H0                ' �\��ς�
'------------------------------------------------
                                                ' �L�[�O
    P_NYU_Speck.ks0.keypos = 1                  ' �L�[�|�W�V����
    P_NYU_Speck.ks0.keyleng = 1                 ' �L�[��
    P_NYU_Speck.ks0.keyflag = BtKfExt + BtKfSeg ' �L�[�t���O
    P_NYU_Speck.ks0.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    P_NYU_Speck.ks0.reserve = &H0               ' �\��ς�
                                                ' �L�[�O
    P_NYU_Speck.ks1.keypos = 2                  ' �L�[�|�W�V����
    P_NYU_Speck.ks1.keyleng = 1                 ' �L�[��
    P_NYU_Speck.ks1.keyflag = BtKfExt + BtKfSeg ' �L�[�t���O
    P_NYU_Speck.ks1.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    P_NYU_Speck.ks1.reserve = &H0               ' �\��ς�
                                                ' �L�[�O
    P_NYU_Speck.ks2.keypos = 3                  ' �L�[�|�W�V����
    P_NYU_Speck.ks2.keyleng = 20                ' �L�[��
    P_NYU_Speck.ks2.keyflag = BtKfExt + BtKfSeg ' �L�[�t���O
    P_NYU_Speck.ks2.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    P_NYU_Speck.ks2.reserve = &H0               ' �\��ς�
                                                ' �L�[�O
    P_NYU_Speck.ks3.keypos = 23                 ' �L�[�|�W�V����
    P_NYU_Speck.ks3.keyleng = 8                 ' �L�[��
    P_NYU_Speck.ks3.keyflag = BtKfExt           ' �L�[�t���O
    P_NYU_Speck.ks3.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    P_NYU_Speck.ks3.reserve = &H0               ' �\��ς�

'------------------------------------------------
                                                ' �L�[�P
    P_NYU_Speck.ks4.keypos = 1                  ' �L�[�|�W�V����
    P_NYU_Speck.ks4.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    P_NYU_Speck.ks4.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    P_NYU_Speck.ks4.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    P_NYU_Speck.ks4.reserve = &H0               ' �\��ς�
                                                ' �L�[�P
    P_NYU_Speck.ks5.keypos = 2                  ' �L�[�|�W�V����
    P_NYU_Speck.ks5.keyleng = 1                 ' �L�[��
                                                ' �L�[�t���O
    P_NYU_Speck.ks5.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    P_NYU_Speck.ks5.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    P_NYU_Speck.ks5.reserve = &H0               ' �\��ς�
                                                ' �L�[�P
    P_NYU_Speck.ks6.keypos = 3                  ' �L�[�|�W�V����
    P_NYU_Speck.ks6.keyleng = 20                ' �L�[��
                                                 ' �L�[�t���O
    P_NYU_Speck.ks6.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    P_NYU_Speck.ks6.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    P_NYU_Speck.ks6.reserve = &H0               ' �\��ς�
                                                ' �L�[�P
    P_NYU_Speck.ks7.keypos = 58                 ' �L�[�|�W�V����
    P_NYU_Speck.ks7.keyleng = 5                 ' �L�[��
                                                ' �L�[�t���O
    P_NYU_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    P_NYU_Speck.ks7.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    P_NYU_Speck.ks7.reserve = &H0               ' �\��ς�

                                                ' �L�[�P
    P_NYU_Speck.ks8.keypos = 63                 ' �L�[�|�W�V����
    P_NYU_Speck.ks8.keyleng = 11                ' �L�[��
                                                ' �L�[�t���O
    P_NYU_Speck.ks8.keyflag = BtKfExt + BtKfDup + BtKfChg
    P_NYU_Speck.ks8.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    P_NYU_Speck.ks8.reserve = &H0               ' �\��ς�


'------------------------------------------------

    sts = BTRV(BtOpCreate, P_NYU_POS, P_NYU_Speck, Len(P_NYU_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���ޑO�؃f�[�^")
        Exit Function
    End If
    
    P_NYU_Create = False

End Function
Public Function P_NYU_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              ���ޑO�؃f�[�^�@�n�o�d�m                            *
'*                                                                  *
'*      ��  ��:Open Mode(Btrieve�Q��)                               *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    P_NYU_Open = True
                                        '���ޑO�؃f�[�^�t���p�X�捞��
    sts = GetIni("FILE", P_NYU_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_NYU]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, P_NYU_POS, P_NYUREC, Len(P_NYUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_NYU_Create()        '���ޑO�؃f�[�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_NYU_POS, P_NYUREC, Len(P_NYUREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���ޑO�؃f�[�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���ޑO�؃f�[�^")
                Exit Function
        End Select
    Loop

    P_NYU_Open = False

End Function


