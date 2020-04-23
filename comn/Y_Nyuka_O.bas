Attribute VB_Name = "Y_NYU_O"
Option Explicit
'********************************************************************
'*                                                                  *
'*              ���ח\��f�[�^�i���PC�����j  �t�@�C����`          *
'*                                                                  *
'********************************************************************
'�t�@�C���h�c
Public Const Y_NYU_O_ID$ = "Y_NYU_O"

'�y�[�W�T�C�Y
Public Const Y_NYU_O_PG_SIZ% = 2048

'�|�W�V�����E�u���b�N
Public Y_NYU_O_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type Y_NYUREC_O_Tag
    JGYOBU(0 To 0)              As Byte     '���ƕ�
    SOKO_NO(0 To 1)             As Byte     '�q�ɇ�
    SEQ_NO(0 To 2)              As Byte     'SEQ_NO
    NYUKO_YMD(0 To 7)           As Byte     '���ɓ�(���ד�)
    DEN_NO(0 To 5)              As Byte     '�`�[��
    MAKER_CODE(0 To 5)          As Byte     'Ұ������
    NAIGAI(0 To 0)              As Byte     '�����O
    HIN_NO(0 To 19)             As Byte     '�i��
    Y_SURYO(0 To 7)             As Byte     '�\�萔��
    J_SURYO(0 To 7)             As Byte     '���ѐ���
    TANTO_CODE(0 To 4)          As Byte     '�S���Һ���
    ORDER_NO(0 To 9)            As Byte     '������
    KENPIN_F(0 To 0)            As Byte     '���iF
    WEL_ID(0 To 2)              As Byte     '�g�p�q�@ID
    PRG_ID(0 To 7)              As Byte     '�g�p���v���O����
    FILLER(0 To 165)            As Byte     'FILLER
    
End Type

'�f�[�^�E�o�b�t�@
Public Y_NYU_O_REC                  As Y_NYUREC_O_Tag

'�L�[��`
Type KEY0_Y_NYU_O            '�j�d�x�O
    SEQ_NO(0 To 2)              As Byte     'SEQ_NO
End Type

Type KEY1_Y_NYU_O            '�j�d�x�P
    JGYOBU(0 To 0)              As Byte     '���ƕ�
    NAIGAI(0 To 0)              As Byte     '�����O
    HIN_NO(0 To 19)             As Byte     '�i��
End Type

Type KEY2_Y_NYU_O            '�j�d�x�P
    WEL_ID(0 To 2)              As Byte     '�g�p�q�@ID
    PRG_ID(0 To 7)              As Byte     '�g�p���v���O����
End Type



'�L�[�E�f�[�^
Public K0_Y_NYU_O               As KEY0_Y_NYU_O
Public K1_Y_NYU_O               As KEY1_Y_NYU_O
Public K2_Y_NYU_O               As KEY2_Y_NYU_O

Private Type Y_NYU_O_FSpeck
    fs      As BtFileSpeck              ' ̧�� ��߯��\����
    ks0     As BtKeySpeck               ' �� ��߯��\����
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck               ' �� ��߯��\����
    ks4     As BtKeySpeck
    ks5     As BtKeySpeck
End Type

Private Y_NYU_O_Speck As Y_NYU_O_FSpeck

Private Function Y_NYU_O_Create() As Integer
'********************************************************************
'*                                                                  *
'*              ���ח\��f�[�^(���PC����)  �b�q�d�`�s�d            *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    Y_NYU_O_Create = True
                                            '���ח\��f�[�^�t���p�X�捞��
    sts = GetIni("FILE", Y_NYU_O_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [Y_NYU_O]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    Y_NYU_O_Speck.fs.recoleng = Len(Y_NYU_O_REC)    ' ���R�[�h��
    Y_NYU_O_Speck.fs.PageSize = Y_NYU_O_PG_SIZ      ' �y�[�W�T�C�Y
    Y_NYU_O_Speck.fs.idexnumb = 3                   ' �C���f�b�N�X��
    Y_NYU_O_Speck.fs.fileflag = 0                   ' �t�@�C���t���O
    Y_NYU_O_Speck.fs.reserve = &H0                  ' �\��ς�
    '-------------------------------------------
                                                
    Y_NYU_O_Speck.ks0.keypos = 4                    ' �L�[�|�W�V����
    Y_NYU_O_Speck.ks0.keyleng = 3                   ' �L�[��
                                                    ' �L�[�t���O
    Y_NYU_O_Speck.ks0.keyflag = BtKfExt
    Y_NYU_O_Speck.ks0.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    Y_NYU_O_Speck.ks0.reserve = &H0                 ' �\��ς�
                                                
                                                
                                                ' �L�[�O
    '-------------------------------------------
    
    '-------------------------------------------
                                                ' �L�[�P
    Y_NYU_O_Speck.ks1.keypos = 1                    ' �L�[�|�W�V����
    Y_NYU_O_Speck.ks1.keyleng = 1                   ' �L�[��
                                                    ' �L�[�t���O
    Y_NYU_O_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_NYU_O_Speck.ks1.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    Y_NYU_O_Speck.ks1.reserve = &H0                 ' �\��ς�
                                                
    Y_NYU_O_Speck.ks2.keypos = 27                    ' �L�[�|�W�V����
    Y_NYU_O_Speck.ks2.keyleng = 1                   ' �L�[��
                                                    ' �L�[�t���O
    Y_NYU_O_Speck.ks2.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_NYU_O_Speck.ks2.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    Y_NYU_O_Speck.ks2.reserve = &H0                 ' �\��ς�
                                                
    Y_NYU_O_Speck.ks3.keypos = 28                    ' �L�[�|�W�V����
    Y_NYU_O_Speck.ks3.keyleng = 20                   ' �L�[��
                                                    ' �L�[�t���O
    Y_NYU_O_Speck.ks3.keyflag = BtKfExt + BtKfChg + BtKfDup
    Y_NYU_O_Speck.ks3.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    Y_NYU_O_Speck.ks3.reserve = &H0                 ' �\��ς�
                                                
                                                
                                                ' �L�[�P
    '-------------------------------------------
    
    
    '-------------------------------------------
                                                ' �L�[�Q
    Y_NYU_O_Speck.ks4.keypos = 80                   ' �L�[�|�W�V����
    Y_NYU_O_Speck.ks4.keyleng = 3                   ' �L�[��
                                                    ' �L�[�t���O
    Y_NYU_O_Speck.ks4.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_NYU_O_Speck.ks4.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    Y_NYU_O_Speck.ks4.reserve = &H0                 ' �\��ς�
                                                
    Y_NYU_O_Speck.ks5.keypos = 83                   ' �L�[�|�W�V����
    Y_NYU_O_Speck.ks5.keyleng = 8                   ' �L�[��
                                                    ' �L�[�t���O
    Y_NYU_O_Speck.ks5.keyflag = BtKfExt + BtKfChg + BtKfDup
    Y_NYU_O_Speck.ks5.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    Y_NYU_O_Speck.ks5.reserve = &H0                 ' �\��ς�
                                                
                                                
                                                ' �L�[�Q
    '-------------------------------------------
    
    
    sts = BTRV(BtOpCreate, Y_NYU_O_POS, Y_NYU_O_Speck, Len(Y_NYU_O_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���ח\��f�[�^")
        Y_NYU_O_Create = True
        Exit Function
    End If

    Y_NYU_O_Create = False

End Function

Function Y_NYU_O_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              ���ח\��f�[�^(���PC����)  �n�o�d�m                *
'*                                                                  *
'*      ��  ��:Open Mode(Btrieve�Q��)                               *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    Y_NYU_O_Open = True
                                            '���ח\��f�[�^�t���p�X�捞��
    sts = GetIni("FILE", Y_NYU_O_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [Y_NYU]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, Y_NYU_O_POS, Y_NYU_O_REC, Len(Y_NYU_O_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = Y_NYU_O_Create()        '���ח\��f�[�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, Y_NYU_O_POS, Y_NYU_O_REC, Len(Y_NYU_O_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���ח\��f�[�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���ח\��f�[�^")
                Exit Function
        End Select
    Loop
    
    Y_NYU_O_Open = False

End Function


