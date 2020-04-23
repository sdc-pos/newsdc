Attribute VB_Name = "OLD_FURIKAE"
Option Explicit
'********************************************************************
'*
'*              �i�ԐU�ւl�@�t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const OLD_FURIKAE_ID$ = "OLD_FURIKAE"

'�y�[�W�T�C�Y
Public Const OLD_FURIKAE_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public OLD_FURIKAE_POS As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type OLD_FURIKAEREC_Tag
    HIN_MAE(0 To 19)            As Byte     '�U�֑O�i�ԁi�O���j
    HIN_GO(0 To 19)             As Byte     '�U�֌�i�ԁi�O���j
    BIKOU(0 To 39)              As Byte     '���l
    
    FILLER(0 To 31)             As Byte    '
    
    INS_TANTO(0 To 9)           As Byte     '�ǉ��@�S����
    Ins_DateTime(0 To 13)       As Byte     '�ǉ��@����

    UPD_TANTO(0 To 9)           As Byte     '�X�V�@�S����
    UPD_DATETIME(0 To 13)       As Byte     '�X�V�@����

End Type

'�f�[�^�E�o�b�t�@
Public OLD_FURIKAEREC           As OLD_FURIKAEREC_Tag

'�L�[��`
Type KEY0_OLD_FURIKAE           '�j�d�x�O
    HIN_MAE(0 To 19)                    As Byte     '�U�֑O�i�ԁi�O���j
    HIN_GO(0 To 19)                     As Byte     '�U�֌�i�ԁi�O���j
End Type

Type KEY1_OLD_FURIKAE           '�j�d�x�P
    HIN_GO(0 To 19)                     As Byte     '�U�֌�i�ԁi�O���j
    HIN_MAE(0 To 19)                    As Byte     '�U�֑O�i�ԁi�O���j
End Type


'�L�[�E�f�[�^
Public K0_OLD_FURIKAE               As KEY0_OLD_FURIKAE
Public K1_OLD_FURIKAE               As KEY1_OLD_FURIKAE

Type OLD_FURIKAE_FSpeck
    fs      As BtFileSpeck          ' ̧�� ��߯��\����
    ks0     As BtKeySpeck           ' �� ��߯��\����
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck

End Type

Private OLD_FURIKAE_Speck               As OLD_FURIKAE_FSpeck
Private Function OLD_FURIKAE_Create() As Integer
'********************************************************************
'*
'*              �i�ԐU�ւl�@�b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    OLD_FURIKAE_Create = True
                                            '�i�ԐU�ւl�t���p�X�捞��
    sts = GetIni("FILE", OLD_FURIKAE_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [" & OLD_FURIKAE_ID & "]�ǂݍ��݃G���[")
        Exit Function
    End If
     
    FullPath = RTrim(c)
    
    OLD_FURIKAE_Speck.fs.recoleng = Len(OLD_FURIKAEREC)         ' ���R�[�h��
    OLD_FURIKAE_Speck.fs.PageSize = OLD_FURIKAE_PG_SIZ          ' �y�[�W�T�C�Y
    OLD_FURIKAE_Speck.fs.idexnumb = 2                   ' �C���f�b�N�X��
    OLD_FURIKAE_Speck.fs.fileflag = 0                   ' �t�@�C���t���O
    OLD_FURIKAE_Speck.fs.reserve = &H0                  ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�O
    OLD_FURIKAE_Speck.ks0.keypos = 1                   ' �L�[�|�W�V����
                                                ' �L�[��
    OLD_FURIKAE_Speck.ks0.keyleng = 20
                                                ' �L�[�t���O
    OLD_FURIKAE_Speck.ks0.keyflag = BtKfExt + BtKfSeg '+ BtKfDup
    OLD_FURIKAE_Speck.ks0.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    OLD_FURIKAE_Speck.ks0.reserve = &H0                 ' �\��ς�


    OLD_FURIKAE_Speck.ks1.keypos = 21               ' �L�[�|�W�V����
                                                ' �L�[��
    OLD_FURIKAE_Speck.ks1.keyleng = 20
                                                ' �L�[�t���O
    OLD_FURIKAE_Speck.ks1.keyflag = BtKfExt  '+ BtKfDup
    OLD_FURIKAE_Speck.ks1.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    OLD_FURIKAE_Speck.ks1.reserve = &H0                 ' �\��ς�


'-----------------------------------------------
                                                ' �L�[�P
    OLD_FURIKAE_Speck.ks2.keypos = 21                   ' �L�[�|�W�V����
    OLD_FURIKAE_Speck.ks2.keyleng = 20                   ' �L�[��
                                                ' �L�[�t���O
    OLD_FURIKAE_Speck.ks2.keyflag = BtKfExt + BtKfDup + BtKfSeg
    OLD_FURIKAE_Speck.ks2.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    OLD_FURIKAE_Speck.ks2.reserve = &H0                 ' �\��ς�

    OLD_FURIKAE_Speck.ks3.keypos = 1                   ' �L�[�|�W�V����
    OLD_FURIKAE_Speck.ks3.keyleng = 20                   ' �L�[��
                                                ' �L�[�t���O
    OLD_FURIKAE_Speck.ks3.keyflag = BtKfExt + BtKfDup
    OLD_FURIKAE_Speck.ks3.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    OLD_FURIKAE_Speck.ks3.reserve = &H0                 ' �\��ς�


'-----------------------------------------------

    sts = BTRV(BtOpCreate, OLD_FURIKAE_POS, OLD_FURIKAE_Speck, Len(OLD_FURIKAE_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���@�i�ԐU�ւl")
        Exit Function
    End If

    OLD_FURIKAE_Create = False

End Function

Public Function OLD_FURIKAE_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �i�ԐU�ւl�@�n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    OLD_FURIKAE_Open = True
                                            '�i�ԐU�ւl�t���p�X�捞��
    sts = GetIni("FILE", OLD_FURIKAE_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [OLD_FURIKAE]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, OLD_FURIKAE_POS, OLD_FURIKAEREC, Len(OLD_FURIKAEREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = OLD_FURIKAE_Create()        '�i�ԐU�ւl�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, OLD_FURIKAE_POS, OLD_FURIKAEREC, Len(OLD_FURIKAEREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���@�i�ԐU�ւl")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���@�i�ԐU�ւl")
                Exit Function
        End Select
    Loop
    OLD_FURIKAE_Open = False
End Function


