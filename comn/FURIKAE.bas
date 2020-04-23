Attribute VB_Name = "FURIKAE"
Option Explicit
'********************************************************************
'*
'*              �i�ԐU�ւl�@�t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const FURIKAE_ID$ = "FURIKAE"

'�y�[�W�T�C�Y
Public Const FURIKAE_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public FURIKAE_POS As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type FURIKAEREC_Tag
    JGYOBU_MAE(0 To 0)          As Byte     '�U�֑O���ƕ�           2012.03.13
    NAIGAI_MAE(0 To 0)          As Byte     '�U�֑O�����O           2012.03.13
    HIN_MAE(0 To 19)            As Byte     '�U�֑O�i�ԁi�O���j
    JGYOBU_GO(0 To 0)           As Byte     '�U�֌㎖�ƕ�           2012.03.13
    NAIGAI_GO(0 To 0)           As Byte     '�U�֌㍑���O           2012.03.13
    HIN_GO(0 To 19)             As Byte     '�U�֌�i�ԁi�O���j
    BIKOU(0 To 39)              As Byte     '���l
    
    CUT_SU(0 To 2)              As Byte     '�ؒf��                 2012.03.14
    
    
    MOTO_LEN(0 To 2)            As Byte     '���̒���               2012.12.26
    
    
    KO_QTY(0 To 3)              As Byte     '����                   2013.02.22
    
    
    FILLER(0 To 17)             As Byte    '                        2013.02.22 �����ύX
    
    INS_TANTO(0 To 9)           As Byte     '�ǉ��@�S����
    Ins_DateTime(0 To 13)       As Byte     '�ǉ��@����

    UPD_TANTO(0 To 9)           As Byte     '�X�V�@�S����
    UPD_DATETIME(0 To 13)       As Byte     '�X�V�@����

End Type

'�f�[�^�E�o�b�t�@
Public FURIKAEREC   As FURIKAEREC_Tag

'�L�[��`
Type KEY0_FURIKAE            '�j�d�x�O
    JGYOBU_MAE(0 To 0)                  As Byte     '�U�֑O���ƕ�           2012.03.13
    NAIGAI_MAE(0 To 0)                  As Byte     '�U�֑O�����O           2012.03.13
    HIN_MAE(0 To 19)                    As Byte     '�U�֑O�i�ԁi�O���j
    JGYOBU_GO(0 To 0)                   As Byte     '�U�֌㎖�ƕ�           2012.03.13
    NAIGAI_GO(0 To 0)                   As Byte     '�U�֌㍑���O           2012.03.13
    HIN_GO(0 To 19)                     As Byte     '�U�֌�i�ԁi�O���j
End Type

Type KEY1_FURIKAE            '�j�d�x�P
    JGYOBU_GO(0 To 0)                   As Byte     '�U�֌㎖�ƕ�           2012.03.13
    NAIGAI_GO(0 To 0)                   As Byte     '�U�֌㍑���O           2012.03.13
    HIN_GO(0 To 19)                     As Byte     '�U�֌�i�ԁi�O���j
    JGYOBU_MAE(0 To 0)                  As Byte     '�U�֑O���ƕ�           2012.03.13
    NAIGAI_MAE(0 To 0)                  As Byte     '�U�֑O�����O           2012.03.13
    HIN_MAE(0 To 19)                    As Byte     '�U�֑O�i�ԁi�O���j
End Type


'�L�[�E�f�[�^
Public K0_FURIKAE                   As KEY0_FURIKAE
Public K1_FURIKAE                   As KEY1_FURIKAE

Type FURIKAE_FSpeck
    fs      As BtFileSpeck          ' ̧�� ��߯��\����
    ks0     As BtKeySpeck           ' �� ��߯��\����
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck
    ks4     As BtKeySpeck           '2012.03.13
    ks5     As BtKeySpeck           '2012.03.13
    ks6     As BtKeySpeck           '2012.03.13
    ks7     As BtKeySpeck           '2012.03.13
    ks8     As BtKeySpeck           '2012.03.13
    ks9     As BtKeySpeck           '2012.03.13
    ks10    As BtKeySpeck           '2012.03.13
    ks11    As BtKeySpeck           '2012.03.13

End Type

Private FURIKAE_Speck               As FURIKAE_FSpeck
Private Function FURIKAE_Create() As Integer
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

    FURIKAE_Create = True
                                            '�i�ԐU�ւl�t���p�X�捞��
    sts = GetIni("FILE", FURIKAE_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [" & FURIKAE_ID & "]�ǂݍ��݃G���[")
        Exit Function
    End If
     
    FullPath = RTrim(c)
    
    FURIKAE_Speck.fs.recoleng = Len(FURIKAEREC)         ' ���R�[�h��
    FURIKAE_Speck.fs.PageSize = FURIKAE_PG_SIZ          ' �y�[�W�T�C�Y
    FURIKAE_Speck.fs.idexnumb = 2                   ' �C���f�b�N�X��
    FURIKAE_Speck.fs.fileflag = 0                   ' �t�@�C���t���O
    FURIKAE_Speck.fs.reserve = &H0                  ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�O
    FURIKAE_Speck.ks0.keypos = 1                ' �L�[�|�W�V����
                                                ' �L�[��
    FURIKAE_Speck.ks0.keyleng = 1
                                                ' �L�[�t���O
    FURIKAE_Speck.ks0.keyflag = BtKfExt + BtKfSeg '+ BtKfDup
    FURIKAE_Speck.ks0.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    FURIKAE_Speck.ks0.reserve = &H0                 ' �\��ς�


    FURIKAE_Speck.ks1.keypos = 2                ' �L�[�|�W�V����
                                                ' �L�[��
    FURIKAE_Speck.ks1.keyleng = 1
                                                ' �L�[�t���O
    FURIKAE_Speck.ks1.keyflag = BtKfExt + BtKfSeg '+ BtKfDup
    FURIKAE_Speck.ks1.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    FURIKAE_Speck.ks1.reserve = &H0                 ' �\��ς�

    FURIKAE_Speck.ks2.keypos = 3                ' �L�[�|�W�V����
                                                ' �L�[��
    FURIKAE_Speck.ks2.keyleng = 20
                                                ' �L�[�t���O
    FURIKAE_Speck.ks2.keyflag = BtKfExt + BtKfSeg '+ BtKfDup
    FURIKAE_Speck.ks2.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    FURIKAE_Speck.ks2.reserve = &H0                 ' �\��ς�

    FURIKAE_Speck.ks3.keypos = 23                ' �L�[�|�W�V����
                                                ' �L�[��
    FURIKAE_Speck.ks3.keyleng = 1
                                                ' �L�[�t���O
    FURIKAE_Speck.ks3.keyflag = BtKfExt + BtKfSeg '+ BtKfDup
    FURIKAE_Speck.ks3.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    FURIKAE_Speck.ks3.reserve = &H0                 ' �\��ς�


    FURIKAE_Speck.ks4.keypos = 24                ' �L�[�|�W�V����
                                                ' �L�[��
    FURIKAE_Speck.ks4.keyleng = 1
                                                ' �L�[�t���O
    FURIKAE_Speck.ks4.keyflag = BtKfExt + BtKfSeg '+ BtKfDup
    FURIKAE_Speck.ks4.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    FURIKAE_Speck.ks4.reserve = &H0                 ' �\��ς�


    FURIKAE_Speck.ks5.keypos = 25                ' �L�[�|�W�V����
                                                ' �L�[��
    FURIKAE_Speck.ks5.keyleng = 20
                                                ' �L�[�t���O
    FURIKAE_Speck.ks5.keyflag = BtKfExt  '+ BtKfDup
    FURIKAE_Speck.ks5.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    FURIKAE_Speck.ks5.reserve = &H0                 ' �\��ς�

'-----------------------------------------------
                                                ' �L�[�P
    FURIKAE_Speck.ks6.keypos = 23                   ' �L�[�|�W�V����
    FURIKAE_Speck.ks6.keyleng = 1                   ' �L�[��
                                                ' �L�[�t���O
    FURIKAE_Speck.ks6.keyflag = BtKfExt + BtKfSeg
    FURIKAE_Speck.ks6.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    FURIKAE_Speck.ks6.reserve = &H0                 ' �\��ς�

    FURIKAE_Speck.ks7.keypos = 24                   ' �L�[�|�W�V����
    FURIKAE_Speck.ks7.keyleng = 1                   ' �L�[��
                                                ' �L�[�t���O
    FURIKAE_Speck.ks7.keyflag = BtKfExt + BtKfSeg
    FURIKAE_Speck.ks7.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    FURIKAE_Speck.ks7.reserve = &H0                 ' �\��ς�

    FURIKAE_Speck.ks8.keypos = 25                   ' �L�[�|�W�V����
    FURIKAE_Speck.ks8.keyleng = 20                   ' �L�[��
                                                ' �L�[�t���O
    FURIKAE_Speck.ks8.keyflag = BtKfExt + BtKfSeg
    FURIKAE_Speck.ks8.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    FURIKAE_Speck.ks8.reserve = &H0                 ' �\��ς�



    FURIKAE_Speck.ks9.keypos = 1                   ' �L�[�|�W�V����
    FURIKAE_Speck.ks9.keyleng = 1                   ' �L�[��
                                                ' �L�[�t���O
    FURIKAE_Speck.ks9.keyflag = BtKfExt + BtKfSeg
    FURIKAE_Speck.ks9.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    FURIKAE_Speck.ks9.reserve = &H0                 ' �\��ς�

    FURIKAE_Speck.ks10.keypos = 2                   ' �L�[�|�W�V����
    FURIKAE_Speck.ks10.keyleng = 1                  ' �L�[��
                                                    ' �L�[�t���O
    FURIKAE_Speck.ks10.keyflag = BtKfExt + BtKfSeg
    FURIKAE_Speck.ks10.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    FURIKAE_Speck.ks10.reserve = &H0                ' �\��ς�

    FURIKAE_Speck.ks11.keypos = 3                   ' �L�[�|�W�V����
    FURIKAE_Speck.ks11.keyleng = 20                 ' �L�[��
                                                    ' �L�[�t���O
    FURIKAE_Speck.ks11.keyflag = BtKfExt
    FURIKAE_Speck.ks11.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    FURIKAE_Speck.ks11.reserve = &H0                ' �\��ς�


'-----------------------------------------------

    sts = BTRV(BtOpCreate, FURIKAE_POS, FURIKAE_Speck, Len(FURIKAE_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�i�ԐU�ւl")
        Exit Function
    End If

    FURIKAE_Create = False

End Function

Public Function FURIKAE_Open(Mode As Integer) As Integer
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
    
    FURIKAE_Open = True
                                            '�i�ԐU�ւl�t���p�X�捞��
    sts = GetIni("FILE", FURIKAE_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [FURIKAE]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, FURIKAE_POS, FURIKAEREC, Len(FURIKAEREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = FURIKAE_Create()        '�i�ԐU�ւl�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, FURIKAE_POS, FURIKAEREC, Len(FURIKAEREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�i�ԐU�ւl")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�i�ԐU�ւl")
                Exit Function
        End Select
    Loop
    FURIKAE_Open = False
End Function


Function FURIKAE_Get(JGYOBU As String, NAIGAI As String, HIN_MAE As String, HIN_GO As String, Locked As Integer)
'----------------------------------------------------------------------------
'                   �i�ԐU�ւl�t�@�C���f����

'       Locked      :False=NormalGet,ۯ�����Btrieve���ڰ��݂�ۯ��萔
'----------------------------------------------------------------------------
Dim com As Integer
Dim sts As Integer
Dim yn As Integer

    FURIKAE_Get = True
    
    Call UniCode_Conv(K0_FURIKAE.JGYOBU_MAE, JGYOBU)    '2012.03.13
    Call UniCode_Conv(K0_FURIKAE.NAIGAI_MAE, NAIGAI)    '2012.03.13
    Call UniCode_Conv(K0_FURIKAE.HIN_MAE, HIN_MAE)
    
    Call UniCode_Conv(K0_FURIKAE.JGYOBU_GO, JGYOBU)     '2012.03.13
    Call UniCode_Conv(K0_FURIKAE.NAIGAI_GO, NAIGAI)     '2012.03.13
    Call UniCode_Conv(K0_FURIKAE.HIN_GO, HIN_GO)
    com = BtOpGetEqual + Locked
Do
    sts = BTRV(com, FURIKAE_POS, FURIKAEREC, Len(FURIKAEREC), K0_FURIKAE, Len(K0_FURIKAE), 0)
    
    Select Case sts
        Case BtNoErr
            Exit Do
        Case BtErrKeyNotFound       '���R�[�h����
            
            'MsgBox "�w�肳�ꂽ�H��������܂���B"
            Exit Function
        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
            yn = MsgBox("���Ŏg�p���ł��I<FURIKAE>" & Chr(13) & Chr(10) & _
                        "�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
            If yn = vbNo Then Exit Function
        Case Else
            Call File_Error(sts, com, "�i�ԐU�ւl")
            Exit Function
    End Select
Loop

    FURIKAE_Get = False

End Function
Sub FURIKAE_CLOSE()
Dim sts As Integer

    sts = BTRV(BtOpClose, FURIKAE_POS, FURIKAEREC, Len(FURIKAEREC), K0_FURIKAE, Len(K0_FURIKAE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ԐU�ւl")
        End If
    End If

End Sub
Sub FURIKAE_CLR()


    Call UniCode_Conv(FURIKAEREC.JGYOBU_MAE, "")        '2012.03.13
    Call UniCode_Conv(FURIKAEREC.NAIGAI_MAE, "")        '2012.03.13
    Call UniCode_Conv(FURIKAEREC.HIN_MAE, "")
    
    Call UniCode_Conv(FURIKAEREC.JGYOBU_GO, "")         '2012.03.13
    Call UniCode_Conv(FURIKAEREC.NAIGAI_GO, "")         '2012.03.13
    Call UniCode_Conv(FURIKAEREC.HIN_GO, "")            '2012.03.14
    Call UniCode_Conv(FURIKAEREC.BIKOU, "")
    
    Call UniCode_Conv(FURIKAEREC.CUT_SU, "")
    
    
    Call UniCode_Conv(FURIKAEREC.FILLER, "")
    
    Call UniCode_Conv(FURIKAEREC.INS_TANTO, "")
    Call UniCode_Conv(FURIKAEREC.Ins_DateTime, "")
    Call UniCode_Conv(FURIKAEREC.UPD_TANTO, "")
    Call UniCode_Conv(FURIKAEREC.UPD_DATETIME, "")
    
    'Call UniCode_Conv(FURIKAEREC.FILLER, String(UBound(FURIKAEREC.FILLER) + 1, "0"))

End Sub

