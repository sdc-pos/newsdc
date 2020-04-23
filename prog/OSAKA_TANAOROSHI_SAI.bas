Attribute VB_Name = "OSAKA_TANAOROSHI_SAI"
Option Explicit

'********************************************************************
'*
'*              ���PC�@�I������F  �t�@�C����`
'*
'*          CREATE 2012.04.17
'********************************************************************
'�t�@�C���h�c
Public Const OSAKA_TANAOROSHI_SAI_ID$ = "OSAKA_TANAOROSHI_SAI"

'�y�[�W�T�C�Y
Private Const OSAKA_TANAOROSHI_SAI_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public OSAKA_TANAOROSHI_SAI_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Public Type OSAKA_TANAOROSHI_SAI_REC_Tag
    
    HIN_GAI(0 To 19)            As Byte         '���ޕi��
                            
    ST_SOKO(0 To 1)             As Byte         '�W�����ɑq�� �q��
    ST_RETU(0 To 1)             As Byte         '             ��
    ST_REN(0 To 1)              As Byte         '             �A
    ST_DAN(0 To 1)              As Byte         '             �i
    
    SHIZAI_ZAIKO_QTY(0 To 7)    As Byte         '���ލ݌ɐ�
    BUZAI_ZAIKO_QTY(0 To 7)     As Byte         '���ރZ���^�[�݌ɐ�

    SAI_SU(0 To 7)              As Byte         '���ِ�

    FILLER(0 To 75)             As Byte






End Type
'�f�[�^�E�o�b�t�@
Public OSAKA_TANAOROSHI_SAI_REC As OSAKA_TANAOROSHI_SAI_REC_Tag

'�L�[��`
    
Public Type KEY0_OSAKA_TANAOROSHI_SAI           '�j�d�x�O
    
    HIN_GAI(0 To 19)            As Byte         '���ޕi��
    
    
End Type
    
Public Type KEY1_OSAKA_TANAOROSHI_SAI           '�j�d�x�P
    
    ST_SOKO(0 To 1)             As Byte         '�W�����ɑq�� �q��
    ST_RETU(0 To 1)             As Byte         '             ��
    ST_REN(0 To 1)              As Byte         '             �A
    ST_DAN(0 To 1)              As Byte         '             �i
    
    HIN_GAI(0 To 19)            As Byte         '���ޕi��
    
    
End Type
    
    
    
    
    
'�L�[�E�f�[�^
Public K0_OSAKA_TANAOROSHI_SAI  As KEY0_OSAKA_TANAOROSHI_SAI
Public K1_OSAKA_TANAOROSHI_SAI  As KEY1_OSAKA_TANAOROSHI_SAI


Type OSAKA_TANAOROSHI_SAI_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
    ks1                     As BtKeySpeck   ' �� ��߯��\����
    ks2                     As BtKeySpeck   ' �� ��߯��\����
    ks3                     As BtKeySpeck   ' �� ��߯��\����
    ks4                     As BtKeySpeck   ' �� ��߯��\����
    ks5                     As BtKeySpeck   ' �� ��߯��\����

End Type

Private OSAKA_TANAOROSHI_SAI_Speck  As OSAKA_TANAOROSHI_SAI_FSpeck

Private Function OSAKA_TANAOROSHI_SAI_Create() As Integer
'********************************************************************
'*
'*              ���PC�@�I������F  �t�@�C����`
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

Dim Ret             As Long     '2007.11.13




    OSAKA_TANAOROSHI_SAI_Create = True
                                            '���PC�@�I������F  �t���p�X�捞��
    sts = GetIni("FILE", OSAKA_TANAOROSHI_SAI_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [OSAKA_TANAOROSHI_SAI]�@�ǂݍ��݃G���[")
        Exit Function
    End If

    FullPath = RTrim(c)


    OSAKA_TANAOROSHI_SAI_Speck.fs.recoleng = Len(OSAKA_TANAOROSHI_SAI_REC)      ' ���R�[�h��
    OSAKA_TANAOROSHI_SAI_Speck.fs.PageSize = OSAKA_TANAOROSHI_SAI_PG_SIZ        ' �y�[�W�T�C�Y
    OSAKA_TANAOROSHI_SAI_Speck.fs.idexnumb = 2                                  ' �C���f�b�N�X��
    OSAKA_TANAOROSHI_SAI_Speck.fs.fileflag = 0                                  ' �t�@�C���t���O
    OSAKA_TANAOROSHI_SAI_Speck.fs.reserve = &H0                                 ' �\��ς�
    
    '--------------------------------------------------- �L�[�O ��
    
    OSAKA_TANAOROSHI_SAI_Speck.ks0.keypos = 1                                   ' �L�[�|�W�V����
    OSAKA_TANAOROSHI_SAI_Speck.ks0.keyleng = 20                                 ' �L�[��
    OSAKA_TANAOROSHI_SAI_Speck.ks0.keyflag = BtKfExt + BtKfChg                  ' �L�[�t���O
    OSAKA_TANAOROSHI_SAI_Speck.ks0.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    OSAKA_TANAOROSHI_SAI_Speck.ks0.reserve = &H0                     ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    
    '--------------------------------------------------- �L�[�P ��
    OSAKA_TANAOROSHI_SAI_Speck.ks1.keypos = 21                                  ' �L�[�|�W�V����
    OSAKA_TANAOROSHI_SAI_Speck.ks1.keyleng = 2                                  ' �L�[��
    OSAKA_TANAOROSHI_SAI_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfSeg        ' �L�[�t���O
    OSAKA_TANAOROSHI_SAI_Speck.ks1.keytype = Chr(BtKtString)                    ' �L�[�^�C�v
    OSAKA_TANAOROSHI_SAI_Speck.ks1.reserve = &H0                                ' �\��ς�
    
    OSAKA_TANAOROSHI_SAI_Speck.ks2.keypos = 23                                  ' �L�[�|�W�V����
    OSAKA_TANAOROSHI_SAI_Speck.ks2.keyleng = 2                                  ' �L�[��
    OSAKA_TANAOROSHI_SAI_Speck.ks2.keyflag = BtKfExt + BtKfChg + BtKfSeg        ' �L�[�t���O
    OSAKA_TANAOROSHI_SAI_Speck.ks2.keytype = Chr(BtKtString)                    ' �L�[�^�C�v
    OSAKA_TANAOROSHI_SAI_Speck.ks2.reserve = &H0                                ' �\��ς�
    
    OSAKA_TANAOROSHI_SAI_Speck.ks3.keypos = 25                                  ' �L�[�|�W�V����
    OSAKA_TANAOROSHI_SAI_Speck.ks3.keyleng = 2                                  ' �L�[��
    OSAKA_TANAOROSHI_SAI_Speck.ks3.keyflag = BtKfExt + BtKfChg + BtKfSeg        ' �L�[�t���O
    OSAKA_TANAOROSHI_SAI_Speck.ks3.keytype = Chr(BtKtString)                    ' �L�[�^�C�v
    OSAKA_TANAOROSHI_SAI_Speck.ks3.reserve = &H0                                ' �\��ς�
    
    OSAKA_TANAOROSHI_SAI_Speck.ks4.keypos = 27                                  ' �L�[�|�W�V����
    OSAKA_TANAOROSHI_SAI_Speck.ks4.keyleng = 2                                  ' �L�[��
    OSAKA_TANAOROSHI_SAI_Speck.ks4.keyflag = BtKfExt + BtKfChg + BtKfSeg        ' �L�[�t���O
    OSAKA_TANAOROSHI_SAI_Speck.ks4.keytype = Chr(BtKtString)                    ' �L�[�^�C�v
    OSAKA_TANAOROSHI_SAI_Speck.ks4.reserve = &H0                                ' �\��ς�
    
    OSAKA_TANAOROSHI_SAI_Speck.ks5.keypos = 1                                   ' �L�[�|�W�V����
    OSAKA_TANAOROSHI_SAI_Speck.ks5.keyleng = 20                                 ' �L�[��
    OSAKA_TANAOROSHI_SAI_Speck.ks5.keyflag = BtKfExt + BtKfChg                  ' �L�[�t���O
    OSAKA_TANAOROSHI_SAI_Speck.ks5.keytype = Chr(BtKtString)                    ' �L�[�^�C�v
    OSAKA_TANAOROSHI_SAI_Speck.ks5.reserve = &H0                                ' �\��ς�
    '--------------------------------------------------- �L�[�P ��
    sts = BTRV(BtOpCreate, OSAKA_TANAOROSHI_SAI_POS, OSAKA_TANAOROSHI_SAI_Speck, Len(OSAKA_TANAOROSHI_SAI_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���PC�@�I������F")
        Exit Function
    End If
    
    OSAKA_TANAOROSHI_SAI_Create = False

End Function

Public Function OSAKA_TANAOROSHI_SAI_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���PC�@�I������F  �t�@�C����`
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

Dim Ret             As Long     '2007.11.13


    OSAKA_TANAOROSHI_SAI_Open = True
                                            '���PC�@�I������F  �t���p�X�捞��
    sts = GetIni("FILE", OSAKA_TANAOROSHI_SAI_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [OSAKA_TANAOROSHI_SAI]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, OSAKA_TANAOROSHI_SAI_POS, OSAKA_TANAOROSHI_SAI_REC, Len(OSAKA_TANAOROSHI_SAI_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = OSAKA_TANAOROSHI_SAI_Create()   '���PC�@�I������F �ް��쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, OSAKA_TANAOROSHI_SAI_POS, OSAKA_TANAOROSHI_SAI_REC, Len(OSAKA_TANAOROSHI_SAI_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���PC�@�I������F")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���PC�@�I������F")
                Exit Function
        End Select
    Loop
    
    OSAKA_TANAOROSHI_SAI_Open = False

End Function

