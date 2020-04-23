Attribute VB_Name = "PSTOCK"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �͈͓��ړ����݌Ɉꗗ�@�t�@�C����`                  *
'*                                                                  *
'*          CREATE 2004.04.27                                       *
'********************************************************************
'�t�@�C���h�c
Public Const PSTOCK_ID = "PSTOCK"

'�y�[�W�T�C�Y
Public Const PSTOCK_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public PSTOCK_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Private Type PSTOCKREC_Tag
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
    ST_Location(0 To 7) As Byte     '�W���I��
    T_Zai_Qty(0 To 7)   As Byte     '�݌ɑ���
    HS_ZAIQTY(0 To 7)   As Byte     'νč݌ɐ�
    Plus_QTY(0 To 7)    As Byte     '�݌Ɂ{
    Minus_QTY(0 To 7)   As Byte     '�݌Ɂ|
End Type

'�f�[�^�E�o�b�t�@
Public PSTOCKREC        As PSTOCKREC_Tag

'�L�[��`
Private Type KEY0_PSTOCK            '�j�d�x�O
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
End Type

Private Type KEY1_PSTOCK            '�j�d�x�P
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    ST_Location(0 To 7) As Byte     '�W���I��
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
End Type

'�L�[�E�f�[�^
Public K0_PSTOCK        As KEY0_PSTOCK
Public K1_PSTOCK        As KEY1_PSTOCK

Private Type PSTOCK_FSpeck
    fs As BtFileSpeck               ' ̧�� ��߯��\����
    ks0 As BtKeySpeck               ' �� ��߯��\����
    ks1 As BtKeySpeck               ' �� ��߯��\����
    ks2 As BtKeySpeck               ' �� ��߯��\����
    ks3 As BtKeySpeck               ' �� ��߯��\����
    ks4 As BtKeySpeck               ' �� ��߯��\����
    ks5 As BtKeySpeck               ' �� ��߯��\����
    ks6 As BtKeySpeck               ' �� ��߯��\����
End Type

Public PSTOCK_Speck As PSTOCK_FSpeck

Private Function PSTOCK_Create() As Integer
'********************************************************************
'*                                                                  *
'*              �͈͓��ړ����݌Ɉꗗ�f�[�^�@�b�q�d�`�s�d            *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 2004.04.27                                       *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    PSTOCK_Create = True
                                            '�͈͓��ړ����݌Ɉꗗ�f�[�^�t���p�X�捞��
    sts = GetIni("FILE", PSTOCK_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI[PSTOCK] �ǂݍ��݃G���[ ")
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    PSTOCK_Speck.fs.recoleng = Len(PSTOCKREC)       ' ���R�[�h��
    PSTOCK_Speck.fs.PageSize = PSTOCK_PG_SIZ        ' �y�[�W�T�C�Y
    PSTOCK_Speck.fs.idexnumb = 2                    ' �C���f�b�N�X��
    PSTOCK_Speck.fs.fileflag = 0                    ' �t�@�C���t���O
    PSTOCK_Speck.fs.reserve = &H0                   ' �\��ς�
'-----------------------------------------------    ' �L�[�O
    PSTOCK_Speck.ks0.keypos = 1                     ' �L�[�|�W�V����
    PSTOCK_Speck.ks0.keyleng = 1                    ' �L�[��
    PSTOCK_Speck.ks0.keyflag = BtKfExt + BtKfSeg    ' �L�[�t���O
    PSTOCK_Speck.ks0.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    PSTOCK_Speck.ks0.reserve = &H0                  ' �\��ς�

    PSTOCK_Speck.ks1.keypos = 2                     ' �L�[�|�W�V����
    PSTOCK_Speck.ks1.keyleng = 1                    ' �L�[��
    PSTOCK_Speck.ks1.keyflag = BtKfExt + BtKfSeg    ' �L�[�t���O
    PSTOCK_Speck.ks1.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    PSTOCK_Speck.ks1.reserve = &H0                  ' �\��ς�

    PSTOCK_Speck.ks2.keypos = 3                     ' �L�[�|�W�V����
    PSTOCK_Speck.ks2.keyleng = 20                   ' �L�[��
    PSTOCK_Speck.ks2.keyflag = BtKfExt              ' �L�[�t���O
    PSTOCK_Speck.ks2.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    PSTOCK_Speck.ks2.reserve = &H0                  ' �\��ς�
'-----------------------------------------------    ' �L�[�P
    PSTOCK_Speck.ks3.keypos = 1                     ' �L�[�|�W�V����
    PSTOCK_Speck.ks3.keyleng = 1                    ' �L�[��
    PSTOCK_Speck.ks3.keyflag = BtKfExt + BtKfSeg    ' �L�[�t���O
    PSTOCK_Speck.ks3.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    PSTOCK_Speck.ks3.reserve = &H0                  ' �\��ς�

    PSTOCK_Speck.ks4.keypos = 2                     ' �L�[�|�W�V����
    PSTOCK_Speck.ks4.keyleng = 1                    ' �L�[��
    PSTOCK_Speck.ks4.keyflag = BtKfExt + BtKfSeg    ' �L�[�t���O
    PSTOCK_Speck.ks4.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    PSTOCK_Speck.ks4.reserve = &H0                  ' �\��ς�

    PSTOCK_Speck.ks5.keypos = 23                    ' �L�[�|�W�V����
    PSTOCK_Speck.ks5.keyleng = 8                    ' �L�[��
    PSTOCK_Speck.ks5.keyflag = BtKfExt + BtKfSeg    ' �L�[�t���O
    PSTOCK_Speck.ks5.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    PSTOCK_Speck.ks5.reserve = &H0                  ' �\��ς�

    PSTOCK_Speck.ks6.keypos = 3                     ' �L�[�|�W�V����
    PSTOCK_Speck.ks6.keyleng = 20                   ' �L�[��
    PSTOCK_Speck.ks6.keyflag = BtKfExt              ' �L�[�t���O
    PSTOCK_Speck.ks6.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    PSTOCK_Speck.ks6.reserve = &H0                  ' �\��ς�


    sts = BTRV(BtOpCreate, PSTOCK_POS, PSTOCK_Speck, Len(PSTOCK_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�͈͓��ړ����݌Ƀf�[�^")
        Exit Function
    End If
    
    PSTOCK_Create = False

End Function
Public Function PSTOCK_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              �͈͓��ړ����݌Ɉꗗ�f�[�^�@�n�o�d�m                *
'*                                                                  *
'*      ��  ��:Open Mode(Btrieve�Q��)                               *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 2004.04.27                                       *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    PSTOCK_Open = True
                                            '�͈͓��ړ����݌Ɉꗗ�f�[�^�t���p�X�捞��
    sts = GetIni("FILE", PSTOCK_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI[PSTOCK] �ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, PSTOCK_POS, PSTOCKREC, Len(PSTOCKREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = PSTOCK_Create()               '�͈͓��ړ����݌Ɉꗗ�f�[�^�쐬
                If sts Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, PSTOCK_POS, PSTOCKREC, Len(SUMZREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�͈͓��ړ����݌Ɉꗗ�f�[�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�͈͓��ړ����݌Ɉꗗ�f�[�^")
                Exit Function
        End Select
    Loop

    PSTOCK_Open = False

End Function


