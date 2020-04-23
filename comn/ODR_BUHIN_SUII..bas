Attribute VB_Name = "ODR_BUHIN_SUII"
Option Explicit
'********************************************************************
'*
'*              �q���i���ڃf�[�^ �t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const ODR_BUHIN_SUII_ID$ = "ODR_BUHIN_SUII"

'�y�[�W�T�C�Y
Public Const ODR_BUHIN_SUII_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public ODR_BUHIN_SUII_POS As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************




Type ODR_BUHIN_SUII_REC_Tag
    
    SEL_DATE(0 To 7)        As Byte         '�I����t
    
    KO_JGYOBU(0 To 0)       As Byte         '���ƕ�
    KO_NAIGAI(0 To 0)       As Byte         '�����O
    KO_HIN_GAI(0 To 19)     As Byte         '�i�ԁi�O���j

    USE_YM(0 To 5)          As Byte         '�g�p���iYYYYMM)

    ORDER_CODE(0 To 4)      As Byte         '�d����(��\)

    Y_ZAIKO_QTY(0 To 7)     As Byte         '�L���݌�
    HIKIATE_QTY(0 To 7)     As Byte         '�����\��
    NYUKO_QTY(0 To 7)       As Byte         '���ɐ�
    SYUKO_QTY(0 To 7)       As Byte         '�o�ɐ�



End Type

'�f�[�^�E�o�b�t�@
Public ODR_BUHIN_SUII_REC   As ODR_BUHIN_SUII_REC_Tag

'�L�[��`

Type KEY0_ODR_BUHIN_SUII                    '�j�d�x�O
    SEL_DATE(0 To 7)        As Byte         '�I����t
    
    KO_JGYOBU(0 To 0)       As Byte         '���ƕ�
    KO_NAIGAI(0 To 0)       As Byte         '�����O
    KO_HIN_GAI(0 To 19)     As Byte         '�i�ԁi�O���j
    
End Type

Type KEY1_ODR_BUHIN_SUII                    '�j�d�x�P
    
    KO_JGYOBU(0 To 0)       As Byte         '���ƕ�
    KO_NAIGAI(0 To 0)       As Byte         '�����O
    KO_HIN_GAI(0 To 19)     As Byte         '�i�ԁi�O���j
    
    SEL_DATE(0 To 7)        As Byte         '�I����t
    
End Type



'�L�[�E�f�[�^
Public K0_ODR_BUHIN_SUII    As KEY0_ODR_BUHIN_SUII
Public K1_ODR_BUHIN_SUII    As KEY1_ODR_BUHIN_SUII

Type ODR_BUHIN_SUII_FSpeck
    fs      As BtFileSpeck          ' ̧�� ��߯��\����
    ks0     As BtKeySpeck           ' �� ��߯��\����
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck

    ks4     As BtKeySpeck
    ks5     As BtKeySpeck
    ks6     As BtKeySpeck
    ks7     As BtKeySpeck

End Type

Private ODR_BUHIN_SUII_Speck    As ODR_BUHIN_SUII_FSpeck
Private Function ODR_BUHIN_SUII_Create() As Integer
'********************************************************************
'*
'*              �q���i���ڃf�[�^�@�b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

Dim sBuffer         As String * 255
Dim com             As String


Dim Ret             As Integer

    ODR_BUHIN_SUII_Create = True
                                            '�q���i���ڃf�[�^�@�t���p�X�捞��
    sts = GetIni("FILE", ODR_BUHIN_SUII_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_BUHIN_SUII]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)





    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "???"
    End If


    Ret = InStr(1, Trim(c), ".") - 1
    FullPath = Left(Trim(c), Ret) & com & Right(Trim(c), Len(Trim(c)) - Ret)














    ODR_BUHIN_SUII_Speck.fs.recoleng = Len(ODR_BUHIN_SUII_REC)  ' ���R�[�h��
    ODR_BUHIN_SUII_Speck.fs.PageSize = ODR_BUHIN_SUII_PG_SIZ    ' �y�[�W�T�C�Y
    ODR_BUHIN_SUII_Speck.fs.idexnumb = 2                        ' �C���f�b�N�X��
    ODR_BUHIN_SUII_Speck.fs.fileflag = 0                        ' �t�@�C���t���O
    ODR_BUHIN_SUII_Speck.fs.reserve = &H0                       ' �\��ς�
'---------------------------------------------------'
                                                    ' �L�[�O
    ODR_BUHIN_SUII_Speck.ks0.keypos = 1                         ' �L�[�|�W�V����
    ODR_BUHIN_SUII_Speck.ks0.keyleng = 8                        ' �L�[��
    ODR_BUHIN_SUII_Speck.ks0.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    ODR_BUHIN_SUII_Speck.ks0.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    ODR_BUHIN_SUII_Speck.ks0.reserve = &H0                      ' �\��ς�
                                                    
                                                    
    ODR_BUHIN_SUII_Speck.ks1.keypos = 9                         ' �L�[�|�W�V����
    ODR_BUHIN_SUII_Speck.ks1.keyleng = 1                        ' �L�[��
    ODR_BUHIN_SUII_Speck.ks1.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    ODR_BUHIN_SUII_Speck.ks1.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    ODR_BUHIN_SUII_Speck.ks1.reserve = &H0                      ' �\��ς�
                                                    
    ODR_BUHIN_SUII_Speck.ks2.keypos = 10                        ' �L�[�|�W�V����
    ODR_BUHIN_SUII_Speck.ks2.keyleng = 1                        ' �L�[��
    ODR_BUHIN_SUII_Speck.ks2.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    ODR_BUHIN_SUII_Speck.ks2.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    ODR_BUHIN_SUII_Speck.ks2.reserve = &H0                      ' �\��ς�
                                                    
    ODR_BUHIN_SUII_Speck.ks3.keypos = 11                        ' �L�[�|�W�V����
    ODR_BUHIN_SUII_Speck.ks3.keyleng = 20                       ' �L�[��
    ODR_BUHIN_SUII_Speck.ks3.keyflag = BtKfExt                  ' �L�[�t���O
    ODR_BUHIN_SUII_Speck.ks3.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    ODR_BUHIN_SUII_Speck.ks3.reserve = &H0                      ' �\��ς�
                                                    
                                                    ' �L�[�O
'---------------------------------------------------'
    
    
'---------------------------------------------------'
    ODR_BUHIN_SUII_Speck.ks4.keypos = 9                         ' �L�[�|�W�V����
    ODR_BUHIN_SUII_Speck.ks4.keyleng = 1                        ' �L�[��
    ODR_BUHIN_SUII_Speck.ks4.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    ODR_BUHIN_SUII_Speck.ks4.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    ODR_BUHIN_SUII_Speck.ks4.reserve = &H0                      ' �\��ς�
                                                    
    ODR_BUHIN_SUII_Speck.ks5.keypos = 10                        ' �L�[�|�W�V����
    ODR_BUHIN_SUII_Speck.ks5.keyleng = 1                        ' �L�[��
    ODR_BUHIN_SUII_Speck.ks5.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    ODR_BUHIN_SUII_Speck.ks5.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    ODR_BUHIN_SUII_Speck.ks5.reserve = &H0                      ' �\��ς�
                                                    
    ODR_BUHIN_SUII_Speck.ks6.keypos = 11                        ' �L�[�|�W�V����
    ODR_BUHIN_SUII_Speck.ks6.keyleng = 20                       ' �L�[��
    ODR_BUHIN_SUII_Speck.ks6.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    ODR_BUHIN_SUII_Speck.ks6.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    ODR_BUHIN_SUII_Speck.ks6.reserve = &H0                      ' �\��ς�
                                                    
    ODR_BUHIN_SUII_Speck.ks7.keypos = 1                         ' �L�[�|�W�V����
    ODR_BUHIN_SUII_Speck.ks7.keyleng = 8                        ' �L�[��
    ODR_BUHIN_SUII_Speck.ks7.keyflag = BtKfExt                  ' �L�[�t���O
    ODR_BUHIN_SUII_Speck.ks7.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    ODR_BUHIN_SUII_Speck.ks7.reserve = &H0                      ' �\��ς�
                                                    
                                                    
                                                    ' �L�[�P
'---------------------------------------------------'
    
    
    sts = BTRV(BtOpCreate, ODR_BUHIN_SUII_POS, ODR_BUHIN_SUII_Speck, Len(ODR_BUHIN_SUII_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�q���i���ڃf�[�^")
        Exit Function
    End If
    ODR_BUHIN_SUII_Create = False
End Function
Public Function ODR_BUHIN_SUII_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �q���i���ڃf�[�^�@�n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
Dim sBuffer         As String * 255
Dim com             As String


Dim Ret             As Integer
    
    ODR_BUHIN_SUII_Open = True
                                                        '�q���i���ڃf�[�^�@�t���p�X�捞��
    sts = GetIni("FILE", ODR_BUHIN_SUII_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_BUHIN_SUII_]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "???"
    End If


    Ret = InStr(1, Trim(c), ".") - 1
    FullPath = Left(Trim(c), Ret) & com & Right(Trim(c), Len(Trim(c)) - Ret)
    
    
    
    
    
    
    Do
        sts = BTRV(BtOpOpen, ODR_BUHIN_SUII_POS, ODR_BUHIN_SUII_REC, Len(ODR_BUHIN_SUII_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = ODR_BUHIN_SUII_Create()           '�q���i���ڃf�[�^�@�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ODR_BUHIN_SUII_POS, ODR_BUHIN_SUII_REC, Len(ODR_BUHIN_SUII_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�q���i���ڃf�[�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�q���i���ڃf�[�^")
                Exit Function
        End Select
    Loop
    ODR_BUHIN_SUII_Open = False

End Function

