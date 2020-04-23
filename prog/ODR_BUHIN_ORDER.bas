Attribute VB_Name = "ODR_BUHIN_ORDER"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �q���i�@�����e �t�@�C����`                         *
'*                                                                  *
'*          CREATE 2008.02.19                                       *
'********************************************************************
'�t�@�C���h�c
Public Const ODR_BUHIN_ORDER_ID$ = "ODR_BUHIN_ORDER"

'�y�[�W�T�C�Y
Private Const ODR_BUHIN_ORDER_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public ODR_BUHIN_ORDER_POS      As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Public Type ODR_BUHIN_ORDER_REC_Tag
    SEL_DATE(0 To 7)            As Byte         '�Ώۓ��t
    
    JGYOBU(0 To 0)              As Byte         '���ƕ�
    NAIGAI(0 To 0)              As Byte         '�����O
    HIN_GAI(0 To 19)            As Byte         '�e�i��
    DATA_KBN(0 To 0)            As Byte         '�f�[�^�敪 1:�\�� 2:����
    USE_YM(0 To 5)              As Byte         '�g�p���iYYYYMM)
    NYUKO_QTY(0 To 7)           As Byte         '������

End Type
'�f�[�^�E�o�b�t�@
Public ODR_BUHIN_ORDER_REC            As ODR_BUHIN_ORDER_REC_Tag



'�L�[��`

Type KEY0_ODR_BUHIN_ORDER                           '�j�d�x�O
    SEL_DATE(0 To 7)            As Byte         '�Ώۓ��t
    
    JGYOBU(0 To 0)              As Byte         '���ƕ�
    NAIGAI(0 To 0)              As Byte         '�����O
    HIN_GAI(0 To 19)            As Byte         '�e�i��

    DATA_KBN(0 To 0)            As Byte         '�f�[�^�敪 1:�\�� 2:����

End Type

Type KEY1_ODR_BUHIN_ORDER                           '�j�d�x�P
    JGYOBU(0 To 0)              As Byte         '���ƕ�
    NAIGAI(0 To 0)              As Byte         '�����O
    HIN_GAI(0 To 19)            As Byte         '�e�i��

    SEL_DATE(0 To 7)            As Byte         '�Ώۓ��t

    DATA_KBN(0 To 0)            As Byte         '�f�[�^�敪 1:�\�� 2:����

End Type




'�L�[�E�f�[�^
Public K0_ODR_BUHIN_ORDER           As KEY0_ODR_BUHIN_ORDER
Public K1_ODR_BUHIN_ORDER           As KEY1_ODR_BUHIN_ORDER

Type ODR_BUHIN_ORDER_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
    ks1                     As BtKeySpeck   ' �� ��߯��\����
    ks2                     As BtKeySpeck   ' �� ��߯��\����
    ks3                     As BtKeySpeck   ' �� ��߯��\����
    ks4                     As BtKeySpeck   ' �� ��߯��\����
    
    ks5                     As BtKeySpeck   ' �� ��߯��\����
    ks6                     As BtKeySpeck   ' �� ��߯��\����
    ks7                     As BtKeySpeck   ' �� ��߯��\����
    ks8                     As BtKeySpeck   ' �� ��߯��\����
    ks9                     As BtKeySpeck   ' �� ��߯��\����
    

End Type

Private ODR_BUHIN_ORDER_Speck       As ODR_BUHIN_ORDER_FSpeck
Private Function ODR_BUHIN_ORDER_Create() As Integer
'********************************************************************
'*                                                                  *
'*              �q���i�Q�����e  �b�q�d�`�s�d                        *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

Dim sBuffer         As String * 255
Dim com             As String


Dim Ret             As Integer


    ODR_BUHIN_ORDER_Create = True
                                            '�q���i�Q�����e�t���p�X�捞��
    sts = GetIni("FILE", ODR_BUHIN_ORDER_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_BUHIN_ORDER]�ǂݍ��݃G���[")
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


    ODR_BUHIN_ORDER_Speck.fs.recoleng = Len(ODR_BUHIN_ORDER_REC)      ' ���R�[�h��
    ODR_BUHIN_ORDER_Speck.fs.PageSize = ODR_BUHIN_ORDER_PG_SIZ        ' �y�[�W�T�C�Y
    ODR_BUHIN_ORDER_Speck.fs.idexnumb = 2                       ' �C���f�b�N�X��
    ODR_BUHIN_ORDER_Speck.fs.fileflag = 0                       ' �t�@�C���t���O
    ODR_BUHIN_ORDER_Speck.fs.reserve = &H0                      ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    ODR_BUHIN_ORDER_Speck.ks0.keypos = 1                        ' �L�[�|�W�V����
    ODR_BUHIN_ORDER_Speck.ks0.keyleng = 8                       ' �L�[��
                                                                ' �L�[�t���O
    ODR_BUHIN_ORDER_Speck.ks0.keyflag = BtKfExt + BtKfSeg
    ODR_BUHIN_ORDER_Speck.ks0.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_BUHIN_ORDER_Speck.ks0.reserve = &H0                     ' �\��ς�
    
    ODR_BUHIN_ORDER_Speck.ks1.keypos = 9                        ' �L�[�|�W�V����
    ODR_BUHIN_ORDER_Speck.ks1.keyleng = 1                       ' �L�[��
                                                                ' �L�[�t���O
    ODR_BUHIN_ORDER_Speck.ks1.keyflag = BtKfExt + BtKfSeg
    ODR_BUHIN_ORDER_Speck.ks1.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_BUHIN_ORDER_Speck.ks1.reserve = &H0                     ' �\��ς�
    
    ODR_BUHIN_ORDER_Speck.ks2.keypos = 10                       ' �L�[�|�W�V����
    ODR_BUHIN_ORDER_Speck.ks2.keyleng = 1                       ' �L�[��
                                                                ' �L�[�t���O
    ODR_BUHIN_ORDER_Speck.ks2.keyflag = BtKfExt + BtKfSeg
    ODR_BUHIN_ORDER_Speck.ks2.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_BUHIN_ORDER_Speck.ks2.reserve = &H0                     ' �\��ς�
    
    ODR_BUHIN_ORDER_Speck.ks3.keypos = 11                       ' �L�[�|�W�V����
    ODR_BUHIN_ORDER_Speck.ks3.keyleng = 20                      ' �L�[��
                                                                ' �L�[�t���O
    ODR_BUHIN_ORDER_Speck.ks3.keyflag = BtKfExt + BtKfSeg
    ODR_BUHIN_ORDER_Speck.ks3.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_BUHIN_ORDER_Speck.ks3.reserve = &H0                     ' �\��ς�
    
    ODR_BUHIN_ORDER_Speck.ks4.keypos = 31                       ' �L�[�|�W�V����
    ODR_BUHIN_ORDER_Speck.ks4.keyleng = 1                       ' �L�[��
                                                                ' �L�[�t���O
    ODR_BUHIN_ORDER_Speck.ks4.keyflag = BtKfExt
    ODR_BUHIN_ORDER_Speck.ks4.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_BUHIN_ORDER_Speck.ks4.reserve = &H0                     ' �\��ς�
    
    '--------------------------------------------------- �L�[�O ��
    
    '--------------------------------------------------- �L�[�P ��
    ODR_BUHIN_ORDER_Speck.ks5.keypos = 9                        ' �L�[�|�W�V����
    ODR_BUHIN_ORDER_Speck.ks5.keyleng = 1                       ' �L�[��
                                                                ' �L�[�t���O
    ODR_BUHIN_ORDER_Speck.ks5.keyflag = BtKfExt + BtKfSeg
    ODR_BUHIN_ORDER_Speck.ks5.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_BUHIN_ORDER_Speck.ks5.reserve = &H0                     ' �\��ς�
    
    ODR_BUHIN_ORDER_Speck.ks6.keypos = 10                       ' �L�[�|�W�V����
    ODR_BUHIN_ORDER_Speck.ks6.keyleng = 1                       ' �L�[��
                                                                ' �L�[�t���O
    ODR_BUHIN_ORDER_Speck.ks6.keyflag = BtKfExt + BtKfSeg
    ODR_BUHIN_ORDER_Speck.ks6.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_BUHIN_ORDER_Speck.ks6.reserve = &H0                     ' �\��ς�
    
    ODR_BUHIN_ORDER_Speck.ks7.keypos = 11                       ' �L�[�|�W�V����
    ODR_BUHIN_ORDER_Speck.ks7.keyleng = 20                      ' �L�[��
                                                                ' �L�[�t���O
    ODR_BUHIN_ORDER_Speck.ks7.keyflag = BtKfExt + BtKfSeg
    ODR_BUHIN_ORDER_Speck.ks7.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_BUHIN_ORDER_Speck.ks7.reserve = &H0                     ' �\��ς�
    
    ODR_BUHIN_ORDER_Speck.ks8.keypos = 1                        ' �L�[�|�W�V����
    ODR_BUHIN_ORDER_Speck.ks8.keyleng = 8                       ' �L�[��
                                                                ' �L�[�t���O
    ODR_BUHIN_ORDER_Speck.ks8.keyflag = BtKfExt + BtKfSeg
    ODR_BUHIN_ORDER_Speck.ks8.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_BUHIN_ORDER_Speck.ks8.reserve = &H0                     ' �\��ς�
    
    
    ODR_BUHIN_ORDER_Speck.ks9.keypos = 31                       ' �L�[�|�W�V����
    ODR_BUHIN_ORDER_Speck.ks9.keyleng = 1                       ' �L�[��
                                                                ' �L�[�t���O
    ODR_BUHIN_ORDER_Speck.ks9.keyflag = BtKfExt
    ODR_BUHIN_ORDER_Speck.ks9.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    '--------------------------------------------------- �L�[�P ��
    
    
    
    
    sts = BTRV(BtOpCreate, ODR_BUHIN_ORDER_POS, ODR_BUHIN_ORDER_Speck, Len(ODR_BUHIN_ORDER_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�q���i�Q�����e")
        Exit Function
    End If
    
    ODR_BUHIN_ORDER_Create = False

End Function

Public Function ODR_BUHIN_ORDER_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �q���i�Q�����e  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

Dim sBuffer     As String * 255
Dim com         As String


Dim Ret         As Integer


    ODR_BUHIN_ORDER_Open = True
                                            '�q���i�Q�����e�t���p�X�捞��
    sts = GetIni("FILE", ODR_BUHIN_ORDER_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_BUHIN_ORDER]�ǂݍ��݃G���[")
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
        sts = BTRV(BtOpOpen, ODR_BUHIN_ORDER_POS, ODR_BUHIN_ORDER_REC, Len(ODR_BUHIN_ORDER_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = ODR_BUHIN_ORDER_Create()      '�q���i�Q�����e�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ODR_BUHIN_ORDER_POS, ODR_BUHIN_ORDER_REC, Len(ODR_BUHIN_ORDER_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�q���i �����e")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�q���i �����e")
                Exit Function
        End Select
    Loop
    
    ODR_BUHIN_ORDER_Open = False
    
End Function
