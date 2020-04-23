Attribute VB_Name = "OYA_ITEM"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �e���i�W�J�f�[�^�@�t�@�C����`                      *
'*                                                                  *
'*          CREATE 2008.11.05                                       *
'********************************************************************
'�t�@�C���h�c
Public Const OYA_ITEM_ID$ = "OYA_ITEM"

'�y�[�W�T�C�Y
Public Const OYA_ITEM_PG_SIZ% = 2048

'�|�W�V�����E�u���b�N
Public OYA_ITEM_POS         As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                              *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type OYA_ITEMREC_Tag
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j

    AVE_SYUKA(0 To 7)       As Byte     '���Ϗo�א�

    ST_SOKO(0 To 1)         As Byte     '�W�����ɑq�� �q��
    ST_RETU(0 To 1)         As Byte     '             ��
    ST_REN(0 To 1)          As Byte     '             �A
    ST_DAN(0 To 1)          As Byte     '             �i




End Type

'�f�[�^�E�o�b�t�@
Public OYA_ITEMREC          As OYA_ITEMREC_Tag

'�L�[��`
Private Type KEY0_OYA_ITEM          '�j�d�x�O
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j
End Type

Private Type KEY1_OYA_ITEM          '�j�d�x�P
    AVE_SYUKA(0 To 7)       As Byte     '���Ϗo�א�
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j
End Type

'�L�[�E�f�[�^
Public K0_OYA_ITEM          As KEY0_OYA_ITEM
Public K1_OYA_ITEM          As KEY1_OYA_ITEM

Private Type OYA_ITEM_FSpeck
    fs As BtFileSpeck               ' ̧�� ��߯��\����
    ks0 As BtKeySpeck               ' �� ��߯��\����
    ks1 As BtKeySpeck               ' �� ��߯��\����
    ks2 As BtKeySpeck               ' �� ��߯��\����
    ks3 As BtKeySpeck               ' �� ��߯��\����
    ks4 As BtKeySpeck               ' �� ��߯��\����
    ks5 As BtKeySpeck               ' �� ��߯��\����
    ks6 As BtKeySpeck               ' �� ��߯��\����
End Type

Private OYA_ITEM_Speck      As OYA_ITEM_FSpeck

Private Function OYA_ITEM_Create() As Integer
'********************************************************************
'*                                                                  *
'*              �e���i�W�J�f�[�^�@�b�q�d�`�s�d                      *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 2008.11.05                                       *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128


Dim sBuffer     As String * 255
Dim com         As String


Dim Ret         As Integer




    OYA_ITEM_Create = True
                                            '�݌ɏW�v�f�[�^�t���p�X�捞��
    sts = GetIni("FILE", OYA_ITEM_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI[OYA_ITEM] �ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "???"
    End If


    Ret = InStr(1, FullPath, ".") - 1
    FullPath = Left(FullPath, Ret) & com & Right(FullPath, Len(FullPath) - Ret)
    
    
    
    
    
    OYA_ITEM_Speck.fs.recoleng = Len(OYA_ITEMREC)       ' ���R�[�h��
    OYA_ITEM_Speck.fs.PageSize = OYA_ITEM_PG_SIZ        ' �y�[�W�T�C�Y
    OYA_ITEM_Speck.fs.idexnumb = 2                      ' �C���f�b�N�X��
    OYA_ITEM_Speck.fs.fileflag = 0                      ' �t�@�C���t���O
    OYA_ITEM_Speck.fs.reserve = &H0                     ' �\��ς�
'-----------------------------------------------' �L�[�O
    OYA_ITEM_Speck.ks0.keypos = 1                       ' �L�[�|�W�V����
    OYA_ITEM_Speck.ks0.keyleng = 1                      ' �L�[��
    OYA_ITEM_Speck.ks0.keyflag = BtKfExt + BtKfSeg      ' �L�[�t���O
    OYA_ITEM_Speck.ks0.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    OYA_ITEM_Speck.ks0.reserve = &H0                    ' �\��ς�

    OYA_ITEM_Speck.ks1.keypos = 2                       ' �L�[�|�W�V����
    OYA_ITEM_Speck.ks1.keyleng = 1                      ' �L�[��
    OYA_ITEM_Speck.ks1.keyflag = BtKfExt + BtKfSeg      ' �L�[�t���O
    OYA_ITEM_Speck.ks1.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    OYA_ITEM_Speck.ks1.reserve = &H0                    ' �\��ς�

    OYA_ITEM_Speck.ks2.keypos = 3                       ' �L�[�|�W�V����
    OYA_ITEM_Speck.ks2.keyleng = 20                     ' �L�[��
    OYA_ITEM_Speck.ks2.keyflag = BtKfExt                ' �L�[�t���O
    OYA_ITEM_Speck.ks2.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    OYA_ITEM_Speck.ks2.reserve = &H0                    ' �\��ς�
'-----------------------------------------------' �L�[�P

    OYA_ITEM_Speck.ks3.keypos = 23                      ' �L�[�|�W�V����
    OYA_ITEM_Speck.ks3.keyleng = 8                      ' �L�[��
    OYA_ITEM_Speck.ks3.keyflag = BtKfExt + BtKfSeg      ' �L�[�t���O
    OYA_ITEM_Speck.ks3.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    OYA_ITEM_Speck.ks3.reserve = &H0                    ' �\��ς�
    
    OYA_ITEM_Speck.ks4.keypos = 1                       ' �L�[�|�W�V����
    OYA_ITEM_Speck.ks4.keyleng = 1                      ' �L�[��
    OYA_ITEM_Speck.ks4.keyflag = BtKfExt + BtKfSeg      ' �L�[�t���O
    OYA_ITEM_Speck.ks4.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    OYA_ITEM_Speck.ks4.reserve = &H0                    ' �\��ς�

    OYA_ITEM_Speck.ks5.keypos = 2                       ' �L�[�|�W�V����
    OYA_ITEM_Speck.ks5.keyleng = 1                      ' �L�[��
    OYA_ITEM_Speck.ks5.keyflag = BtKfExt + BtKfSeg      ' �L�[�t���O
    OYA_ITEM_Speck.ks5.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    OYA_ITEM_Speck.ks5.reserve = &H0                    ' �\��ς�

    OYA_ITEM_Speck.ks6.keypos = 3                       ' �L�[�|�W�V����
    OYA_ITEM_Speck.ks6.keyleng = 20                     ' �L�[��
    OYA_ITEM_Speck.ks6.keyflag = BtKfExt                ' �L�[�t���O
    OYA_ITEM_Speck.ks6.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    OYA_ITEM_Speck.ks6.reserve = &H0                    ' �\��ς�

    sts = BTRV(BtOpCreate, OYA_ITEM_POS, OYA_ITEM_Speck, Len(OYA_ITEM_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�e���i�W�J�f�[�^")
        Exit Function
    End If
    
    OYA_ITEM_Create = False

End Function

Function OYA_ITEM_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              �e���i�W�J�f�[�^�@�n�o�d�m                          *
'*                                                                  *
'*      ��  ��:Open Mode(Btrieve�Q��)                               *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 2008.11.05                                       *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    
Dim sBuffer     As String * 255
Dim com         As String


Dim Ret         As Integer
    
    
    OYA_ITEM_Open = True
                                            '�݌ɏW�v�f�[�^�t���p�X�捞��
    sts = GetIni("FILE", OYA_ITEM_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI[OYA_ITEM] �ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    
    On Error Resume Next
    
    
    Kill (FullPath)
    
    On Error GoTo 0
    
    
    
    
    
    Do
        sts = BTRV(BtOpOpen, OYA_ITEM_POS, OYA_ITEMREC, Len(SUMZREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = OYA_ITEM_Create()        '�݌ɏW�v�f�[�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, OYA_ITEM_POS, OYA_ITEMREC, Len(OYA_ITEMREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�e���i�W�J�f�[�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�e���i�W�J�f�[�^")
                Exit Function
        End Select
    Loop

    OYA_ITEM_Open = False
End Function


