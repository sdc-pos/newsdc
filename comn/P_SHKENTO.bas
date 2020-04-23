Attribute VB_Name = "P_SHKENTO"
Option Explicit

'********************************************************************
'*
'*              ��������̧��  �t�@�C����`
'*
'*          CREATE 2006.11.17
'********************************************************************
'�t�@�C���h�c
Public Const P_SHKENTO_ID$ = "P_SHKENTO"

'�y�[�W�T�C�Y
Private Const P_SHKENTO_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public P_SHKENTO_POS       As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`

Private Type JITU_TBL_Tag
    JITU_YM(0 To 6)        As Byte
    JITU_QTY(0 To 7)        As Byte
End Type





Public Type P_SHKENTO_REC_Tag
    
    
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j
                                        '���я����
    JITU_TBL(0 To 2)        As JITU_TBL_Tag
    
    LT_CODE(0 To 0)         As Byte     'ذ����с@����
    LT_DAYS(0 To 2)         As Byte     'ذ����с@����
    
    SYUSHI_CODE(0 To 2)     As Byte     '���x�P��
    
    ZAIKO_STANDARD(0 To 7)  As Byte     '��݌�
    ZAIKO_QTY(0 To 7)       As Byte     '���݌�
    
    LOT(0 To 7)             As Byte     '����ۯ�
    ORDER_CODE(0 To 4)      As Byte     '�����溰��
        
    SHIJI_Z_QTY(0 To 7)     As Byte     '�����c
    SHIJI_Z_CODE(0 To 0)    As Byte     '�����c����
        
    SHIJI_QTY_R(0 To 7)     As Byte     '���������_
    SHIJI_QTY_K(0 To 7)     As Byte     '�������m��
    SHIJI_CODE(0 To 0)      As Byte     '��������

    TANKA(0 To 10)          As Byte     '����P��(9(8)V99)
    KINGAKU(0 To 9)         As Byte     '������z(S9(9))

    SORT_KEY(0 To 9)        As Byte
    
    S_YMD(0 To 7)           As Byte     '�w��@�J�n�N����
    E_YMD(0 To 7)           As Byte
    
    
    FILLER(0 To 15)         As Byte     '�w��@�I���N����

End Type
'�f�[�^�E�o�b�t�@
Public P_SHKENTO_REC        As P_SHKENTO_REC_Tag

'�L�[��`

Public Type KEY0_P_SHKENTO                     '�j�d�x�O
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j
End Type

Public Type KEY1_P_SHKENTO                     '�j�d�x�P
    SORT_KEY(0 To 9)        As Byte
End Type
    
'�L�[�E�f�[�^
Public K0_P_SHKENTO         As KEY0_P_SHKENTO
Public K1_P_SHKENTO         As KEY1_P_SHKENTO

Type P_SHKENTO_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
    ks1                     As BtKeySpeck   ' �� ��߯��\����
    ks2                     As BtKeySpeck   ' �� ��߯��\����
    ks3                     As BtKeySpeck   ' �� ��߯��\����
End Type

Private P_SHKENTO_Speck     As P_SHKENTO_FSpeck
Private Function P_SHKENTO_Create() As Integer
'********************************************************************
'*
'*              ���ޔ�������̧��  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

Dim sBuffer         As String * 255
Dim com             As String

Dim Ret             As Integer


    P_SHKENTO_Create = True
                                            '���ގ�������ް��t���p�X�捞��
    sts = GetIni("FILE", P_SHKENTO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SHKENTO]�ǂݍ��݃G���[")
        Exit Function
    End If


    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "???"
    End If

    Ret = InStr(1, Trim(c), ".") - 1
    FullPath = Left(Trim(c), Ret) & com & Right(Trim(c), Len(Trim(c)) - Ret)


    P_SHKENTO_Speck.fs.recoleng = Len(P_SHKENTO_REC)    ' ���R�[�h��
    P_SHKENTO_Speck.fs.PageSize = P_SHKENTO_PG_SIZ      ' �y�[�W�T�C�Y
    P_SHKENTO_Speck.fs.idexnumb = 2                     ' �C���f�b�N�X��
    P_SHKENTO_Speck.fs.fileflag = 0                     ' �t�@�C���t���O
    P_SHKENTO_Speck.fs.reserve = &H0                    ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    P_SHKENTO_Speck.ks0.keypos = 1                      ' �L�[�|�W�V����
    P_SHKENTO_Speck.ks0.keyleng = 1                     ' �L�[��
    P_SHKENTO_Speck.ks0.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    P_SHKENTO_Speck.ks0.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    P_SHKENTO_Speck.ks0.reserve = &H0                   ' �\��ς�
    
    P_SHKENTO_Speck.ks1.keypos = 2                      ' �L�[�|�W�V����
    P_SHKENTO_Speck.ks1.keyleng = 1                     ' �L�[��
    P_SHKENTO_Speck.ks1.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    P_SHKENTO_Speck.ks1.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    P_SHKENTO_Speck.ks1.reserve = &H0                   ' �\��ς�
    
    P_SHKENTO_Speck.ks2.keypos = 3                      ' �L�[�|�W�V����
    P_SHKENTO_Speck.ks2.keyleng = 20                    ' �L�[��
    P_SHKENTO_Speck.ks2.keyflag = BtKfExt               ' �L�[�t���O
    P_SHKENTO_Speck.ks2.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    P_SHKENTO_Speck.ks2.reserve = &H0                   ' �\��ς�
    
    
    '--------------------------------------------------- �L�[�O ��
    
    '--------------------------------------------------- �L�[�P ��
    P_SHKENTO_Speck.ks3.keypos = 151                    ' �L�[�|�W�V����
    P_SHKENTO_Speck.ks3.keyleng = 10                    ' �L�[��
                                                        ' �L�[�t���O
    P_SHKENTO_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfChg
    P_SHKENTO_Speck.ks3.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    P_SHKENTO_Speck.ks3.reserve = &H0                   ' �\��ς�
    
    '--------------------------------------------------- �L�[�P ��
    
    sts = BTRV(BtOpCreate, P_SHKENTO_POS, P_SHKENTO_Speck, Len(P_SHKENTO_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���ޔ�������̧��")
        Exit Function
    End If
    
    P_SHKENTO_Create = False

End Function

Public Function P_SHKENTO_Open(mode As Integer) As Integer
'********************************************************************
'*
'*              ���ޔ�������̧��  �n�o�d�m
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

    P_SHKENTO_Open = True
                                            '���ޔ�������̧�كt���p�X�捞��
    sts = GetIni("FILE", P_SHKENTO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SHKENTO]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "???"
    End If

    Ret = InStr(1, Trim(c), ".") - 1
    FullPath = Left(Trim(c), Ret) & com & Right(Trim(c), Len(Trim(c)) - Ret)
    
    

    Do
        sts = BTRV(BtOpOpen, P_SHKENTO_POS, P_SHKENTO_REC, Len(P_SHKENTO_REC), ByVal FullPath, Len(FullPath), mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_SHKENTO_Create()   '���ޔ�������̧�ٍ쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_SHKENTO_POS, P_SHKENTO_REC, Len(P_SHKENTO_REC), ByVal FullPath, Len(FullPath), mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���ޔ�������̧��")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���ޔ�������̧��")
                Exit Function
        End Select
    Loop
    
    P_SHKENTO_Open = False

End Function

