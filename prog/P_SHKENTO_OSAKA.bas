Attribute VB_Name = "P_SHKENTO_OSAKA"
Option Explicit

'********************************************************************
'*
'*              ���������@���PC����  �t�@�C����`
'*
'*          CREATE 2005.11.11
'********************************************************************
'�t�@�C���h�c
Public Const P_SHKENTO_OSAKA_ID$ = "P_SHKENTO_OSAKA"

'�y�[�W�T�C�Y
Private Const P_SHKENTO_OSAKA_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public P_SHKENTO_OSAKA_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Public Type P_SHKENTO_OSAKA_REC_Tag
    
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '���ޕi��
    
    SO_SUU(0 To 10)         As Byte         '���K�v��(9(8)V99)
    TANKA(0 To 10)          As Byte         '�d���P��(9(8)V99)
    
    ST_SOKO(0 To 1)         As Byte         '�W���I�ԁ@�q��
    ST_RETU(0 To 1)         As Byte         '�W���I�ԁ@��
    ST_REN(0 To 1)          As Byte         '�W���I�ԁ@�A
    ST_DAN(0 To 1)          As Byte         '�W���I�ԁ@�i
    
    ZAIKO_QTY(0 To 7)       As Byte         '�݌ɐ�
    
    SHIJI_Z_QTY(0 To 10)    As Byte         '�����c(9(8)V99)
    
    HIKIATE_Z_QTY(0 To 10)  As Byte         '�����c(9(8)V99)
    
    FUSOKU_QTY(0 To 10)     As Byte         '�s��(9(8)V99)
    
    ORDER_QTY(0 To 10)      As Byte         '������(9(8)V99)
    
    LOT(0 To 7)             As Byte         '����ۯ�
    
    ORDER_CODE(0 To 4)      As Byte         '�����溰��
    
    LT(0 To 2)              As Byte         'ذ�����
    
    
    Y_NOUKI_DT(0 To 7)      As Byte         '�\��[��
    
    REC_NO(0 To 3)          As Byte         'ں��އ�
    
    
    FILLER(0 To 59)         As Byte         'Filler

End Type
'�f�[�^�E�o�b�t�@
Public P_SHKENTO_OSAKA_REC  As P_SHKENTO_OSAKA_REC_Tag

'�L�[��`

Public Type KEY0_P_SHKENTO_OSAKA            '�j�d�x�O
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '���ޕi��
End Type
    
Public Type KEY1_P_SHKENTO_OSAKA            '�j�d�x�P
    REC_NO(0 To 3)          As Byte         'ں��އ�
End Type
    
Public Type KEY2_P_SHKENTO_OSAKA            '�j�d�x�Q
    FUSOKU_QTY(0 To 10)     As Byte         '�s��(9(8)V99)
End Type
    
    
    
'�L�[�E�f�[�^
Public K0_P_SHKENTO_OSAKA   As KEY0_P_SHKENTO_OSAKA
Public K1_P_SHKENTO_OSAKA   As KEY1_P_SHKENTO_OSAKA
Public K2_P_SHKENTO_OSAKA   As KEY2_P_SHKENTO_OSAKA


Type P_SHKENTO_OSAKA_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
    ks1                     As BtKeySpeck   ' �� ��߯��\����
    ks2                     As BtKeySpeck   ' �� ��߯��\����
    ks3                     As BtKeySpeck   ' �� ��߯��\����
    ks4                     As BtKeySpeck   ' �� ��߯��\����

End Type

Private P_SHKENTO_OSAKA_Speck   As P_SHKENTO_OSAKA_FSpeck
Private Function P_SHKENTO_OSAKA_Create(Optional F_NAME As String = " ") As Integer
'********************************************************************
'*
'*              ���������@���PC����  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

Dim Ret             As Integer
    
    P_SHKENTO_OSAKA_Create = True
                                            '���������@���PC�����t���p�X�捞��
    sts = GetIni("FILE", P_SHKENTO_OSAKA_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SHKENTO_OSAKA]�ǂݍ��݃G���[")
        Exit Function
    End If

    If Trim(F_NAME) = "" Then
        FullPath = RTrim(c)
    Else
        Ret = InStr(1, Trim(c), ".") - 1
        FullPath = Left(Trim(c), Ret) & Trim(F_NAME) & Right(Trim(c), Len(Trim(c)) - Ret)
    End If

    P_SHKENTO_OSAKA_Speck.fs.recoleng = Len(P_SHKENTO_OSAKA_REC)    ' ���R�[�h��
    P_SHKENTO_OSAKA_Speck.fs.PageSize = P_SHKENTO_OSAKA_PG_SIZ      ' �y�[�W�T�C�Y
    P_SHKENTO_OSAKA_Speck.fs.idexnumb = 3                           ' �C���f�b�N�X��
    P_SHKENTO_OSAKA_Speck.fs.fileflag = 0                           ' �t�@�C���t���O
    P_SHKENTO_OSAKA_Speck.fs.reserve = &H0                          ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    P_SHKENTO_OSAKA_Speck.ks0.keypos = 1                            ' �L�[�|�W�V����
    P_SHKENTO_OSAKA_Speck.ks0.keyleng = 1                           ' �L�[��
    P_SHKENTO_OSAKA_Speck.ks0.keyflag = BtKfExt + BtKfSeg           ' �L�[�t���O
    P_SHKENTO_OSAKA_Speck.ks0.keytype = Chr(BtKtString)             ' �L�[�^�C�v
    P_SHKENTO_OSAKA_Speck.ks0.reserve = &H0                         ' �\��ς�
    
    P_SHKENTO_OSAKA_Speck.ks1.keypos = 2                            ' �L�[�|�W�V����
    P_SHKENTO_OSAKA_Speck.ks1.keyleng = 1                           ' �L�[��
    P_SHKENTO_OSAKA_Speck.ks1.keyflag = BtKfExt + BtKfSeg           ' �L�[�t���O
    P_SHKENTO_OSAKA_Speck.ks1.keytype = Chr(BtKtString)             ' �L�[�^�C�v
    P_SHKENTO_OSAKA_Speck.ks1.reserve = &H0                         ' �\��ς�
    
    P_SHKENTO_OSAKA_Speck.ks2.keypos = 3                            ' �L�[�|�W�V����
    P_SHKENTO_OSAKA_Speck.ks2.keyleng = 20                          ' �L�[��
    P_SHKENTO_OSAKA_Speck.ks2.keyflag = BtKfExt                     ' �L�[�t���O
    P_SHKENTO_OSAKA_Speck.ks2.keytype = Chr(BtKtString)             ' �L�[�^�C�v
    P_SHKENTO_OSAKA_Speck.ks2.reserve = &H0                         ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    
    '--------------------------------------------------- �L�[�P ��
    P_SHKENTO_OSAKA_Speck.ks3.keypos = 129                          ' �L�[�|�W�V����
    P_SHKENTO_OSAKA_Speck.ks3.keyleng = 4                           ' �L�[��
    P_SHKENTO_OSAKA_Speck.ks3.keyflag = BtKfExt + BtKfChg + BtKfDup ' �L�[�t���O
    P_SHKENTO_OSAKA_Speck.ks3.keytype = Chr(BtKtString)             ' �L�[�^�C�v
    P_SHKENTO_OSAKA_Speck.ks3.reserve = &H0                         ' �\��ς�
    '--------------------------------------------------- �L�[�P ��
    
    '--------------------------------------------------- �L�[�Q ��
    P_SHKENTO_OSAKA_Speck.ks4.keypos = 83                          ' �L�[�|�W�V����
    P_SHKENTO_OSAKA_Speck.ks4.keyleng = 11                           ' �L�[��
    P_SHKENTO_OSAKA_Speck.ks4.keyflag = BtKfExt + BtKfChg + BtKfDup ' �L�[�t���O
    P_SHKENTO_OSAKA_Speck.ks4.keytype = Chr(BtKtString)             ' �L�[�^�C�v
    P_SHKENTO_OSAKA_Speck.ks4.reserve = &H0                         ' �\��ς�
    '--------------------------------------------------- �L�[�Q ��
    
    
    
    sts = BTRV(BtOpCreate, P_SHKENTO_OSAKA_POS, P_SHKENTO_OSAKA_Speck, Len(P_SHKENTO_OSAKA_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���������@���PC�����ް�")
        Exit Function
    End If
    
    P_SHKENTO_OSAKA_Create = False

End Function

Public Function P_SHKENTO_OSAKA_Open(mode As Integer, Optional F_NAME As String = " ") As Integer
'********************************************************************
'*
'*              ���������@���PC����  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

Dim Ret         As Integer

    P_SHKENTO_OSAKA_Open = True
                                                        '���������@���PC�����t���p�X�捞��
    sts = GetIni("FILE", P_SHKENTO_OSAKA_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SHKENTO_OSAKA]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    If Trim(F_NAME) = "" Then
        FullPath = RTrim(c)
    Else
        Ret = InStr(1, Trim(c), ".") - 1
        FullPath = Left(Trim(c), Ret) & Trim(F_NAME) & Right(Trim(c), Len(Trim(c)) - Ret)
    End If
    
    On Error Resume Next
    Kill (FullPath)
    On Error GoTo 0

    Do
        sts = BTRV(BtOpOpen, P_SHKENTO_OSAKA_POS, P_SHKENTO_OSAKA_REC, Len(P_SHKENTO_OSAKA_REC), ByVal FullPath, Len(FullPath), mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_SHKENTO_OSAKA_Create(F_NAME)          '���������@���PC�����쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_SHKENTO_OSAKA_POS, P_SHKENTO_OSAKA_REC, Len(P_SHKENTO_OSAKA_REC), ByVal FullPath, Len(FullPath), mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���������@���PC����")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���������@���PC����")
                Exit Function
        End Select
    Loop
    
    P_SHKENTO_OSAKA_Open = False

End Function

