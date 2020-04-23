Attribute VB_Name = "P_SEISAN_DET"
Option Explicit

'********************************************************************
'*
'*              ���Y���і����ް�  �t�@�C����`
'*
'*          CREATE 2005.11.11
'********************************************************************
'�t�@�C���h�c
Public Const P_SEISAN_DET_ID$ = "P_SEISAN_DET"

'�y�[�W�T�C�Y
Private Const P_SEISAN_DET_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public P_SEISAN_DET_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************

'���R�[�h��`
Public Type P_SEISAN_DET_REC_Tag
    
    TORI_KBN(0 To 0)        As Byte         '�����敪
    TORI_CODE(0 To 4)       As Byte         '����溰��
    UKEIRE_DT(0 To 7)       As Byte         '�����
    SHIJI_NO(0 To 4)        As Byte         '�w�}�[��
    SHIMUKE_CODE(0 To 1)    As Byte         '�d�����溰��
    HIN_GAI(0 To 19)        As Byte         '�e�i��
    UKEIRE_QTY(0 To 10)      As Byte         '�����(9(8)V99)
    S_CLASS_CODE(0 To 19)   As Byte         '���i���׽
    F_CLASS_CODE(0 To 19)   As Byte         '�t���׽
    N_CLASS_CODE(0 To 19)   As Byte         '���E�׽
    KOURYOU(0 To 10)        As Byte         '�P�� 9(8)V99
    KIN(0 To 10)            As Byte         '���z


End Type
'�f�[�^�E�o�b�t�@
Public P_SEISAN_DET_REC     As P_SEISAN_DET_REC_Tag

'�L�[��`
Public Type KEY0_P_SEISAN_DET               '�j�d�x�O
    TORI_CODE(0 To 4)       As Byte         '����溰��
End Type
    
    
'�L�[�E�f�[�^
Public K0_P_SEISAN_DET      As KEY0_P_SEISAN_DET

Type P_SEISAN_DET_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
End Type

Private P_SEISAN_DET_Speck  As P_SEISAN_DET_FSpeck
Private Function P_SEISAN_DET_Create() As Integer
'********************************************************************
'*
'*              ���Y���і����ް�  �b�q�d�`�s�d
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

    P_SEISAN_DET_Create = True
                                            '���Y���і����ް��t���p�X�捞��
    sts = GetIni("FILE", P_SEISAN_DET_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SEISAN_DET]�ǂݍ��݃G���[")
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

    P_SEISAN_DET_Speck.fs.recoleng = Len(P_SEISAN_DET_REC)  ' ���R�[�h��
    P_SEISAN_DET_Speck.fs.PageSize = P_SEISAN_DET_PG_SIZ    ' �y�[�W�T�C�Y
    P_SEISAN_DET_Speck.fs.idexnumb = 1                      ' �C���f�b�N�X��
    P_SEISAN_DET_Speck.fs.fileflag = 0                      ' �t�@�C���t���O
    P_SEISAN_DET_Speck.fs.reserve = &H0                     ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    
    P_SEISAN_DET_Speck.ks0.keypos = 2                       ' �L�[�|�W�V����
    P_SEISAN_DET_Speck.ks0.keyleng = 5                      ' �L�[��
    P_SEISAN_DET_Speck.ks0.keyflag = BtKfExt + BtKfDup      ' �L�[�t���O
    P_SEISAN_DET_Speck.ks0.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    P_SEISAN_DET_Speck.ks0.reserve = &H0                    ' �\��ς�
    
    
    
    '--------------------------------------------------- �L�[�O ��
    
    
    
    
    sts = BTRV(BtOpCreate, P_SEISAN_DET_POS, P_SEISAN_DET_Speck, Len(P_SEISAN_DET_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���Y���і����ް�")
        Exit Function
    End If
    
    P_SEISAN_DET_Create = False

End Function

Public Function P_SEISAN_DET_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���Y���і����ް�  �n�o�d�m
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


    P_SEISAN_DET_Open = True
                                            '���Y���і��׃f�[�^�t���p�X�捞��
    sts = GetIni("FILE", P_SEISAN_DET_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SEISAN_DET]�ǂݍ��݃G���[")
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
        sts = BTRV(BtOpOpen, P_SEISAN_DET_POS, P_SEISAN_DET_REC, Len(P_SEISAN_DET_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_SEISAN_DET_Create()     '���Y���і����ް��쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_SEISAN_DET_POS, P_SEISAN_DET_REC, Len(P_SEISAN_DET_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���Y���і����ް�")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���Y���і����ް�")
                Exit Function
        End Select
    Loop
    
    P_SEISAN_DET_Open = False

End Function

