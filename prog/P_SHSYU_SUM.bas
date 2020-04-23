Attribute VB_Name = "P_SHSYU_SUM"
Option Explicit

'********************************************************************
'*
'*              ���ގd���W�v(���x�P�ʕ�)�ް�  �t�@�C����`
'*
'*          CREATE 2007.04.01
'********************************************************************
'�t�@�C���h�c
Public Const P_SHSYU_SUM_ID$ = "P_SHSYU_SUM"

'�y�[�W�T�C�Y
Private Const P_SHSYU_SUM_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public P_SHSYU_SUM_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
Private Type SHIIRE_TBL_Tag
    SHIIRE(0 To 9)          As Byte
End Type

'���R�[�h��`
Public Type P_SHSYU_SUM_REC_Tag
    
    G_SYUSHI(0 To 2)        As Byte             '���x����
    SHIIRE_TBL(0 To 6)      As SHIIRE_TBL_Tag

End Type
'�f�[�^�E�o�b�t�@
Public P_SHSYU_SUM_REC      As P_SHSYU_SUM_REC_Tag

'�L�[��`
Public Type KEY0_P_SHSYU_SUM            '�j�d�x�O
    G_SYUSHI(0 To 2)        As Byte             '���x����
End Type
    
'�L�[�E�f�[�^
Public K0_P_SHSYU_SUM       As KEY0_P_SHSYU_SUM

Type P_SHSYU_SUM_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
End Type

Private P_SHSYU_SUM_Speck   As P_SHSYU_SUM_FSpeck
Private Function P_SHSYU_SUM_Create() As Integer
'********************************************************************
'*
'*              ���ގd���W�v�ް�    �b�q�d�`�s�d
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


    P_SHSYU_SUM_Create = True
                                            '���ގd���W�v�ް��t���p�X�捞��
    sts = GetIni("FILE", P_SHSYU_SUM_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SHSYU_SUM]�ǂݍ��݃G���[")
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

    P_SHSYU_SUM_Speck.fs.recoleng = Len(P_SHSYU_SUM_REC)    ' ���R�[�h��
    P_SHSYU_SUM_Speck.fs.PageSize = P_SHSYU_SUM_PG_SIZ      ' �y�[�W�T�C�Y
    P_SHSYU_SUM_Speck.fs.idexnumb = 1                       ' �C���f�b�N�X��
    P_SHSYU_SUM_Speck.fs.fileflag = 0                       ' �t�@�C���t���O
    P_SHSYU_SUM_Speck.fs.reserve = &H0                      ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    P_SHSYU_SUM_Speck.ks0.keypos = 1                        ' �L�[�|�W�V����
    P_SHSYU_SUM_Speck.ks0.keyleng = 3                       ' �L�[��
    P_SHSYU_SUM_Speck.ks0.keyflag = BtKfExt                 ' �L�[�t���O
    P_SHSYU_SUM_Speck.ks0.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_SHSYU_SUM_Speck.ks0.reserve = &H0                     ' �\��ς�
    
    
    '--------------------------------------------------- �L�[�O ��
    
    sts = BTRV(BtOpCreate, P_SHSYU_SUM_POS, P_SHSYU_SUM_Speck, Len(P_SHSYU_SUM_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���ގd���W�v�ް�")
        Exit Function
    End If
    
    P_SHSYU_SUM_Create = False

End Function

Public Function P_SHSYU_SUM_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���ގd���W�v�ް�  �n�o�d�m
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

    P_SHSYU_SUM_Open = True
                                            '���ގd���W�v�ް��t���p�X�捞��
    sts = GetIni("FILE", P_SHSYU_SUM_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SHSYU_SUM]�ǂݍ��݃G���[")
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
        sts = BTRV(BtOpOpen, P_SHSYU_SUM_POS, P_SHSYU_SUM_REC, Len(P_SHSYU_SUM_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_SHSYU_SUM_Create()  '���ގd���W�v�ް��쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_SHSYU_SUM_POS, P_SHSYU_SUM_REC, Len(P_SHSYU_SUM_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���ގd���W�v�ް�")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���ގd���W�v�ް�")
                Exit Function
        End Select
    Loop
    
    P_SHSYU_SUM_Open = False

End Function

