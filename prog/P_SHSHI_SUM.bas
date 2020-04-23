Attribute VB_Name = "P_SH_SHI_SUM"
Option Explicit

'********************************************************************
'*
'*              ���ގd���W�v�ް�  �t�@�C����`
'*
'*          CREATE 2005.11.11
'********************************************************************
'�t�@�C���h�c
Public Const P_SHSHI_SUM_ID$ = "P_SHSHI_SUM"

'�y�[�W�T�C�Y
Private Const P_SHSHI_SUM_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public P_SHSHI_SUM_POS As POSBLK
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
Public Type P_SHSHI_SUM_REC_Tag
    
    SHIIRE_CODE(0 To 4)     As Byte             '�d���溰��
    TORI_KBN(0 To 0)        As Byte             '�����敪
    SHIIRE_TBL(0 To 6)      As SHIIRE_TBL_Tag

End Type
'�f�[�^�E�o�b�t�@
Public P_SHSHI_SUM_REC      As P_SHSHI_SUM_REC_Tag

'�L�[��`
Public Type KEY0_P_SHSHI_SUM            '�j�d�x�O
    SHIIRE_CODE(0 To 4)     As Byte             '�d���溰��
    TORI_KBN(0 To 0)        As Byte             '�����敪
End Type
    
'�L�[�E�f�[�^
Public K0_P_SHSHI_SUM       As KEY0_P_SHSHI_SUM

Type P_SHSHI_SUM_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
    ks1                     As BtKeySpeck   ' �� ��߯��\����
End Type

Private P_SHSHI_SUM_Speck   As P_SHSHI_SUM_FSpeck
Private Function P_SHSHI_SUM_Create() As Integer
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


    P_SHSHI_SUM_Create = True
                                            '���ގd���W�v�ް��t���p�X�捞��
    sts = GetIni("FILE", P_SHSHI_SUM_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SHSHI_SUM]�ǂݍ��݃G���[")
        Exit Function
    End If

    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "???"
    End If


'    Ret = InStr(1, Trim(c), ".") - 1
    
    Ret = InStrRev(Trim(c), ".") - 1
    
    
    
    FullPath = Left(Trim(c), Ret) & com & Right(Trim(c), Len(Trim(c)) - Ret)

    P_SHSHI_SUM_Speck.fs.recoleng = Len(P_SHSHI_SUM_REC)    ' ���R�[�h��
    P_SHSHI_SUM_Speck.fs.PageSize = P_SHSHI_SUM_PG_SIZ      ' �y�[�W�T�C�Y
    P_SHSHI_SUM_Speck.fs.idexnumb = 1                       ' �C���f�b�N�X��
    P_SHSHI_SUM_Speck.fs.fileflag = 0                       ' �t�@�C���t���O
    P_SHSHI_SUM_Speck.fs.reserve = &H0                      ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    P_SHSHI_SUM_Speck.ks0.keypos = 1                       ' �L�[�|�W�V����
    P_SHSHI_SUM_Speck.ks0.keyleng = 5                      ' �L�[��
    P_SHSHI_SUM_Speck.ks0.keyflag = BtKfExt + BtKfSeg      ' �L�[�t���O
    P_SHSHI_SUM_Speck.ks0.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    P_SHSHI_SUM_Speck.ks0.reserve = &H0                    ' �\��ς�
    
    
    P_SHSHI_SUM_Speck.ks1.keypos = 6                       ' �L�[�|�W�V����
    P_SHSHI_SUM_Speck.ks1.keyleng = 1                      ' �L�[��
    P_SHSHI_SUM_Speck.ks1.keyflag = BtKfExt                ' �L�[�t���O
    P_SHSHI_SUM_Speck.ks1.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    P_SHSHI_SUM_Speck.ks1.reserve = &H0                    ' �\��ς�
    
    '--------------------------------------------------- �L�[�O ��
    
    sts = BTRV(BtOpCreate, P_SHSHI_SUM_POS, P_SHSHI_SUM_Speck, Len(P_SHSHI_SUM_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���ގd���W�v�ް�")
        Exit Function
    End If
    
    P_SHSHI_SUM_Create = False

End Function

Public Function P_SHSHI_SUM_Open(Mode As Integer) As Integer
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

    P_SHSHI_SUM_Open = True
                                            '���ގd���W�v�ް��t���p�X�捞��
    sts = GetIni("FILE", P_SHSHI_SUM_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SHSHI_SUM]�ǂݍ��݃G���[")
        Exit Function
    End If
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "???"
    End If


'    Ret = InStr(1, Trim(c), ".") - 1
    
    
    
    Ret = InStrRev(Trim(c), ".") - 1
    
    
    FullPath = Left(Trim(c), Ret) & com & Right(Trim(c), Len(Trim(c)) - Ret)

    Do
        sts = BTRV(BtOpOpen, P_SHSHI_SUM_POS, P_SHSHI_SUM_REC, Len(P_SHSHI_SUM_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_SHSHI_SUM_Create()  '���ގd���W�v�ް��쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_SHSHI_SUM_POS, P_SHSHI_SUM_REC, Len(P_SHSHI_SUM_REC), ByVal FullPath, Len(FullPath), Mode)
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
    
    P_SHSHI_SUM_Open = False

End Function

